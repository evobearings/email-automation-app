from datetime import datetime, timedelta
from email import encoders
from email.header import decode_header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
import re
import smtplib
import sqlite3
import time
import imapclient
import pandas as pd
import streamlit as st

# ==========================================
# 1. DATABASE MANAGEMENT (SQLite)
# ==========================================


def init_db():
  conn = sqlite3.connect('campaigns.db')
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS sent_emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT UNIQUE,
            initial_sent_date TIMESTAMP,
            status TEXT DEFAULT 'Pending',
            followup_sent_date TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


def log_initial_email(name: str, email: str):
  conn = sqlite3.connect('campaigns.db')
  cursor = conn.cursor()
  now = datetime.now()
  cursor.execute(
      """
        INSERT INTO sent_emails (name, email, initial_sent_date, status)
        VALUES (?, ?, ?, 'Pending')
        ON CONFLICT(email) DO UPDATE SET
            initial_sent_date=?,
            status='Pending'
    """,
      (name, email, now, now),
  )
  conn.commit()
  conn.close()


def log_followup_email(email: str):
  conn = sqlite3.connect('campaigns.db')
  cursor = conn.cursor()
  now = datetime.now()
  cursor.execute(
      """
        UPDATE sent_emails
        SET status = 'Follow-up Sent', followup_sent_date = ?
        WHERE email = ?
    """,
      (now, email),
  )
  conn.commit()
  conn.close()


def get_emails_sent_in_last_24h() -> set:
  """Retrieves all email addresses sent an initial or follow-up email in the last 24 hours."""
  conn = sqlite3.connect('campaigns.db')
  cutoff_time = datetime.now() - timedelta(hours=24)
  cursor = conn.cursor()

  cursor.execute(
      """
        SELECT LOWER(email) FROM sent_emails
        WHERE initial_sent_date >= ? OR followup_sent_date >= ?
    """,
      (cutoff_time, cutoff_time),
  )

  rows = cursor.fetchall()
  conn.close()
  return {row[0] for row in rows}


def update_reply_status(replied_emails: list):
  if not replied_emails:
    return 0
  conn = sqlite3.connect('campaigns.db')
  cursor = conn.cursor()
  count = 0
  for email in replied_emails:
    cursor.execute(
        """
            UPDATE sent_emails
            SET status = 'Replied'
            WHERE LOWER(email) = LOWER(?) AND status != 'Replied'
        """,
        (email,),
    )
    count += cursor.rowcount
  conn.commit()
  conn.close()
  return count


def fetch_followup_candidates(days_threshold=15):
  conn = sqlite3.connect('campaigns.db')
  cutoff_date = datetime.now() - timedelta(days=days_threshold)
  df = pd.read_sql_query(
      """
        SELECT id, name, email, initial_sent_date, status
        FROM sent_emails
        WHERE status = 'Pending' AND initial_sent_date <= ?
    """,
      conn,
      params=(cutoff_date,),
  )
  conn.close()
  return df


init_db()

# ==========================================
# 2. CONTACT CLEANING ENGINE WITH 24H SAFEGUARD
# ==========================================

EMAIL_REGEX = re.compile(
    r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', re.IGNORECASE
)


def extract_and_clean_contacts(uploaded_file):
  records = []
  if uploaded_file.name.endswith('.csv'):
    df_raw = pd.read_csv(uploaded_file)
    sheet_data = [('CSV_Data', df_raw)]
  else:
    xls = pd.ExcelFile(uploaded_file)
    sheet_data = [
        (s, pd.read_excel(xls, sheet_name=s, header=None))
        for s in xls.sheet_names
    ]

  for sheet_name, df in sheet_data:
    if df.dropna(how='all').empty:
      continue
    rows = df.fillna('').astype(str).values.tolist()
    header_idx = -1
    for idx, row in enumerate(rows[:5]):
      row_str = ' '.join(row).lower()
      if any(
          k in row_str
          for k in [
              'company',
              'contact',
              'mail',
              'email',
              'name',
              'person',
              'member',
          ]
      ):
        header_idx = idx
        break

    headers = (
        [c.strip().lower() for c in rows[header_idx]]
        if header_idx != -1
        else []
    )
    data_rows = rows[header_idx + 1 :] if header_idx != -1 else rows

    name_col = -1
    for c_idx, h in enumerate(headers):
      if any(
          k in h
          for k in [
              'name',
              'company',
              'contact',
              'person',
              'firm',
              'organization',
          ]
      ):
        name_col = c_idx
        break

    for row in data_rows:
      row_str = ' '.join(row)
      emails = list(dict.fromkeys(EMAIL_REGEX.findall(row_str)))
      email_val = '; '.join(emails) if emails else ''
      name_val = 'N/A'
      if name_col != -1 and name_col < len(row):
        cand = row[name_col].strip()
        if cand and cand.lower() not in ['sl no', 'sl.no', 'name', 'company']:
          name_val = cand

      if name_val == 'N/A':
        for cell in row:
          c_s = cell.strip()
          if (
              c_s
              and not c_s.isdigit()
              and c_s.lower() not in ['sl no', 'sl.no', 'name', 'company']
              and not EMAIL_REGEX.search(c_s)
              and len(c_s) < 90
          ):
            name_val = c_s
            break

      if name_val != 'N/A' or email_val != '':
        records.append({'Name': name_val, 'Email': email_val})

  if not records:
    return pd.DataFrame(columns=['Name', 'Email']), 0

  df_clean = pd.DataFrame(records)
  df_clean['Name'] = df_clean['Name'].fillna('').astype(str).str.strip()
  df_clean['Email'] = (
      df_clean['Email'].fillna('').astype(str).str.strip().str.lower()
  )
  df_clean = df_clean.drop_duplicates(subset=['Name', 'Email']).copy()
  df_clean = df_clean[(df_clean['Name'] != 'N/A') | (df_clean['Email'] != '')]

  # --- 24-HOUR DUPLICATE SAFEGUARD ---
  recent_emails = get_emails_sent_in_last_24h()

  def is_recent(email_str):
    extracted_emails = [
        e.strip().lower() for e in email_str.split(';') if '@' in e
    ]
    return any(e in recent_emails for e in extracted_emails)

  initial_count = len(df_clean)
  df_filtered = df_clean[~df_clean['Email'].apply(is_recent)].copy()
  skipped_count = initial_count - len(df_filtered)

  return df_filtered, skipped_count


# ==========================================
# 3. IMAP AUTO-REPLY DETECTOR
# ==========================================


def check_inbox_for_replies(
    user_email, user_password, imap_server='imap.gmail.com'
):
  replied_senders = set()
  try:
    with imapclient.IMAPClient(imap_server, ssl=True) as client:
      client.login(user_email, user_password)
      client.select_folder('INBOX')

      since_date = (datetime.now() - timedelta(days=60)).strftime('%d-%b-%Y')
      messages = client.search(['SINCE', since_date])

      for msg_id, data in client.fetch(messages, ['ENVELOPE']).items():
        envelope = data[b'ENVELOPE']
        if envelope.from_:
          for addr in envelope.from_:
            mailbox = addr.mailbox.decode() if addr.mailbox else ''
            host = addr.host.decode() if addr.host else ''
            if mailbox and host:
              replied_senders.add(f'{mailbox}@{host}'.lower())

    return list(replied_senders), None
  except Exception as e:
    return [], str(e)


# ==========================================
# 4. EMAIL FORMATTER (HTML Builder with Bold/Highlight)
# ==========================================


def build_html_body(
    text_content,
    gdrive_links=None,
    catalog_links=None,
    video_links=None,
    video_names=None,
):
  formatted_text = text_content.replace('\n', '<br>')
  links_html = ''

  highlight_style = 'font-weight: bold; background-color: #fffacd; padding: 3px 6px; border-radius: 4px; text-decoration: underline;'

  if gdrive_links:
    drive_items = ''
    for idx, link in enumerate(gdrive_links, 1):
      if link and link.strip():
        drive_items += f"""
                <li style="margin-bottom: 8px;">
                    <a href="{link.strip()}" target="_blank" style="color: #1a73e8; {highlight_style}">
                        Google Drive File #{idx}
                    </a>
                </li>
                """
    if drive_items:
      links_html += f"""
            <div style="margin-top: 15px;">
                📁 <b>Google Drive Attachments:</b>
                <ul style="margin: 5px 0 0 20px; padding: 0;">
                    {drive_items}
                </ul>
            </div>
            """

  if catalog_links:
    catalog_items = ''
    for idx, link in enumerate(catalog_links, 1):
      if link and link.strip():
        catalog_items += f"""
                <li style="margin-bottom: 8px;">
                    <a href="{link.strip()}" target="_blank" style="color: #008080; {highlight_style}">
                        Catalog / Website Link #{idx}
                    </a>
                </li>
                """
    if catalog_items:
      links_html += f"""
            <div style="margin-top: 15px;">
                📖 <b>Catalog / Website Links:</b>
                <ul style="margin: 5px 0 0 20px; padding: 0;">
                    {catalog_items}
                </ul>
            </div>
            """

  if video_links:
    video_link_items = ''
    for idx, link in enumerate(video_links, 1):
      if link and link.strip():
        video_link_items += f"""
                <li style="margin-bottom: 8px;">
                    <a href="{link.strip()}" target="_blank" style="color: #d9534f; {highlight_style}">
                        Video Link #{idx}
                    </a>
                </li>
                """
    if video_link_items:
      links_html += f"""
            <div style="margin-top: 15px;">
                🎥 <b>Video Links:</b>
                <ul style="margin: 5px 0 0 20px; padding: 0;">
                    {video_link_items}
                </ul>
            </div>
            """

  if video_names:
    video_items = ''
    for v_name in video_names:
      video_items += f'<li style="margin-bottom: 8px;"><span style="{highlight_style}">🎥 {v_name} (Attached)</span></li>'
    if video_items:
      links_html += f"""
            <div style="margin-top: 15px;">
                ▶️ <b>Attached Video Files:</b>
                <ul style="margin: 5px 0 0 20px; padding: 0;">
                    {video_items}
                </ul>
            </div>
            """

  return f"""
    <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6;">
            <div>{formatted_text}</div>
            {links_html}
        </body>
    </html>
    """


# ==========================================
# 5. STREAMLIT INTERFACE & CONTROLS
# ==========================================

st.set_page_config(
    page_title='Email Automation & CRM', page_icon='📧', layout='wide'
)

st.title('📧 Email Automation & Follow-Up Portal')

if 'is_sending' not in st.session_state:
  st.session_state.is_sending = False

st.sidebar.header('⚙️ Sender Credentials')
sender_email = st.sidebar.text_input('Sender Email', value='')
sender_password = st.sidebar.text_input(
    'App Password (Gmail/SMTP)', type='password'
)
smtp_server = st.sidebar.text_input('SMTP Server', value='smtp.gmail.com')
smtp_port = st.sidebar.number_input('SMTP Port', value=587)
imap_server = st.sidebar.text_input('IMAP Server', value='imap.gmail.com')

tab1, tab2, tab3 = st.tabs(
    ['🚀 Launch Campaign', '⏰ Follow-Up Manager (15 Days)', '📊 Database Log']
)

# --- TAB 1: LAUNCH CAMPAIGN ---
with tab1:
  st.header('1. Upload Contact Spreadsheet')
  uploaded_file = st.file_uploader(
      'Upload Excel (.xlsx, .xls) or CSV',
      type=['xlsx', 'xls', 'csv'],
      key='up1',
  )

  if uploaded_file:
    df_clean, skipped_count = extract_and_clean_contacts(uploaded_file)
    valid_contacts = df_clean[
        (df_clean['Email'] != '') & (df_clean['Email'] != 'N/A')
    ]

    st.success(f'Extracted {len(valid_contacts)} valid contact emails.')
    if skipped_count > 0:
      st.warning(
          f'🛡️ **24-Hour Safeguard Active:** Automatically skipped'
          f' {skipped_count} recipient(s) who were already emailed in the last'
          ' 24 hours.'
      )

    with st.expander('Preview Contact List'):
      st.dataframe(df_clean)

    st.header('2. Compose Initial Email')

    randomize_subjects = st.checkbox(
        '🎲 Randomize Subject Lines across emails',
        value=False,
        key='rand_sub1',
    )

    if randomize_subjects:
      st.info(
          'Provide up to 5 subject lines below. The script will randomly pick'
          ' one for each recipient.'
      )
      sub1 = st.text_input(
          'Subject Line 1', value='Partnership Proposal', key='s1'
      )
      sub2 = st.text_input(
          'Subject Line 2', value='Collaboration Opportunity', key='s2'
      )
      sub3 = st.text_input(
          'Subject Line 3', value='Quick Question / Intro', key='s3'
      )
      sub4 = st.text_input(
          'Subject Line 4 (Optional)', value='', key='sub4_opt'
      )
      sub5 = st.text_input(
          'Subject Line 5 (Optional)', value='', key='sub5_opt'
      )
    else:
      main_subject = st.text_input(
          'Subject Line', value='Introduction & Partnership Proposal', key='main_s1'
      )

    template_body = st.text_area(
        'Email Body',
        value='Dear {name},\n\nWe are following up regarding our initial discussion.\n\nBest regards,\nSales Team',
        height=150,
        key='body1',
    )

    # --- EXPANDER FOR ATTACHMENTS & EXTERNAL RESOURCES ---
    with st.expander(
        '📎 3. Attachments & External Resources (Optional)', expanded=False
    ):
      col_att1, col_att2 = st.columns(2)

      with col_att1:
        st.markdown('📁 **Google Drive Links**')
        gdrive_link_1 = st.text_input(
            'Google Drive Link 1',
            value='',
            placeholder='https://drive.google.com/...',
            key='gd1',
        )
        gdrive_link_2 = st.text_input(
            'Google Drive Link 2',
            value='',
            placeholder='https://drive.google.com/...',
            key='gd2',
        )

        st.markdown('---')
        st.markdown('📖 **Catalog / Website Links**')
        catalog_link_1 = st.text_input(
            'Catalog Link 1',
            value='',
            placeholder='https://yourwebsite.com/...',
            key='cat1',
        )
        catalog_link_2 = st.text_input(
            'Catalog Link 2',
            value='',
            placeholder='https://yourwebsite.com/...',
            key='cat2',
        )

      with col_att2:
        st.markdown('🎥 **Video Links**')
        video_link_1 = st.text_input(
            'Video Link 1',
            value='',
            placeholder='https://youtube.com/...',
            key='vlink1',
        )
        video_link_2 = st.text_input(
            'Video Link 2',
            value='',
            placeholder='https://youtube.com/...',
            key='vlink2',
        )

        st.markdown('---')
        video_attachments = st.file_uploader(
            '🎥 Attach Video File(s)',
            type=['mp4', 'mov', 'avi', 'mkv', 'webm'],
            accept_multiple_files=True,
            key='vid_att1',
        )

      st.markdown('---')
      attachments = st.file_uploader(
          '📎 Attach Document File(s) (PDF, Images, etc.)',
          accept_multiple_files=True,
          key='att1',
      )

    st.header('4. Campaign Controls & Delay Settings')

    delay_range = st.slider(
        '⏱️ Random Delay Range Between Emails (seconds):',
        min_value=0,
        max_value=300,
        value=(30, 60),
        step=1,
        key='delay1',
    )

    col_btn1, col_btn2 = st.columns(2)

    with col_btn1:
      start_clicked = st.button(
          '🚀 Start Campaign',
          disabled=st.session_state.is_sending,
          key='start1',
      )
    with col_btn2:
      stop_clicked = st.button('🛑 Stop Campaign', key='stop1')

    if stop_clicked:
      st.session_state.is_sending = False
      st.warning('Campaign manually stopped.')

    if start_clicked:
      if not sender_email or not sender_password:
        st.error('Please complete your credentials in the sidebar.')
      elif valid_contacts.empty:
        st.error('No valid contacts found.')
      else:
        if randomize_subjects:
          subject_pool = [
              s.strip()
              for s in [sub1, sub2, sub3, sub4, sub5]
              if s.strip() != ''
          ]
          if not subject_pool:
            subject_pool = ['Introduction & Partnership Proposal']
        else:
          subject_pool = [main_subject]

        active_gdrive_links = [
            gdrive_link_1 if 'gdrive_link_1' in locals() else '',
            gdrive_link_2 if 'gdrive_link_2' in locals() else '',
        ]
        active_catalog_links = [
            catalog_link_1 if 'catalog_link_1' in locals() else '',
            catalog_link_2 if 'catalog_link_2' in locals() else '',
        ]
        active_video_links = [
            video_link_1 if 'video_link_1' in locals() else '',
            video_link_2 if 'video_link_2' in locals() else '',
        ]
        video_names = (
            [v.name for v in video_attachments]
            if 'video_attachments' in locals() and video_attachments
            else []
        )

        st.session_state.is_sending = True

        try:
          server = smtplib.SMTP(smtp_server, smtp_port, timeout=120)
          server.starttls()
          server.login(sender_email, sender_password)

          progress = st.progress(0)
          status_msg = st.empty()
          total = len(valid_contacts)

          min_sec, max_sec = delay_range

          for i, (_, row) in enumerate(valid_contacts.iterrows()):
            if not st.session_state.is_sending:
              status_msg.warning('Process interrupted and stopped.')
              break

            recipients = [
                e.strip() for e in str(row['Email']).split(';') if '@' in e
            ]
            name_val = (
                row['Name'] if row['Name'] != 'N/A' else 'Valued Customer'
            )
            p_body_text = template_body.replace('{name}', name_val)

            html_body = build_html_body(
                p_body_text,
                active_gdrive_links,
                active_catalog_links,
                active_video_links,
                video_names,
            )

            current_subject = random.choice(subject_pool)

            for recipient in recipients:
              msg = MIMEMultipart('alternative')
              msg['From'] = sender_email
              msg['To'] = recipient
              msg['Subject'] = current_subject

              msg.attach(MIMEText(html_body, 'html'))

              if 'attachments' in locals() and attachments:
                for att in attachments:
                  part = MIMEBase('application', 'octet-stream')
                  part.set_payload(att.getvalue())
                  encoders.encode_base64(part)
                  part.add_header(
                      'Content-Disposition',
                      f'attachment; filename="{att.name}"',
                  )
                  msg.attach(part)

              if 'video_attachments' in locals() and video_attachments:
                for vid in video_attachments:
                  vid_part = MIMEBase('video', 'octet-stream')
                  vid_part.set_payload(vid.getvalue())
                  encoders.encode_base64(vid_part)
                  vid_part.add_header(
                      'Content-Disposition',
                      f'attachment; filename="{vid.name}"',
                  )
                  msg.attach(vid_part)

              server.sendmail(sender_email, [recipient], msg.as_string())
              log_initial_email(name_val, recipient)

              actual_delay = round(random.uniform(min_sec, max_sec), 1)
              status_msg.text(
                  f'[{i+1}/{total}] Sent to: {recipient} | Subject:'
                  f' "{current_subject}" (Waiting {actual_delay}s)'
              )
              time.sleep(actual_delay)

            progress.progress((i + 1) / total)

          server.quit()
          if st.session_state.is_sending:
            st.success(
                '🎉 Campaign fully completed and logged in local database!'
            )
          st.session_state.is_sending = False

        except Exception as e:
          st.error(f'Error sending emails: {e}')
          st.session_state.is_sending = False

# --- TAB 2: FOLLOW-UP MANAGER ---
with tab2:
  st.header('⏰ 15-Day Auto Follow-Up Engine')

  col1, col2 = st.columns(2)
  with col1:
    days_thresh = st.number_input(
        'Days Without Response', value=15, min_value=0, key='days_th'
    )
  with col2:
    if st.button('🔄 Sync Inbox & Detect Replies (IMAP)', key='sync_btn'):
      if not sender_email or not sender_password:
        st.error('Provide credentials in sidebar.')
      else:
        st.info('Scanning inbox for customer replies...')
        senders, err = check_inbox_for_replies(
            sender_email, sender_password, imap_server
        )
        if err:
          st.error(f'IMAP Sync Failed: {err}')
        else:
          updated_count = update_reply_status(senders)
          st.success(
              f'Sync Complete! Detected replies and updated {updated_count}'
              ' records in database.'
          )

  candidates = fetch_followup_candidates(days_threshold=days_thresh)

  st.subheader(
      f'Contacts Pending Follow-Up (Sent >= {days_thresh} days ago without'
      ' reply)'
  )
  st.dataframe(candidates)

  if not candidates.empty:
    fu_randomize_subjects = st.checkbox(
        '🎲 Randomize Follow-Up Subject Lines across emails',
        value=False,
        key='rand_sub2',
    )

    if fu_randomize_subjects:
      st.info(
          'Provide up to 5 subject lines below for follow-ups. The script will'
          ' randomly select one per recipient.'
      )
      fu_sub1 = st.text_input(
          'Follow-Up Subject Line 1',
          value='Following up on my previous message',
          key='fs1',
      )
      fu_sub2 = st.text_input(
          'Follow-Up Subject Line 2', value='Checking back in', key='fs2'
      )
      fu_sub3 = st.text_input(
          'Follow-Up Subject Line 3',
          value='Any thoughts on this proposal?',
          key='fs3',
      )
      fu_sub4 = st.text_input(
          'Follow-Up Subject Line 4 (Optional)', value='', key='fu_sub4_opt'
      )
      fu_sub5 = st.text_input(
          'Follow-Up Subject Line 5 (Optional)', value='', key='fu_sub5_opt'
      )
    else:
      fu_subject = st.text_input(
          'Follow-Up Subject Line',
          value='Following up on my previous message',
          key='main_fs1',
      )

    fu_body = st.text_area(
        'Follow-Up Body Template',
        value='Hi {name},\n\nI am following up on my previous email sent a couple of weeks ago. Let me know if you are open to discussing this.\n\nBest,',
        height=150,
        key='body2',
    )

    # --- EXPANDER FOR FOLLOW-UP ATTACHMENTS & RESOURCES ---
    with st.expander(
        '📎 Follow-Up Attachments & External Resources (Optional)',
        expanded=False,
    ):
      col_fu_att1, col_fu_att2 = st.columns(2)

      with col_fu_att1:
        st.markdown('📁 **Follow-Up Google Drive Links**')
        fu_gdrive_link_1 = st.text_input(
            'Google Drive Link 1',
            value='',
            placeholder='https://drive.google.com/...',
            key='fu_gd1',
        )
        fu_gdrive_link_2 = st.text_input(
            'Google Drive Link 2',
            value='',
            placeholder='https://drive.google.com/...',
            key='fu_gd2',
        )

        st.markdown('---')
        st.markdown('📖 **Follow-Up Catalog / Website Links**')
        fu_catalog_link_1 = st.text_input(
            'Catalog Link 1',
            value='',
            placeholder='https://yourwebsite.com/...',
            key='fu_cat1',
        )
        fu_catalog_link_2 = st.text_input(
            'Catalog Link 2',
            value='',
            placeholder='https://yourwebsite.com/...',
            key='fu_cat2',
        )

      with col_fu_att2:
        st.markdown('🎥 **Follow-Up Video Links**')
        fu_video_link_1 = st.text_input(
            'Video Link 1',
            value='',
            placeholder='https://youtube.com/...',
            key='fu_vlink1',
        )
        fu_video_link_2 = st.text_input(
            'Video Link 2',
            value='',
            placeholder='https://youtube.com/...',
            key='fu_vlink2',
        )

        st.markdown('---')
        fu_video_attachments = st.file_uploader(
            '🎥 Attach Follow-Up Video File(s)',
            type=['mp4', 'mov', 'avi', 'mkv', 'webm'],
            accept_multiple_files=True,
            key='fu_vid_att',
        )

    st.subheader('⏱️ Follow-Up Sending Speed')
    fu_delay_range = st.slider(
        '⏱️ Random Delay Range Between Follow-Ups (seconds):',
        min_value=0,
        max_value=300,
        value=(30, 60),
        step=1,
        key='delay2',
    )

    col_fu1, col_fu2 = st.columns(2)

    with col_fu1:
      fu_start = st.button(
          '🚀 Start Follow-Ups',
          disabled=st.session_state.is_sending,
          key='start2',
      )
    with col_fu2:
      fu_stop = st.button('🛑 Stop Follow-Ups', key='stop2')

    if fu_stop:
      st.session_state.is_sending = False
      st.warning('Follow-up process stopped.')

    if fu_start:
      if not sender_email or not sender_password:
        st.error('Provide credentials in sidebar.')
      else:
        if fu_randomize_subjects:
          fu_subject_pool = [
              s.strip()
              for s in [fu_sub1, fu_sub2, fu_sub3, fu_sub4, fu_sub5]
              if s.strip() != ''
          ]
          if not fu_subject_pool:
            fu_subject_pool = ['Following up on my previous message']
        else:
          fu_subject_pool = [fu_subject]

        fu_active_gdrive_links = [
            fu_gdrive_link_1 if 'fu_gdrive_link_1' in locals() else '',
            fu_gdrive_link_2 if 'fu_gdrive_link_2' in locals() else '',
        ]
        fu_active_catalog_links = [
            fu_catalog_link_1 if 'fu_catalog_link_1' in locals() else '',
            fu_catalog_link_2 if 'fu_catalog_link_2' in locals() else '',
        ]
        fu_active_video_links = [
            fu_video_link_1 if 'fu_video_link_1' in locals() else '',
            fu_video_link_2 if 'fu_video_link_2' in locals() else '',
        ]
        fu_video_names = (
            [v.name for v in fu_video_attachments]
            if 'fu_video_attachments' in locals() and fu_video_attachments
            else []
        )

        st.session_state.is_sending = True

        try:
          server = smtplib.SMTP(smtp_server, smtp_port, timeout=120)
          server.starttls()
          server.login(sender_email, sender_password)

          fu_progress = st.progress(0)
          fu_status_msg = st.empty()
          total_fu = len(candidates)

          fu_min_sec, fu_max_sec = fu_delay_range

          for i, (_, row) in enumerate(candidates.iterrows()):
            if not st.session_state.is_sending:
              fu_status_msg.warning('Follow-up interrupted and stopped.')
              break

            recipient = row['email']
            name_val = row['name']
            p_body_text = fu_body.replace('{name}', name_val)

            html_body = build_html_body(
                p_body_text,
                fu_active_gdrive_links,
                fu_active_catalog_links,
                fu_active_video_links,
                fu_video_names,
            )

            current_fu_subject = random.choice(fu_subject_pool)

            msg = MIMEMultipart('alternative')
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = current_fu_subject
            msg.attach(MIMEText(html_body, 'html'))

            if 'fu_video_attachments' in locals() and fu_video_attachments:
              for vid in fu_video_attachments:
                vid_part = MIMEBase('video', 'octet-stream')
                vid_part.set_payload(vid.getvalue())
                encoders.encode_base64(vid_part)
                vid_part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{vid.name}"',
                )
                msg.attach(vid_part)

            server.sendmail(sender_email, [recipient], msg.as_string())
            log_followup_email(recipient)

            fu_actual_delay = round(random.uniform(fu_min_sec, fu_max_sec), 1)
            fu_status_msg.text(
                f'[{i+1}/{total_fu}] Sent follow-up to: {recipient} | Subject:'
                f' "{current_fu_subject}" (Waiting {fu_actual_delay}s)'
            )

            fu_progress.progress((i + 1) / total_fu)
            time.sleep(fu_actual_delay)

          server.quit()
          if st.session_state.is_sending:
            st.success('🎉 Follow-up campaign completed successfully!')
          st.session_state.is_sending = False

        except Exception as e:
          st.error(f'SMTP Error: {e}')
          st.session_state.is_sending = False

# --- TAB 3: DATABASE LOG ---
with tab3:
  st.header('📊 Campaign History & Records')
  conn = sqlite3.connect('campaigns.db')
  all_logs = pd.read_sql_query('SELECT * FROM sent_emails', conn)
  conn.close()

  st.dataframe(all_logs)