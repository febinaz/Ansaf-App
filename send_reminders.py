import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

EXCEL_FILE = 'gears.xlsx'
COLUMNS = [
    'Gear',
    'TRTR Date', 'TRTR Next Open', 'TRTR Reminder Sent',
    'MNTT Date', 'MNTT Next Open', 'MNTT Reminder Sent'
]

EMAIL_TO = 'ansaf999@gmail.com'
EMAIL_FROM = 'automaticmail293@gmail.com'  # <-- Change this to your Gmail
EMAIL_PASS = ''     # <-- Use an App Password, not your Gmail password


def load_data():
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=COLUMNS)
    return pd.read_excel(EXCEL_FILE, engine='openpyxl')

def get_due_reminders(df):
    today = datetime.today().date()
    due = []
    # TRTR
    trtr_due = df[(df['TRTR Next Open'].notna()) & (df['TRTR Reminder Sent'] != 'Yes') & (pd.to_datetime(df['TRTR Next Open'], errors='coerce').dt.date == today)]
    for _, row in trtr_due.iterrows():
        due.append(f"Gear: {row['Gear']} - TRTR needs to be opened today! (Next Open: {row['TRTR Next Open']})")
    # MNTT
    mntt_due = df[(df['MNTT Next Open'].notna()) & (df['MNTT Reminder Sent'] != 'Yes') & (pd.to_datetime(df['MNTT Next Open'], errors='coerce').dt.date == today)]
    for _, row in mntt_due.iterrows():
        due.append(f"Gear: {row['Gear']} - MNTT needs to be opened today! (Next Open: {row['MNTT Next Open']})")
    return due, trtr_due.index, mntt_due.index

def send_email(subject, body):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_FROM, EMAIL_PASS)
            server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        print('Email sent!')
    except Exception as e:
        print('Failed to send email:', e)

def add_today_due_sample():
    today = datetime.today().date()
    df = load_data()
    # Prepare new rows as DataFrame
    new_rows = pd.DataFrame([
        {
            'Gear': 'SAMPLE TRTR DUE',
            'TRTR Date': (today - pd.Timedelta(days=14)).strftime('%Y-%m-%d'),
            'TRTR Next Open': today.strftime('%Y-%m-%d'),
            'TRTR Reminder Sent': 'No',
            'MNTT Date': '',
            'MNTT Next Open': '',
            'MNTT Reminder Sent': ''
        },
        {
            'Gear': 'SAMPLE MNTT DUE',
            'TRTR Date': '',
            'TRTR Next Open': '',
            'TRTR Reminder Sent': '',
            'MNTT Date': (today - pd.Timedelta(days=14)).strftime('%Y-%m-%d'),
            'MNTT Next Open': today.strftime('%Y-%m-%d'),
            'MNTT Reminder Sent': 'No'
        }
    ], columns=COLUMNS)
    df = pd.concat([df, new_rows], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

def main():
    df = load_data()
    due, trtr_idx, mntt_idx = get_due_reminders(df)
    if due:
        body = '\n'.join(due)
        send_email('Gear Reminder - Action Needed Today', body)
        # Mark reminders as sent
        if len(trtr_idx) > 0:
            df.loc[trtr_idx, 'TRTR Reminder Sent'] = 'Yes'
        if len(mntt_idx) > 0:
            df.loc[mntt_idx, 'MNTT Reminder Sent'] = 'Yes'
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
    else:
        print('No reminders due today.')

if __name__ == '__main__':
    add_today_due_sample()
    main()
