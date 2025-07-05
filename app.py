import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

EXCEL_FILE = 'gears.xlsx'
COLUMNS = [
    'Gear',
    'TRTR Date', 'TRTR Next Open', 'TRTR Reminder Sent',
    'MNTT Date', 'MNTT Next Open', 'MNTT Reminder Sent'
]

EMAIL_TO = 'febinaz@gmail.com'  # Change as needed
EMAIL_FROM = 'automaticmail293@gmail.com'  # Change as needed
EMAIL_PASS = st.secrets["EMAIL_PASS"]  # Use Streamlit secrets for security

# Initialize Excel file if not exists
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

def load_data():
    try:
        return pd.read_excel(EXCEL_FILE, engine='openpyxl')
    except Exception:
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        return df

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

def migrate_excel():
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        # If columns are not as expected, migrate
        if set(df.columns) != set(COLUMNS):
            new_df = pd.DataFrame(columns=COLUMNS)
            for _, row in df.iterrows():
                gear = row.get('Gear', '')
                # Try to map old columns to new columns
                trtr_date = row.get('TRTR Date', '')
                mntt_date = row.get('MNTT Date', '')
                # Calculate next open dates if possible
                trtr_next = pd.to_datetime(trtr_date, errors='coerce') + timedelta(days=14) if pd.notna(trtr_date) and trtr_date != '' else ''
                mntt_next = pd.to_datetime(mntt_date, errors='coerce') + timedelta(days=14) if pd.notna(mntt_date) and mntt_date != '' else ''
                new_df = new_df.append({
                    'Gear': gear,
                    'TRTR Date': trtr_date,
                    'TRTR Next Open': trtr_next if trtr_next != '' else '',
                    'TRTR Reminder Sent': 'No' if trtr_next != '' else '',
                    'MNTT Date': mntt_date,
                    'MNTT Next Open': mntt_next if mntt_next != '' else '',
                    'MNTT Reminder Sent': 'No' if mntt_next != '' else ''
                }, ignore_index=True)
            new_df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
    except Exception:
        # If file is unreadable, create a new one with sample data
        new_df = pd.DataFrame([
            {'Gear': 'SampleGear1', 'TRTR Date': '2025-07-01', 'TRTR Next Open': '2025-07-15', 'TRTR Reminder Sent': 'No',
             'MNTT Date': '2025-07-02', 'MNTT Next Open': '2025-07-16', 'MNTT Reminder Sent': 'No'},
            {'Gear': 'SampleGear2', 'TRTR Date': '', 'TRTR Next Open': '', 'TRTR Reminder Sent': '',
             'MNTT Date': '', 'MNTT Next Open': '', 'MNTT Reminder Sent': ''}
        ], columns=COLUMNS)
        new_df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

def populate_sample_data():
    sample = [
        {'Gear': 'POINTS - ERS SIDE', 'TRTR Date': '2025-06-09', 'MNTT Date': '2025-06-20'},
        {'Gear': 'POINTS - KTYM SIDE', 'TRTR Date': '2025-06-17', 'MNTT Date': '2025-06-20'},
        {'Gear': 'TC & SIG - KTYM SIDE', 'TRTR Date': '2025-06-24', 'MNTT Date': '2025-06-20'},
        {'Gear': 'TC & SIG - ERS PF SIDE', 'TRTR Date': '2025-06-08', 'MNTT Date': '2025-06-27'},
        {'Gear': 'TC & SIG - LC-7 SIDE', 'TRTR Date': '2025-06-27', 'MNTT Date': ''},
        {'Gear': 'BLOCK - KTYM SIDE', 'TRTR Date': '2025-06-26', 'MNTT Date': ''},
        {'Gear': 'BLOCK - ERS SIDE', 'TRTR Date': '2025-06-25', 'MNTT Date': '2025-06-28'},
        {'Gear': 'RELAY ROOM & PANEL', 'TRTR Date': '2025-06-06', 'MNTT Date': '2025-06-06'},
        {'Gear': 'IPS & BATTERY', 'TRTR Date': '2025-06-15', 'MNTT Date': '2025-06-28'},
        {'Gear': 'HASSDAC - ERS SIDE', 'TRTR Date': '2025-06-27', 'MNTT Date': '2025-06-27'},
        {'Gear': 'HASSDAC - KTYM SIDE', 'TRTR Date': '2025-06-24', 'MNTT Date': ''},
        {'Gear': 'DATALOGGER', 'TRTR Date': '2025-06-25', 'MNTT Date': '2025-06-28'},
        {'Gear': 'ELD', 'TRTR Date': '2025-06-15', 'MNTT Date': '2025-06-28'},
        {'Gear': 'FIRE ALARM', 'TRTR Date': '2025-06-25', 'MNTT Date': '2025-06-28'},
        {'Gear': 'CRANK HANDLE', 'TRTR Date': '2025-01-23', 'MNTT Date': '2025-01-19'},
        {'Gear': 'ERS - D', 'TRTR Date': '', 'MNTT Date': '2025-06-27'},
        {'Gear': 'LC GATES', 'TRTR Date': '', 'MNTT Date': 'DATES'},
        {'Gear': 'LC - 5', 'TRTR Date': '', 'MNTT Date': '2025-06-03'},
        {'Gear': 'LC - 6', 'TRTR Date': '', 'MNTT Date': '2025-06-03'},
        {'Gear': 'LC - 7 (TF)', 'TRTR Date': '', 'MNTT Date': '2025-04-20'},
        {'Gear': 'LC - 9', 'TRTR Date': '', 'MNTT Date': '2025-06-22'},
        {'Gear': 'LC - 10', 'TRTR Date': '', 'MNTT Date': '2025-06-22'},
        {'Gear': 'LC - 11', 'TRTR Date': '', 'MNTT Date': '2025-06-22'},
        {'Gear': 'LC - 12', 'TRTR Date': '', 'MNTT Date': '2025-05-26'},
    ]
    rows = []
    for entry in sample:
        trtr_date = pd.to_datetime(entry['TRTR Date'], errors='coerce')
        mntt_date = pd.to_datetime(entry['MNTT Date'], errors='coerce')
        trtr_next = (trtr_date + timedelta(days=14)).strftime('%Y-%m-%d') if not pd.isna(trtr_date) else ''
        mntt_next = (mntt_date + timedelta(days=14)).strftime('%Y-%m-%d') if not pd.isna(mntt_date) else ''
        rows.append({
            'Gear': entry['Gear'],
            'TRTR Date': entry['TRTR Date'],
            'TRTR Next Open': trtr_next,
            'TRTR Reminder Sent': 'No' if trtr_next else '',
            'MNTT Date': entry['MNTT Date'],
            'MNTT Next Open': mntt_next,
            'MNTT Reminder Sent': 'No' if mntt_next else ''
        })
    df = pd.DataFrame(rows, columns=COLUMNS)
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

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
        return True, 'Email sent!'
    except Exception as e:
        return False, f'Failed to send email: {e}'

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

# Call migration before loading data
migrate_excel()
# Populate with sample data if file is empty
if pd.read_excel(EXCEL_FILE, engine='openpyxl').empty:
    populate_sample_data()
# Removed sample data addition for testing

init_excel()
df = load_data()

# Remove the block that resets reminder sent flags based on date, as this is now handled only when marking as opened

st.title('Gear Reminder App')

# Dropdown for gears
gears = df['Gear'].dropna().unique().tolist()
new_gear = st.text_input('Add new gear (optional)')
if new_gear:
    if new_gear not in gears:
        # Add new gear to DataFrame and Excel
        new_row = {
            'Gear': new_gear,
            'TRTR Date': '',
            'TRTR Next Open': '',
            'TRTR Reminder Sent': '',
            'MNTT Date': '',
            'MNTT Next Open': '',
            'MNTT Reminder Sent': ''
        }
        df = df.append(new_row, ignore_index=True)
        save_data(df)
        gears.append(new_gear)
        st.success(f'Added new gear: {new_gear}')
selected_gear = st.selectbox('Select Gear', gears)

# Date pickers
trtr_date = st.date_input('TRTR Date', value=datetime.today())
mntt_date = st.date_input('MNTT Date', value=datetime.today())

col1, col2 = st.columns(2)
with col1:
    if st.button('Mark TRTR as Opened'):
        idx = df[df['Gear'] == selected_gear].index
        next_trtr = trtr_date + timedelta(days=14)
        if len(idx) > 0:
            df.loc[idx, 'TRTR Date'] = trtr_date
            df.loc[idx, 'TRTR Next Open'] = next_trtr
            # Reset reminder sent flag to 'No' when new open date is set
            df.loc[idx, 'TRTR Reminder Sent'] = 'No'
        else:
            df = df.append({
                'Gear': selected_gear,
                'TRTR Date': trtr_date,
                'TRTR Next Open': next_trtr,
                'TRTR Reminder Sent': 'No',
                'MNTT Date': '',
                'MNTT Next Open': '',
                'MNTT Reminder Sent': ''
            }, ignore_index=True)
        save_data(df)
        st.success(f'{selected_gear} TRTR marked as opened. Next open: {next_trtr.strftime("%Y-%m-%d")}')

with col2:
    if st.button('Mark MNTT as Opened'):
        idx = df[df['Gear'] == selected_gear].index
        next_mntt = mntt_date + timedelta(days=14)
        if len(idx) > 0:
            df.loc[idx, 'MNTT Date'] = mntt_date
            df.loc[idx, 'MNTT Next Open'] = next_mntt
            # Reset reminder sent flag to 'No' when new open date is set
            df.loc[idx, 'MNTT Reminder Sent'] = 'No'
        else:
            df = df.append({
                'Gear': selected_gear,
                'TRTR Date': '',
                'TRTR Next Open': '',
                'TRTR Reminder Sent': '',
                'MNTT Date': mntt_date,
                'MNTT Next Open': next_mntt,
                'MNTT Reminder Sent': 'No'
            }, ignore_index=True)
        save_data(df)
        st.success(f'{selected_gear} MNTT marked as opened. Next open: {next_mntt.strftime("%Y-%m-%d")}')

# Reminders for TRTR
reminders_trtr = df[(df['TRTR Next Open'].notna()) & (df['TRTR Reminder Sent'] != 'Yes') & (pd.to_datetime(df['TRTR Next Open'], errors='coerce').dt.date == datetime.today().date())]
reminders_mntt = df[(df['MNTT Next Open'].notna()) & (df['MNTT Reminder Sent'] != 'Yes') & (pd.to_datetime(df['MNTT Next Open'], errors='coerce').dt.date == datetime.today().date())]
reminder_msgs = []
if not reminders_trtr.empty:
    st.warning('TRTR Reminders Due:')
    for _, row in reminders_trtr.iterrows():
        msg = f"Gear: {row['Gear']} - TRTR needs to be opened today! (Next Open: {row['TRTR Next Open']})"
        st.write(msg)
        reminder_msgs.append(msg)
    df.loc[reminders_trtr.index, 'TRTR Reminder Sent'] = 'Yes'
    save_data(df)
if not reminders_mntt.empty:
    st.warning('MNTT Reminders Due:')
    for _, row in reminders_mntt.iterrows():
        msg = f"Gear: {row['Gear']} - MNTT needs to be opened today! (Next Open: {row['MNTT Next Open']})"
        st.write(msg)
        reminder_msgs.append(msg)
    df.loc[reminders_mntt.index, 'MNTT Reminder Sent'] = 'Yes'
    save_data(df)

if st.button('Run Reminder Check and Send Email'):
    if reminder_msgs:
        subject = 'Gear Reminder - Action Needed Today'
        body = '\n'.join(reminder_msgs)
        success, info = send_email(subject, body)
        if success:
            st.success(info)
        else:
            st.error(info)
    else:
        st.info('No reminders due today.')

# Show all data
st.subheader('All Gear Data')
st.dataframe(df)
