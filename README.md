# Streamlit Excel Gear Reminder App

This app allows you to track gears, mark when they are opened in TRTR and MNTT, and receive reminders after 14 days. Data is stored in an Excel file.

## Features
- Dropdown to select gear
- Date pickers for TRTR and MNTT
- Mark as opened (updates Excel)
- Automatic reminder after 14 days
- Reminders shown in the app UI

## How to Run
1. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```
2. Start the app:
   ```powershell
   streamlit run app.py
   ```

## Excel Template
The app uses `gears.xlsx` with these columns:
- Gear
- TRTR Date
- MNTT Date
- Marked (Yes/No)
- Next Reminder Date
- Reminder Sent (Yes/No)

The app will create this file if it does not exist.
