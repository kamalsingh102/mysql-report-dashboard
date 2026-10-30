# MySQL Report Mailer (Streamlit)

A local dashboard to run an authorized stored procedure, pass JSON arguments, create a generic formatted Excel workbook, and email it to one or more recipients. It runs only when launched manually; nothing is scheduled.

## Windows setup
1. Copy your existing working `.env` into this folder. Alternatively copy `.env.example` to `.env` and fill in the credentials.
2. Keep database and SMTP credentials in `.env` only. Use the Yandex app password that already worked for SMTP.
3. Open Command Prompt in this folder and run:

   ```bat
   py -m venv .venv
   .venv\Scripts\activate
   python -m pip install -r requirements.txt
   streamlit run app.py
   ```
4. In the browser page, enter the stored procedure name, JSON array arguments, one or more recipient addresses, and an optional filename label. Click **Generate and send report**.

## Input examples
- No arguments: `[]`
- One date string: `["2026-10-01"]`
- Multiple values: `["2026-10-01", "2026-10-31", 1]`

The app accepts recipient addresses separated by newlines, commas, or semicolons. It removes duplicate addresses.

## Notes
- Only enter procedures you are authorized to execute.
- The app generates and emails the workbook only after you click the button.
- SMTP acceptance means the server accepted the message; it does not guarantee inbox delivery.
- Confirm company policy before sending company data to personal email addresses.
- Never share or commit `.env`.
