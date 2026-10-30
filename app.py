import json
import re
from email.utils import parseaddr
from pathlib import Path
import streamlit as st
from report_engine import create_excel, fetch_procedure_result_sets, send_report_email

st.set_page_config(page_title="MySQL Report Mailer", page_icon="📊", layout="wide")
st.title("MySQL Report Mailer")
st.caption("Run a stored procedure, create a formatted Excel report, and email it to selected recipients.")
st.info("Database and SMTP credentials are read from your local .env file; they are not entered in this page.")

with st.form("report_form"):
    st.subheader("Report configuration")
    procedure = st.text_input("Stored procedure name", placeholder="e.g. udsp_organicpaidleads")
    args_text = st.text_area("Procedure arguments (JSON array)", value="[]", help='Examples: [] or ["2026-10-01", "2026-10-31"]')
    recipients_text = st.text_area("Recipient email addresses", placeholder="person1@company.com\nperson2@company.com", help="One address per line; commas and semicolons also work.")
    report_label = st.text_input("Report filename label (optional)", placeholder="e.g. organic_paid_leads")
    submitted = st.form_submit_button("Generate and send report", type="primary", use_container_width=True)

if submitted:
    errors = []
    procedure = procedure.strip()
    if not procedure:
        errors.append("Enter a stored procedure name.")
    elif not re.fullmatch(r"[A-Za-z0-9_]+", procedure):
        errors.append("Procedure name can contain only letters, numbers, and underscores.")
    try:
        args = json.loads(args_text)
        if not isinstance(args, list):
            errors.append("Procedure arguments must be a JSON array, for example [].")
    except json.JSONDecodeError as exc:
        args = None
        errors.append(f"Arguments are not valid JSON: {exc.msg} (position {exc.pos}).")

    recipients = []
    for item in re.split(r"[\n,;]+", recipients_text):
        address = item.strip()
        if not address:
            continue
        _, parsed = parseaddr(address)
        if parsed != address or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", address):
            errors.append(f"Invalid email address: {address}")
        elif address.lower() not in {r.lower() for r in recipients}:
            recipients.append(address)
    if not recipients:
        errors.append("Enter at least one recipient email address.")

    if errors:
        for error in errors:
            st.error(error)
    else:
        status = st.status("Starting report generation…", expanded=True)
        try:
            status.write("Checking database connection and executing stored procedure.")
            result_sets = fetch_procedure_result_sets(procedure, tuple(args))
            status.write(f"Retrieved {len(result_sets)} tabular result set(s).")
            status.write("Creating and formatting Excel workbook.")
            path = create_excel(result_sets, report_label.strip() or procedure)
            status.write(f"Created {path.name}. Sending to {len(recipients)} recipient(s).")
            send_report_email(path, recipients)
            status.update(label="Report generated and email submission succeeded.", state="complete")
            st.success("The SMTP server accepted the report email.")
            st.write(f"**File:** `{path.name}`")
            st.write(f"**Recipients:** {', '.join(recipients)}")
            st.download_button("Download generated Excel report", data=path.read_bytes(), file_name=path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as exc:
            status.update(label="Report run failed.", state="error")
            st.error(f"The report could not be completed: {exc}")
            st.caption("Check logs/automation.log for details. Never share your .env file or credentials.")
