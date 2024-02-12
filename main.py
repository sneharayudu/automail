import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def read_body_from_file(body_file):
    body = body_file.read().decode("utf-8")
    return body

def send_email(sender_email, sender_password, sender_name, receiver_email, receiver_name, resume_file, body_text, subject, progress_bar):
    # Email content
    message = MIMEMultipart()
    message['From'] = f'{sender_name} <{sender_email}>'
    message['To'] = f'{receiver_name} <{receiver_email}>'
    message['Subject'] = subject #"Application for Reactjs developer position"
    body_text = "Hi {receiver_name},\n\n" + body_text
    body = body_text.format(receiver_name=receiver_name)
    message.attach(MIMEText(body, 'plain'))

    # Attachment
    resume_path = resume_file.name
    part = MIMEBase("application", "octet-stream")
    part.set_payload(resume_file.getvalue())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {resume_path}",
    )
    message.attach(part)

    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        progress_bar.progress(1)

def main():
    st.title("MSS Techno Email Sender")
    sender_email = st.text_input("Enter your email address:")
    sender_password = st.text_input("Enter your email password:", type="password")
    sender_name = st.text_input("Enter your name:")
    subject = st.text_input("Enter Subject")
    body_file = st.file_uploader("Select body file (body.txt)", type=["txt"])

    if body_file is not None:
        body_text = read_body_from_file(body_file)

        excel_file = st.file_uploader("Select Excel file with Name and Email columns", type=["xlsx", "csv"])

        if excel_file is not None:
            df = pd.read_excel(excel_file) if excel_file.name.endswith('.xlsx') else pd.read_csv(excel_file)
            if 'name' not in df.columns or 'mail' not in df.columns:
                st.error("Error: Excel is in the wrong format. Columns should be named 'name' and 'mail'.")
                return

            name_column = 'name'
            email_column = 'mail'

            resume_file = st.file_uploader("Select resume file (resume.docx)", type=["docx"])

            if st.button("Send Emails"):
                progress_bar = st.progress(0)
                total_emails = len(df)
                emails_sent = 0
                for index, row in df.iterrows():
                    receiver_name = row[name_column]
                    receiver_email = row[email_column]
                    send_email(sender_email, sender_password, sender_name, receiver_email, receiver_name, resume_file, body_text, subject, progress_bar)
                    emails_sent += 1
                    progress_bar.progress(emails_sent / total_emails)
                st.success("Emails sent successfully!")

if __name__ == "__main__":
    main()
