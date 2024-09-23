import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import os
import pandas as pd

# Function to send email with QR code and logo
def send_email(recipient, name, reg_no, sender_email, sender_password, logo_file, qr_code_file):
    subject = "🚀 GenAI Hackathon Invitation and Instruction Mail 🥳"
    body = f"""
    <html>
    <body>
    <img src="cid:logo" alt="Logo" style="width: 100%; max-width: 200px; height: 100%; max-height: 200px;" />
    <p>Dear {name},</p>
    <p>Welcome to the GenAI Hackathon 🤖! We are excited to have you with us. Here are the details and instructions:</p>
    <h4>Instructions:</h4>
    <ul>
        <li>Participants are asked to assemble in the venue (RNG Hall, SNSCT) by 8:45AM. Since verification process will be there, participants are asked to be come early and be seated.</li>
        <li>1st round (Technical MCQ) will be from 9:30AM to 10:30AM, conducted through Unstop. Make sure to use either a mobile or laptop (Unstop is not compatible with iOS).</li>
        <li>Shortlisted participants will move to the 2nd round (Coding) starting at 11:00AM. It will consist of Leetcode-style questions (easy, medium, hard). Participants must attend this on their laptops and ensure to bring all necessary items, including a charger, internet facility, and a junction box.</li>
        <li>Shortlisted participants will proceed to the live hackathon from 1:00PM to 4:30PM. You will be given 5 use cases to solve using your preferred tech stack, and upload your solutions to GitHub.</li>
        <li>The final results will be announced after evaluation.</li>
        <li><b>Please find your QR code attached below; it is essential for event entry and attendance.</b></li>
    </ul>
    <p>For more information, keep an eye on the WhatsApp group shared with you. If you're not in the group yet, please join as soon as possible to stay updated.</p>
    <p>Best of luck,<br>Team SNS iHUB</p>
    </body>
    </html>
    """
    

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = subject

    # Attach HTML body
    msg.attach(MIMEText(body, 'html'))

    # Attach logo
    with open(logo_file, 'rb') as f:
        logo = MIMEImage(f.read())
        logo.add_header('Content-ID', '<logo>')
        msg.attach(logo)

    # Attach QR code
    with open(qr_code_file, 'rb') as f:
        qr_code = MIMEApplication(f.read())
        qr_code.add_header('Content-Disposition', f'attachment; filename="{reg_no}_QRCode.png"')
        msg.attach(qr_code)

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Failed to send email to {recipient}: {e}")
        return False

# Streamlit app setup
st.title("GenAI Hackathon Mail Automation")

# Upload Excel file
excel_file = st.file_uploader("Upload the Excel file with participant data", type=["xlsx"])
qr_code_folder = "./dummy-qrcodes/"
logo_file = st.file_uploader("Upload the logo file (PNG)", type=["png"])

# Email credentials
# sender_email = st.text_input("Your Email", "sukanthkannansk@gmail.com")
# sender_password = st.text_input("Your Email Password", "qlbf kwqu nbwx bejo", type="password")

sender_email = st.text_input("Your Email", "Princestanlyvictor23@gmail.com")
sender_password = st.text_input("Your Email Password", "bcuu cfyr iiup zgrh", type="password")

if st.button("Send Emails"):
    if excel_file and qr_code_folder and logo_file and sender_email and sender_password:
        df = pd.read_excel(excel_file)

        # Prepare list to track participants with missing or invalid data
        invalid_participants = []

        # Save uploaded logo to a temporary file
        with open("temp_logo.png", "wb") as f:
            f.write(logo_file.read())

        # Iterate over each participant and send email
        for index, row in df.iterrows():
            name = row['Name']
            reg_no = row['Reg No']
            recipient = row['Domain Email ID (College ID)']

            # Check if Reg No or Email is missing
            if pd.isna(reg_no) or pd.isna(recipient):
                invalid_participants.append({
                    'Name': name,
                    'Reg No': reg_no,
                    'College': row['College'],
                    'Department': row['Department'],
                    'Whatsapp Number': row['Whatsapp Number']
                })
                st.warning(f"Skipping {name}: Missing Reg No or Email")
                continue

            # Path to QR code for the participant
            qr_code_file = os.path.join(qr_code_folder, f"{reg_no}.png")

            if os.path.exists(qr_code_file):
                success = send_email(recipient, name, reg_no, sender_email, sender_password, "temp_logo.png", qr_code_file)
                if success:
                    st.success(f"Email sent successfully to {name} ({recipient})")
                else:
                    st.error(f"Failed to send email to {name} ({recipient})")
            else:
                st.warning(f"QR code not found for {name} ({reg_no})")

        # If there are invalid participants, save them to an Excel file
        if invalid_participants:
            invalid_df = pd.DataFrame(invalid_participants)
            invalid_df.to_excel('invalid_participants.xlsx', index=False)
            st.success(f"Saved invalid participants to 'invalid_participants.xlsx'")
        else:
            st.info("No invalid participants found.")
    else:
        st.error("Please upload all required files and enter your email credentials.")
