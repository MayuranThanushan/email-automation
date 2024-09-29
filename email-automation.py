import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

file_path = 'recipients_list.xlsx' 
data = pd.read_excel(file_path)

sender_email = "your_email@gmail.com"
sender_password = "your_app_password" 

smtp_server = "smtp.gmail.com"
smtp_port = 587

def send_email(receiver_email, receiver_name, attachment_path=None):
    
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Subject of the email"

        body = f"Hello {receiver_name},\n\nThis is an automated email with an attachment.\n\nBest regards,\nYour Name"
        msg.attach(MIMEText(body, 'plain'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as attachment:
                # Create MIMEBase object
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)

        else:
            print(f"Attachment not found for {receiver_email}, skipping attachment.")

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()

        print(f"Email sent successfully to {receiver_email}")

    except Exception as e:
        print(f"Failed to send email to {receiver_email}. Error: {str(e)}")

for index, row in data.iterrows():
    receiver_email = row['Email']
    receiver_name = row['Name']
    attachment_path = row['Attachment'] if 'Attachment' in row else None
    
    send_email(receiver_email, receiver_name, attachment_path)

print("All emails have been processed.")
