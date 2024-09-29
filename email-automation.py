import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

file_path = 'recipients_list.xlsx' # Location of the recipients list file
data = pd.read_excel(file_path)

sender_email = "your_email@gmail.com"
sender_password = "your_app_password" # Get the password from app password

smtp_server = "smtp.gmail.com"
smtp_port = 587

def send_email(receiver_email, receiver_name, attachment_path=None):
    """
    Sends an email to a specified recipient with an optional attachment.

    Parameters:
    receiver_email (str): The email address of the recipient.
    receiver_name (str): The name of the recipient.
    attachment_path (str, optional): The path to the attachment file. Defaults to None.

    Returns:
    None
    """
    
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Subject"

        # Change the body as per your wish.
        body = f'''Hello {receiver_name},
        
        Body.
        
        Best regards,
        Name'''

        msg.attach(MIMEText(body, 'plain'))

        # Checks the list and adds attachment if specified.
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as attachment:
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

# Process the recipients list and send emails.
for index, row in data.iterrows():
    receiver_email = row['Email']
    receiver_name = row['Name']
    attachment_path = row['Attachment'] if 'Attachment' in row else None
    
    send_email(receiver_email, receiver_name, attachment_path)

print("All emails have been processed.")
