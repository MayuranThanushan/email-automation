# Automated Email Sending with Attachments

This Python script automates sending emails with optional attachments using data from an Excel file. The script uses the `smtplib` library to send emails via Gmail's SMTP server and `pandas` to read recipient details from an Excel sheet.

## Prerequisites

Before running the script, make sure you have the following:

1. **Python** installed on your machine.
2. Required Python libraries:
   - `pandas`: To read data from the Excel sheet.
   - `smtplib`, `email`: To handle email sending.
   
   You can install these packages by running:
   ```bash
   pip install pandas
   ```

3. A **Gmail account** with an **App Password**:
   - You need to generate an App Password in Gmail for this script to work because Gmail blocks less secure apps.
   - To generate an App Password:
     - Go to your Google Account.
     - Navigate to **Security** > **App Passwords**.
     - Generate a password for **Mail**.

## File Structure

- `recipients_list.xlsx`: An Excel file that contains the list of email recipients and the file paths of the attachments (if any).
- `email_automation.py`: The Python script to send emails using the details from the Excel file.

## Excel File Format

The `recipients_list.xlsx` file should contain the following columns:

| Name        | Email                | Attachment                              |
|-------------|----------------------|-----------------------------------------|
| John Doe    | john@example.com      | /path/to/john_attachment.xlsx           |
| Jane Smith  | jane@example.com      | /path/to/jane_attachment.pdf            |
| ...         | ...                   | ...                                     |

- **Name**: The name of the recipient (used in the email body).
- **Email**: The recipient's email address.
- **Attachment**: The absolute path to the file to be attached to the email. If there is no attachment for a recipient, leave the cell blank.

## Script Breakdown

### `send_email(receiver_email, receiver_name, attachment_path=None)`

This function sends an email to a specified recipient with an optional attachment.

- **Parameters**:
  - `receiver_email`: The recipient's email address.
  - `receiver_name`: The name of the recipient (used in the email body).
  - `attachment_path`: The file path of the attachment. If not provided, the email is sent without an attachment.

- **Email Body**:
  The body of the email can be customized in the script where the `body` variable is defined.

- **Email Subject**: You can change the subject of the email by modifying this line:
  ```python
  msg['Subject'] = "Your Subject Here"
  ```

### `for index, row in data.iterrows()`

This loop processes each row in the Excel sheet, extracting the recipient's name, email address, and attachment path. It then calls `send_email()` for each recipient.


- If an attachment is not found or there is an error while sending the email, the script will print an error message in the console.

- This script automates the process of sending emails with optional attachments using data from an Excel sheet. It is designed to work with Gmail but can be adapted to work with other email providers by modifying the SMTP server settings.
