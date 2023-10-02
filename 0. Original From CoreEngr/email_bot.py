import pandas as pd
import smtplib
import ssl

# Email credentials
sender_email = 'Fill_in_here'
password = 'Fill_in_here'

# Read the spreadsheet
df = pd.read_excel('recipients.xlsx')  # Modify the filename and path if needed
recipients = df[['Name', 'Email']]

# Email content
subject = 'Your Subject'

# SMTP server configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Connect to the SMTP server
context = ssl.create_default_context()
with smtplib.SMTP(smtp_server, smtp_port) as smtp_obj:
    smtp_obj.starttls(context=context)
    smtp_obj.login(sender_email, password)

    # Send email to each recipient
    for index, recipient in recipients.iterrows():
        try:
            # Extract name and email from the DataFrame
            name = recipient['Name']
            email = recipient['Email']

            # Personalized message
            message = f"Dear {name},\n\nYour personalized message goes here."

            # Create the email message
            email_message = f"Subject: {subject}\n\n{message}"

            # Send the email
            smtp_obj.sendmail(sender_email, email, email_message)
            print(f"Email sent successfully to {name} ({email})")
        except smtplib.SMTPException as e:
            print(f"Failed to send email to {name} ({email}): {str(e)}")
