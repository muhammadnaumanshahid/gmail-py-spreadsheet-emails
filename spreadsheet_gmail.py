import random
import pandas as pd
import smtplib
import time
import configparser
import re

config = configparser.ConfigParser()
config.read('config.ini')

# User inputs
your_email = config.get('user', 'email')
your_password = config.get('user', 'password')
subject = config.get('email', 'subject')

# Validate inputs
if not re.match(r"[^@]+@[^@]+\.[^@]+", your_email):
    print("Invalid email address")
    exit()

if len(your_password) < 8:
    print("Password is too short")
    exit()

try:
    # establishing connection with gmail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.ehlo()
        server.login(your_email, your_password)

        # reading the spreadsheet
        email_list = pd.read_excel('student_emails.xlsx')

        # getting the names and the emails
        names = email_list['Message']
        emails = email_list['Email']

        # iterate through the records
        total_emails = len(emails)
        sent_emails = 0
        success_emails = []

        for i, email in enumerate(emails):
            # for every record get the name and the email addresses
            name = names[i]

            # the message to be emailed
            message = f"Subject: {subject}\n\n{name}".encode('utf-8')

            # sending the email
            try:
                server.sendmail(your_email, [email], message)
                success_emails.append(email)
                print(f"Email sent to: {email}")
            except smtplib.SMTPException as e:
                print(f"An error occurred while sending email to {email}: {e}")

            sent_emails += 1
            print(f"{sent_emails} of {total_emails} emails sent")
            time.sleep(random.randint(20, 60))

    # Save successful emails to Excel file
    success_df = pd.DataFrame({'Email': success_emails})
    success_df.to_excel('success_emails.xlsx', index=False)
    print(f"{len(success_emails)} emails successfully sent and saved to 'success_emails.xlsx'")

except smtplib.SMTPAuthenticationError:
    print("Incorrect email address or password")
except smtplib.SMTPConnectError:
    print("Could not connect to SMTP server")
except pd.errors.EmptyDataError:
    print("The spreadsheet is empty or could not be read")
