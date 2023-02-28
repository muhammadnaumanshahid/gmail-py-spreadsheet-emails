import random
import pandas as pd
import smtplib
import time

# User inputs
your_email = input("Enter your Gmail address: ")
your_password = input("Enter your Gmail app password: ")
subject = input("Enter the email subject: ")

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

        for i, email in enumerate(emails):
            # for every record get the name and the email addresses
            name = names[i]

            # the message to be emailed
            message = f"Subject: {subject}\n\n{name}".encode('utf-8')

            # sending the email
            try:
                server.sendmail(your_email, [email], message)
            except smtplib.SMTPException as e:
                print(f"An error occurred while sending email to {email}: {e}")
                continue

            sent_emails += 1
            print(f"{sent_emails} emails sent out of {total_emails}")
            print(f"Email sent to: {email}")

            # adding a random delay between 20 and 60 seconds before sending the next email
            delay = random.randint(20, 60)
            time.sleep(delay)

    print("All emails sent successfully")

except Exception as e:
    print("An error occurred: ", e)
