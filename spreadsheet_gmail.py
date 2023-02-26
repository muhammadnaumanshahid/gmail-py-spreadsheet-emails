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
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
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

    for i in range(len(emails)):
        # for every record get the name and the email addresses
        name = names[i]
        email = emails[i]

        # the message to be emailed
        message = "Subject: {}\n\n{}".format(subject, name).encode('utf-8')

        # sending the email
        server.sendmail(your_email, [email], message)
        sent_emails += 1
        print("{} emails sent out of {}".format(sent_emails, total_emails))
        print("Email sent to: {}".format(email))

        # adding a random delay between 20 and 60 seconds before sending the next email
        delay = random.randint(20, 60)
        time.sleep(delay)

    # close the smtp server
    server.close()

except Exception as e:
    print("An error occurred: ", e)
