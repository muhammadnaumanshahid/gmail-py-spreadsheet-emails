import pandas as pd
import smtplib
import time

# change these as per use
your_email = "your gmail"
your_password = "your gmail app password"

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
        subject = "Result (BT1101): Tutorial 1 (Part-II)"
        message = "Subject: {}\n\n{}".format(subject, name)

        # sending the email
        server.sendmail(your_email, [email], message)
        sent_emails += 1
        print("{} emails sent out of {}".format(sent_emails, total_emails))

        # adding a delay of 10 seconds before sending the next email
        time.sleep(10)

    # close the smtp server
    server.close()

except Exception as e:
    print("An error occurred: ", e)
