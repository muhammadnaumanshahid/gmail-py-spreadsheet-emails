# Python code to send email to a list of emails from a spreadsheet

# import the required libraries
import pandas as pd
import smtplib

# change these as per use
your_email = "<Insert your Gmail address here>"
your_password = "<Insert app password here. Please check ReadMe file to know how to generate password>"

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
for i in range(len(emails)):

	# for every record get the name and the email addresses
	name = names[i]
	email = emails[i]

	# the message to be emailed
	subject = "Assignment BT1101 Result"
	message = "Subject: {}\n\n{}".format(subject, name)

	# sending the email
	server.sendmail(your_email, [email], message)

# close the smtp server
server.close()
