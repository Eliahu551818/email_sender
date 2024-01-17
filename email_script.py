# Python code to send email to a list of
# emails from a spreadsheet

# import the required libraries
import pandas as pd
import smtplib
from email.message import EmailMessage

# change these as per use
your_name = "My Name"
your_email = "myemail@gmail.com"
your_password = "My_Password_123"

# Python code to send email to a list of
# emails from a spreadsheet

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# reading the spreadsheet
email_list = pd.read_excel('emails_list.xlsx')

# getting the names and the emails
names = email_list['name']
emails = email_list['email']

subject = """my subject"""
content = ("content").encode('utf8')

# iterate through the records
for i in range(len(emails)):

	# for every record get the name and the email addresses
	name = names[i]
	email = emails[i]

	# the message to be emailed
	message = (
    f"From: {your_name} <{your_email}>\n"
    f"To: {name} <{email}>\n"
    f"""Subject: {subject}"""
	content
	# sending the email
	server.sendmail(your_email, [email], message)

# close the smtp server
server.close()
