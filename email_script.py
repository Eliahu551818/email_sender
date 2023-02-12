# Python code to send email to a list of
# emails from a spreadsheet

# import the required libraries
import pandas as pd
import smtplib
from email.message import EmailMessage

# change these as per use
your_name = "Beit Mashiach LA"
your_email = "beitmashiachla@gmail.com"
your_password = "nddszshajvwmymmo"

# Python code to send email to a list of
# emails from a spreadsheet

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# reading the spreadsheet
email_list = pd.read_excel('/Users/eliahubenamotz/Downloads/coding/python/hi.xlsx')

# getting the names and the emails
names = email_list['name']
emails = email_list['email']



# iterate through the records
for i in range(len(emails)):

	# for every record get the name and the email addresses
	name = names[i]
	email = emails[i]

	# the message to be emailed
	message = (
    f"From: {your_name} <{your_email}>\n"
    f"To: {name} <{email}>\n"
    """Subject: טיסה לרבי - י' שבט תשפ"ג\n\n"""
    f"""\nשלום וברכה\nלכבוד {name}\n\nאנחנו שמחים ונרגשים להודיע לכם על טיסה המאורגנת שלנו ל770 - בית משיח\nבשעה טובה, מתוכנן לו"ז כדלקמן\n- יוצאים ביום רביעי בשעה 12, אחרי שחרית וארוחת בוקר\n- מגיעים לשם בשעה 9 בערב ומתוודעים עד 3 בבוקר\n- הולכים לישון\n- קמים למקווה ושחרית עם המלך\n- ארוחת בוקר כיד המלך\n- טיסת חזור בשעה 2\n\nבאהבה, הרב מנדי ורוחמה דיין - בית משיח לוס-אנג'לס""").encode('utf8')

	# sending the email
	server.sendmail(your_email, [email], message)

# close the smtp server
server.close()
