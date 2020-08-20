
import os #import environment variables

gmail_user = os.environ.get('GMAIL_USER') #this is my gmail log-in
gmail_password = os.environ.get('GMAIL_PASS') #this is my gmail special password

from datetime import date #imports the date
today = date.today()      #define variable with the date
d1 = today.strftime("%B %d, %Y") #format the date

 

import smtplib
#using the smtp lib call (location for gmail, port) as smtp is the variable
with smtplib.SMTP('smtp.gmail.com', 587) as smtp:  
	smtp.ehlo() #identifies ourselves with the mail server
	smtp.starttls() #encrypt traffic
	smtp.ehlo() #re-introduce ourselves as an encrypted connection

	smtp.login(gmail_user, gmail_password)#log-in using environment variables

	subject = 'Test Email using Python'
	body = "Dear Parents,\n\n\tYour child did not"\
							" submit their homework assignment on"\
							f" {d1}. It will be recorded as a zero. "\
							"Please encourage them to be more organized with "\
							"their assignments.\n\n\nBest Regards, "\
							"\n\n"\
							"Anthony Grear\nMath Instructor\n\n"\
							"Singapore International School"\
							" @ Gamuda Gardens\n"\
							"Accredited by the Western"\
							" Association of Schools"\
							" and Colleges (WASC)"

	msg = f'Subject:{subject}\n\n{body}'

	smtp.sendmail(gmail_user, 'anthony.j.grear@gmail.com', msg)


