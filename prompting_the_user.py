"""Pandas is a package for working with labeled data. The smtplib module 

defines an SMTP client session object that can be used to send mail 

to any Internet machine with an SMTP listener daemon. The email.message is a 

module that we import EmailMessage class. This will allow us to setting and 

querying header fields, for accessing message bodies, and for creating or 

modifying structured messages. Import os (operating system interfaces) 

in order to get the username and password stored as environment variables. 

Usernames and passwords should not be stored inside code. 
"""
import pandas as pd 
import smtplib
from email.message import EmailMessage
import os
office_user = os.environ.get('OFFICE_USER')
office_pass = os.environ.get('OFFICE_PASS')

# import each email excel file into 4 different dataframes
df1 = pd.read_excel(r'C:\Users\Anthony\Documents\python_work'\
					   r'\Personal Projects\excel data'\
					   r'\7LB Parent Email List.xlsx')

df2 = pd.read_excel(r'C:\Users\Anthony\Documents\python_work'\
					   r'\Personal Projects\excel data'\
					   r'\IGCSE 2C Parent Email.xlsx')

df3 = pd.read_excel(r'C:\Users\Anthony\Documents\python_work'\
					   r'\Personal Projects\excel data'\
					   r'\IGCSE 2B Parent Email.xlsx')

df4 = pd.read_excel(r'C:\Users\Anthony\Documents\python_work'\
					   r'\Personal Projects\excel data'\
					   r'\Parent Email IGCSE 1C.xlsx')

#imports the date, defines the variable and formats it
from datetime import date 
today = date.today()      
date = today.strftime("%B %d, %Y") 

choose_again = 'yes'
while choose_again == 'yes' or choose_again == 'y':
	class_name = input('Press 1 for 7LB\nPress 2 for 2C\n'\
	               'Press 3 for 2B\nPress 4 for 1C\n')
	# print (f'You have chosen {class_name}!')

	if class_name == '1':
		#set student number as the index
		df1 = df1.set_index('Student Number') 
		print(df1)							  	
		print ()  

		#user inputs student number and change to integer value
		s_num = input('Enter student number: ') 
		s_num = int(s_num)						
		
		#filter by user entered student number and 3 column string tags
		print(df1.loc[df1.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
		print()

		#filter by student number and column string
		father_email = df1.loc[s_num, "Father Email"]
		mother_email = df1.loc[s_num, "Mother Email"]
		print(f'Father Email: {father_email}\n')
		print(f'Mother Email: {mother_email}\n')

		pa_or_ma = input('Press 1 to email the father.\n'
						 'Press 2 to email the mother.\n')
		if pa_or_ma == '1':
			to_address = father_email
		elif pa_or_ma == '2':
			to_address = mother_email

		#The email message is contained in a plain text file.
		text_file = open("C:\\Users\\Anthony\\Documents\\python_work\\"
	     "Personal Projects\\Email Txt Files\\"
	     "Missing Homework Email.txt", "r")
		msg = EmailMessage()
		msg.set_content(text_file.read())

		msg['Subject']='Email Testing with Python'
		msg['From'] = office_user
		msg['To'] = to_address

		"""Set the host and port. Identify yourself to ESMTP server using EHLO.
		Start tls encryption. Login. Send message. Close connection.
		"""
		smtpsrv = 'smtp.office365.com'
		smtpserver = smtplib.SMTP(smtpsrv,587)
		smtpserver.ehlo()
		smtpserver.starttls()
		smtpserver.login(office_user,office_pass)
		smtpserver.send_message (msg)
		smtpserver.close()

		#reset the dataframe back to using integers as the index
		df1 = df1.reset_index()

	elif class_name == '2':
		df2 = df2.set_index('Student Number')
		print(df2)
		print()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df2.loc[df2.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
		print()
		father_email = df2.loc[s_num, "Father Email"]
		mother_email = df2.loc[s_num, "Mother Email"]
		print(f'Father Email: {father_email}\n')
		print(f'Mother Email: {mother_email}\n')

		pa_or_ma = input('Press 1 to email the father.\n'
						 'Press 2 to email the mother.\n')
		if pa_or_ma == '1':
			to_address = father_email
		elif pa_or_ma == '2':
			to_address = mother_email

		text_file = open("C:\\Users\\Anthony\\Documents\\python_work\\"
	     "Personal Projects\\Email Txt Files\\"
	     "Missing Homework Email.txt", "r")
		msg = EmailMessage()
		msg.set_content(text_file.read())

		msg['Subject']='Email Testing with Python'
		msg['From'] = office_user
		msg['To'] = to_address

		smtpsrv = 'smtp.office365.com'
		smtpserver = smtplib.SMTP(smtpsrv,587)
		smtpserver.ehlo()
		smtpserver.starttls()
		smtpserver.login(office_user,office_pass)
		smtpserver.send_message (msg)
		smtpserver.close()

		df2 = df2.reset_index()

	elif class_name == '3':
		df3 = df3.set_index('Student Number')
		print(df3)
		print()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df3.loc[df3.index[s_num - 1:s_num],["Student Name",
										"Father Email","Mother Email"]])
		print()
		father_email = df3.loc[s_num, "Father Email"]
		mother_email = df3.loc[s_num, "Mother Email"]
		print(f'Father Email: {father_email}\n')
		print(f'Mother Email: {mother_email}\n')

		pa_or_ma = input('Press 1 to email the father.\n'
						 'Press 2 to email the mother.\n')
		if pa_or_ma == '1':
			to_address = father_email
		elif pa_or_ma == '2':
			to_address = mother_email

		text_file = open("C:\\Users\\Anthony\\Documents\\python_work\\"
	     "Personal Projects\\Email Txt Files\\"
	     "Missing Homework Email.txt", "r")
		msg = EmailMessage()
		msg.set_content(text_file.read())

		msg['Subject']='Email Testing with Python'
		msg['From'] = office_user
		msg['To'] = to_address

		smtpsrv = 'smtp.office365.com'
		smtpserver = smtplib.SMTP(smtpsrv,587)
		smtpserver.ehlo()
		smtpserver.starttls()
		smtpserver.login(office_user,office_pass)
		smtpserver.send_message (msg)
		smtpserver.close()

		df3 = df3.reset_index()

	elif class_name == '4':
		df4 = df4.set_index('Student Number') 
		print(df4)
		print ()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df4.loc[df4.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
		print()
		father_email = df4.loc[s_num, "Father Email"]
		mother_email = df4.loc[s_num, "Mother Email"]
		print(f'Father Email: {father_email}\n')
		print(f'Mother Email: {mother_email}\n')

		pa_or_ma = input('Press 1 to email the father.\n'
						 'Press 2 to email the mother.\n')
		if pa_or_ma == '1':
			to_address = father_email
		elif pa_or_ma == '2':
			to_address = mother_email

		text_file = open("C:\\Users\\Anthony\\Documents\\python_work\\"
	     "Personal Projects\\Email Txt Files\\"
	     "Missing Homework Email.txt", "r")
		msg = EmailMessage()
		msg.set_content(text_file.read())

		msg['Subject']='Email Testing with Python'
		msg['From'] = office_user
		msg['To'] = to_address

		smtpsrv = 'smtp.office365.com'
		smtpserver = smtplib.SMTP(smtpsrv,587)
		smtpserver.ehlo()
		smtpserver.starttls()
		smtpserver.login(office_user,office_pass)
		smtpserver.send_message (msg)
		smtpserver.close()

		df4 = df4.reset_index()

	choose_again = input('Do you want to send another email? (yes or no)')
	choose_again = choose_again.lower() 
	
x = input('Press enter to end the program.')