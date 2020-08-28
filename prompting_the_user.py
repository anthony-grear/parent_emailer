import pandas as pd #package for working with labeled data

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

from datetime import date #imports the date
today = date.today()      #define variable with the date
date = today.strftime("%B %d, %Y") #format the date

subject = 'Missing Homework Assignment'
body = "Dear Parents,\n\n\tYour child did not"\
							" submit their homework assignment on"\
							f" {date}. It will be recorded as a zero. "\
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

choose_again = 'yes'

while choose_again == 'yes' or choose_again == 'y':
	class_name = input('Choose a class:\n2B\n2C\n1C\n7LB\n\n')
	print (f'You have chosen {class_name}!')

	if class_name == '7LB' or class_name == '7lb':
		
		df1 = df1.set_index('Student Number') #set student number as the index
		print(df1)							  #print whole class list	
		print ()                              #skip a line
		
		s_num = input('Enter student number: ') #user inputs student number
		s_num = int(s_num)						#change to integer value
		
		#filter by user entered student number and 3 column string tags
		print(df1.loc[df1.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
		print()
		#filter by student number and column string
		father_email = df1.loc[s_num, "Father Email"]
		mother_email = df1.loc[s_num, "Mother Email"]
		print(f'Father Email: {father_email}\n')
		print(f'Mother Email: {mother_email}\n')
		#reset the dataframe back to using integers as the index
		df1 = df1.reset_index()

	elif class_name == '2C' or class_name == '2c':
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
		df2 = df2.reset_index()

	elif class_name == '2B' or class_name == '2b':
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
		df3 = df3.reset_index()

	elif class_name == '1C' or class_name == '1c':
		df4 = df4.set_index('Student Number') #set index to student number
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
		df4 = df4.reset_index()

	choose_again = input('Do you want to send another email? (yes or no)')
	choose_again = choose_again.lower() #changes string to lowercase
	
x = input('Press enter to end the program.')