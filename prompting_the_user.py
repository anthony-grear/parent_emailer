import pandas as pd
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

# print(df1.columns)
# print(df2.columns)
# print(df3.columns)
# print(df4.columns)
choose_again = 'yes'

while choose_again == 'yes':
	class_name = input('Choose a class:\n2B\n2C\n1C\n7LB\n\n')
	print (f'You have chosen {class_name}!')
	if class_name == '7LB':
		df1 = df1.set_index('Student Number') #set index to student number
		print(df1)
		print ()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df1.loc[df1.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
	elif class_name == '2C':
		df2 = df2.set_index('Student Number')
		print(df2)
		print()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df2.loc[df2.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])
	elif class_name == '2B':
		df3 = df3.set_index('Student Number')
		print(df3)
		print()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df3.loc[df3.index[s_num - 1:s_num],["Student Name",
										"Father Email","Mother Email"]])
	elif class_name == '1C':
		df4 = df4.set_index('Student Number') #set index to student number
		print(df4)
		print ()
		s_num = input('Enter student number: ')
		s_num = int(s_num)
		print(df4.loc[df4.index[s_num - 1:s_num],["Student Name",
											"Father Email","Mother Email"]])

	choose_again = input('Do you want to send another email? (yes or no)')

x = input('Press enter to end the program.')