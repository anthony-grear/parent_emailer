import pandas as pd

#using the r before the string converted to a raw string
df = pd.read_excel(r'C:\Users\Anthony\Documents\python_work'\
					r'\Personal Projects\import_data.xlsx') 
print(df)

print(df.to_dict()) #converts the dataframe to a dictionary and prints it

