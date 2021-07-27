import xlwings as xw
import smtplib
from email.message import EmailMessage
from datetime import date
import os
office_user = os.environ.get('OFFICE_USER')
office_pass = os.environ.get('OFFICE_PASS')
   
def send_email():
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    
    msg = EmailMessage()
    msg.set_content(sheet.range("B8").value)    
    msg['Subject'] = (sheet.range("B6").value)
    msg['From'] = office_user
    msg['To'] = (sheet.range("B3").value)
    msg['Cc'] = (sheet.range("B4").value)
    
    smtpsrv = 'smtp.office365.com'
    smtpserver = smtplib.SMTP(smtpsrv, 587)
    smtpserver.ehlo()
    smtpserver.starttls()
    smtpserver.login(office_user, office_pass) 
    smtpserver.send_message(msg)
    smtpserver.close()
    sheet.range("B2:B8").options(transpose = True).clear()
    
def get_parent_email():
    wb = xw.Book.caller()
    sheet = wb.sheets[1]
    active_cell = wb.app.selection
    active_cell_row = active_cell.row
    sheet = wb.sheets[2]
    sheet.range("B3").value = active_cell.value

def get_student_name():   
    wb = xw.Book.caller()
    sheet = wb.sheets[1]
    active_cell = wb.app.selection
    active_cell_row = active_cell.row
    sheet = wb.sheets[2]
    sheet.range("B2").value = active_cell.value
    
def get_absent_message():
    today = date.today()
    d1 = today.strftime("%d/%m/%Y")
    path = 'C:\\Python\\python_work\\Personal Projects\\parent_emailer\\parent_emailer_V2\\absent'
    f = open(path, 'r')
    content = f.read()
    wb = xw.Book.caller()    
    sheet = wb.sheets[2]
    student = sheet.range("B2").value
    sheet.range("B6").value = f"Regarding {student} absent from online class on {d1}."
    sheet.range("B8").value = content
    
def get_missing_work_message():
    today = date.today()
    d1 = today.strftime("%d/%m/%Y")
    path = 'C:\\Python\\python_work\\Personal Projects\\parent_emailer\\parent_emailer_V2\\missing_work'
    f = open(path, 'r')
    content = f.read()
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    student = sheet.range("B2").value
    sheet.range("B6").value = f"Regarding {student} and missing work on {d1}."
    sheet.range("B8").value = content

def get_no_communication_message():
    today = date.today()
    d1 = today.strftime("%d/%m/%Y")
    path = 'C:\\Python\\python_work\\Personal Projects\\parent_emailer\\parent_emailer_V2\\no_communication'
    f = open(path, 'r')
    content = f.read()
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    student = sheet.range("B2").value
    sheet.range("B6").value = f"Regarding {student} and no communication online class on {d1}."
    sheet.range("B8").value = content
    
def get_carbon_copy():
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    active_cell = wb.app.selection
    sheet.range("B4").value = f"{active_cell.value}@gamudagardens.sis.edu.vn"
   
def clear_form():
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    sheet.range("B2:B8").options(transpose = True).clear()