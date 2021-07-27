import xlwings as xw
import smtplib
from email.message import EmailMessage
from datetime import date
from datetime import datetime
import os
import sqlite3
office_user = os.environ.get('OFFICE_USER')
office_pass = os.environ.get('OFFICE_PASS')
old_value_path = 'C:\\Python\\python_work\\Personal Projects\\parent_emailer\\parent_emailer_V2\\old_value'

   
def send_email():
    wb = xw.Book.caller()
    sheet = wb.sheets[2]
    s_number = sheet.range("B1").value
    s_name = sheet.range("B2").value
    parent_email = (sheet.range("B3").value)
    the_subject = (sheet.range("B6").value)    
    the_content = sheet.range("B8").value
    the_cc = sheet.range("B4").value    
    date_now = datetime.now()
    
    msg = EmailMessage()
    msg.set_content(the_content)    
    msg['Subject'] = the_subject
    msg['From'] = office_user
    msg['To'] = parent_email
    msg['Cc'] = the_cc
    
    smtpsrv = 'smtp.office365.com'
    smtpserver = smtplib.SMTP(smtpsrv, 587)
    smtpserver.ehlo()
    smtpserver.starttls()
    smtpserver.login(office_user, office_pass) 
    smtpserver.send_message(msg)
    smtpserver.close()
    
    file = open(old_value_path, 'r')
    old_content = file.read()
    
    content_no_footer = the_content.replace(old_content, "")
    
    
    conn = sqlite3.connect('C:\\Python\\python_work\\Personal Projects\\parent_emailer\\parent_emailer_V2\\parent_email_log.db')
    c = conn.cursor()
    c.execute("INSERT INTO 'messages' VALUES (?, ?, ?, ?, ?, ?, ?)", (s_number, s_name, parent_email, the_subject, content_no_footer, the_cc, date_now))
    conn.commit()
    conn.close()
    
    
    sheet.range("B1:B8").options(transpose = True).clear()
    
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
    active_cell_column = active_cell.column
    active_cell_column -= 1
    student_number_value = sheet.range((active_cell_row, active_cell_column)).value
    sheet = wb.sheets[2]
    sheet.range("B2").value = active_cell.value
    sheet.range("B1").value = student_number_value
    
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
    sheet.range("B1:B8").options(transpose = True).clear()