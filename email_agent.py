# -*- coding: utf-8 -*-
"""
Created on Sun Feb 11 20:41:01 2024

@author: user
"""

# Imports
# 1. GUI
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk, scrolledtext, Menu

# 2. Server communication
import smtplib # libraria pentru trimitere e-mail
from email.mime.multipart import MIMEMultipart # libraria pentru formatarea mesajului
from email.mime.text import MIMEText # pentru textul scrisorii
from email.mime.application import MIMEApplication # pentru attachment

# 3. Other libs
import openpyxl, re, os, time, winsound


""" ************************** Main Functions ******************************"""


""" ************************** Back END  ***********************************"""

""" ************************** Server communication ************************"""

def connect_server(from_addr, password):
    """ Functia care stabileste conexiunea cu serverul"""
    global server
    # Vom verifica ce fel de posta utilizeaza user-ul
    tail = from_addr.split("@")[-1] # coada adresei (yahoo.com, gmail.com, mt.utm.md, mail.ru)
    flag = None
    if tail == "yahoo.com":
        server_adress = 'smtp.mail.yahoo.com'; server_port = 465
        flag = 1
        
    elif tail == "gmail.com":
        server_adress = 'smtp.gmail.com'; server_port = 465
        flag = 1
    
    elif tail == "mail.ru":
        server_adress = 'smtp.mail.ru'; server_port = 465
        flag = 1

    else:
        info_msg = "Your email is not available now, please contact the developper"
        popup_msg(info_msg)
        flag = 0
        
    server = None
    if flag == 1: # Pentru alte servere de mail
        try:
            server = smtplib.SMTP_SSL(server_adress, server_port)
        except:
            popup_msg("Failed to connect server. Check your login or internet connection")
        try:
            server.ehlo() # Trimitem un fel de hello serverului, sa vedem daca raspunde
        except:
            popup_msg("Server not responding ")
            quit
        try:
            server.starttls()
        except:
            popup_msg("Server not responding ")
            quit
    # Daca am reusit conexiunea cu serverul, atunci mergem mai departe
    if server:
        try:
            server.login(from_addr, password) # Ne logam pe account-ul nostru
            b2.configure(bg='light blue',font=fnt, text = "Connected...")
            b4['state'] = 'normal'; b4['bg'] = 'yellow'
            b5['state'] = 'normal'; b5['bg'] = 'green'
        except:
            popup_msg("Failed to login. Check your login and password")


""" ************************** Email Extractor **********************************"""

def open_excel(path_to_excel):
    try:
        print("Reading excel file ...")
        t1 = time.time()
        excel_obj = openpyxl.load_workbook(path_to_excel, read_only=True, data_only=True)
        t2 = time.time()
        print("Done in {t:.1f} sec.\n".format(t = t2-t1))
    except:
        popup_msg("Unable to read excel file. \ Please close the program and try again")
    # Then, return the excel_object
    return excel_obj

def read_excel(excel_obj):
    # Extracting sheets from excel object
    sheets = excel_obj.worksheets
    # Extracting sheet names (only to be printed in console)
    sheet_names = excel_obj.sheetnames
    # Prepare an empty list to be filled with content extracted from sheets
    emails = []
    # Initiate a counter, to count sheets extracted (only to print in console)
    sheet_counter = 0
    # Initiate timer, to print total extracting time in console
    # Declare regex pattern for emails
    regex = re.compile(r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])") 
    t1 = time.time()
    print("Start extracting cell values from every sheet")
    for sheet in sheets:
        t11 = time.time()
        # Get maximum number of rows with data in every sheet
        rows = 0
        for max_row, row in enumerate(sheet,1):
            if not all(col.value is None for col in row):
                rows +=1
        # Extract all values from indicated columns
        print(f"Extracting cell values from sheet {sheet_names[sheet_counter]} ({rows} rows)")
        for row in sheet.iter_rows(min_row=1, min_col = 1, max_row = rows, max_col = 11):
            for cell in row:
                cell_val = str(cell.value).split("\n")
                for i in cell_val:
                    if re.fullmatch(regex, i.strip()):
                       emails.append(i.strip())
        t22 = time.time()
        print("Done in {t:.1f} sec.".format(t = t22-t11))
        sheet_counter+=1
    #garbage collector
    excel_obj.close()
    # Removing duplicates from list
    # ! Set is a collection of unique items
    emails = set(emails)
    # Generate a sound when finishes
    freq = 1250
    dur = 100
    for i in range(0, 5):    
        winsound.Beep(freq, dur)    
        freq+= 100   
    t2 = time.time()
    print("\n All cells from columns A to K, from all sheets were read in {t:.1f} sec.".format(t = t2-t1))
    print(f"\n Exporting {len(emails)} emails... \n")
    # Call function export emails
    t1 = time.time()
    export_emails(emails)
    t2 = time.time()
    print("\n Done, emails are updated! in {t:.1f} sec.".format(t = t2-t1))
    
    
def export_emails(path_to_excel, emails):
    # check if there exists an excel book for emails:
    # A little modification of path,in order to point to new exmails.xlsx document
    dir_path = os.path.dirname(path_to_excel)
    export_path = os.path.abspath(f"{dir_path}/emails.xlsx")
    if not os.path.isfile(export_path):
        # create a new workbook
        wb = openpyxl.Workbook()
        sheet = wb.active
        for i in range(len(emails)):
            #Cell object is created by  using sheet object's cell() method.
            c = sheet.cell(row = i+1, column = 1) 
            # writing values to cells 
            c.value = list(emails)[i]
        wb.save(export_path)
        # garbage collector
        wb.close()
    else:
        # if the file already exists, read emails and update list
        # Read excel object
        excel_obj = openpyxl.load_workbook(export_path,read_only=True)
        sheet = excel_obj.worksheets[0]
        # Get max num of rows
        rows = 0
        for max_row, row in enumerate(sheet,1):
            if not all(col.value is None for col in row):
                rows +=1
        # Extract all values from indicated columns
        old_emails = []
        for row in sheet.iter_rows(min_row=1, min_col = 1, max_row = rows, max_col = 1):
            for cell in row:
                cell_val = cell.value 
                if cell_val is not None: # Remove none from list, if exists
                    old_emails.append(cell_val)
        # Now, we will join emails list with old_emails and will remove duplicates
        new_emails = set(old_emails + list(emails))
        # And we will rewrite all into the same excel file
        wb = openpyxl.Workbook()
        sheet = wb.active
        for i in range(len(new_emails)):
            #Cell object is created by  using sheet object's cell() method.
            c = sheet.cell(row = i+1, column = 1) 
            # writing values to cells 
            c.value = list(new_emails)[i]
        wb.save(export_path)
        # Garbage collector
        excel_obj.close()
        wb.close()


""" ************************** Email Sender **********************************"""

def send_mail(from_addr, password, to_addr, msg):
    """Functia care expediaza e-mail de pe gmail"""
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465) # Definim datele serverului
    server.ehlo() # Trimitem un fel de hello serverului, sa vedem daca raspunde
    server.login(from_addr, password) # Ne logam pe account-ul nostru
    server.sendmail(from_addr, to_addr, msg.as_string()) # Trimitem e-mail de pe account-ul nostru
    server.quit() # Spunem la revedere serverului

def create_message_content(from_addr, to_addr, subject, content):
    """ Functia care formeaza continutul e-mail-ului"""
    # Mesajul inclus in email:
    msg = MIMEMultipart() # cream obiectul msg
    # 1. Antetetul scrisorii
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    # 2. Continutul de text al scrisorii
    body = MIMEText(content, 'html') # creaza obiectul html - body
    msg.attach(body)
    # 3. Attachment la scrisoare
    if attachment_path:
        file_to_attach = MIMEApplication(open(attachment_path, "rb").read())
        # 3.3. Indicam concret tipul de attachment:
        # ! Vom extrage numele fisierului pentru a fi inclus in filename
        file_to_attach.add_header('Content-Disposition', 'attachment',\
                                    filename=os.path.basename(attachment_path))
        #    unde: add_header(_name, _value, **_params) - 'Content-Disposition' indica
        #          codului htmp cum va fi afisat contentul. In cazul dat ca attachment
        # 3.4. Atasam si aceasta la mesaj:
        msg.attach(file_to_attach)
    return msg
        
def main_sender(from_addr, password, subject, content, recipient_emails):
    for item in recipient_emails:
        try:
            to_addr = item
            send_mail(from_addr, password, to_addr, \
                      create_message_content(from_addr, to_addr, subject, content))
            log_mess = 'Email successfully sent to -- ' + to_addr
            print(log_mess)
        except:
            log_mess = 'Email not sent to -- '+ to_addr 
            print(log_mess)


""" ************************* GUI ************************ """

""" GUI-linked functions"""

def open_emails():
    global recipient_emails
    # Get excel path
    excel_path = os.path.abspath(askopenfilename(title="Select a File",\
                                 filetype=(("Excel", "*.xlsx"), ("Excel", "*.xls"))))
    # Read excel with emails
    excel_obj = openpyxl.load_workbook(excel_path,read_only=True)
    sheet = excel_obj.worksheets[0]
    # Prepare an empty list to be filled with content extracted from sheets
    recipient_emails = []
    # Get max num of rows
    rows = 0
    for max_row, row in enumerate(sheet,1):
        if not all(col.value is None for col in row):
            rows +=1
    # Declare regex pattern for emails
    regex = re.compile(r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])")
    for row in sheet.iter_rows(min_row=1, min_col = 1, max_row = rows, max_col = 1):
        for cell in row:
            cell_val=str(cell.value)
            if re.fullmatch(regex, cell_val):
               recipient_emails.append(cell_val)        
    b1.configure(background='blue')
        

def set_expeditor():
    global from_addr
    from_addr = e2.get()
    b2.configure(background='blue')
    
def set_password():
    global password
    password = e3.get()
    b3.configure(background='blue')


def set_subject():
    global subject
    subject = e4.get()
    b4.configure(background='blue')

def set_content():
    global content
    content = e5.get("1.0","end-1c")
    b5.configure(background='blue')
    

def set_attachment():
    global attachment_path
    attachment_path = os.path.abspath(askopenfilename(title="Select a File",\
                                 filetype=(('JPG Files', '*.jpg'),
                                           ('JPEG Files', '*.jpeg'),
                                           ('PNG Files', '*.png'),
                                           ('All files','*.*'))))
    b6.configure(background='blue')

def send():
    # Reset background color for all buttons
    for i in [b1,b2,b3,b4,b5,b6]:
        i.configure(background = 'light gray')
    try:
        main_sender(from_addr, password, subject, content, recipient_emails)
    except:
        popup_msg('Eroare, verificati datele introduse,verificati ENTER')
        
def send_test_email():
    """Functia pentru a expedia mesaj pe propriul account"""
    try:
        main_sender(from_addr, password, subject, content, [from_addr])
        popup_msg('Pe adresa Dvs. a fost expediat un email de testare')        
    except:
        popup_msg('Eroare, verificati datele introduse,verificati ENTER')
        
def popup_msg(message):
    popup = tk.Tk()
    popup.wm_title("Info")
    label = tk.Label(popup, text=message, fg = 'blue' )
    label.pack(side="top", fill="x", pady=20)
    label.config(font=("Times New Roman", 14))
    B1 = tk.Button(popup, text="Ok", padx = 30, pady = 5, borderwidth = 5, \
                   command = popup.destroy)
    B1.pack()
    


""" GUI graphical elements """

root = tk.Tk()
root.title("Email Agent")
root.geometry("850x600")

# Tabs
tab1 = ttk.Frame
tabControl = ttk.Notebook(root) 
  
tab1 = ttk.Frame(tabControl) 
tab2 = ttk.Frame(tabControl) 
  
tabControl.add(tab1, text ='Email Sender') 
tabControl.add(tab2, text ='Email Extractor') 
tabControl.grid(row = 0, column = 0)

# Font specifications
fnt = ("Arial", 12, "bold")

# Menu
menubar = Menu(root) # generam un obiect menubar in root
filemenu = Menu(menubar, tearoff=0) # Definim obiectul pentru File
filemenu.add_command(label="New", command=lambda: popup_msg('under process...')) # Adaugam comanda New
filemenu.add_separator()
filemenu.add_command(label = 'Exit', command = None)
menubar.add_cascade(label="File", menu=filemenu) #In bara de meniuri includem obiectul filemenu

editmenu = Menu(menubar, tearoff=0)
editmenu.add_command(label="Delete", command=lambda: popup_msg('under process...'))
menubar.add_cascade(label="Edit", menu=editmenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Update", command=None)
helpmenu.add_command(label="About...", command=lambda: popup_msg('A simple application for extracting and sending emails'))
menubar.add_cascade(label="Help", menu=helpmenu)

root.config(menu = menubar) # Configuram root-ul ca sa stie ca in menu avem obiectul menubar


# GUI elements definition, Email Sender
l1 = tk.Label(tab1, text = "Excel file with email adresses   ",font = fnt)
b1 = tk.Button(tab1, text = "browse", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = open_emails)

    
l2 = tk.Label(tab1, text = "Sender email", font=fnt)
e2 = tk.Entry(tab1,  width = 45, borderwidth = 5, font=fnt)
b2 = tk.Button(tab1, text = "Enter", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = set_expeditor)
# Temporar, inseram un email de testare
e2.insert(0,'')
    
l3 = tk.Label(tab1, text = "Sender password", font=fnt)
e3 = tk.Entry(tab1,  width = 45, borderwidth = 5, font=fnt, show = "*")
b3 = tk.Button(tab1, text = "Enter", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = set_password)
# Temporar, inseram parola
e3.insert(0,'')
    
l4 = tk.Label(tab1, text = "Subject", font=fnt)
e4 = tk.Entry(tab1, width = 45, borderwidth = 5,font = fnt)
e4.insert(tk.END,'Please, consider this message')
b4 = tk.Button(tab1, text = "Enter", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = set_subject)

l5 = tk.Label(tab1, text = "Message content ", font=fnt)
e5 = tk.Text(tab1, height = 10, width = 45,  borderwidth = 5, font=fnt)
b5 = tk.Button(tab1, text = "Enter", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = set_content)
# Temporar, inseram ceva content
e5.insert('1.0','Content pentru testare')


l6 = tk.Label(tab1, text = "Want to attach some images? ", font=fnt)
b6 = tk.Button(tab1, text = "Browse", padx = 30, pady = 5,borderwidth = 5, font=fnt,\
               command = set_attachment)
    
l7 = tk.Label(tab1, text = " "*45, font=fnt)

b7 = tk.Button(tab1, text = "SEND",fg = 'black',font=fnt,\
               padx = 50, pady = 20, borderwidth = 5,\
               bg = 'green',command = send)
    
b8 = tk.Button(tab1, text = "EXIT", fg = 'black', font=fnt,\
              padx = 50, pady = 10, borderwidth = 5,\
              bg = 'red',command =  root.destroy) 

b9 = tk.Button(tab1, text = "Send test email", fg = 'black', font=fnt,\
              padx = 30, pady = 10, borderwidth = 5,\
              bg = 'yellow',command = send_test_email) 
    
    
# GUI elements definition, Email Extractor

l1_1 = tk.Label(tab2, text = "Select an excel file to scan and extract emails",font = fnt)

l2_1 = tk.Label(tab2, text = "Excel file: ",font = fnt)
b1_1 = tk.Button(tab2, text = "browse", padx = 30, pady = 2,borderwidth = 5, font=fnt,\
               command = open_excel)

"""************* Window layout **************"""

# Email Sender
l1.grid(row = 0, column = 1)
b1.grid(row = 0, column = 2)

l2.grid(row = 1, column = 0)
e2.grid(row = 1, column = 1)
b2.grid(row = 1, column = 2)

l3.grid(row = 2, column = 0)
e3.grid(row = 2, column = 1)
b3.grid(row = 2, column = 2)

l4.grid(row = 3, column = 0)
e4.grid(row = 3, column = 1)
b4.grid(row = 3, column = 2)

l5.grid(row = 4, column = 0)
e5.grid(row = 4, column = 1)
b5.grid(row = 4, column = 2)

l6.grid(row = 5, column = 1)
b6.grid(row = 5, column = 2)

l7.grid(row = 6, column = 0, columnspan = 5)

b7.grid(row = 7, column = 2)


b8.grid(row = 7, column = 1)

b9.grid(row = 7, column = 0)

# Email Extractor
l1_1.grid(sticky="W", row = 0, column = 0, columnspan = 4)
l2_1.grid(sticky="E", row = 1, column = 0, rowspan = 2)
b1_1.grid(sticky="E",row = 1, column = 2, rowspan = 2)    

root.mainloop()
    
    
