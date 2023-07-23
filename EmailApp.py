#!/usr/bin/python3
# Development platform: Windows
# Python version: 3.10.5

from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import messagebox
from tkinter import filedialog
import re
import os
import pathlib
import sys
import base64
import socket

#
# Global variables
#

# Replace this variable with your CS email address
YOUREMAIL = ""
# Replace this variable with your student number
MARKER = ''

# The Email SMTP Server
SERVER = "testmail.cs.hku.hk"   #SMTP Email Server
SPORT = 25                      #SMTP listening port

# For storing the attachment file information
fileobj = None                  #For pointing to the opened file
filename = ''                   #For keeping the filename


#
# For the SMTP communication
#
def do_Send():
  
  #TASK 1 =================================================
  send_to = get_TO() 
  subject = get_Subject() 
  mssg = get_Msg() 
  send_cc = get_CC()
  send_bcc = get_BCC() 

  #checking necessary fields 
  if not send_to:
    return alertbox('Must enter the recipient\'s email')
  
  elif not subject: 
    return alertbox('Email must contain a subject')
  
  elif not mssg: 
    return alertbox('Email must contain a message')

  #all emails are separated by a comma ','
  email_to = send_to.split(',') 
  email_cc = send_cc.split(',') 
  email_bcc = send_bcc.split(',')

  #check for invalid email and email exists
  for i in email_to: 
    if not echeck(i) and i: 
      return alertbox('Invalid Receiver, Email: {}'.format(i))
    
  for i in email_cc: 
    if not echeck(i) and i: 
      return alertbox('Invalid CC, Email: {}'.format(i))
  
  for i in email_bcc: 
    if not echeck(i) and i: 
      return alertbox('Invalid BCC, Email: {}'.format(i))
  
  #TASK 2 =================================================
  header_from = 'From: {} \r\n'.format(YOUREMAIL) 
  header_to = 'To: {} \r\n'.format(send_to) 
  header_subject = 'Subject: {} \r\n'.format(subject) 
  header = header_from + header_to + header_subject
  if send_cc: #if there is a cc, add it to the header
    header_cc = 'Cc: {} \r\n'.format(send_cc) 
    header += header_cc 
  
  body = mssg 
  only_text_data =  header + '\n' + body + '.\r\n' #used if no attachments are there 
  isAttachment = False #default value 

  #check for attachments 
  if fileobj: 
    isAttachment = True 
    mime_mssg = base64.encodebytes(fileobj.read()) #file encoded to base 64 
    header_mime = 'MIME-Version: 1.0\r\n'
    header_content_type = 'Content-Type: multipart/mixed; boundary={}\r\n'.format(MARKER) 
    header += header_mime + header_content_type

    marker_start = '--{}\r\n'.format(MARKER) 
    marker_end = '--{}--\r\n'.format(MARKER) 

    #Part 1: Text content 
    text_header_content_type = 'Content-Type: text/plain\r\n' 
    text_header_transfer_encoding = 'Content-Transfer-Encoding: 7bit\r\n'
    text_header = text_header_content_type + text_header_transfer_encoding
    text = marker_start + text_header + '\n' + mssg + '\n'

    #Part 2: Attachment 
    att_header_content_type = 'Content-Type: application/octet-stream\r\n'
    att_header_transfer_encoding = 'Content-Transfer-Encoding: base64\r\n' 
    att_header_content_disposition = 'Content-Disposition: attachment; filename={}\r\n'.format(filename) 
    att_header = att_header_content_type + att_header_transfer_encoding + att_header_content_disposition
    att_start = marker_start + att_header + '\n'
    
    #if attachments are present, we send the data in the following order
    # body_start -> mime_msg (the encoded attachment) -> body_end
    body_start = header + '\n' + text + att_start
    body_end = '\n\n' + marker_end + '.\r\n'


  #TASK 3 =================================================
  #Initiate connection with the testmail server
  try: 
    tcp_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM) 
    tcp_socket.settimeout(10) # 10 second Timeout 
    tcp_socket.connect((SERVER, SPORT)) 
    tcp_address = tcp_socket.getsockname()
  except socket.error as emsg: 
    tcp_socket.close()
    alertbox("Socket error: {}".format(emsg))
    return 

  #220 connection established 
  try: 
    rcv_msg = tcp_socket.recv(1024).decode()
    print(rcv_msg) 
    if rcv_msg[:3] != '220': 
        tcp_socket.close()
        return alertbox('Wrong Code Received: ', rcv_msg)
  except socket.error as emsg: 
    tcp_socket.close()
    alertbox('Socket error: {}'.format(emsg))
    return

  #EHLO statement 
  try: 
    hello_msg = 'EHLO {x} \r\n'.format(x=tcp_address[0]) 
    print(hello_msg)
    tcp_socket.sendall(hello_msg.encode())
    rcv_msg = tcp_socket.recv(1024).decode() 
    if rcv_msg[:3] != '250': 
      tcp_socket.close()
      return alertbox('Wrong Code Received: ', rcv_msg)
    print(rcv_msg)
  except socket.error as emsg:
    tcp_socket.close()
    alertbox("Socket error: {}".format(emsg))           
    return

  #MAIL FROM 
  try: 
    from_msg = 'MAIL FROM: <{x}> \r\n'.format(x=YOUREMAIL)  
    print(from_msg)
    tcp_socket.sendall(from_msg.encode())
    rcv_msg = tcp_socket.recv(1024).decode() 
    if rcv_msg[:3] != '250': 
      tcp_socket.close()
      return alertbox('Wrong Code Received: {}'.format(rcv_msg))
    print(rcv_msg)
  except socket.error as emsg:
    alertbox("Socket error: {}".format(emsg))
    tcp_socket.close()           
    return

  #RCPT TO 
  try: 
    for receiver in email_to + email_cc + email_bcc: 
      if receiver: 
        from_msg = 'RCPT TO: <{x}> \r\n'.format(x=receiver)  
        print(from_msg)
        tcp_socket.sendall(from_msg.encode())
        rcv_msg = tcp_socket.recv(1024).decode() 
        if rcv_msg[:3] != '250': 
          tcp_socket.close()
          return alertbox('Wrong Code Received: {}'.format(rcv_msg))
        print(rcv_msg)
  except socket.error as emsg:
    alertbox("Socket error: {}".format(emsg))
    tcp_socket.close()           
    return

  #DATA
  try: 
      data_msg = 'DATA\r\n'
      print(data_msg)
      tcp_socket.sendall(data_msg.encode())
      rcv_msg = tcp_socket.recv(1024).decode() 
      if rcv_msg[:3] != '354': 
        tcp_socket.close()
        return alertbox('Wrong Code Received: {}'.format(rcv_msg))
      print(rcv_msg)
  except socket.error as emsg:
    alertbox("Socket error: {}".format(emsg))
    tcp_socket.close()           
    return 

  #Sending headers and body of email
  try: 
    if isAttachment: 
      tcp_socket.sendall(body_start.encode())
      tcp_socket.sendall(mime_mssg)
      tcp_socket.sendall(body_end.encode())
    else: 
      tcp_socket.sendall(only_text_data.encode())
    rcv_msg = tcp_socket.recv(1024).decode() 
    if rcv_msg[:3] != '250': 
      tcp_socket.close()
      return alertbox('Wrong Code Received: {}'.format(rcv_msg))
    print(rcv_msg)
  except socket.error as emsg:
    alertbox("Socket error: {}".format(emsg))
    tcp_socket.close()           
    return
  
  #QUIT
  try: 
      quit_msg = 'QUIT\r\n'
      print(quit_msg)
      tcp_socket.sendall(quit_msg.encode())
      rcv_msg = tcp_socket.recv(1024).decode() 
      if rcv_msg[:3] != '221': 
        tcp_socket.close()
        return alertbox('Wrong Code Received: {}'.format(rcv_msg))
      print(rcv_msg)
  except socket.error as emsg:
    alertbox("Socket error: {}".format(emsg))
    tcp_socket.close()           
    return

  #END of Send. We close the socket, and the program as well. 
  tcp_socket.close() 
  alertbox('Successful')
  sys.exit(1)

#
# Utility functions
#

#This set of functions is for getting the user's inputs
def get_TO():
  return tofield.get()

def get_CC():
  return ccfield.get()

def get_BCC():
  return bccfield.get()

def get_Subject():
  return subjfield.get()

def get_Msg():
  return SendMsg.get(1.0, END)

#This function checks whether the input is a valid email
def echeck(email):   
  regex = '^([A-Za-z0-9]+[.\-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+'  
  if(re.fullmatch(regex,email)):   
    return True  
  else:   
    return False

#This function displays an alert box with the provided message
def alertbox(msg):
  messagebox.showwarning(message=msg, icon='warning', title='Alert', parent=win)

#This function calls the file dialog for selecting the attachment file.
#If successful, it stores the opened file object to the global
#variable fileobj and the filename (without the path) to the global
#variable filename. It displays the filename below the Attach button.
def do_Select():
  global fileobj, filename
  if fileobj:
    fileobj.close()
  fileobj = None
  filename = ''
  filepath = filedialog.askopenfilename(parent=win)
  if (not filepath):
    return
  print(filepath)
  if sys.platform.startswith('win32'):
    filename = pathlib.PureWindowsPath(filepath).name
  else:
    filename = pathlib.PurePosixPath(filepath).name
  try:
    fileobj = open(filepath,'rb')
  except OSError as emsg:
    print('Error in open the file: %s' % str(emsg))
    fileobj = None
    filename = ''
  if (filename):
    showfile.set(filename)
  else:
    alertbox('Cannot open the selected file')

#################################################################################
#################################################################################

#
# Set up of Basic UI
#
win = Tk()
win.title("EmailApp")

#Special font settings
boldfont = font.Font(weight="bold")

#Frame for displaying connection parameters
frame1 = ttk.Frame(win, borderwidth=1)
frame1.grid(column=0,row=0,sticky="w")
ttk.Label(frame1, text="SERVER", padding="5" ).grid(column=0, row=0)
ttk.Label(frame1, text=SERVER, foreground="green", padding="5", font=boldfont).grid(column=1,row=0)
ttk.Label(frame1, text="PORT", padding="5" ).grid(column=2, row=0)
ttk.Label(frame1, text=str(SPORT), foreground="green", padding="5", font=boldfont).grid(column=3,row=0)

#Frame for From:, To:, CC:, Bcc:, Subject: fields
frame2 = ttk.Frame(win, borderwidth=0)
frame2.grid(column=0,row=2,padx=8,sticky="ew")
frame2.grid_columnconfigure(1,weight=1)
#From 
ttk.Label(frame2, text="From: ", padding='1', font=boldfont).grid(column=0,row=0,padx=5,pady=3,sticky="w")
fromfield = StringVar(value=YOUREMAIL)
ttk.Entry(frame2, textvariable=fromfield, state=DISABLED).grid(column=1,row=0,sticky="ew")
#To
ttk.Label(frame2, text="To: ", padding='1', font=boldfont).grid(column=0,row=1,padx=5,pady=3,sticky="w")
tofield = StringVar()
ttk.Entry(frame2, textvariable=tofield).grid(column=1,row=1,sticky="ew")
#Cc
ttk.Label(frame2, text="Cc: ", padding='1', font=boldfont).grid(column=0,row=2,padx=5,pady=3,sticky="w")
ccfield = StringVar()
ttk.Entry(frame2, textvariable=ccfield).grid(column=1,row=2,sticky="ew")
#Bcc
ttk.Label(frame2, text="Bcc: ", padding='1', font=boldfont).grid(column=0,row=3,padx=5,pady=3,sticky="w")
bccfield = StringVar()
ttk.Entry(frame2, textvariable=bccfield).grid(column=1,row=3,sticky="ew")
#Subject
ttk.Label(frame2, text="Subject: ", padding='1', font=boldfont).grid(column=0,row=4,padx=5,pady=3,sticky="w")
subjfield = StringVar()
ttk.Entry(frame2, textvariable=subjfield).grid(column=1,row=4,sticky="ew")

#frame for user to enter the outgoing message
frame3 = ttk.Frame(win, borderwidth=0)
frame3.grid(column=0,row=4,sticky="ew")
frame3.grid_columnconfigure(0,weight=1)
scrollbar = ttk.Scrollbar(frame3)
scrollbar.grid(column=1,row=1,sticky="ns")
ttk.Label(frame3, text="Message:", padding='1', font=boldfont).grid(column=0, row=0,padx=5,pady=3,sticky="w")
SendMsg = Text(frame3, height='10', padx=5, pady=5)
SendMsg.grid(column=0,row=1,padx=5,sticky="ew")
SendMsg.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=SendMsg.yview)

#frame for the button
frame4 = ttk.Frame(win,borderwidth=0)
frame4.grid(column=0,row=6,sticky="ew")
frame4.grid_columnconfigure(1,weight=1)
Sbutt = Button(frame4, width=5,relief=RAISED,text="SEND",command=do_Send).grid(column=0,row=0,pady=8,padx=5,sticky="w")
Atbutt = Button(frame4, width=5,relief=RAISED,text="Attach",command=do_Select).grid(column=1,row=0,pady=8,padx=10,sticky="e")
showfile = StringVar()
ttk.Label(frame4, textvariable=showfile).grid(column=1, row=1,padx=10,pady=3,sticky="e")

win.mainloop()
