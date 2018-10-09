# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 14:45:25 2018
Takes a .csv file as exported from Zipgrade (selecting the standard option) 
and sends an email for each student in the file using Outlook for Windows.

@author: jfand
"""
import pandas as pd
import os
import shutil
import sys
import win32com.client as win32 

####### Change the path to the folder where the zipgrade csv file is.#########
folderpath = "C:/Users/jfand/Downloads/ZIPGRADE"
os.chdir(folderpath)

########## Enter the name of the ZipGrade csv file   #################
zipgrade_filename = "quiz-Q05-standard20180510.csv"

df = pd.read_csv(zipgrade_filename)
quiz_name = df.iloc[0,0]

# Checking for data integrity. Sometimes a grade is recorded by the mobile app 
# without assigning it to a particular student. That causes a problem since 
# there is not an email address associated.
if df['External Id'].isnull().sum(axis = 0)> 0:
    print("Warning! Data integrity issues. Check the source file and try again")
    sys.exit(0)

#Search for an existing temp folder, if found delete it, then create a new temp folder
temp_dir = folderpath + "\/temp"
if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
    shutil.rmtree(temp_dir)
    
os.makedirs(temp_dir)

# Here the text message could be customized.
msg_text = "Hi there, here is your feedback"

msg_subject = quiz_name + " Feedback"
msg_auto = True

# Sending emails
for name, group in df.groupby('External Id'): 
    temp_filename = temp_dir + "\/"+ quiz_name + "_"+ str(name).split('@')[0] + '.csv'
    group.to_csv(temp_filename, index=False)
    msg_recipients = str(name)
    msg_attach = temp_filename
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Recipients.Add(msg_recipients)
    mail.Subject = msg_subject
    mail.HtmlBody = msg_text
    mail.Attachments.Add(msg_attach)

    if msg_auto:
        mail.send
    else:
        mail.Display(True)

# Cleanning up     
if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
    shutil.rmtree(temp_dir)
