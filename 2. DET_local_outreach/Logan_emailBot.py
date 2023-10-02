import pandas as pd
# import win32com.client as win32
import xlwings as xw

"""This file contains updates to work with DET automated emails reaching out to local potential sponsorship partners"""

path2xl_list=r"C:\Users\logan\Desktop\Capstone\0 Funding\Email Bot R2 - local sponsors\EmailBot\2. DET_local_outreach\EmailPeeps.xlsx"
Email_content=r"C:\Users\logan\Desktop\Capstone\0 Funding\Email Bot R2 - local sponsors\EmailBot\2. DET_local_outreach\EmailContent.txt" #this is the email textp
path2attachment=r"C:\Users\logan\Desktop\Capstone\0 Funding\Email Bot R2 - local sponsors\EmailBot\2. DET_local_outreach\DETechnologies_information_package.pdf"
pathto_Macro4Email=r"C:\Users\logan\Desktop\Capstone\0 Funding\Email Bot R2 - local sponsors\EmailBot\2. DET_local_outreach\Macro4Email.xlsm"
attachment=r"C:\Users\logan\Desktop\Capstone\0 Funding\Email Bot R2 - local sponsors\EmailBot\2. DET_local_outreach\DETechnologies_information_package.pdf"

with open(Email_content,"r") as file:
    mainBody=file.read()

df = pd.read_excel(path2xl_list)

num_rows=df.shape[0]
# print(num_rows)

Company_col=1-1 #remember counting starts at 0
names_col=2-1 #column 2 has names-- remember counting starts at 0
emails_col=3-1 #column 3 has names-- remember counting starts at 0

for row in range(0,num_rows):
    # get data for the email personalization ----------------------------
    
    #get name
    name=df.iloc[row,names_col]
    if pd.isna(name): 
        name = "all"
    print(name)
    #get email
    email=df.iloc[row,emails_col]
    print(email)
    #get company name
    company=df.iloc[row,Company_col]
    print(company)

    #now drop these info bits into the template! ----------------------------------
    n=mainBody.replace("<<insert name>>",name)
    finalEmail=n.replace("<<insert company>>",company)

    ##email things ------------------------
    wb=xw.Book(pathto_Macro4Email)
    macro=wb.macro('ThisWorkbook.SendEmail')

    subject="Request for support - DETechnologies - help us build a rocket engine!"
    body=finalEmail
    
    macro(email,subject,body,attachment)
    print("Sent!")