import pandas as pd
# import win32com.client as win32
import xlwings as xw

path2xl_list=r"C:\Users\logan\Desktop\Capstone\0 Funding\EmailBot\ITB_folks.xlsx"
txt_file=r"C:\Users\logan\Desktop\Capstone\0 Funding\EmailBot\EmailText.txt" #this is the email text
signature=r"C:\Users\logan\Desktop\Capstone\0 Funding\EmailBot\signature.txt"

with open(txt_file,"r") as file:
    email_body=file.read()
with open(signature,"r") as f:
    sign=f.read()
mainBody=email_body+sign


df = pd.read_excel(path2xl_list)

num_rows=df.shape[0]
# print(num_rows)

Company_col=1-1
projects_col=2-1
names_col=9-1 #column 9 has names and emails -- remember counting starts at 0

for row in range(0,num_rows):
    #get name and email
    name_email=df.iloc[row,names_col]
    t=name_email.split('\n')
    name=t[0]
    email=t[1]
    print(name,email)

    #get company name
    company=df.iloc[row,Company_col]
    print(company)

    #get project name
    project=df.iloc[row,projects_col]
    print(project)
    
    #now drop these info bits into the template!
    n=mainBody.replace("<<insert name>>",name)
    p=n.replace("<<insert project>>",project)
    finalEmail=p.replace("<<insert company>>",company)

    ##email things ------------------------
    wb=xw.Book(r"C:\Users\logan\Desktop\Capstone\0 Funding\EmailBot\Macro4Email.xlsm")
    macro=wb.macro('ThisWorkbook.SendEmail')

    subject="Collaboration Opportunity: Advancing Research on Rotating Detonation Engines (RDE) - DETechnologies R&D Group"
    body=finalEmail
    attachment=r"C:\Users\logan\Desktop\Capstone\0 Funding\EmailBot\DET_info_package.pdf"

    macro(email,subject,body,attachment)
    print("Sent!")

