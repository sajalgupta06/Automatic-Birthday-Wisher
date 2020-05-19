import pandas as pd
import datetime
import smtplib

GMAIL_ID=' '
GMAIL_PASS=' '

def sendemail(to,sub,msg):
    
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASS)
    s.sendmail(GMAIL_ID,to,f"subject:{sub}\n\n{msg}")
    print("message Sent")
    s.quit()

df=pd.read_excel("data.xlsx")
today=datetime.datetime.now().strftime("%d-%m")
yearnow=datetime.datetime.now().strftime("%Y")
update=[]
msg="bhot bhot mubaarak janamdin ki"
for index,item in df.iterrows():
     bday=item['Birthday'].strftime("%d-%m")
     if (today==bday) and yearnow not in str(item["Year"]):
         sendemail(item["Email"],item['Dialogue'],msg)
         update.append(index)

for i in update:
    yr=df.loc[i,'Year']
    df.loc[i,'Year']=str(yr) +',' + str(yearnow)
    df.to_excel('data.xlsx', index=False)

    
