import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import itertools
import pandas as pd

def sendEmail(receiver_address,attach_jpg=None,body=''):
    #The mail addresses and password
    sender_address = 'jyotiubaheti@gmail.com'
    sender_pass = 'jayyoubee'
    # receiver_address = 'nikhilubaheti@gmail.com'
    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Geeta Pariwar Summer Camp 2021: Roll Number'
    message.attach(MIMEText(body,'html'))
    #The subject line
    #The body and the attachments for the mail
    if attach_jpg is not None:
        ImgFileName = attach_jpg
        img_data = open(ImgFileName, 'rb').read()
        message.attach( MIMEImage(img_data, name=os.path.basename(ImgFileName)))

    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent',receiver_address,attach_jpg)

df = pd.read_csv("Summer Camp 2021 - Phone.csv")
Name = df['Name'].tolist()
for i,name in enumerate(Name):
    if i==0: continue
    rollNo = name.split("_")[-1]
    message = 'Greetings From Geeta Pariwar!! Your assigned *Roll number* for the summer camp scheduled this 3rd May 2021 is *' + str(int(rollNo)) + '* for Participant Name *' + df['Full Name'].loc[i].strip() + '*. Kindly set you zoom name with your full name and this roll number. Also, kindly follow this throughout the camp for every session.'
    sendEmail(df['E-Mail'].iloc[i],body=message)
    print ("Done:",name)