import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import itertools
import pandas as pd
import urllib.request


sender_address = 'jyotiubaheti@gmail.com'
sender_pass = 'jayyoubee'
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password

def sendEmail(receiver_address,attach_jpg=None,body=''):
    #The mail addresses and password

    # receiver_address = 'nikhilubaheti@gmail.com'
    #Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Geeta Pariwar Summer Camp 2021: Certificate'
    message.attach(MIMEText(body,'html'))
    #The subject line
    #The body and the attachments for the mail
    if attach_jpg is not None:
        ImgFileName = attach_jpg
        img_data = open(ImgFileName, 'rb').read()
        message.attach( MIMEImage(img_data, name=os.path.basename(ImgFileName)))

    #Create SMTP session for sending the mail
    print ("Mail To be sent",receiver_address)
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    print('Mail Sent',receiver_address,attach_jpg)


dfAttendance = pd.read_excel('zoomAttendance.xlsx')
attendents = dfAttendance['Roll Number'].to_list()
dfCertificates = pd.read_excel('Certificate Links.xlsx')
rollDict = dfCertificates.set_index('Roll No.').T.to_dict('list')
# print (rollDict)
# urllib.request.urlretrieve("https://docs.google.com/presentation/d/1I2DE8x-69vN1vH1fVIhM4AqBLRsa68ZUIkvF5ydj8So/export/png?id=1I2DE8x-69vN1vH1fVIhM4AqBLRsa68ZUIkvF5ydj8So&pageid=SLIDES_API444664923_18", "local-filename.jpg")
for i,roll in enumerate(attendents):
    if int(roll)<204: continue
    strRoll = "GPSC2021_" + str(roll).zfill(3)
    link = rollDict[strRoll][1]
    photoName = "certificates/" + strRoll + ".jpg"
    urllib.request.urlretrieve(link, photoName)    
    message = 'Greetings From Geeta Pariwar!! The previous certificates have some errors so kindly delete them. Attaching the corrected certificate.'
    sendEmail(rollDict[strRoll][2],body=message,attach_jpg=photoName)
    print ("Done:",photoName)
session.quit()
    