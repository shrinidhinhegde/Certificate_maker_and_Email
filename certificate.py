import os
import re
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from PIL import Image, ImageDraw, ImageFont
import pandas as pd


def email(d):
    for key, value in d.items():
        img_data = open('C:\\Users\\Shrinidhi N Hegde\\PycharmProjects\\untitled1\\images\\' + key + '.PNG',
                        'rb').read()  # change path
        smtpServer = 'smtp.gmail.com'
        subject = 'Certificate for attending the webinar on "Hopping into Entrepreneurship Wagon'  # change subject
        message = 'Hello ' + key + ',\n' + 'Greetings from IEEE BMSIT&M SB\nThank you for attending the webinar on " Hopping into Entrepreneurship Wagon".\nKindly find your certificate attached below.\nRegards.'  # change body
        sender = 'ieee_certificates@bmsit.in'  # change email ID
        password = '2017s3775.'  #change password
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = value
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))
        image = MIMEImage(img_data, name=os.path.basename(key + '.JPG'))
        msg.attach(image)
        server = smtplib.SMTP(smtpServer, 587)
        server.ehlo()
        server.starttls()
        server.login(sender, password)
        text = msg.as_string()
        server.sendmail(sender, value, text)
        server.quit()
        print('email sent to ' + value)


df = pd.ExcelFile('Book (1).xlsx').parse('Sheet1')  # Change xlsx file name
x = []
y = []
x.append(str(df['Names']))  # Column of names
y.append(str(df['Email']))  # column of emails
emails = re.findall("[A-Za-z].*[a-z]", y[0])
names = re.findall("[A-Za-z].*[A-Za-z]", x[0])
emails.pop()
names.pop()
dic = dict(zip(names, emails))
for i in names:
    image = Image.open(
        'image.png')  # change image

    draw = ImageDraw.Draw(image)

    font = ImageFont.truetype('Ubuntu-Regular.ttf', size=32)  # use the font of your choice

    (x, y) = (860, 670)  # set coordinates (use ms paint)

    message = i

    color = 'rgb(0,0,0)'

    draw.text((x, y), message, fill=color, font=font)

    save = message + '.PNG'

    dir_path = "C:\\Users\\Shrinidhi N Hegde\\PycharmProjects\\untitled1\\images"  # change path
    file_path = os.path.join(dir_path, save)

    image.save(file_path)
    print(save + ' image saved')

email(dic)  # call email function
print('Success!')
