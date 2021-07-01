"""
Have the participant data with email, name and events (, separated) in columns
of the spreadsheet resp. Put the certificate template and the spreadsheet in the
images directory.
"""

import smtplib
import openpyxl
import owncloud
import pyqrcode
from PIL import Image, ImageFont, ImageDraw
from subhogay import mail
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Logging in to CodeTech drive
oc = owncloud.Client('https://drive.teamcodetech.in/')
oc.login('', '')
oc.mkdir('bbb-v3-certificates')
links = []

wb = openpyxl.load_workbook(filename='images/BitByBit v3.0 _ Live Event Registration.xlsx')
sheet = wb['Form Responses 1']

names = []
email = []
events = []

EMAIL_ADDRESS = ''
PASSWORD = ''


def sendmail(receiver, html):
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Your BitByBit v3.0 participation certificate"
    msg['From'] = ''
    msg['To'] = receiver
    msg.attach(html)
    server.login(EMAIL_ADDRESS, PASSWORD)
    server.sendmail(EMAIL_ADDRESS, receiver, msg.as_string())
    server.quit()


# fetching data into empty list variables
for row in sheet.values:
    count = 0
    for value in row:
        count += 1
        if value:
            if count == 1:
                email.append(value)
            elif count == 2:
                names.append(value.title())
            elif count == 3:
                value = "".join(x for x in value if (x.isalnum() or x in "._-, "))
                events.append(value)

participants = len(names)

for participant in range(1, participants):
    template = Image.open('images/bbb_cert_template.png')
    name_font = ImageFont.truetype('fonts/GlacialIndifference-Bold.otf', 100)
    event_font = ImageFont.truetype('fonts/GlacialIndifference-Regular.otf', 40)
    draw = ImageDraw.Draw(template)

    # in-case of multiple registration form submissions put the certificate into same folder
    try:
        oc.mkdir(f'bbb-v3-certificates/{names[participant]}')
        email = False
    except:
        email = True
    link = oc.share_file_with_link(f'bbb-v3-certificates/{names[participant]}').get_link()
    links.append(link)
    qr = pyqrcode.create(link)
    qr.png(f'QRCodes/{names[participant]}.png', scale=10)
    qr = Image.open(f'QRCodes/{names[participant]}.png')
    qr = qr.resize((261, 261))
    print(link)
    mail_content = mail(link)
    html = MIMEText(mail_content, 'html')

    # checking if participant was a part of multiple events
    if events[participant].find(',') != -1:
        participant_events = events[participant].split(',')
        for event in participant_events:
            if event == 'Gaming':
                pass
            else:
                template = Image.open('images/bbb_cert_template.png')
                draw = ImageDraw.Draw(template)
                draw.text((273, 670), names[participant], (30, 54, 92), font=name_font)
                draw.text((629, 871), event, (30, 54, 92), font=event_font)
                template.paste(qr, (1643, 574))
                template.save(f'bbbv3-certs/{names[participant]}_{event}.png')
                filename = f'{names[participant]}_{event}.pdf'
                png = Image.open(f'bbbv3-certs/{names[participant]}_{event}.png')
                png.load()
                background = Image.new("RGB", png.size, (255, 255, 255))
                background.paste(png, mask=png.split()[3])
                background.save(f'bbbv3-certs/{names[participant]}_{event}.pdf')
                oc.put_file(f'bbb-v3-certificates/{names[participant]}/{filename}',
                            f'bbbv3-certs/{filename}')
        if email:
            sendmail(receiver=email[participant], html=html)

    else:
        if events[participant] == 'Gaming':
            pass
        else:
            draw.text((273, 670), names[participant], (30, 54, 92), font=name_font)
            draw.text((629, 871), events[participant], (30, 54, 92), font=event_font)
            template.paste(qr, (1643, 574))
            template.save(f'bbbv3-certs/{names[participant]}_{events[participant]}.png')
            filename = f'{names[participant]}_{events[participant]}.pdf'
            png = Image.open(f'bbbv3-certs/{names[participant]}_{events[participant]}.png')
            png.load()
            background = Image.new("RGB", png.size, (255, 255, 255))
            background.paste(png, mask=png.split()[3])
            background.save(f'bbbv3-certs/{names[participant]}_{events[participant]}.pdf')
            oc.put_file(f'bbb-v3-certificates/{names[participant]}/{filename}',
                        f'bbbv3-certs/{filename}')
        if email:
            sendmail(receiver=email[participant], html=html)
