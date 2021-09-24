from subprocess import Popen
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto import mouse
from PIL import Image, ImageGrab
import json
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage 
from email.header import Header
import os
import win32com
import win32com.client as win32

os.startfile("outlook")
accounts = {}

with open("./accounts.json") as accounts_f:
    accounts = json.load(accounts_f)

cc = accounts["cc"]
for account in accounts["accounts"]:
  url = "https://shipready.vsdata.io/vs/whodogfoods/wdmain?alias={0}".format(account["alias"]) 
  app = Application().start("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
  edge = Desktop(backend="uia")["New tab - Work - Microsoft Edge"]
  edge.type_keys(url)
  edge.type_keys("{ENTER}")
  time.sleep(25)
  label = edge["{0} reports".format(account["alias"])]
  
  content = ""
  with open("./email.html") as email_f:
    content = email_f.read()
  content = content.replace("ALIASIMAGE", account["alias"])
  parent = label.parent()
  children = parent.children()
  for index in range(len(children)):
    if "{0} reports".format(account["alias"]) in children[index].element_info.name:
      break
  elem = children[index+1]
  rect = elem.rectangle()
  aliasRect = label.rectangle()
  mouse.move(coords=(int((rect.left + rect.right)/2), int((rect.top + rect.bottom)/2)))
  time.sleep(2)
  mouse.click(button='left', coords=(int((rect.left + rect.right)/2), int((rect.top + rect.bottom)/2)))
  time.sleep(1)
  mouse.click(button='left', coords=(int((rect.left + rect.right)/2), int((rect.top + rect.bottom)/2)))
  time.sleep(1)
  img = ImageGrab.grab((aliasRect.left, rect.top - 300, rect.left + 600, rect.top + 350))
  img.save("./images/{0}.png".format(account["alias"]))
  mouse.move(coords=(100, 100))
  edge["Close"].click()

  subject = "VS 2022 Dogfood Adoption Rates {0} for week ending {1}".format(
  account["alias"], time.strftime('%m/%d/%y' , time.localtime()))


  outlook = win32.Dispatch('Outlook.Application')
  mail = outlook.CreateItem(0)
  mail.Recipients.Add(account["email"])
  mail.Subject = subject
  if(len(account["fullName"]) == 0):
    mail.CC = cc
  else:
    mail.CC = cc + ";" + account["cc"]
  mail.BodyFormat = 2
  mail.Attachments.Add(os.path.abspath("./images/{0}.png".format(account["alias"])))
  mail.HtmlBody = content
  mail.Display()
  mail.Send()
  '''
  message = MIMEMultipart('related')
  message['From'] = "lvniao11235@163.com"
  message['To'] = account["email"]
  message['Subject'] = subject
  message.attach(MIMEText(content,'html','utf-8'))
  if(len(account["fullName"]) == 0):
    message['CC'] = accounts["cc"]
  else:
    message['CC'] = account["cc"]
  
  with open("./images/{0}.png".format(account["alias"]), 'rb') as img:
    image = MIMEImage(img.read())
    image.add_header('Content-ID', '<image1>')
    message.attach(image)
  
  smtpObj = smtplib.SMTP_SSL(accounts["emailServerUrl"], accounts["emailServerPort"])
  smtpObj.login(accounts["emailUsername"], accounts["emailPassword"])
  smtpObj.sendmail("lvniao11235@163.com", "lvniao11235@163.com", message.as_string())
  time.sleep(2)
  '''
  


