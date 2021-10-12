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
import openpyxl

os.startfile("outlook")
accounts = {}

with open("./accounts.json") as accounts_f:
    accounts = json.load(accounts_f)

cc = accounts["cc"]
totals = []

wb = openpyxl.load_workbook("DevDiv Dogfooding By M2.xlsx")
try:
  sheet = wb[time.strftime('%y-%m-%d' , time.localtime())]
except KeyError:
  sheet = wb.create_sheet(time.strftime('%y-%m-%d' , time.localtime()))
  sheet["A1"] = "Manager"
  sheet["B1"] = "Week Of 9/27 %Dogfooding"

i = 2
while True:
  if sheet["A{0}".format(i)].value is None:
    break
  else:
    i+=1

Managers = "lubomirb,manishj,sacalla,barryta,anthc,dondr,anirudhg,laguiler,zhiszhan,sameal,batul,danmose,faijaz,galini,kevinpi,paulku,srivatsn,ansonh,neerajar,bogdanm,yuvalm,rhadley,artl,skofman,gregar,sumitg,qingye,zhenjiao,shawnr,jeffschw,masafa,brandonb,gboland,ruisun,pfeldman,moabdu,joncart,cweining,grwheele,pchapman,mandywhaley,heathar,stefsch,mayurid"
lists = Managers.split(",")


for account in accounts["accounts"]:
  if len(lists) > 0 and account["alias"] not in lists:
    continue
  url = "https://shipready.vsdata.io/vs/whodogfoods/wdmain?alias={0}".format(account["alias"]) 
  app = Application().start("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
  edge = Desktop(backend="uia")["New tab - Work - Microsoft Edge"]
  time.sleep(2)
  edge.type_keys(url)
  edge.type_keys("{ENTER}")
  time.sleep(12)
  label = edge["{0} reports".format(account["alias"])]
  
  content = ""
  with open("./email.html", encoding="utf-8") as email_f:
    content = email_f.read()
  content = content.replace("ALIASIMAGE", account["alias"])
  parent = label.parent()
  children = parent.children()
  for index in range(len(children)):
    if "{0} reports".format(account["alias"]) in children[index].element_info.name:
      break
  elem = children[index+1]

  sheet["A{0}".format(i)] = account["fullName"]
  sheet["B{0}".format(i)] = elem.element_info.name
  i += 1
  wb.save("DevDiv Dogfooding By M2.xlsx")

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
 
  outlook = win32.Dispatch("Outlook.Application")
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