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
  content = '<p><b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Your Team’s VS 2022 Dogfood Adoption Rates </span></b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p></o:p></span></p>'
  content += '<p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">&nbsp;</span><img width="650" height="600" src="{0}.png" ><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p></o:p></span></p>'.format(account["alias"])
  content += '<p><b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Your Action</span></b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"><o:p></o:p></span></p>'
  content += '<p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">As you are aware VS 2022 is close to shipping release candidate and we need your help to ensure a quality release.&nbsp; <b>The goal is for each Director/M2 to exceed and maintain a dogfooding rate of 80%.</b> As such we will be tracking Director/M2 across Julia\'s orgs VS 2022 dogfooding adoption numbers, reviewing weekly in VS Tactics and the LT Project Meeting.&nbsp; <o:p></o:p></span></p>'
  content += '<ul style="margin-top:0in" type="disc"><li class="MsoNormal" style="mso-list:l4 level1 lfo1;vertical-align:middle">Please encourage your team to use VS 2022 as their primary IDE for their day to day work and report feedback for any issues they run into.&nbsp; Get more information about the builds and report feedback from&nbsp;<a href="https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Faka.ms%2Fvsdogfood&amp;data=04%7C01%7Cv-joejiao%40microsoft.com%7Cdc74cd53b2fc4f3ae92608d97d6d08b7%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637678729513090057%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=Cq8xeTKacW2NIhvypATfigJ8B7e8jh0EWFcfGN9fgJU%3D&amp;reserved=0"            originalsrc="https://aka.ms/vsdogfood" shash="GJmxdCQO9WpReM+ZfiIej+G0L3FPUualxKRO9+f9r9DwE6iJYlF97r0ZagdZw/z36Vam0igTe/RNr9Y65or5PdDsfdIyE8MxJX1wetLPm6l/mrH0+Sdu3GIl6b1Zr0i5VbZsjiXpp5Z8pmBKzlXlPlIlkDl/PoHwLJTLyFnFh8g=">https://aka.ms/vsdogfood</a><o:p></o:p></li><li class="MsoNormal" style="mso-list:l4 level1 lfo1;vertical-align:middle">Install using one of the links -<o:p></o:p></li><ul style="margin-top:0in" type="circle"><li class="MsoNormal" style="mso-list:l4 level2 lfo1;vertical-align:middle">Nightly internal builds:<a href="https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Faka.ms%2Fvs%2F17%2Fintpreview%2Fvs_enterprise.exe&amp;data=04%7C01%7Cv-joejiao%40microsoft.com%7Cdc74cd53b2fc4f3ae92608d97d6d08b7%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637678729513090057%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=YxdcDBJXYLVS1QZl4JepJRxQQZlRRswanWX4vgulnXc%3D&amp;reserved=0"                originalsrc="https://aka.ms/vs/17/intpreview/vs_enterprise.exe" shash="LOAJIM6MVk3yHmr72R6xlm8bkFSYYK5Lso4gTIoXsHj+AG+XtavfJc7WfJaIN21iRSTI2971fJ+pgcN0lQBv63wZRflvv2+8kaxV9wjh/2RKYR2JWUcBYxe+tprJGNgfB7+Uw6tNsB1ycVoOIKT53GuqsH8Qji+8+QLYZICkl/w=">https://aka.ms/vs/17/intpreview/vs_enterprise.exe</a><o:p></o:p></li><li class="MsoNormal" style="mso-list:l4 level2 lfo1;vertical-align:middle">Public Preview builds:<a href="https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Faka.ms%2Fvs%2F17%2Fpre%2Fvs_enterprise.exe&amp;data=04%7C01%7Cv-joejiao%40microsoft.com%7Cdc74cd53b2fc4f3ae92608d97d6d08b7%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637678729513100045%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=HkzbL40VRVSMbwkCIKX%2FLAdHcv%2BsUbjmNU1IycuzJPY%3D&amp;reserved=0"                originalsrc="https://aka.ms/vs/17/pre/vs_enterprise.exe" shash="wbI33dj8gGI2pR6F7KK3dhIAjAkWS81Ws8Jl/H/dz/wyW8jJQNhktjHEeMBaAwwA95Af2BQ1bAWw0CKCWhM1gtPCeudIVl3xpTPBw1PLUVfeQKRa429Aw3qlwoBfKngA5RIhG/Vkjj3S3vYM5PoBAGfIACoUcQjFU9GIFfPF3xo=">https://aka.ms/vs/17/pre/vs_enterprise.exe</a><o:p></o:p></li><li class="MsoNormal" style="mso-list:l4 level2 lfo1;vertical-align:middle">For teams outside of DougAm or JoC\'s org, Public Preview builds is the recommended option.<o:p></o:p></li></ul><li class="MsoNormal" style="mso-list:l4 level1 lfo1;vertical-align:middle">Exceed a dogfooding rate of 80%. You can use this<a href="https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwhodogfoods.vsdata.io%2F&amp;data=04%7C01%7Cv-joejiao%40microsoft.com%7Cdc74cd53b2fc4f3ae92608d97d6d08b7%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637678729513110039%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C1000&amp;sdata=n1OHnbKdMuB1AOItKObclM1f%2FKmAG6dEhpWM2tD4SLw%3D&amp;reserved=0"            originalsrc="https://whodogfoods.vsdata.io/" shash="dABxvavN5B4CWTyaAIaatpVXCQks3wLWN8WMpOjxQFNoa19WDKM5W2mFEcCMfZ7HYP+0mJNYqXqCtvsM073OH+OUlZdyta4jHyoglHvfXxRGCIOoD8wXutbhOEnfrzJg3MM0CvJ0zi7g81iw13aniqO8wFxuhJ8BGDJ9aTjVS+U=">link</a> to track        adoption in your team.&nbsp; &quot;Primarily using builds from this channel&quot;&nbsp; and &quot;on the latest build&quot; are the key metrics.<o:p></o:p></li><li class="MsoNormal" style="mso-list:l4 level1 lfo1;vertical-align:middle">Follow up with engineers who are not &quot;on the latest build.” If they do not have a valid business reason to be on latest, then encourage them to update asap. If they are blocked open blocking bugs and/or reach out to scottph, riteshp or varung.<o:p></o:p></li></ul><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">&nbsp;<o:p></o:p></span></p><p><b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">FAQ</span></b><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif"></span></p><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">For any questions, please reach out to<i>scottph, riteshp or varung</i>.</span></p><ul style="margin-top:0in" type="disc"><li class="MsoNormal" style="mso-list:l2 level1 lfo2;vertical-align:middle">What does Primarily using build mean?<o:p></o:p></li></ul><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Primary usage: In case a user is using multiple versions of VS in last 7 days, the version that has the most usage seen in telemetry is considered as the primarily used version. This is indicated by the purple star above. This is a rolling 7 day metric and if the user has just started using VS 2022 dogfood build, it might take some usage from the user on the dogfood builds to change the Primary IDE indication on the site from the older version that the user was using.<o:p></o:p></span></p><ul style="margin-top:0in" type="disc"><li class="MsoNormal" style="mso-list:l1 level1 lfo3;vertical-align:middle">Are employees&nbsp; who do not use VS at all expected to dogfood?<o:p></o:p></li></ul><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">No. The primary metric takes this into account.<o:p></o:p></span></p><ul style="margin-top:0in" type="disc"><li class="MsoNormal" style="mso-list:l0 level1 lfo4;vertical-align:middle">What builds qualify as dogfood builds?<o:p></o:p></li></ul><p><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Any build newer than (N-1) public Preview is considered as dogfood build. However, it would be great if users can update to the latest build once a week at least.<o:p></o:p></span></p><p style="margin:0in;margin-bottom:.0001pt"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">&nbsp;<o:p></o:p></span></p><ul style="margin-top:0in" type="disc"><li class="MsoNormal" style="mso-list:l3 level1 lfo5;vertical-align:middle">How frequently should one update the build?<o:p></o:p></li></ul><p style="margin:0in;margin-bottom:.0001pt"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Nightly internal builds – please upgrade at least once a week.<o:p></o:p></span></p><p style="margin:0in;margin-bottom:.0001pt"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Public Preview builds – anytime a new build is released to this channel.<o:p></o:p></span></p><p style="margin:0in;margin-bottom:.0001pt"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">&nbsp;<o:p></o:p></span></p><p style="margin:0in;margin-bottom:.0001pt"><span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Thanks!<o:p></o:p></span></p><p class="MsoNormal"><o:p>&nbsp;</o:p></p></p>'

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
  


