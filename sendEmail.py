import os
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

def send_report():
   #Config Reading Part
   configfile = 'external.config'
   config_file = open(configfile).read()
   config = eval(config_file)
   fromaddr = config['email_sender']
   toaddr = config['email_receiver']
   password = config['email_login_password']

   msg = MIMEMultipart()
   msg['From'] = fromaddr
   msg['To'] = toaddr
   msg['Subject'] = "Selenium Automation Test Result by Jeevan Paudel"
   body = MIMEText("Dear sir, <br> <br>"
                   "This is automation assignment test result send from selenium. <br> <br>"
                   "Please find the attachment below for the test result and test summary. <br> <br>"
                   "Best Regards, <br>"
                   "Jeevan Paudel. <br>"
                   ,'html', 'utf-8')
   msg.attach(body)

   filename = 'Output_Result/test_result/TestResult.xlsx'
   attachment = open(filename, "rb")
   part = MIMEBase('application', 'octet-stream')
   part.set_payload((attachment).read())
   encoders.encode_base64(part)
   part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
   msg.attach(part)
   server = smtplib.SMTP('smtp.gmail.com', 587)
   server.starttls()
   server.login(fromaddr, password)
   text = msg.as_string()
   server.sendmail(fromaddr, toaddr, text)
   print("Email Sent Successfully")
   server.quit()

