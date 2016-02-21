#  -*- coding: utf-8 -*-
#  coded by @magic_coding (icoder@mail.com)
#                   __               __            
#    ____  _____   / /   ____  _____/ /_____  _____
#   / __ \/ ___/  / /   / __ \/ ___/ //_/ _ \/ ___/
#  / /_/ / /__   / /___/ /_/ / /__/ ,< /  __/ /    
# / .___/\___/  /_____/\____/\___/_/|_|\___/_/     
#/_/                                                

import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import imaplib
import email
import os
import time
from time import gmtime, strftime
import getpass

CONFIG = 'config.txt'
fp = file(CONFIG)
exec fp in globals()
fp.close()

_admin_email = admin_email
_program_email = program_email
_email_password = email_password

print """                   __               __            
    ____  _____   / /   ____  _____/ /_____  _____
   / __ \/ ___/  / /   / __ \/ ___/ //_/ _ \/ ___/
  / /_/ / /__   / /___/ /_/ / /__/ ,< /  __/ /    
 / .___/\___/  /_____/\____/\___/_/|_|\___/_/     
/_/ """
print "\nPc control started.."
print "\nKeep this window open to receive your command from your ", _admin_email
print "\nOR you can minimise this window."
print """\n|--------------------------------------------------------|
|                                                        |
| I'm ready to receive your command.                     |
| Please send me (/lock) command to lock this computer.  |
|                                                        |
|--------------------------------------------------------|"""


def check():
	mail = imaplib.IMAP4_SSL('imap.gmail.com')
	mail.login(_program_email, _email_password)
	mail.list()
	mail.select("inbox")
	result, data = mail.search(None, '(UNSEEN)')
	if data != [""]:
		ids = data[0]
		id_list = ids.split()
		latest_email_id = id_list[-1]
		result, data = mail.fetch(latest_email_id, "(RFC822)") 
		raw_email = data[0][1]
		b = email.message_from_string(raw_email)
		print "Message received!"
		if b.is_multipart():
			for payload in b.get_payload():
				body = payload.get_payload(None, True).split('<td>',1)[-1].split('</td>')[0].strip()
				if body.count("/lock"):
					sender = email.utils.parseaddr(b['From'])[1]
					if sender.count(_admin_email):
						send_mail()
						time.sleep(2)
						os.popen("rundll32.exe user32.dll,LockWorkStation")
						print "computer locked!"
					else:
						print "Email not have access to use this program."
				else:
					print "Wrong command was sent!"
		else:
			text = b.get_payload()
			if text.count("/lock"):
				sender = email.utils.parseaddr(b['From'])[1]
				if sender.count(_admin_email):
					send_mail()
					time.sleep(2)
					os.popen("rundll32.exe user32.dll,LockWorkStation")
					print "computer locked!"
				else:
					print "Email not have access to use this program."
			else:
				print "Wrong command was sent!"
	else:
		pass


def send_mail():
	msg = MIMEMultipart()
	msg['From'] = _program_email
	msg['To'] = _admin_email
	msg['Subject'] = "Pc Control"

	body = "Hi Admin,\nYour computer locked successfully.\nDate/Time: "+str(strftime("%Y-%m-%d %H:%M:%S", gmtime()))+"\nPC Name: "+getpass.getuser()+"\n\n-------\nCoded by @magic_coding - www.twitter.com/magic_coding"
	msg.attach(MIMEText(body, 'plain'))

	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(_program_email, _email_password)
	text = msg.as_string()
	server.sendmail(_program_email, _admin_email, text)
	server.quit()
	print "Email sent to: ", _admin_email

while True:
	check()