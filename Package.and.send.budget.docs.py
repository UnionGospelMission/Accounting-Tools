import smtplib,getpass,os
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.mime.text import MIMEText
from email import Encoders
from email.utils import formatdate

servername = False
while not servername:
	try:
		servername = raw_input('email server?\n')
		server = smtplib.SMTP(servername)
	except (socket.gaierror, socket.error):
		print 'Connection to email server failed'
		server = False
server.starttls()

success = False
while not success:
	try:
		username = raw_input('username?\n')
		password = getpass.getpass('password?\n')
		server.login(username,password)
		success = True
	except smtplib.SMTPAuthenticationError:
		print 'Invalid Login'

return_test = [501]
while return_test[0]==501:
	return_address = raw_input('return address?\n')
	return_test = server.docmd('mail from:',return_address)
	if return_test[0] == 501:
		print 'Invalid Return Address'

server.close()

subject = 'Budget Process 2016-2017'

body1 = \
'''Hi %s,

Here are your budget worksheet and salaries worksheet for the current budget process.  Please look them over carefully and let me know if you see any mistakes.

The budget process this year will be:
1) look over your personnel requirements for the upcomming year, including positions, hours, and wages to create a proposal
2) meet with your senior leader to review the proposal
3) return the updated salaries worksheet to Dean Bitz by ***MARCH 31st***

4) look over your budget worksheet for an idea of your expenses the past 4 years.
5) decide what changes to your program you would like to see and estimate the costs.
6) meet with your senior leader to review your proposal
7) return your proposal to me by ***APRIL 29th***

8) we will schedule a time to meet in early May to review your proposal together an get it into the budget proposal

9) the board will review the budget proposal in total and make recommendations

10) I will send you whatever the board approves for your next year budget

Please let me know if you have any questions.

Thanks!

--Luke Perkins--'''

body2 = \
'''Hi %s,

Here is your budget worksheet for the current budget process.  Please look it over carefully and let me know if you see any mistakes.

The budget process this year will be:
1) discuss and salaries changes you would like with your senior leader 
2) inform Dean of the outcome by ***MARCH 31st***

3) look over your budget worksheet for an idea of your expenses the past 4 years.
4) decide what changes to your program you would like to see and estimate the costs.
5) meet with your senior leader to review your proposal
6) return your proposal to me by ***APRIL 29th***

7) we will schedule a time to meet in early May to review your proposal together an get it into the budget proposal

8) the board will review the budget proposal in total and make recommendations

9) I will send you whatever the board approves for your next year budget

Please let me know if you have any questions.

Thanks!

--Luke Perkins--'''


def sendEmail(addresses,body=body1,attachments=[],translation={},autosend = False):
    global servername,username,password
    server = smtplib.SMTP(servername)
    server.starttls()
    server.login(username,password)
    message = MIMEMultipart()
    message['Return-Receipt-To'] = return_address
    message['From'] = return_address
    message['To'] = ', '.join(addresses)
    message['Date'] = formatdate(localtime=True)
    message['Subject'] = subject
    message.attach(MIMEText(body))
    for each_file in attachments:
        attachment = MIMEBase('application', "octet-stream")
        attachment.set_payload(open(each_file,'rb').read())
        Encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment; filename="%s"'%os.path.basename(each_file))
        message.attach(attachment)
    server.sendmail(return_address,addresses,message.as_string())
    server.close()

os.chdir('o:Budget/2016-2017')
a=os.listdir('.')

sentto = []
try:
	for i in a:
		if os.path.isdir(i):
			print i
			b = os.listdir(i)
			head = raw_input('department head?\n')
			if len(b)<2:
				body = body2%(head,)
			else:
				body = body1%(head,)
			attachments = [i+'\\'+z for z in b]
			email = raw_input('email?\n')
			if email != '':
				sendEmail([email],body,attachments)
				sentto.append(email)
				
finally:
	sendEmail([return_address],body='\n'.join(sentto))
