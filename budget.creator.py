import sqlite3,subprocess,os,time,json,shutil,openpyxl,fnmatch,datetime,pyodbc
subprocess.list2cmdline = lambda x: " ".join(x)



def issueSQLCommand(command):
	global server
	global username
	global password
	global database
	global cursor
#	p = subprocess.Popen(['sqlcmd','-S%s'%server,'-U%s'%username,'-P%s'%password,'-d%s'%database,'-Q"%s"'%command], stdout = subprocess.PIPE)
#	lines = p.stdout.readlines()
#	p.wait()
	lines=cursor.execute(command).fetchall()
	return lines

def outputLines(lines,filepath = ''):
	if not filepath:
		filepath = 'out.txt'
	f = open(filepath,'w')
	for i in lines:
		f.write(json.dumps(lines)+'\n')
	f.close()
	
def runReport(mylist,filepath,date,numberofyears):
	columnoffset=101
	command="SELECT Sub,Descr FROM SubAcct"
	lines=issueSQLCommand(command)
	subacctdict={}
	#subacctdict={sub:name,...}
	for i in lines:
		sub=i[0].strip()
		desc=i[1].strip()
		subacctdict[sub]=desc
	fiscalstart = 9
	year = int(date.split('/')[2])
	month = int(date.split('/')[0])
	day = int(date.split("/")[1])
	if month >= fiscalstart:
		myyear = year + 1
		mymonth = month + 1 - fiscalstart
	else:
		myyear=year
		mymonth = 13 - fiscalstart + month
	endperiod=int('%s%s'%(str(myyear),format(mymonth,'02d')))
	if mymonth == 12:
		myyear+=1
		mymonth=0
	startperiod=int('%s%s'%(str(myyear-int(numberofyears)),format(mymonth,'02d')))+1
	headermerge=chr(columnoffset+(int(startperiod/100)+int(numberofyears)-int(startperiod/100))*3-1).upper()
	for i in mylist:
		expensetotal={}
		#expensetotal={period1:total,...}
		sheetname=i[0]
		print "Working on %s"%sheetname
		myfilepath = "%s\%s"%(filepath,sheetname)
		if not os.path.isdir(myfilepath):
			os.mkdir(myfilepath)
		myfilepath="%s\%s Budget Worksheet.xlsx"%(myfilepath,sheetname)
		mysublist=''
		for a in range(2,len(i)):
			mysub=i[a][2].replace('-','')
			if ',' in mysub:
				mysub=mysub.split(',')
			else:
				mysub=[mysub]
			for b in mysub:
				if b not in mysublist:
					if len(mysublist)!=0:
						mysublist='%s%s'%(mysublist," OR ")
					mysublist = "%ssub LIKE '%s'"%(mysublist,b)
		command="SELECT Acct,PerPost,sub,dramt, Cramt,TranDesc from gltran WHERE PerPost>'%s' AND PerPost<='%s' AND (%s)"%(startperiod,endperiod,mysublist.replace('?','%'))
		mydict={}
		#mydict={subacct:{acct:{perpost:{vendid:amt},...},...},...}
		lines = issueSQLCommand(command)
		print "Processing Transactions"
		counter=0
		for a in lines:
			counter+=1
			if counter==100:
				print ".",
				counter=0
			acct=a[0].strip()
			perpost=detPeriod(int(a[1].strip()),startperiod)
			sub=a[2].strip()
			dramt=a[3]
			cramt=a[4]
			vendid=a[5].split(' ')[0].strip()
			mydict[sub]=mydict.get(sub,{})
			mydict[sub][acct]=mydict[sub].get(acct,{})
			mydict[sub][acct][perpost]=mydict[sub][acct].get(perpost,{})
			mydict[sub][acct][perpost][vendid]=mydict[sub][acct][perpost].get(vendid,0.0) + float(dramt) - float(cramt)
		print "\nCreating workbook"
		wb = openpyxl.workbook.Workbook()
		ws=wb.worksheets[0]
		ws.title = sheetname
		ws.cell('A1').value="%s Budget Worksheet"%sheetname
		ws.cell('A1').style.alignment.horizontal='center'
		ws.merge_cells('A1:%s1'%headermerge)
		ws.cell('A2').value="For the %s Years"%numberofyears
		ws.cell('A2').style.alignment.horizontal='center'
		ws.merge_cells('A2:%s2'%headermerge)
		ws.cell('A3').value="Ended %s"%date
		ws.cell('A3').style.alignment.horizontal='center'
		ws.merge_cells('A3:%s3'%headermerge)
		for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
			ws.cell('%s4'%chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1)).style.alignment.horizontal='left'
			ws.merge_cells('%s4:%s4'%(chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1),chr((columnoffset+3)+(eachperiod-int(startperiod/100))*3-1)))
			ws.cell('%s4'%chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1).upper()).value=eachperiod
		rowoffset=5
		acctdict={}
		#acctdict={acct:{subacct:total,...},...}
		print "Processing Row Layouts"
		counter=0
		for eachrownum in range(2,len(i)):
			counter+=1
			if counter==10:
				print ".",
				counter=0
			myrowtotals={}
			#myrowtotals={period1:value,...}
			if ',' in i[eachrownum][1]:
				myacctlist=i[eachrownum][1].split(',')
			else:
				myacctlist=[i[eachrownum][1]]
			myaccttotal=0.0
			titlerow=int(rowoffset)
			ws.cell('A%s'%rowoffset).value=i[eachrownum][0]
			rowoffset+=1
			myrowlist=[]
			#myrowlist=[subacct1,subacct2,...]
			if ',' in i[eachrownum][2]:
				mysubacctlist=i[eachrownum][2].split(',')
			else:
				mysubacctlist=[i[eachrownum][2]]
			mysubmatches=[]
			for eachsubacct in mysubacctlist:
				filtered=fnmatch.filter(mydict.keys(),eachsubacct.replace('-',''))
				for eachfilteredsubacct in filtered:
					if eachfilteredsubacct not in mysubmatches:
						mysubmatches.append(eachfilteredsubacct)
			for eachsubacct in mysubmatches:
				mysubacctrow=int(rowoffset)
				ws.cell('B%s'%rowoffset).value=subacctdict.get(eachsubacct,eachsubacct)
				mysubaccttotals=[]
				mysubacctdict={}
				#mysubacctdict={vendor:{period:total,...},...}
				for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
					mycurtotal=0.0
					for eachacct in myacctlist:
						for eachvendor,eachtotal in mydict.get(eachsubacct,{}).get(eachacct,{}).get(eachperiod,{}).iteritems():
							mycurtotal = mycurtotal+eachtotal
							myrowtotals[eachperiod]=myrowtotals.get(eachperiod,0.0)+eachtotal
							expensetotal[eachperiod]=expensetotal.get(eachperiod,0.0)+eachtotal
							mysubacctdict[eachvendor]=mysubacctdict.get(eachvendor,{})
							mysubacctdict[eachvendor][eachperiod]=mysubacctdict[eachvendor].get(eachperiod,0.0)+eachtotal
					mysubaccttotals.append(mycurtotal)
				for a in range(0,len(mysubaccttotals)):
					ws.cell('%s%s'%(chr((columnoffset+2)+a*3-1),rowoffset)).value=mysubaccttotals[a]
				rowoffset+=1
				myvendorlist=mysubacctdict.keys()
				myvendorlist=sorted(myvendorlist)
				for eachvendor in myvendorlist:
					eachperiodsdict=mysubacctdict[eachvendor]
					ws.cell('C%s'%rowoffset).value=eachvendor
					for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
						ws.cell('%s%s'%(chr((columnoffset+3)+(eachperiod-int(startperiod/100))*3-1),rowoffset)).value=mysubacctdict[eachvendor].get(eachperiod,"")
					rowoffset+=1
			for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
				ws.cell('%s%s'%(chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1),titlerow)).value=myrowtotals.get(eachperiod,"")
		for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
			ws.cell('%s%s'%(chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1),rowoffset-1)).value=expensetotal.get(eachperiod,"")
		saved=False
		while not saved:
			try:
				wb.save(filename = myfilepath)
				saved=True
			except IOError:
				print 'Output File Open\n  Close file and try again'
				time.sleep(1)
		print "DONE"
		time.sleep(1)

def detPeriod(period,startperiod):
	curper = [int(str(period)[:4]),int(str(period)[4:])]
	startper=[int(str(startperiod)[:4]),int(str(startperiod)[4:])]
	breaks=int(str(startperiod)[4:])
	startyear=startper[0]
	while startper!=curper:
		startper=[startper[0],startper[1]+1]
		if startper[1]>12:
			startper=[startper[0]+1,1]
		if startper[1]==breaks:
			startyear+=1
	return startyear
	

def writeFile(sheetname,exportgrid,filepath=''):
	if len(sheetname) > 30:
		sheetname = sheetname[:31]
	wb = openpyxl.workbook.Workbook()
	ws=wb.worksheets[0]
	ws.title = sheetname
	for i in range(0,exportgrid.GetNumberRows()):
		for a in range(0,exportgrid.GetNumberCols()):
			ws.cell('%s%s'%(chr(a+97),i+1)).value = exportgrid.GetCellValue(i,a)
	try:
		wb.save(filename = filepath)
	except IOError:
		print 'Output File Open\n  Close file and try again'
		time.sleep(1)

#mydatabase=[[title,description,[Row Title,acct num,sub acct],...],...]
try:
	i = open('./Database/filepath.txt','r')
	filepath = json.loads(i.read())
	i.close()
except:
	filepath=""
try:
	i = open('./Database/Database.txt','r')
	mydatabase = json.loads(i.read())
	i.close()
except:
	mydatabase = []

os.system('cls')
server = raw_input('Server\n?...')
database=raw_input('Database\n?...')
username=raw_input('Username\n?...')
password=raw_input('Password\n?...')
date=raw_input('Date\n?...')
numberofyears=raw_input('Number of Years\n?...')

connection = pyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s'%(server,database,username,password))
cursor = connection.cursor()


while True:
	try:
		os.system('cls')
		command = raw_input('1: view list of reports\n2: run all reports\n3: run report #\n4: Set File Output\n5: Change Inputs\n6: Quit without save\n?...')
		if command == '1':
			os.system('cls')
			for i in range(0,len(mydatabase)):
				print "%s: %s, %s"%(i,mydatabase[i][0],mydatabase[i][1])
			try:
				command=raw_input('\n\n1:update report #\n2:Create New Report\n3:Delete Report\n4:View Report\n?...')
				if command=='1':
					report = raw_input('Report #?\n')
					os.system('cls')
					print "0: %s\n1: %s"%(mydatabase[int(report)][0],mydatabase[int(report)][1])
					for i in range(2,len(mydatabase[int(report)])):
						print '%s: %s| %s->%s'%(i,mydatabase[int(report)][i][0],mydatabase[int(report)][i][1],mydatabase[int(report)][i][2])
					print '%s: ...'%str(i+1)
					line=raw_input('\n Line #...')
					os.system('cls')
					if int(line)==len(mydatabase[int(report)]):
						print 'Adding lines'
						while True:
							try:
								print 'use ^c to stop'
								mydatabase[int(report)].append([raw_input('\nTitle...'),raw_input('\nAcct...'),raw_input('Sub Acct...')])
							except KeyboardInterrupt:
								break
					else:
						print '%s | %s'%(mydatabase[int(report)][int(line)][0],mydatabase[int(report)][int(line)][1])
						mydatabase[int(report)][int(line)]=[raw_input('New Title...'),raw_input('New Acct...'),raw_input('New Sub Acct...')]
				elif command=='2':
					myreport=[raw_input('\nNew Report\n----------\nName...'),raw_input('Description...')]
					while True:
						try:
							print 'use ^c to stop'
							myreport.append([raw_input('\nTitle...'),raw_input('\nAcct...'),raw_input('Sub Acct...')])
						except KeyboardInterrupt:
							break
					mydatabase.append(myreport)
				elif command=='3':
					report = raw_input('delete #\n?...')
					mydatabase.pop(int(report))
				elif command=='4':
					report = raw_input('Report #...')
					os.system('cls')
					print '%s: %s'%(mydatabase[int(report)][0],mydatabase[int(report)][1])
					for i in range(2,len(mydatabase[int(report)])):
						print '%s \t| %s \t| %s'%(mydatabase[int(report)][i][0],mydatabase[int(report)][i][1],mydatabase[int(report)][i][2])
					raw_input('Press Enter to continue')
			except KeyboardInterrupt:
				pass
		elif command == '2':
			runReport(mydatabase,filepath,date,numberofyears)
		elif command == '3':
			os.system('cls')
			for i in range(0,len(mydatabase)):
				print "%s: %s, %s"%(i,mydatabase[i][0],mydatabase[i][1])
			report=raw_input("Which report?\n...")
			myreport=[mydatabase[int(report)]]
			runReport(myreport,filepath,date,numberofyears)
		elif command=='4':
			os.system('cls')
			filepath=raw_input("enter complete file path\n...")
			o=open('./Database/filepath.txt','w')
			o.write(json.dumps(filepath))
			o.close()
		elif command=='5':
			os.system('cls')
			server = raw_input('Server\n?...')
			database=raw_input('Database\n?...')
			username=raw_input('Username\n?...')
			password=raw_input('Password\n?...')
			date=raw_input('Date\n?...')
			numberofyears=raw_input('Number of Years\n?...')
		elif command=='6':
			exit()
		else:
			confirm=raw_input('Really Quit? (y/n)')
			if confirm=='y':
				break
		time.sleep(1)
	except KeyboardInterrupt:
		break
try:
	backupnumber=len(os.listdir("./Database"))
	shutil.move('./Database/Database.txt','./Database/Database.txt.bak%s'%backupnumber)
except:
	confirm=raw_input('\n\nno database found\ncontinue without backup? (y/n)\n')
	if not confirm=='y':
		o=open('./Database/CurrentSessionDatabase.txt','w')
		o.write(json.dumps(mydatabase))
		o.close()
		exit()
o=open('./Database/Database.txt','w')
o.write(json.dumps(mydatabase))
o.close()

print '\n\nGood-bye'

'''
command="""select * from SubAcct"""

lines = issueSQLCommand(command,server,username,password,database)
outputLines(lines)
'''

