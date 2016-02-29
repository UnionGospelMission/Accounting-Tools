import os,openpyxl,pyodbc,getpass,shutil

dbs = False
while not dbs:
	try:
		db = pyodbc.connect('DRIVER={SQL Server};SERVER=mission3;DATABASE=slugm;UID=sa;PWD=%s'%(
																						#raw_input('server\n'),
																						#raw_input('database\n'),
																						#raw_input('username\n'),
																						getpass.getpass('password\n'),))
		dbc=db.cursor()
		dbs = True
	except KeyboardInterrupt:
		raise SystemExit
	except pyodbc.Error:
		print 'connection error'
	except:
		print 'error'
		
empdict = {i[0].strip():[i[1],round(i[2]/2080,2) if i[2] > 0 else i[3],i[4]] for i in dbc.execute('select empid,status,stdslry,stdunitrate,name from employee').fetchall()}


os.chdir('o:Budget/2016-2017')
a=os.listdir('.')
for i in a:
	if os.path.isdir(i):
		for b in os.listdir(i):
			if 'Salary' in b and '~' not in b:
				print b
				updated = False
				c = openpyxl.load_workbook(i+'\\'+b)
				row = 12
				while row:
					empid = c.worksheets[0]['A%s'%row].value
					if empid:
						emp = empdict.get(str(empid),False)
						if not emp:
							print empid,"error"
							raw_input()
						else:
							if emp[0] != 'A' or str(emp[1])!=str(c.worksheets[0]['F%s'%row].value):
								print empid,emp[2].strip(),':',emp[0],emp[1],c.worksheets[0]['F%s'%row].value
								test = True
								if emp[0] != 'A':
									test = raw_input("remove? (Y/n)")
									if test =='Y' or test=='y' or not test:
										c.worksheets[0]['A%s'%row].value=''
										c.worksheets[0]['B%s'%row].value=''
										updated = True
								if test !='Y' and test!='y' and test and str(emp[1])!=str(c.worksheets[0]['F%s'%row].value):
									test = raw_input("update? (Y/n)")
									if test =='Y' or test=='y' or not test:
										c.worksheets[0]['F%s'%row].value=emp[1]
										updated = True
						row+=1
					else:
						row=0
				if updated:
					test = raw_input("save? (Y/n)")
					if test =='Y' or test=='y' or not test:
						c.save(i+'\\new_'+b)
						shutil.move(i+'\\new_'+b,i+'\\'+b)
				
