import shutil,os

backupnumber=len(os.listdir("."))
shutil.copy2("Database.txt","MakeReadableDatabaseBackup%s.txt"%backupnumber)

i = open('./Database.txt','r')
mydatabase = i.read()
i.close()

o=open('./Database.txt','w')
closebracket=False
extraline=False
for i in mydatabase:
	if i =='[':
		o.write('\n')
		if extraline:
			o.write('\n')
			extraline=False
	if i==']' and closebracket:
		extraline=True
	if i ==']':
		closebracket=True
	else:
		closebracket=False
	o.write(i)
o.close()
