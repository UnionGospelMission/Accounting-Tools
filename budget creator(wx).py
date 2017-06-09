#!/usr/bin/env python
# -*- coding: CP1252 -*-
#
# generated by wxGlade 0.6.8 (standalone edition) on Thu Feb 02 10:51:10 2017
#

import wx, wx.grid, json, pyodbc, os, openpyxl, fnmatch, time

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode
# end wxGlade

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



class MyFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MyFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.Settings = wx.Panel(self, wx.ID_ANY)
        self.label_2_copy = wx.StaticText(self.Settings, wx.ID_ANY, "Input Path")
        self.button_3_copy = wx.Button(self.Settings, wx.ID_ANY, "Set Input File")
        self.OUTPUTPATH_copy = wx.StaticText(self.Settings, wx.ID_ANY, "")
        self.panel_2 = wx.Panel(self, wx.ID_ANY)
        self.label_2 = wx.StaticText(self.panel_2, wx.ID_ANY, "Output Path")
        self.button_3 = wx.Button(self.panel_2, wx.ID_ANY, "Set Output File")
        self.OUTPUTPATH = wx.StaticText(self.panel_2, wx.ID_ANY, "")
        self.panel_4 = wx.Panel(self, wx.ID_ANY)
        self.label_2_copy_1 = wx.StaticText(self.panel_4, wx.ID_ANY, "Period to Post")
        self.PERIOD = wx.TextCtrl(self.panel_4, wx.ID_ANY, "")
        self.OUTPUTPATH_copy_1 = wx.StaticText(self.panel_4, wx.ID_ANY, "")
        self.panel_1 = wx.Panel(self, wx.ID_ANY)
        self.button_3_copy_1 = wx.Button(self.panel_1, wx.ID_ANY, "Process")
        self.panel_3 = wx.Panel(self, wx.ID_ANY)
        self.label_1 = wx.StaticText(self.panel_3, wx.ID_ANY, "GL Journal Entry Vendors")
        self.REPORTVIEWER = wx.grid.Grid(self, wx.ID_ANY, size=(1, 1))

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.setInput, self.button_3_copy)
        self.Bind(wx.EVT_BUTTON, self.setOutput, self.button_3)
        self.Bind(wx.EVT_TEXT_ENTER, self.validatePTP, self.PERIOD)
        self.Bind(wx.EVT_BUTTON, self.processFile, self.button_3_copy_1)
        self.Bind(wx.grid.EVT_GRID_CMD_CELL_CHANGE, self.saveReport, self.REPORTVIEWER)
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MyFrame.__set_properties
        self.SetTitle("Transaction Importer")
        self.SetSize((427, 404))
        self.panel_3.SetMinSize((422, 13))
        self.REPORTVIEWER.CreateGrid(0, 1)
        self.REPORTVIEWER.EnableDragColSize(0)
        self.REPORTVIEWER.EnableDragRowSize(0)
        self.REPORTVIEWER.EnableDragGridSize(0)
        self.REPORTVIEWER.SetColLabelValue(0, "Vendor Name")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MyFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_3_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8_copy_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8_copy.Add(self.label_2_copy, 0, 0, 0)
        sizer_8_copy.Add(self.button_3_copy, 0, 0, 0)
        sizer_8_copy.Add(self.OUTPUTPATH_copy, 0, 0, 0)
        self.Settings.SetSizer(sizer_8_copy)
        sizer_1.Add(self.Settings, 0, wx.EXPAND, 0)
        sizer_8.Add(self.label_2, 0, 0, 0)
        sizer_8.Add(self.button_3, 0, 0, 0)
        sizer_8.Add(self.OUTPUTPATH, 0, 0, 0)
        self.panel_2.SetSizer(sizer_8)
        sizer_1.Add(self.panel_2, 0, wx.EXPAND, 0)
        sizer_8_copy_1.Add(self.label_2_copy_1, 0, 0, 0)
        sizer_8_copy_1.Add(self.PERIOD, 0, 0, 0)
        sizer_8_copy_1.Add(self.OUTPUTPATH_copy_1, 0, 0, 0)
        self.panel_4.SetSizer(sizer_8_copy_1)
        sizer_1.Add(self.panel_4, 0, wx.EXPAND, 0)
        sizer_2.Add(self.button_3_copy_1, 0, 0, 0)
        self.panel_1.SetSizer(sizer_2)
        sizer_1.Add(self.panel_1, 0, wx.EXPAND, 0)
        sizer_3_copy.Add(self.label_1, 0, 0, 0)
        self.panel_3.SetSizer(sizer_3_copy)
        sizer_1.Add(self.panel_3, 0, wx.EXPAND, 0)
        sizer_1.Add(self.REPORTVIEWER, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade

    def setInput(self, event):  # wxGlade: MyFrame.<event_handler>
        print "Event handler 'setInput' not implemented!"
        event.Skip()

    def setOutput(self, event):  # wxGlade: MyFrame.<event_handler>
        print "Event handler 'setOutput' not implemented!"
        event.Skip()

    def processFile(self, event):  # wxGlade: MyFrame.<event_handler>
        print "Event handler 'processFile' not implemented!"
        event.Skip()

    def saveReport(self, event):  # wxGlade: MyFrame.<event_handler>
        print "Event handler 'saveReport' not implemented!"
        event.Skip()

    def validatePTP(self, event):  # wxGlade: MyFrame.<event_handler>
        print "Event handler 'validatePTP' not implemented!"
        event.Skip()
# end of class MyFrame
class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # Content of this block not found. Did you rename this class?
        pass
        self.db = { "SERVER":"",
                    "DATABASE":"",
                    "USERNAME":"",
                    "YEARS":0,
                    "REPORTLIST":[],
                  }
        try:
            i=open('./Database/Database.txt','r')
            a=json.loads(i.read())
            i.close()
            self.db.update(a)
        except IOError:
            pass
        self.load_report = False
        self.save_inputs = False
        self.SERVER.Value=self.db['SERVER']
        self.DATABASE.Value=self.db['DATABASE']
        self.USERNAME.Value=self.db['USERNAME']
        self.YEARS.Value=self.db['YEARS']
        self.reportdb = {}
        for i in range(0,len(self.db['REPORTLIST'])):
            self.REPORTS.AppendRows()
            self.REPORTS.SetCellValue(i,0,self.db['REPORTLIST'][i][0])
            self.REPORTS.SetCellValue(i,1,self.db['REPORTLIST'][i][1])
            self.reportdb[i] = self.db['REPORTLIST'][i]
        self.REPORTS.AutoSizeColumns()
        self.load_report = True
        self.save_inputs = True
        self.save_report = True
            

    def __set_properties(self):
        # Content of this block not found. Did you rename this class?
        pass

    def __do_layout(self):
        # Content of this block not found. Did you rename this class?
        pass

    def saveInputs(self, event):  # wxGlade: MainFrame.<event_handler>
        if self.save_inputs:
            self.db['SERVER']=self.SERVER.Value
            self.db['DATABASE']=self.DATABASE.Value
            self.db['USERNAME']=self.USERNAME.Value
            self.db['YEARS']=self.YEARS.Value
            self.saveDB()
        event.Skip()

    def saveDB(self):
        report_list = []
        for k,v in self.reportdb.iteritems():
            report_list.append(v)
        self.db['REPORTLIST'] = report_list
        o=open('./Database/Database.txt','w')
        o.write(json.dumps(self.db))
        o.close()

    def addReportList(self, event):  # wxGlade: MainFrame.<event_handler>
        self.REPORTS.AppendRows()
        event.Skip()

    def saveReportList(self, event):  # wxGlade: MainFrame.<event_handler>
        new_reports = {}
        for i in range(0,self.REPORTS.GetNumberRows()):
            if self.REPORTS.GetCellValue(i,0) or self.REPORTS.GetCellValue(i,1):
                new_reports[i] = [self.REPORTS.GetCellValue(i,0), self.REPORTS.GetCellValue(i,1), self.reportdb.get(i,['','',[]])[2]]
        self.reportdb = new_reports
        self.saveDB()
        event.Skip()

    def loadReport(self, event=None):  # wxGlade: MainFrame.<event_handler>
        if self.load_report:
            self.report_number = event
            if hasattr(event,"GetRow"):
                self.report_number = event.GetRow()
            report = self.reportdb.get(self.report_number,None)
            for i in range(self.REPORTVIEWER.GetNumberRows()-1,-1,-1):
                self.REPORTVIEWER.DeleteRows()
            for i in range(0,self.REPORTS.GetNumberRows()):
                self.REPORTS.SetCellBackgroundColour(i,0,wx.NamedColour('white'))
                self.REPORTS.SetCellBackgroundColour(i,1,wx.NamedColour('white'))
            if report:
                self.REPORTS.SetCellBackgroundColour(self.report_number,0,wx.NamedColour('red'))
                self.REPORTS.SetCellBackgroundColour(self.report_number,1,wx.NamedColour('red'))
                for i in range(0,len(report[2])):
                    self.REPORTVIEWER.AppendRows()
                    self.REPORTVIEWER.SetCellValue(i,0,report[2][i][0])
                    self.REPORTVIEWER.SetCellValue(i,1,report[2][i][1])
                    self.REPORTVIEWER.SetCellValue(i,2,report[2][i][2])
            self.REPORTVIEWER.AppendRows()
            self.REPORTVIEWER.AutoSizeColumns()
            self.REPORTS.ForceRefresh()
        if hasattr(event,'Skip'):
            event.Skip()

    def saveReport(self, event=None):  # wxGlade: MainFrame.<event_handler>
        if self.save_report:
            self.save_report = False
            report = []
            for i in range(0,self.REPORTVIEWER.GetNumberRows()):
                if self.REPORTVIEWER.GetCellValue(i,0):
                    report.append([self.REPORTVIEWER.GetCellValue(i,0),self.REPORTVIEWER.GetCellValue(i,1),self.REPORTVIEWER.GetCellValue(i,2)])
            self.reportdb[self.report_number][2] = report
            self.saveDB()
            if self.REPORTVIEWER.GetCellValue(self.REPORTVIEWER.GetNumberRows()-1,0):
                self.REPORTVIEWER.AppendRows()
            self.REPORTVIEWER.AutoSizeColumns()
            self.save_report = True
        if event:
            event.Skip()

    def insertRow(self, event):  # wxGlade: MainFrame.<event_handler>
        if self.ROWINSERT.Value:
            self.REPORTVIEWER.InsertRows(self.ROWINSERT.Value-1)
        event.Skip()
    
    def runReport(self, event):  # wxGlade: MainFrame.<event_handler>
        if not self.SERVER.Value or not self.DATABASE.Value or not self.USERNAME.Value or not self.PASSWORD.Value or self.YEARS.Value == 0 or not self.OUTPUTPATH.Label:
            wx.MessageBox('Please complete the server connection information','Warning',wx.ICON_WARNING)
            event.Skip()
            return
        try:
            con = pyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s'%(self.SERVER.Value,self.DATABASE.Value,self.USERNAME.Value,self.PASSWORD.Value))
            cur=con.cursor()
        except pyodbc.Error:
            wx.MessageBox('Bad Connection','Connection Failed',wx.ICON_WARNING)
            event.Skip()
            return
        columnoffset=101
        command="SELECT Sub,Descr FROM SubAcct"
        lines=cur.execute(command).fetchall()
        subacctdict={}
        #subacctdict={sub:name,...}
        for i in lines:
            sub=i[0].strip()
            desc=i[1].strip()
            subacctdict[sub]=desc
        fiscalstart = 9
        year = self.DATE.Value.Year
        month = self.DATE.Value.Month+1
        day = self.DATE.Value.Day
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
        startperiod=int('%s%s'%(str(myyear-self.YEARS.Value),format(mymonth,'02d')))+1
        headermerge=chr(columnoffset+(int(startperiod/100)+self.YEARS.Value-int(startperiod/100))*3-1).upper()
        mylist = [v for k,v in self.reportdb.iteritems()]
        if self.REPORTTORUN.Value.isdigit() and int(self.REPORTTORUN.Value)-1>-1 and int(self.REPORTTORUN.Value)-1<len(self.reportdb):
            mylist = [self.reportdb[int(self.REPORTTORUN.Value)-1]]
        numberofyears = self.YEARS.Value
        total_status_dialog = wx.ProgressDialog ( 'Report Processing Progress', 'Processing Reports', maximum = len(mylist), style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)
        for i in mylist:
            expensetotal={}
            #expensetotal={period1:total,...}
            sheetname=i[0]
            total_status_dialog.Update(mylist.index(i)+1, "On %s."%sheetname)
            myfilepath = "%s\%s"%(self.OUTPUTPATH.Label,sheetname)
            if not os.path.isdir(myfilepath):
                os.mkdir(myfilepath)
            myfilepath="%s\%s Budget Worksheet.xlsx"%(myfilepath,sheetname)
            mysublist=''
            for a in range(0,len(i[2])):
                mysub=i[2][a][2].replace('-','')
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
            lines=cur.execute(command).fetchall()
            trans_status_dialog = wx.ProgressDialog ( 'Transaction Processing Progress', 'Processing Transactions', maximum = len(lines), style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)
            counter=0
            for a in lines:
                counter+=1
                trans_status_dialog.Update(counter)
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
            trans_status_dialog.Destroy()
            work_status_dialog = wx.ProgressDialog ( 'Workbook Creation Progress', 'Creating Workbook', maximum = 1, style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)
            wb = openpyxl.workbook.Workbook()
            ws=wb.worksheets[0]
            ws.title = sheetname
            ws.cell('A1').value="%s Budget Worksheet"%sheetname
            ws.cell('A1').alignment = openpyxl.styles.Alignment(horizontal='center')
            ws.merge_cells('A1:%s1'%headermerge)
            ws.cell('A2').value="For the %s Years"%str(self.YEARS.Value)
            ws.cell('A2').alignment = openpyxl.styles.Alignment(horizontal='center')
            ws.merge_cells('A2:%s2'%headermerge)
            ws.cell('A3').value="Ended %s"%str(self.DATE.Value).split(' ')[0]
            ws.cell('A3').alignment = openpyxl.styles.Alignment(horizontal='center')
            ws.merge_cells('A3:%s3'%headermerge)
            for eachperiod in range(int(startperiod/100),int(startperiod/100)+int(numberofyears)):
                ws.cell('%s4'%chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1)).alignment = openpyxl.styles.Alignment(horizontal='left')
                z = '%s4:%s4'%(chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1),chr((columnoffset+3)+(eachperiod-int(startperiod/100))*3-1))
                ws.merge_cells(z.upper())
                ws.cell('%s4'%chr((columnoffset+1)+(eachperiod-int(startperiod/100))*3-1).upper()).value=eachperiod
            rowoffset=5
            acctdict={}
            #acctdict={acct:{subacct:total,...},...}
            row_status_dialog = wx.ProgressDialog ( 'Row Layout Processing Progress', 'Processing Row Layouts', maximum = len(i[2]), style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)
            counter=0
            for eachrownum in range(0,len(i[2])):
                counter+=1
                row_status_dialog.Update(counter)
                myrowtotals={}
                #myrowtotals={period1:value,...}
                if ',' in i[2][eachrownum][1]:
                    myacctlist=i[2][eachrownum][1].split(',')
                else:
                    myacctlist=[i[2][eachrownum][1]]
                myaccttotal=0.0
                titlerow=int(rowoffset)
                ws.cell('A%s'%rowoffset).value=i[2][eachrownum][0]
                rowoffset+=1
                myrowlist=[]
                #myrowlist=[subacct1,subacct2,...]
                if ',' in i[2][eachrownum][2]:
                    mysubacctlist=i[2][eachrownum][2].split(',')
                else:
                    mysubacctlist=[i[2][eachrownum][2]]
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
            row_status_dialog.Destroy()
            saved=False
            while not saved:
                try:
                    wb.save(filename = myfilepath)
                    saved=True
                except IOError:
                    wx.MessageBox('Output File Open\n  Close file and close this box to try again','Connection Failed',wx.ICON_WARNING)
            work_status_dialog.Update(1)
            work_status_dialog.Destroy()
        total_status_dialog.Destroy()

        event.Skip()
    
    def updateRunReport(self, event):  # wxGlade: MainFrame.<event_handler>
        if self.REPORTTORUN.Value.isdigit():
            self.RUNREPORT.Label = "Run Report %s"%self.REPORTTORUN.Value
        else:
            self.RUNREPORT.Label = "Run All Reports"
        event.Skip()
    
    def setOutput(self, event):  # wxGlade: MainFrame.<event_handler>
        dia = wx.DirDialog(None,message = "Find Current Year Accounting Directory", style = wx.OPEN)
        if dia.ShowModal() == wx.ID_OK:
            self.OUTPUTPATH.Label = dia.GetPath()
        event.Skip()
        
    def copyReport(self, event):  # wxGlade: MainFrame.<event_handler>
        if not self.reportdb.get(self.report_number,None):
            wx.MessageBox('Give your report a name and description, then try copying again','Create Report Warning',wx.ICON_WARNING)
            event.Skip()
            return
        r = self.reportdb.get(self.COPYREPORT.Value-1,None)
        if r:
            dia = wx.MessageDialog(self, "Do you really want to copy report %s?"%self.COPYREPORT.Value, "Confirm", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
            confirm = dia.ShowModal()
            dia.Destroy()
            if confirm == wx.ID_OK:
                self.reportdb[self.report_number][2] = r[2]
                self.loadReport(self.report_number)
                self.saveReport()
        else:
            wx.MessageBox('No Such Report','Invalid Report Number',wx.ICON_WARNING)
        event.Skip()
        
    def replaceSubs(self, event):  # wxGlade: MainFrame.<event_handler>
        dia = wx.MessageDialog(self, "Do you really want to replace all subaccounts?", "Confirm", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        confirm = dia.ShowModal()
        dia.Destroy()
        if confirm == wx.ID_OK:
            self.save_report = False
            for i in range(0,self.REPORTVIEWER.GetNumberRows()):
                if self.REPORTVIEWER.GetCellValue(i,2):
                    self.REPORTVIEWER.SetCellValue(i,2,self.REPLACESUBACCOUNTS.Value)
            self.save_report = True
            self.saveReport()
        event.Skip()
# end of class MainFrame
if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frame_1 = MainFrame(None, wx.ID_ANY, "")
    app.SetTopWindow(frame_1)
    frame_1.Show()
    app.MainLoop()
