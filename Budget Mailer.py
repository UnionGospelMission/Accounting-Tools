#!/usr/bin/env python
# -*- coding: CP1252 -*-
#
# generated by wxGlade 0.6.8 (standalone edition) on Tue Mar 21 08:45:21 2017
#

import wx,wx.grid,json,socket,smtplib,os
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.mime.text import MIMEText
from email import Encoders
from email.utils import formatdate

# begin wxGlade: dependencies
# end wxGlade

class Globals():
    def __init__(self):
        self.settings={"SERVER":"","USERNAME":"","RETURNADDRESS":"","VIEWER":{},"SUBJECT":"","EMAIL":"","ATTACHMENTS":"","FOLDER":""}
        try:
            i=open("settings.txt")
            self.settings=json.loads(i.read())
        except IOError:
            i=open("settings.txt","w")
        except:
            pass
        i.close()
        self.password=""

    def warnUser(self,message,heading = 'Warning',icon = wx.ICON_WARNING):
        wx.MessageBox(message,heading,icon)
        
    def save(self):
        o=open("settings.txt","w")
        o.write(json.dumps(self.settings))
        o.close()



Globals=Globals()


class ServerConnect(wx.Dialog):
    def __init__(self, *args, **kwds):
        # begin wxGlade: ServerConnect.__init__
        kwds["style"] = wx.DEFAULT_DIALOG_STYLE
        wx.Dialog.__init__(self, *args, **kwds)
        self.label_1 = wx.StaticText(self, wx.ID_ANY, "Server")
        self.SERVER = wx.TextCtrl(self, wx.ID_ANY, "")
        self.label_1_copy = wx.StaticText(self, wx.ID_ANY, "Username")
        self.USERNAME = wx.TextCtrl(self, wx.ID_ANY, "")
        self.label_1_copy_1 = wx.StaticText(self, wx.ID_ANY, "Password")
        self.PASSWORD = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_PROCESS_ENTER | wx.TE_PASSWORD)
        self.label_2 = wx.StaticText(self, wx.ID_ANY, "Return Address")
        self.RETURNADDRESS = wx.TextCtrl(self, wx.ID_ANY, "")
        self.button_1 = wx.Button(self, wx.ID_ANY, "Connect")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_TEXT_ENTER, self.connectServer, self.PASSWORD)
        self.Bind(wx.EVT_BUTTON, self.connectServer, self.button_1)
        # end wxGlade
        self.Bind(wx.EVT_CLOSE,self.onClose)
        self.SERVER.Value=Globals.settings["SERVER"]
        self.USERNAME.Value=Globals.settings["USERNAME"]
        self.RETURNADDRESS.Value=Globals.settings["RETURNADDRESS"]
        self.PASSWORD.Value=Globals.password

    def __set_properties(self):
        # begin wxGlade: ServerConnect.__set_properties
        self.SetTitle("Email Connection")
        self.SERVER.SetToolTipString("Server Name")
        self.USERNAME.SetToolTipString("User Name")
        self.PASSWORD.SetToolTipString("Password")
        self.PASSWORD.SetFocus()
        self.RETURNADDRESS.SetToolTipString("Return Address")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: ServerConnect.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_3_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add(self.label_1, 0, 0, 0)
        sizer_2.Add(self.SERVER, 0, 0, 0)
        sizer_1.Add(sizer_2, 0, wx.EXPAND, 0)
        sizer_5.Add(self.label_1_copy, 0, 0, 0)
        sizer_5.Add(self.USERNAME, 0, 0, 0)
        sizer_1.Add(sizer_5, 0, wx.EXPAND, 0)
        sizer_4.Add(self.label_1_copy_1, 0, 0, 0)
        sizer_4.Add(self.PASSWORD, 0, 0, 0)
        sizer_1.Add(sizer_4, 0, wx.EXPAND, 0)
        sizer_3_copy.Add(self.label_2, 0, 0, 0)
        sizer_3_copy.Add(self.RETURNADDRESS, 1, wx.EXPAND, 0)
        sizer_1.Add(sizer_3_copy, 0, wx.EXPAND, 5)
        sizer_1.Add(self.button_1, 0, 0, 0)
        self.SetSizer(sizer_1)
        sizer_1.Fit(self)
        self.Layout()
        # end wxGlade
        
    def onClose(self,event=None):
        self.Destroy()
        if event:
            event.Skip()

    def connectServer(self, event):  # wxGlade: ServerConnect.<event_handler>
        if not self.SERVER.Value or not self.USERNAME.Value or not self.PASSWORD.Value or not self.RETURNADDRESS.Value:
            Globals.warnUser("Missing Connection Information")
            event.Skip()
            return
        try:
            server = smtplib.SMTP(self.SERVER.Value)
        except (socket.gaierror, socket.error):
            Globals.warnUser("Invalid Server")
            event.Skip()
            return
        try:
            server.login(self.USERNAME.Value,self.PASSWORD.Value)
        except smtplib.SMTPAuthenticationError:
            Globals.warnUser("Invalid Login")
            event.Skip()
            return
        return_test = server.docmd('mail from:',self.RETURNADDRESS.Value)
        if return_test[0] == 501:
            server.close()
            Globals.warnUser("Invalid Return Address")
            event.Skip()
            return
        server.close()
        Globals.settings["SERVER"]=self.SERVER.Value
        Globals.settings["USERNAME"]=self.USERNAME.Value
        Globals.settings["RETURNADDRESS"]=self.RETURNADDRESS.Value
        Globals.password = self.PASSWORD.Value
        Globals.save()
        self.Destroy()
        event.Skip()

# end of class ServerConnect
class BudgetMailer(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: BudgetMailer.__init__
        kwds["style"] = wx.CAPTION | wx.CLOSE_BOX | wx.MINIMIZE_BOX | wx.MAXIMIZE_BOX | wx.SYSTEM_MENU | wx.RESIZE_BORDER | wx.FULL_REPAINT_ON_RESIZE | wx.TAB_TRAVERSAL | wx.CLIP_CHILDREN
        wx.Frame.__init__(self, *args, **kwds)
        self.button_2 = wx.Button(self, wx.ID_ANY, "Server Connection")
        self.button_3 = wx.Button(self, wx.ID_ANY, "Send")
        self.CONFIRM = wx.CheckBox(self, wx.ID_ANY, "Confirm Each Email")
        self.label_5 = wx.StaticText(self, wx.ID_ANY, "Budget\nFolder")
        self.button_5 = wx.Button(self, wx.ID_ANY, "Set")
        self.FOLDER = wx.StaticText(self, wx.ID_ANY, "")
        self.label_3 = wx.StaticText(self, wx.ID_ANY, "Subject")
        self.SUBJECT = wx.TextCtrl(self, wx.ID_ANY, "")
        self.EMAIL = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_MULTILINE)
        self.label_4 = wx.StaticText(self, wx.ID_ANY, "Extra\nAttachments")
        self.button_4 = wx.Button(self, wx.ID_ANY, "Add")
        self.ATTACHMENTS = wx.TextCtrl(self, wx.ID_ANY, "")
        self.VIEWER = wx.grid.Grid(self, wx.ID_ANY, size=(1, 1))

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.serverConnection, self.button_2)
        self.Bind(wx.EVT_BUTTON, self.sendEmails, self.button_3)
        self.Bind(wx.EVT_BUTTON, self.setFolder, self.button_5)
        self.Bind(wx.EVT_TEXT, self.saveEmail, self.SUBJECT)
        self.Bind(wx.EVT_TEXT, self.saveEmail, self.EMAIL)
        self.Bind(wx.EVT_BUTTON, self.addAttachment, self.button_4)
        self.Bind(wx.EVT_TEXT, self.saveEmail, self.ATTACHMENTS)
        self.Bind(wx.grid.EVT_GRID_CMD_CELL_CHANGE, self.saveEmail, self.VIEWER)
        # end wxGlade
        self.save=False
        self.SUBJECT.Value=Globals.settings['SUBJECT']
        self.EMAIL.Value=Globals.settings['EMAIL']
        self.ATTACHMENTS.Value=Globals.settings['ATTACHMENTS']
        self.FOLDER.Label = Globals.settings['FOLDER']
        self.loadViewer()
        self.save=True

    def __set_properties(self):
        # begin wxGlade: BudgetMailer.__set_properties
        self.SetTitle("Budget Mailer")
        self.SetSize((707, 599))
        self.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.VIEWER.CreateGrid(0, 3)
        self.VIEWER.SetRowLabelSize(0)
        self.VIEWER.SetColLabelValue(0, "Folder")
        self.VIEWER.SetColLabelValue(1, "Head")
        self.VIEWER.SetColLabelValue(2, "Email")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: BudgetMailer.__do_layout
        sizer_6 = wx.BoxSizer(wx.VERTICAL)
        sizer_9_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7.Add(self.button_2, 0, 0, 0)
        sizer_7.Add(self.button_3, 0, 0, 0)
        sizer_7.Add(self.CONFIRM, 0, 0, 0)
        sizer_6.Add(sizer_7, 0, wx.EXPAND, 0)
        sizer_3.Add(self.label_5, 0, 0, 0)
        sizer_3.Add(self.button_5, 0, 0, 0)
        sizer_3.Add(self.FOLDER, 0, 0, 0)
        sizer_6.Add(sizer_3, 0, wx.EXPAND, 0)
        sizer_8.Add(self.label_3, 0, 0, 0)
        sizer_8.Add(self.SUBJECT, 1, wx.EXPAND, 0)
        sizer_6.Add(sizer_8, 0, wx.EXPAND, 0)
        sizer_6.Add(self.EMAIL, 1, wx.EXPAND, 0)
        sizer_9_copy.Add(self.label_4, 0, 0, 0)
        sizer_9_copy.Add(self.button_4, 0, 0, 0)
        sizer_9_copy.Add(self.ATTACHMENTS, 1, wx.EXPAND, 0)
        sizer_6.Add(sizer_9_copy, 0, wx.EXPAND, 0)
        sizer_6.Add(self.VIEWER, 1, wx.EXPAND, 0)
        self.SetSizer(sizer_6)
        self.Layout()
        # end wxGlade
    
    def loadViewer(self):
        viewer=Globals.settings['VIEWER']
        for k,i in viewer.iteritems():
            self.VIEWER.AppendRows()
            self.VIEWER.SetCellValue(self.VIEWER.GetNumberRows()-1,0,k)
            self.VIEWER.SetCellValue(self.VIEWER.GetNumberRows()-1,1,i[0])
            self.VIEWER.SetCellValue(self.VIEWER.GetNumberRows()-1,2,i[1])
        self.VIEWER.AppendRows()
        self.VIEWER.AutoSizeColumns()
            
    def serverConnection(self, event=None):  # wxGlade: BudgetMailer.<event_handler>
        dia = ServerConnect(None, -1, "",).ShowModal()
        if event:
            event.Skip()
    def sendEmails(self, event):  # wxGlade: BudgetMailer.<event_handler>
        if not Globals.password:
            self.serverConnection()
        if not Globals.password:
            event.Skip()
            return
        if not self.FOLDER.Label or not self.SUBJECT.Value or not self.EMAIL.Value:
            Globals.warnUser("Email not ready to send")
            event.Skip()
            return
        emails=[]
        sentto=[]
        os.chdir(self.FOLDER.Label)
        folders=os.listdir('.')
        total_status_dialog = wx.ProgressDialog ( 'Email Preparation Progress', 'Preparing Emails', maximum = len(folders), style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)

        for folder in folders:
            total_status_dialog.Update(folders.index(folder)+1, "On %s."%folder)
            instructions=Globals.settings['VIEWER'].get(folder,False)
            if instructions:
                if os.path.isdir(folder):
                    files = os.listdir(folder)
                    body = self.EMAIL.Value.replace("|HEAD|",instructions[0])
                    attachments = [folder+'\\'+f for f in files]
                    addresslist = instructions[1].split(",")
                    test = True
                    if self.CONFIRM.Value:
                        dia = wx.MessageDialog(self, "Send email to {} for {}".format(instructions[0],folder),"Confirm Send", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
                        test=dia.ShowModal() == wx.ID_OK
                    if test:
                        emails.append((addresslist,body,attachments,instructions[0]))
                        sentto.append(str(addresslist))
        total_status_dialog.Destroy()
        
        server = smtplib.SMTP(Globals.settings['SERVER'])
        server.starttls()
        server.login(Globals.settings['USERNAME'],Globals.password)
        total_status_dialog = wx.ProgressDialog ( 'Email Sending Progress', 'Sending Emails', maximum = len(emails), style = wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME|wx.STAY_ON_TOP|wx.PD_AUTO_HIDE)
        for email in emails:
            total_status_dialog.Update(emails.index(email)+1, "Sending to %s."%email[3])
            message = MIMEMultipart()
            message['Return-Receipt-To'] = Globals.settings['RETURNADDRESS']
            message['From'] = Globals.settings['RETURNADDRESS']
            message['To'] = ', '.join(email[0])
            message['Date'] = formatdate(localtime=True)
            message['Subject'] = self.SUBJECT.Value
            message.attach(MIMEText(email[1]))
            for each_file in email[2]+self.ATTACHMENTS.Label.split(","):
                attachment = MIMEBase('application', "octet-stream")
                attachment.set_payload(open(each_file,'rb').read())
                Encoders.encode_base64(attachment)
                attachment.add_header('Content-Disposition', 'attachment; filename="%s"'%os.path.basename(each_file))
                message.attach(attachment)
            server.sendmail(Globals.settings['RETURNADDRESS'],email[0],message.as_string())

        total_status_dialog.Destroy()

        message = MIMEMultipart()
        message['Return-Receipt-To'] = Globals.settings['RETURNADDRESS']
        message['From'] = Globals.settings['RETURNADDRESS']
        message['To'] = Globals.settings['RETURNADDRESS']
        message['Date'] = formatdate(localtime=True)
        message['Subject'] = "sent budgets to"
        message.attach(MIMEText('\n'.join(sentto)))
        server.sendmail(Globals.settings['RETURNADDRESS'],Globals.settings['RETURNADDRESS'],message.as_string())
        server.close()

        Globals.warnUser("Done!","Confirmation")
        event.Skip()
        
    def saveEmail(self, event):  # wxGlade: BudgetMailer.<event_handler>
        if self.save:
            Globals.settings["SUBJECT"]=self.SUBJECT.Value
            Globals.settings["EMAIL"]=self.EMAIL.Value
            Globals.settings["ATTACHMENTS"]=self.ATTACHMENTS.Value
            viewer = {}
            for i in range(0,self.VIEWER.GetNumberRows()):
                if self.VIEWER.GetCellValue(i,0):
                    viewer[self.VIEWER.GetCellValue(i,0)]=[self.VIEWER.GetCellValue(i,1),self.VIEWER.GetCellValue(i,2)]
            Globals.settings["VIEWER"]=viewer
            Globals.save()
            if self.VIEWER.GetCellValue(self.VIEWER.GetNumberRows()-1,0):
                self.VIEWER.AppendRows()
            self.save = False
            self.VIEWER.AutoSizeColumns()
            self.save = True
        event.Skip()
    def addAttachment(self, event):  # wxGlade: BudgetMailer.<event_handler>
        dia = wx.FileDialog(None,message = "Select File", wildcard = 'All files (*.*)|*.*', style = wx.OPEN)
        if dia.ShowModal() == wx.ID_OK:
            filepath = dia.GetPaths()[0]
            if ":" not in filepath:
                Globals.warnUser("UNC Paths not supported.  Please use your drives")
                event.Skip()
                return
            if self.ATTACHMENTS.Value:
                self.ATTACHMENTS.Value+=","
            self.ATTACHMENTS.Value+=filepath
        event.Skip()
    def setFolder(self, event):  # wxGlade: BudgetMailer.<event_handler>
        dia = wx.DirDialog(None,message = "Find Current Budget Directory", style = wx.OPEN)
        if dia.ShowModal() == wx.ID_OK:
            self.FOLDER.Label = dia.GetPath()
            if ':' not in self.FOLDER.Label:
                self.FOLDER.Label=""
                Globals.warnUser("UNC Paths not supported")
            Globals.settings['FOLDER']=self.FOLDER.Label
            Globals.save()
        event.Skip()
# end of class BudgetMailer


if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    BudgetMailer = BudgetMailer(None, wx.ID_ANY, "")
    app.SetTopWindow(BudgetMailer)
    BudgetMailer.Show()
    app.MainLoop()
