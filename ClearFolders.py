#!/usr/bin/env python
# -*- coding: ISO-8859-15 -*-
#
# generated by wxGlade 0.6.8 (standalone edition) on Tue Nov 12 07:55:06 2013
#

import wx,os,shutil

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode
# end wxGlade


class ClearFolders(wx.Frame):
	def __init__(self, *args, **kwds):
		# begin wxGlade: ClearFolders.__init__
		kwds["style"] = wx.DEFAULT_FRAME_STYLE
		wx.Frame.__init__(self, *args, **kwds)
		self.button_1 = wx.Button(self, wx.ID_ANY, "Set Directory")
		self.TARGET = wx.StaticText(self, wx.ID_ANY, "label_1")
		self.button_2 = wx.Button(self, wx.ID_ANY, "Clear Old Reports")

		self.__set_properties()
		self.__do_layout()

		self.Bind(wx.EVT_BUTTON, self.setTarget, self.button_1)
		self.Bind(wx.EVT_BUTTON, self.clearOldReports, self.button_2)
		# end wxGlade

	def __set_properties(self):
		# begin wxGlade: ClearFolders.__set_properties
		self.SetTitle("Clear Folders")
		self.SetSize((400, 112))
		# end wxGlade

	def __do_layout(self):
		# begin wxGlade: ClearFolders.__do_layout
		sizer_1 = wx.BoxSizer(wx.VERTICAL)
		sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
		sizer_2.Add(self.button_1, 0, 0, 0)
		sizer_2.Add(self.TARGET, 0, 0, 0)
		sizer_1.Add(sizer_2, 0, wx.EXPAND, 0)
		sizer_1.Add(self.button_2, 0, 0, 0)
		self.SetSizer(sizer_1)
		self.Layout()
		# end wxGlade

	def onClose(self,event = ''):
		if event:
			event.Skip()
		self.Destroy()

	def setTarget(self, event): # wxGlade: ClearFolders.<event_handler>
		dia = wx.DirDialog(None,message = "Find Accounting Directory to Prepare", style = wx.OPEN)
		if dia.ShowModal() == wx.ID_OK:
			self.TARGET.Label = dia.GetPath()
		event.Skip()

	def clearOldReports(self, event): # wxGlade: ClearFolders.<event_handler>
		src = self.TARGET.Label
		dest = '%s%s'%(src,' backup')
		if not os.path.isdir(src):
			wx.MessageBox('Target directory is not a valid directory\nExiting')
			return
		if os.path.isdir(dest):
			wx.MessageBox('Backup directory already exists\nExiting')
			return
		totalstatusdialog = wx.ProgressDialog( 'Progress', 'Creating Backup1111111111111111111111111111111111111')
		totalstatusdialog.Update (0,"Creating Backup")
		shutil.copytree(src,dest)
		totalstatusdialog.Update (0,"Deleting old files")
		for i in os.listdir(src):
			if os.path.isdir('%s\\%s'%(src,i)):
				totalstatusdialog.Update (0,"Deleting old files from %s"%i)
				for a in os.listdir('%s\\%s'%(src,i)):
					if os.path.isdir('%s\\%s\\%s'%(src,i,a)):
						try:
							shutil.rmtree('%s\\%s\\%s'%(src,i,a))
						except WindowsError:
							pass
					else:
						try:
							os.remove('%s\\%s\\%s'%(src,i,a))
						except WindowsError:
							pass
		totalstatusdialog.Destroy()
		event.Skip()

# end of class ClearFolders
if __name__ == "__main__":
	app = wx.PySimpleApp(0)
	wx.InitAllImageHandlers()
	ClearFolders = ClearFolders(None, wx.ID_ANY, "")
	app.SetTopWindow(ClearFolders)
	ClearFolders.Show()
	app.MainLoop()
