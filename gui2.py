__author__ = 'Andrew'
import wx, os

class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.count = 0

    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        #self.window.notify("\n%d file(s) dropped at %d,%d:\n" %
                              #(len(filenames), x, y))
        #self.testing=filenames
        #print(self.testing)

        for file in filenames:
            self.window.notify(file + '\n', self.count, len(filenames))
            self.count = self.count+1



class MyFrame(wx.Frame):
    def __init__(self, parent,id):
        wx.Frame.__init__(self,parent,id,'Property Report Maker', size=(300,200))
        self.crap = []
        dt1 = MyFileDropTarget(self)
        self.tc_files = wx.TextCtrl(self, -1, "", style = wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
        self.tc_files.SetDropTarget(dt1)


    def notify(self, files, length, maxL):
        self.crap.append(files)
        self.tc_files.WriteText(self.crap[length])
        if length == (maxL-1):
            print(self.crap)




    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# testing gitpush123
if __name__=='__main__':
    app=wx.PySimpleApp()
    frame=MyFrame(parent=None,id=-1)
    frame.Show()
    app.MainLoop()__author__ = 'Andrew'
import wx, os

class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.count = 0

    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        #self.window.notify("\n%d file(s) dropped at %d,%d:\n" %
                              #(len(filenames), x, y))
        #self.testing=filenames
        #print(self.testing)

        for file in filenames:
            self.window.notify(file + '\n', self.count, len(filenames))
            self.count = self.count+1



class MyFrame(wx.Frame):
    def __init__(self, parent,id):
        wx.Frame.__init__(self,parent,id,'Property Report Maker', size=(300,200))
        panel = wx.Panel(self, wx.ID_ANY,)
        self.crap = []
        dt1 = MyFileDropTarget(self)
        self.tc_files = wx.TextCtrl(self, -1, "", style = wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
        self.tc_files.SetDropTarget(dt1)

        button = wx.Button(panel, id=wx.ID_ANY, label="Press Me")
        button.Bind(wx.EVT_BUTTON, self.onButton)

        def onButton(self, event):

            print "Button pressed!"


    def notify(self, files, length, maxL):
        self.crap.append(files)
        self.tc_files.WriteText(self.crap[length])
        if length == (maxL-1):
            print(self.crap)




    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# testing gitpush123
if __name__=='__main__':
    app=wx.PySimpleApp()
    frame=MyFrame(parent=None,id=-1)
    frame.Show()
    app.MainLoop()