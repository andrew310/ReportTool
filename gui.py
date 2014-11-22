__author__ = 'Andrew'
import wx, os

class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        self.window.notify("\n%d file(s) dropped at %d,%d:\n" %
                              (len(filenames), x, y))
        print(filenames)
        for file in filenames:
            self.window.notify(file + '\n')


class MyFrame(wx.Frame):
    def __init__(self, parent,id):
        wx.Frame.__init__(self,parent,id,'Property Report Maker', size=(300,200))
        dt1 = MyFileDropTarget(self)
        self.tc_files = wx.TextCtrl(self, -1, "", style = wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)

        self.tc_files.SetDropTarget(dt1)


    def notify(self, files):
        self.tc_files.WriteText(files)

    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# testing gitpush123
if __name__=='__main__':
    app=wx.PySimpleApp()
    frame=MyFrame(parent=None,id=-1)
    frame.Show()
    app.MainLoop()