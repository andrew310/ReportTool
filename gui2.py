__author__ = 'Andrew'
import wx, os
from win32com import client
from win32com.client import constants

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
            self.window.notify(file, self.count, len(filenames))
            self.count = self.count+1



class MyFrame(wx.Frame):
    def __init__(self, parent,id):
        wx.Frame.__init__(self,parent,id,'Property Report Maker', size=(300,200))
        panel = wx.Panel(self, wx.ID_ANY,)
        self.crap = []
        self.spreadsheet=[]
        self.document=[]
        dt1 = MyFileDropTarget(self)
        self.tc_files = wx.TextCtrl(self, -1, "", style = wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
        self.tc_files.SetDropTarget(dt1)

        button = wx.Button(panel, id=wx.ID_ANY, label="Press Me")
        button.Bind(wx.EVT_BUTTON, self.onButton)

    def onButton(self, event):
        self.tc=0
        for i,v in enumerate(self.crap):
            if self.crap[i].endswith('.xlsx'):
                self.spreadsheet=v
                del self.crap[i]
        for i,v in enumerate(self.crap):
            if self.crap[i].endswith('.docx'):
                self.document=v
                del self.crap[i]
        makeReport(self.spreadsheet, self.document, self.crap)

    def makeReport(self,xlsx, docx, picture ):
        xcel = client.Dispatch("Excel.Application")
        word = client.Dispatch("Word.Application") # opening the template file
        #for creating a new one: doc = wordApp.Documents.Add()sadsad
        book = excel.Workbooks.Open(xlsx)
        doc = word.Documents.Open(docx)
        sheet = book.Worksheets(1)
        doc.SaveAs("C:\Users\Andrew\Documents\Template2.docx")
        frview = doc.Bookmarks("frontpic").Range
        frview.InlineShapes.AddPicture(picture)
        frpic = doc.InlineShapes(1)
        frpic.LockAspectRatio = True
        frpic.Width = 378

        #gets address from excel and puts in word
        address = sheet.Range("G7")
        city = sheet.Range("G8")
        state = sheet.Range("G9")
        zip = sheet.Range("G10")
        fulladd = [str(city), str(state), str(zip)]
        commad = ", ".join(fulladd)
        doc.Bookmarks("front").Range.InsertAfter(address)
        doc.Bookmarks("addyline2").Range.InsertAfter(commad)

        #copies property char from excel
        sheet.Range("F4:G20").Copy()
        doc.Bookmarks("Prop").Range.PasteAndFormat(constants.wdFormatOriginalFormatting)
        tbl = doc.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
        doc.Tables(1).Rows(13).Delete()
        doc.Tables(1).Rows(16).Delete()
        doc.Tables(1).Borders.Enable = True

        #copies repair estimate from excel
        sheet.Range("F21:G37").Copy()
        doc.Bookmarks("repair").Range.PasteAndFormat(constants.wdFormatOriginalFormatting)
        tbl2 = doc.Tables(2).Rows.Alignment = constants.wdAlignRowCenter

        #copies pro forma from excel
        sheet.Range("B1:D40").CopyPicture(constants.xlBitmap)
        doc.Bookmarks("proforma").Range.Paste()
        prfpic = doc.InlineShapes(2)
        prfpic.LockAspectRatio = True
        prfpic.Height = 463.68


    def notify(self, files, length, maxL):
            self.crap.append(files)
            self.tc_files.WriteText(self.crap[length])
            #if length == (maxL-1):
                #print(self.crap)

    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# testing gitpush123
if __name__=='__main__':
    app=wx.PySimpleApp()
    frame=MyFrame(parent=None,id=-1)
    frame.Show()
    app.MainLoop()