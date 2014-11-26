__author__ = 'Andrew'
import wx, os
from win32com import client
from win32com.client import constants
import datetime
import gettext

class MyFileDropTarget(wx.FileDropTarget):
    """"""
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.count = 0

    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        self.window.notify2("\n%d file(s) dropped at %d,%d:\n" %
                              (len(filenames), x, y))
        #self.testing=filenames
        #print(self.testing)

        for file in filenames:
            self.window.notify(file, self.count, len(filenames))
            self.count = self.count+1



class MyFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.crap = []
        self.spreadsheet=[]
        self.document=[]
        dt1 = MyFileDropTarget(self)
        self.tc_files = wx.TextCtrl(self, wx.ID_ANY, "", style=wx.TE_MULTILINE|wx.HSCROLL|wx.TE_READONLY)
        self.tc_files.SetDropTarget(dt1)
        self.button = wx.Button(self, wx.ID_ANY, _("Make Report"))
        self.button.Bind(wx.EVT_BUTTON, self.onButton)

        self.__set_properties()
        self.__do_layout()

    def __set_properties(self):
        # begin wxGlade: MyFrame.__set_properties
        self.SetTitle(_("Genesis Report Maker"))
        _icon = wx.EmptyIcon()
        dn =  os.path.dirname(__file__)
        iconpath = [str(dn), "\o_90ddbcecced809a8-3.bmp"]
        iconpath = "".join(iconpath)

        _icon.CopyFromBitmap(wx.Bitmap(iconpath, wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)
        self.SetSize((395, 347))
        self.SetBackgroundColour(wx.Colour(50, 153, 204))
                # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MyFrame.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(self.tc_files, 1, wx.EXPAND, 0)
        sizer_1.Add(self.button, 0, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()

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

        dialog = wx.DirDialog(None, "Choose a directory:",style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dialog.ShowModal() == wx.ID_OK:
            lepath = dialog.GetPath()
        dialog.Destroy()

        #for creating a new one: doc = wordApp.Documents.Add()sadsad
        excel = client.Dispatch("Excel.Application")
        word = client.Dispatch("Word.Application") # opening the template file
        #for creating a new one: doc = wordApp.Documents.Add()sadsad
        book = excel.Workbooks.Open(self.spreadsheet)
        dn2 =  os.path.dirname(__file__)
        docpath = [str(dn2), "\Template.docx"]
        docpath = "".join(docpath)
        doc = word.Documents.Open(docpath)
        sheet = book.Worksheets(1)

        #doc.SaveAs("D:\Realty\Template2.docx")
        frview = doc.Bookmarks("frontpic").Range
        frview.InlineShapes.AddPicture(self.crap[0])
        frpic = doc.InlineShapes(1)
        frpic.LockAspectRatio = True
        frpic.Width = 378


        #gets address from excel and puts in word
        certnumber = sheet.Range("G5")
        certnumber = str(certnumber)
        certnumber = certnumber[:-2]
        address = sheet.Range("G7")
        city = sheet.Range("G8")
        state = sheet.Range("G9")
        zip = sheet.Range("G10")
        zip = str(zip)
        zip = zip[:-2]
        fulladd = [str(city), str(state), zip]
        commad = ", ".join(fulladd)
        today = datetime.date.today()
        date = [str(today.month), str(today.day), str(today.year)]
        ledate = " ".join(date)
        lesavepath = [lepath, '\\', str(certnumber), " Property Review - ", str(address), ", ", commad, "_Intermediate ", str(ledate), ".docx"]
        lesave = "".join(lesavepath)
        doc.SaveAs(lesave)
        doc.Bookmarks("front").Range.InsertAfter(address)
        doc.Bookmarks("addyline2").Range.InsertAfter(commad)
        doc.Bookmarks("CertNum").Range.InsertAfter(certnumber)

        #copies property char from excel
        sheet.Range("F4:G19").Copy()
        doc.Bookmarks("Prop").Range.PasteAndFormat(constants.wdFormatOriginalFormatting)
        tbl = doc.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
        doc.Tables(1).Borders.Enable = True

        #copies repair estimate from excel
        sheet.Range("F21:G37").Copy()
        doc.Bookmarks("repair").Range.PasteAndFormat(constants.wdFormatOriginalFormatting)
        tbl2 = doc.Tables(2).Rows.Alignment = constants.wdAlignRowCenter

        #copies pro forma from excel
        sheet.Range("B1:D40").CopyPicture(constants.xlBitmap)
        doc.Bookmarks("proforma").Range.Paste()
        doc.InlineShapes(2).Range.Underline = False
        prfpic = doc.InlineShapes(2)
        prfpic.LockAspectRatio = True
        prfpic.Height = 463.68


 #propertyphotos
        for x,z in enumerate(reversed(self.crap[1:])):
            proppics = doc.Bookmarks("photos").Range
            proppics.InlineShapes.AddPicture(z)
            propphotos = doc.InlineShapes(x+3)
            propphotos.LockAspectRatio = True
            propphotos.Height = 222.48
            propphotos.Range.Underline = False


        doc.Close()


    def notify(self, files, length, maxL):
            self.crap.append(files)
            self.tc_files.WriteText(self.crap[length])
            #if length == (maxL-1):
                #print(self.crap)
    def notify2(self,files):
        self.tc_files.WriteText(files)


    def SetInsertionPointEnd(self):
        self.tc_files.SetInsertionPointEnd()

# testing gitpush123
if __name__ == "__main__":
    gettext.install("app") # replace with the appropriate catalog name

    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frame_1 = MyFrame(None, wx.ID_ANY, "")
    app.SetTopWindow(frame_1)
    frame_1.Show()
    app.MainLoop()