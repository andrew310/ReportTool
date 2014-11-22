__author__ = 'Andrew'

from win32com import client
from win32com.client import constants

import gui

excel = client.Dispatch("Excel.Application")
word = client.Dispatch("Word.Application") # opening the template file
#for creating a new one: doc = wordApp.Documents.Add()sadsad


book = excel.Workbooks.Open("C:\Users\Andrew\Documents\Proforma.xlsx")
doc = word.Documents.Open("c:\Users\Andrew\Documents\Template.docx")
sheet = book.Worksheets(1)
doc.SaveAs("C:\Users\Andrew\Documents\Template2.docx")

frview = doc.Bookmarks("frontpic").Range
frview.InlineShapes.AddPicture("C:\Users\Andrew\Documents\exterior.jpg")
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
