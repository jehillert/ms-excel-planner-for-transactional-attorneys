Attribute VB_Name = "File"
Sub saveIfNotSaved()
   If ActiveWorkbook.Saved = False Then ActiveWorkbook.Save
   End Sub
