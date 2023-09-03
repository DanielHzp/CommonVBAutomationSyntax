Attribute VB_Name = "Module4"
Option Explicit



'Export txt files
Sub writefile()
Dim S As String, SaveFile As String
S = InputBox("Please enter a string.")

SaveFile = Application.GetSaveAsFilename(fileFilter:="Text files (*.txt),*.txt")

Open SaveFile For Output As #1

Write #1, S
Close #1



End Sub




'Worksheets Cell selection export
Sub writefile2()
Dim nr As Integer, SaveFile As String, i As Integer
nr = Selection.Rows.Count

SaveFile = Application.GetSaveAsFilename(fileFilter:="Text files (*.txt),*.txt")


Open SaveFile For Output As #1
For i = 1 To nr
    Write #1, Selection.Cells(i, 1)
    Next i
    Close #1
    
    
    
End Sub
