Attribute VB_Name = "Module2"
Option Explicit





'Ask how many sheeets user wants to create in the opened workbook

Sub Addandnamesheets()
Dim i As Integer, n As Integer, PersonName As String
Dim SheetName As String


PersonName = InputBox("What is your name?")

n = InputBox("How mane sheets do you want in your workbook")


For i = 1 To n - 1

    Sheets.Add after:=ActiveSheet
    
Next i
For i = 1 To n


    SheetName = InputBox("What is the name of sheet" & i & "?")
    Worksheets(i).Name = SheetName
    Worksheets(i).Range("A1") = PersonName
    
    
    Next i
    
End Sub

