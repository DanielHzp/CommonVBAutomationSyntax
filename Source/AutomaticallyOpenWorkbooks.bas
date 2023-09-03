Attribute VB_Name = "Module5"
Option Explicit



'Macros used to automatically open workbooks located in fixed folder paths


Sub OpenStaticFile()

Dim FileName As String

FileName = "C:\Users\Documents\GUIDE WORKSHEET FILES\Sample Files\File_1.xlsx"

Workbooks.Open FileName

End Sub



Sub OpenStaticFiles()

Dim FileNames(2) As String, i As Integer

FileNames(1) = "C:\Users\Documents\GUIDE WORKSHEET FILES\Sample Files\File_1.xlsx"

FileNames(2) = "C:\Users\Documents\GUIDE WORKSHEET FILES\Sample Files\File_2.xlsx"

For i = 1 To UBound(FileNames)
    Workbooks.Open FileNames(i)
Next i

End Sub


'Manually let the user choose the workbook to open
Sub OpenUserFiles()
Dim i As Integer
Dim FileNames As Variant
FileNames = Application.GetOpenFilename(FileFilter:="Excel Filter (*.xlsx),*.xlsx", Title:="Open Files", MultiSelect:=True)
Workbooks.Open FileNames
For i = 1 To UBound(FileNames)
    Workbooks.Open FileNames(i)
    Next i
    
End Sub



'Automatically open all workbooks located in the same folder
Sub OpenAllFilesInFolder()

Dim Folder As String, FileName As String

Folder = "C:\Users\Documents\GUIDE WORKSHEET FILES\Sample Files"

FileName = Dir(Folder & "\*.xlsx")

'In this loop, the workbook file names are iterated in order to only open those that start with 'F'
Do

    If Left(FileName, 1) = "F" Then
    
    Workbooks.Open Folder & "\" & FileName
    
    End If
    FileName = Dir
    
Loop Until FileName = ""
    

End Sub








Sub ImportDataworkbooks()
Dim FileNames() As Variant, nw As Integer
Dim i As Integer, A() As Variant
Dim tWB As Workbook, aWB As Workbook
Set tWB = ThisWorkbook


FileNames = Application.GetOpenFilename(FileFilter:="Excel Filter (*.xlsx),*.xlsx", Title:="Open file(s)", MultiSelect:=True)
Application.ScreenUpdating = False

nw = UBound(FileNames)


ReDim A(nw)


For i = 1 To nw

    Workbooks.Open FileNames(i)
    Set aWB = ActiveWorkbook
    A(i) = aWB.Sheets("Sheet1").Range("A1")
    aWB.Close SaveChanges:=False
    Next i
    tWB.Activate
    
    'range A3 through A
    tWB.Sheets("Main").Range("A3:A" & nw + 2) = WorksheetFunction.Transpose(A)
    
End Sub
