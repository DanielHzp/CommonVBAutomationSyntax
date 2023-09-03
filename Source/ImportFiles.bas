Attribute VB_Name = "Module5"
Option Explicit

'Import text separated by commas

Sub ImportText()
Dim FileName As String, S() As String, nr As Integer, i As Integer
Dim tWB As Workbook, aWB As Workbook
Set tWB = ThisWorkbook

On Error GoTo here


FileName = Application.GetOpenFilename(fileFilter:="Text Filter(*.txt),*.txt", Title:="Open File", MultiSelect:=False)
Workbooks.Open FileName
Set aWB = ActiveWorkbook


'nonblank rows
nr = WorksheetFunction.CountA(Columns("A:A"))
ReDim S(nr) As String


For i = 1 To nr
    S(i) = Range("A" & i)
Next i
aWB.Close SaveChanges:=False
tWB.Activate



'output any range
Range("A1:A" & nr) = WorksheetFunction.Transpose(S)

Exit Sub
here:
MsgBox "error"

End Sub






'Import from tab delimited txt files
Option Base 1


Sub ImportText2()
Dim FileName As String, S(6, 2) As String, nr As Integer, i As Integer
Dim tWB As Workbook, aWB As Workbook, nr As Integer, nc As Integer, i As Integer, j As Integer
Set tWB = ThisWorkbook


FileName = Application.GetOpenFilename(fileFilter:="Text Filter(*.txt),*.txt", Title:="Open File", MultiSelect:=False)
Workbooks.Open FileName
Set aWB = ActiveWorkbook


nr = 6
nc = 2
For i = 1 To nr
    For j = 1 To nc
    
    S(i, 1) = Range("A" & i)
    S(i, 2) = Range("B" & i)
    
    Next j
Next i



aWB.Close SaveChanges:=False
tWB.Activate

'Print output in a range of cells
Range("A1:B6") = S

End Sub

