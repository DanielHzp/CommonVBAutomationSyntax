Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

'Sample macros used to manipulate strings importing and exporting string files

'THIS IS THE SOLUTION OF MORSE CODE HACKERRANK CHALLENGE

'EXPORT
Sub MorseCoder()
Dim L As Integer, i As Integer, wrd As String, j As Integer
Dim Letters() As String, coded() As String, SaveFile As String

wrd = UCase(Range("E7"))
L = Len(wrd)

ReDim Letters(L) As String
ReDim coded(L) As String


For i = 1 To L
    Letters(i) = Mid(wrd, i, 1)
    For j = 1 To 36
        If Letters(i) = Range("A1:A36").Cells(j, 1).Text Then
            coded(i) = Range("B1:B36").Cells(j, 1).Text
            Exit For
            End If
    Next j
    Range("E7").Offset(i - 1, 1) = coded(i)
    Next i
    
    
    'Export file
SaveFile = Application.GetSaveAsFilename(fileFilter:="Text files (*.txt),*.txt")
  Open SaveFile For Output As #1
  
  'Print output in exported filed
  For i = 1 To L
  Write #1, coded(i)
  Next i
  Close #1
  
End Sub


Sub Reset()
Columns("F:F").Clear
End Sub



Sub test()

Dim numData As Integer
numData = WorksheetFunction.CountA(Range("E8:E25"))
MsgBox numData


End Sub


'IMPORT
Sub MorseDecoder()
Dim FileName As String, tWB As Workbook, aWB As Workbook, coded() As String, Letters() As String
Dim nr As Integer, i As Integer, j As Integer, wrd As String

Set tWB = ThisWorkbook


FileName = Application.GetOpenFilename(fileFilter:="Text Filter(*.txt), *.txt", Title:="Open File", MultiSelect:=False)
Workbooks.Open FileName
Set aWB = ActiveWorkbook



Range("A1").Select
nr = WorksheetFunction.CountA(Columns("A:A"))


ReDim coded(nr) As String, Letters(nr) As String
For i = 1 To nr

    coded(i) = Range("A1:A" & nr).Cells(i, 1)
    Next i
aWB.Close SaveChanges:=False

For i = 1 To nr
    For j = 1 To 36
        If coded(i) = Range("B1:B36").Cells(j, 1).Text Then
        
        Letters(i) = Range("A1:A36").Cells(j, 1).Text
        Exit For
        End If
        Exit For
        Next j
        wrd = wrd + Letters(i)
Next i
        
        
        
    'Display output
    MsgBox wrd
End Sub




