Attribute VB_Name = "Module1"
Option Explicit
Option Base 1


'This macro show how to unify data of SEPARATE INDEPENDENT EXCEL WORKBOOKS in a single consolidated excel file
Sub ConsolidateSchedulesWorkbooks()

Dim w As Worksheet, i As Integer, j As Integer
Dim Folder As String, FileName As String
Dim aWB As Workbook, tWB As Workbook
Dim nr As Integer, nc As Integer
Dim S() As Integer, mx As Integer, c As Range

Application.ScreenUpdating = False
Set tWB = ThisWorkbook
nr = 9
nc = 5
ReDim S(nr, nc)

'Folder path that contain all workbooks to consolidate in a single file
Folder = "C:\FolderPath\Employee Schedules"
FileName = Dir(Folder & "\*.xlsx")


'Loop over all workbooks located inside a specified folder
Do
    Workbooks.Open Folder & "\" & FileName
    
    Set aWB = ActiveWorkbook
    
    'Navigate inside each workbook iterated
For Each w In Worksheets
    For i = 1 To nr
        For j = 1 To nc
            If w.Range("B4:F12").Cells(i, j) = "X" Then
            S(i, j) = S(i, j) + 1
            End If
            Next j
            Next i
            Next
            aWB.Close SaveChanges:=False
            FileName = Dir
            
Loop Until FileName = ""



tWB.Activate

'Print output in cell range
Worksheets("Summary").Range("B4:F12") = S
mx = WorksheetFunction.Max(S)



'highlight cell green
For Each c In Worksheets("Summary").Range("B4:F12")
    If c.Value = mx Then
    With c.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
     End With
    Else
    With c.Interior
        .Pattern = xlNone
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    End If
    
    
Next


End Sub



'Clean data output printed
Sub reset()
Dim c As Range

For Each c In Worksheets("Summary").Range("B4:F12")
c.Clear
With c.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
    End With
    Next
    
End Sub












