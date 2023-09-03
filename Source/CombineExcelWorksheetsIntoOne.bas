Attribute VB_Name = "Module3"
Option Explicit
Option Base 1



'Combine data located in multiple WORKSHEETS into one worksheet

Sub consolidateData()

Dim A() As String, i As Integer, n As Integer

n = Worksheets.Count


ReDim A(n - 1) As String


For i = 2 To n

    A(i - 1) = Worksheets(i).Range("A1")
    
Next i


Worksheets(1).Range("A1:A" & n - 1) = WorksheetFunction.Transpose(A)

End Sub









Sub CountSevens()
Dim nw As Integer, i As Integer, j As Integer, k As Integer, nc As Integer
Dim nr As Integer, c As Integer
nr = 6
nc = 4
nw = Worksheets.Count
For i = 2 To nw
    For j = 1 To nr
        For k = 1 To nc
        If Worksheets(i).Range("C5:F10").Cells(j, k) = 7 Then
        c = c + 1
        End If
    
        Next k
        Next j
                Next i
                
MsgBox c

End Sub
