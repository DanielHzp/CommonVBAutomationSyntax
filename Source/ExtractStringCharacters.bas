Attribute VB_Name = "Module1"
Option Explicit


'This macro iterates over an input string and extracts each string character
'The code looks for a particular character in a sample string provided by the user
Sub SearchForString()
'Place your code here
Dim nr As Integer, nc As Integer, i As Integer, j As Integer, k As Integer
Dim str As String, s As Integer, g As Integer, wrd As String, switch As Boolean, ws As Integer
Dim W(), rowindex(), colindex()

nr = Selection.Rows.Count
nc = Selection.Columns.Count
On Error GoTo here



str = InputBox("Please enter string to search for in selection:")
s = Len(str)



For i = 1 To nr
    For j = 1 To nc
    
    wrd = Selection.Cells(i, j)
    ws = Len(wrd)
    switch = False
    
    
    For g = 1 To ws - s + 1
    
    If Mid(wrd, g, s) = str Then
    
    switch = True
    
    End If
    Next g
    
    If switch = True Then
    
        k = k + 1
        ReDim Preserve W(k), rowindex(k), colindex(k)
        
    W(k) = Selection.Cells(i, j)
    
    rowindex(k) = i
    colindex(k) = j
    
    End If
    Next j
    Next i


For i = 1 To UBound(W)

'Dynamically Print SEARCH result in cells
Range("E" & i) = W(i)
Range("F" & i) = rowindex(i)
Range("G" & i) = colindex(i)
Next i

Exit Sub


here:
MsgBox "an error has occured"


End Sub

