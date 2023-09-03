Attribute VB_Name = "Module3"
Option Explicit


'Iterate over string arrays and create a string output
Function Email(S As Range) As String
Dim L As Integer, startLoc As Integer, EndLoc As Integer
Dim atLoc As Integer, i As Integer
L = Len(S)
atLoc = InStr(S, "@")
For i = atLoc - 1 To 1 Step -1
    If Mid(S, i, 1) = ":" Or Mid(S, i, 1) = "<" Or Mid(S, i, 1) = " " Then
    startLoc = i + 1
    Exit For
    ElseIf i = 1 Then
    startLoc = 1
    End If
Next i
For i = atLoc + 1 To L
    If Mid(S, i, 1) = "]" Or Mid(S, i, 1) = ">" Then
    EndLoc = i - 1
    Exit For
    ElseIf i = L Then
    EndLoc = L
    End If
    Next i
    Email = Mid(S, startLoc, EndLoc - startLoc + 1)

End Function






'Separate strings into component parts
Option Base 1

Function parts(S As String) As Variant
Dim L As Integer, firstdash As Integer, seconddash As Integer
Dim p(3) As Variant
L = Len(S)
firstdash = InStr(1, S, "-")
seconddash = InStr(firstdash + 1, S, "-")
p(1) = Left(S, firstdash - 1)
p(2) = Mid(S, firstdash + 1, seconddash - firstdash - 1)
p(3) = Right(S, L - seconddash)
parts = p
End Function
