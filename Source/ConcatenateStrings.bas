Attribute VB_Name = "Module2"

Option Explicit
Option Base 1


'Macros that manipulate strings

'String concatenation syntax
Sub joinstrings()
Dim b(3) As String, joinedmsg As String, msg As String
Dim i As Integer
b(1) = "A"
b(2) = "B"
b(3) = "C"
joinedmsg = Join(b)
   For i = 1 To Len(joinedmsg) Step 2
   msg = msg + Mid(joinedmsg, i, 1)
   Next i

End Sub

'Print each component of a string array obtained from a user input text box
Sub SplitAndJoin()
Dim sentence As String, A() As String, i As Integer

sentence = InputBox("please enter sentence.")
A = Split(sentence, " ")
For i = 0 To UBound(A)
    MsgBox A(i)
    
Next i
MsgBox Join(A)
End Sub


'Print each component of a string array obtained from worksheet cells selection
Sub combinecolumn()
Dim b As Object, msg As Variant, i As Integer, nr As Integer
Set b = Selection
nr = Selection.Rows.Count
For i = 1 To nr
    msg = msg + b(i)
    Next i
End Sub

