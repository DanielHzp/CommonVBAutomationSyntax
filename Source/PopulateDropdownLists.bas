Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

' FILL COMBOBOX DROPDOWN OPTIONS

Sub PopulateComboBox1()

'Fill in code here
Dim Names() As String, n As Integer, i As Integer
n = WorksheetFunction.CountA(Columns("A:A"))
ReDim Names(n) As String


For i = 1 To n
    Names(i) = Range("A1:A" & n).Cells(i, 1)
    
    'Dynamically populate dropdown list
    UserForm1.ComboBox1.AddItem Names(i)
    
    Next i
    
    'Default value
    
UserForm1.ComboBox1.Text = Names(1)
    
End Sub

Sub RunForm1()
Call PopulateComboBox1
UserForm1.Show
End Sub
