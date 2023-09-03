Attribute VB_Name = "Module2"
Option Explicit

Sub deleterow()
Attribute deleterow.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Deleterow Macro
'

'
    Rows("9:9").Select
    Selection.Delete Shift:=xlUp
    
    
End Sub
