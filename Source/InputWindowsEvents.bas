Attribute VB_Name = "Module1"
Option Explicit


'Event handling macros
'This macros show how to handle input windows button properties displayed to the user

Sub inputboxtext()
Dim x
x = InputBox("please enter something:", "charlies input", 7)

MsgBox ("you entered" & x & "!")

End Sub

Sub inputboxtest()
Dim x
'type:= 0,1,2,4,8,16,64


Do
    x = Application.InputBox("please enter something", "Charlies Input", 7, Type:=2)
    
    If x <> False And x <> "" Then Exit Do
    
    MsgBox "You didnt enter anything, please try again!"
    
Loop


MsgBox x

End Sub

'buttons :  0 ok only, 1 ok cancel, 2 abortretryignore, 3 yesnocancel , 4 yesno, 5 retrucancel, 16 critical, 32 question, 48 exclamation, 64 information, etc
'msgbox ("prompt", buttons, "title", helpfile, context)

Sub msgboxexamples()
Dim ans As Integer

ans = MsgBox("message any message", 2, "Title of the message")
'return values 1 ok, 2 cancel, 3 abort, retry 4, 5 ignore, 6 yes, 7 no


If ans = 3 Then

MsgBox "You clicked abort!!"

ElseIf ans = 4 Then

MsgBox "youb clicked retry"

Else

MsgBox "you clicked ignore"
End If
End Sub
