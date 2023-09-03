VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Calculator"
   ClientHeight    =   3936
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6072
   OleObjectBlob   =   "UserFormBasicEvents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub CalculateButton_Click()
If multiply Then
    Output = Ainput * Binput
Else
    Output = Ainput / Binput
    End If

End Sub

Private Sub QuitButton_Click()
UserForm1.Hide
'unload userform1
End Sub

Private Sub ResetButton_Click()
Unload UserForm1
UserForm1.Show

End Sub
