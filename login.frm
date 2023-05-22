VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} login 
   Caption         =   "Login"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cm_cancel_Click()
Unload Me

End Sub

Private Sub cm_login_Click()

If txtuser = "user" And txtpass = "user" Then
MsgBox "Berhasil Login Sebagai User", vbInformation
clear_txt
login.Hide
data.Show
ElseIf txtuser = "admin" And txtpass = "admin" Then
MsgBox "Berhasil Login Sebagai Admin", vbInformation
clear_txt
login.Hide
master_data.Show
Else

MsgBox "Usename / Password Salah !!", vbInformation
clear_txt
End If


End Sub

Sub clear_txt()
txtuser.Value = ""
txtpass.Value = ""
End Sub

Private Sub UserForm_Click()

End Sub
