VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Message Box"
   ClientHeight    =   1644
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3468
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Application.OnTime nexttime, "MyMsgBoxHide", , False
    'UserForm1.Hide
    uf1.Hide
    Set uf1 = Nothing
End Sub


