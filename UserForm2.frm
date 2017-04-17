VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Question"
   ClientHeight    =   1632
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   5556
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Application.OnTime nexttime, "MyQuestionBoxHide", False

    uf2.Hide
    confirmation = uf2.CommandButton1.Caption
    Set uf2 = Nothing
End Sub

Private Sub CommandButton2_Click()
    Application.OnTime nexttime, "MyQuestionBoxHide", False

    uf2.Hide
    confirmation = uf2.CommandButton2.Caption
    Set uf2 = Nothing
End Sub


