VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MailForm 
   Caption         =   "Envoi mail"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MailForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton1_Click()
Dim strtfact As Double
Dim refclt As Single
refclt = LcltMail.Value
strtfact = lmtfac.Value
Call mail.mail(refclt, strtfact)
MailForm.Hide
End Sub
