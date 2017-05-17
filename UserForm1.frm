VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Générateur de factures"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10425
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub clientlist_Click()

    On Error GoTo clientlist_Click_Error

    dspCLT.Caption = clientlist.Value
    Const str_sourcefile As String = "C:\Users\p0intz3r0\Desktop\DB_FACTURES_2016.accdb"
    Dim access_db As DAO.Database
    Dim sql_command As String
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.DBEngine.OpenDatabase(str_sourcefile) 'Connection à la BDD '
    sql_command = "SELECT CLTNOM FROM [CLT]"
    Set access_rs = access_db.OpenRecordset(sql_command)
    With clientlist
        .Clear
        Do
            .AddItem access_rs![CLTNOM]
            access_rs.MoveNext
        Loop Until access_rs.EOF
    End With
    sql_command = "SELECT DELAIPAIEMENT FROM [CLT] WHERE CLTNOM = " & Chr(39) & dspCLT.Caption & Chr(39) & ";"  'String SQL de commande : Tout sélectionner de la table factu ATO'
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    txtDELAI.Value = access_rs.Fields(0)


    On Error GoTo 0
    Exit Sub

clientlist_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clientlist_Click of Sub UserForm1"

End Sub
Private Sub UserForm_Activate()

   On Error GoTo UserForm_Activate_Error

 Const str_sourcefile As String = "C:\Users\p0intz3r0\Desktop\DB_FACTURES_2016.accdb"
 Dim access_db As DAO.Database
 Dim sql_command As String
 Dim access_rs As DAO.Recordset
    Set access_db = DAO.DBEngine.OpenDatabase(str_sourcefile) 'Connection à la BDD '
    sql_command = "SELECT DISTINCT CLTNOM FROM [CLT];"
    Set access_rs = access_db.OpenRecordset(sql_command)
       With UserForm1.clientlist
        .Clear
        Do
            .AddItem access_rs![CLTNOM]
            access_rs.MoveNext
        Loop Until access_rs.EOF
        End With
       sql_command = "SELECT DISTINCT NOM FROM [COLLAB];"
    Set access_rs = access_db.OpenRecordset(sql_command)
       With UserForm1.collabBox
        .Clear
        Do
            .AddItem access_rs![NOM]
            access_rs.MoveNext
        Loop Until access_rs.EOF
        End With
        txtDATE.Value = Format(Now(), "dd-mm-yy")

   On Error GoTo 0
   Exit Sub

UserForm_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserForm_Activate of Sub UserForm1"

End Sub
