Attribute VB_Name = "mail"
Option Explicit
'---------------------------------------------------------------------------------------
' Method : mail
' Author : p0intz3r0
' Date   : 01/09/2016
' Purpose: Création brouillon outlook des factures à partir de strtfacrt selon client
'---------------------------------------------------------------------------------------
Sub mail(refclt, strtfact)
   On Error GoTo mail_Error
    Dim olapp As Outlook.Application
    Set olapp = New Outlook.Application
    Dim objMail As Object
    Dim objRcpt As Object
    Dim objPJ As Outlook.Attachments
    Dim strContact As String
    Dim wsDB As Worksheet
    Dim rnNum As Range
    Dim rn As Range
    Dim x As Integer, y As Integer
    Dim wsNew As Worksheet
    Dim ilastrow As Integer
    Const factpath = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\FACTURES 2016\"
    Set wsDB = ActiveWorkbook.Sheets("BDD")
    Set rnNum = wsDB.Range("B8:B3000")
    Set wsNew = Sheets.Add
    wsNew.Name = refclt
    ilastrow = wsNew.Cells(Rows.Count, "B").End(xlUp).Row
    For Each rn In rnNum
        If rn.Value > strtfact Then
            y = rn.Row
            x = rn.Column
            If wsDB.Cells(y, x + 3).Value = refclt Then
                wsNew.Cells(ilastrow, 1) = rn.Value
                wsNew.Cells(ilastrow, 2) = wsDB.Cells(y, x + 10).Value
                wsNew.Cells(ilastrow, 3) = Application.VLookup(wsDB.Cells(y, x + 2).Value, Sheets("BDD MATRICULES").Range("A1:B1000"), 2, 0)
                ilastrow = ilastrow + 1
            End If
        End If
    Next rn
    strContact = Application.VLookup(refclt, Sheets("BDD Clients").Range("A5:I100"), 7, 0)
    If IsEmpty(wsNew.Cells(1, 1)) Then
    Else
    If IsError(strContact) Or strContact = "" Then
        MsgBox ("Adresse e-mail introuvable !")
        Exit Sub
    Else
        Set objMail = olapp.CreateItem(0)
        With objMail
            .To = strContact
            .Subject = "Facturation Cooptalis " & Format((Now() - 1), "mmmm") & ""
            .Body = "Bonjour, veuillez trouver ci-joint le(s) facture(s) du mois d'" & Format((Now() - 1), "mmmm") & " pour l'assistance technique de nos collaborateurs. Cordialement,"
            For Each rn In wsNew.Range("A1:A1000")
                If rn.Value <> "" Then
                    .Attachments.Add factpath & rn.Value & ".pdf"
                End If
            Next rn
            .Save
        End With
    End If
    End If
   On Error GoTo 0
   Exit Sub
mail_Error:
    MsgBox "Erreur " & Err.Number & " (" & Err.Description & ") dans la procedure mail() de Sub mail"
End Sub
Sub automail()
Dim cell As Range
'[p l;['l '; l;'l ';l'
Dim x As Integer, y As Integer
Const firstnumfa = 716132
 For Each cell In Sheets("BDD Clients").Range("A1:A100")
        x = cell.Row
        y = cell.Column
        If Sheets("BDD Clients").Cells(x, y + 7).Value = "IT" Then
            Call mail(cell.Value, firstnumfa)
        End If
    Next cell
    MsgBox "ok"
End Sub
