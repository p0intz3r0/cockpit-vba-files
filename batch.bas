Attribute VB_Name = "batch"
Option Explicit
Sub main()
    Application.ScreenUpdating = False
    Dim isfac As Range
    Dim x, y, n As Integer
    Dim TJM, numfa As Double
    Dim collabname, client As String
    Dim joursfactu As Single
    Dim datefact As Date
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim i As Variant
    Dim firstnumfa As Double
    Dim cell As Range
    datefact = #8/31/2016#
    Sheets("BDD Collabs").Activate
    Select Case MsgBox("Etes vous sur de vouloir éditer toutes les factures IT ? ", vbOKCancel Or vbExclamation, Application.Name)
        Case vbCancel
            Exit Sub
        Case vbOK
            StartTime = Timer
            firstnumfa = Sheets("BDD VBA").Range("K5").Value + 1
            Set isfac = Sheets("BDD Collabs").Range("Q2:Q1500")
            For Each i In isfac
                If i = 1 Then
                    i.Select
                    x = i.Row
                    y = i.Column
                    numfa = Sheets("BDD VBA").Range("K5").Value
                    numfa = numfa + 1
                    Cells(x, y + 4).Value = numfa
                    joursfactu = Cells(x, y - 2).Value
                    TJM = Cells(x, y - 7).Value
                    client = Cells(x, y - 11).Value
                    collabname = Cells(x, y - 13).Value
                    Call fonctionsIT.miseenpageIT(TJM, client, datefact, joursfactu, collabname)
                    n = n + 1
                Else
                End If
                Sheets("BDD Collabs").Activate
            Next i
            Application.ScreenUpdating = True
    End Select
    SecondsElapsed = Round(Timer - StartTime, 2)

  '  For Each cell In Sheets("BDD Clients").Range("A1:A100")
     '   x = cell.Row
     '   y = cell.Column
      '  If Sheets("BDD Clients").Cells(x, y + 7).Value = "IT" Then
     '       Call mail.mail(cell.Value, firstnumfa)
   '     End If
 '   Next cell
        MsgBox "Edité " & n & " factures en " & SecondsElapsed & " secondes"
End Sub
