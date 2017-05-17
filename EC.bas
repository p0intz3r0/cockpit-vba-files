Attribute VB_Name = "EC"
Option Explicit
Sub EC()
    Dim wsEC As Worksheet
    Dim dDate As Date
    Dim sJournal As String
    Dim strLibelle As String
    Dim dMontantHT As Double
    Dim dMontantTTC As Double
    Dim dDateEch As Date
    Dim iNum As String
    Dim wsDB As Worksheet
    Dim rStrt As Range
    Dim strJournal As String
    Dim x As Double
    Dim y As Double
    Dim strFactType As String
    Dim ilastrow As Integer
    Dim strClient As String
    Dim i As Range
    Set wsEC = ActiveWorkbook.Sheets("EC")
    Set wsDB = ActiveWorkbook.Sheets("Facturation ATO 2016")
    Set rStrt = wsDB.Range("B8:B2500")
    ilastrow = wsEC.Cells(Rows.Count, "B").End(xlUp).Row
    For Each i In rStrt
        If i.Value > 716000 Then
            ilastrow = ilastrow + 1
            iNum = i.Value
            x = i.Row
            y = i.Column
            strClient = wsDB.Cells(x, y + 2).Value
            strLibelle = wsDB.Cells(x, y + 3).Value
            dDate = wsDB.Cells(x, y + 7).Value
            dMontantHT = Abs(wsDB.Cells(x, y + 8).Value)
            dMontantTTC = Abs(wsDB.Cells(x, y + 9).Value)
            dDateEch = dDate + wsDB.Cells(x, y + 10).Value
            strClient = Replace(strClient, " ", "")
            strClient = Replace(strClient, Chr(45), "")
            strClient = Replace(strClient, Chr(39), "")
            strClient = "C" + Left(strClient, 11)
            Debug.Print strClient
            If Left(wsDB.Cells(x, y + 1), 1) = "F" Then
                wsEC.Cells(ilastrow, 1) = strClient
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 5) = dMontantTTC
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
                ilastrow = ilastrow + 1
                wsEC.Cells(ilastrow, 1) = 70660400
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 6) = dMontantHT
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
                ilastrow = ilastrow + 1
                wsEC.Cells(ilastrow, 1) = 44571200
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 6) = dMontantTTC - dMontantHT
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
            Else
                 wsEC.Cells(ilastrow, 1) = strClient
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 6) = dMontantTTC
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
                ilastrow = ilastrow + 1
                wsEC.Cells(ilastrow, 1) = 70660400
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 5) = dMontantHT
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
                ilastrow = ilastrow + 1
                wsEC.Cells(ilastrow, 1) = 44571200
                wsEC.Cells(ilastrow, 2) = dDate
                wsEC.Cells(ilastrow, 3) = "VE"
                wsEC.Cells(ilastrow, 4) = strLibelle
                wsEC.Cells(ilastrow, 5) = dMontantTTC - dMontantHT
                wsEC.Cells(ilastrow, 7) = dDateEch
                wsEC.Cells(ilastrow, 8) = iNum
            End If
        End If
    Next i
End Sub
