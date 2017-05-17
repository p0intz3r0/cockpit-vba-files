'---------------------------------------------------------------------------------------
' File   : funct17
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: Sous fonctions pour macro 2017
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' File   : fonctions
' Author : p0intz3r0
' Date   : 28/01/2016
' Purpose: Répertoire de fonctions
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Method : margin17
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: Calcul marges
'---------------------------------------------------------------------------------------
Function margin18(i_row_margin)
    Dim ws_plng As Worksheet
    Set ws_plng = ActiveWorkbook.Sheets("Planning 2016")
    Dim IC As Double
    Dim cotisationspat As Double
    Dim OFII As Double
    Dim CP As Double
    Dim transport As Double
    Dim absence As Double
    Dim SBA As Double
    Dim joursato As Single
    Dim CA As Double, TJM As Double
    Dim txMB As Double
    Dim jferies As Range
    Dim dureecontrat As Single
    Dim Mb As Single, mn As Single
    Set jferies = ActiveWorkbook.Sheets("Notice d'utilisation macro").Range("I3:I13")
    Dim startCtr As Date, fincontrat As Date
    startCtr = ws_plng.Cells(i_row_margin, 12).Value
    fincontrat = ws_plng.Cells(i_row_margin, 13).Value
    'Définir dureecontrat, joursato, TJM '
    dureecontrat = DateDiff("m", startCtr, fincontrat)
    joursato = Application.WorksheetFunction.NetworkDays(startCtr, fincontrat, jferies)
    SBA = ws_plng.Cells(i_row_margin, 16).Value
    TJM = ws_plng.Cells(i_row_margin, 17).Value
    IC = 0
    cotisationspat = 0.4482
    OFII = (SBA / 12) * 0
    CP = 0
    transport = 0 * dureecontrat
    absence = 0
    SBA = SBA / 254
    SBA = SBA * joursato
    CA = TJM * ((1 - CP - absence - IC) * joursato)
    If CA = 0 Then
        txMB = 0
    Else
        Mb = TJM * ((1 - CP - absence - IC) * joursato) - OFII - ((1 + cotisationspat) * SBA) - transport
        txMB = Mb / CA
        ' MB = %age '
    End If
    margin = txMB
End Function

'---------------------------------------------------------------------------------------
' Method : get_iterations
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: Calcul nombre de mois (colonnes) reportés
'---------------------------------------------------------------------------------------
Function get_iterations18(ByRef Ncollab As Variant)
    Dim iterations As Integer
    iterations = DateDiff("m", Ncollab.embauche, Ncollab.finmission)
    get_iterations = Month(Ncollab.embauche) + iterations
End Function
'---------------------------------------------------------------------------------------
' Method : get_Ic
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: Calcul IC
'---------------------------------------------------------------------------------------
Function get_Ic18(ByRef Ncollab As Variant)
    Dim IC As Integer
    'Elimination des dates de démarrage antérieures à celles d'embauche '
    If Ncollab.embauche > Ncollab.demarrage Then
        IC = 0
    Else
        'Affectation des dates d'embauche -- '
        IC = DateDiff("d", Ncollab.embauche, Ncollab.demarrage)
    End If
    get_Ic = IC
End Function
'---------------------------------------------------------------------------------------
' Method : year_switch
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: Switch si début contrat < 2017
'---------------------------------------------------------------------------------------
Function year_switch(ByRef Ncollab As Variant)
    If Year(Ncollab.embauche) < 2017 Then
        Ncollab.embauche = #1/1/2018#
    End If
    If Year(Ncollab.demarrage) < 2017 Then
        Ncollab.demarrage = #1/1/2018#
    End If
End Function
'---------------------------------------------------------------------------------------
' Method : reject
' Author : p0intz3r0
' Date   : 12/01/2017
' Purpose: collab rejeté si période <> 2017
'---------------------------------------------------------------------------------------
Function reject18(ByRef Ncollab As Variant)
    Dim rejet As Boolean
    If Year(Ncollab.finmission) < 2018 Then
        rejet = True
    ElseIf Year(Ncollab.embauche) > 2018 Then
        rejet = True
    ElseIf Year(Ncollab.embauche) < 2018 And Year(Ncollab.finmission) < 2018 Then
        rejet = True
    ElseIf (Year(Ncollab.embauche) > 2018 And Year(Ncollab.finmission)) > 2018 Then
        rejet = True
    Else
        rejet = False
    End If
    reject = rejet
End Function
Function report18(ByRef Ncollab As Variant)
    Dim i As Integer
    Dim jferies As Range
    Set jferies = Getjferies
    Dim debutMois As Date, finMois As Date
    Dim count As Byte
    Dim opendays As Integer
    For i = Month(Ncollab.embauche) To Ncollab.Iteration
        If i > 12 And Month(Ncollab.finmission) <> (i - 12) Then
            debutMois = DateSerial(2019, i - 13, 1)
            finMois = DateSerial(2019, i - 12, 0)
        ElseIf Month(Ncollab.embauche) = Month(Ncollab.finmission) And Year(Ncollab.embauche) = Year(Ncollab.finmission) Then
            debutMois = Ncollab.embauche
            finMois = Ncollab.finmission
        ElseIf Month(Ncollab.finmission) = (i - 12) And i > 12 Then
            debutMois = DateSerial(2019, i - 12, 1)
            finMois = Ncollab.finmission
        Else
            debutMois = DateSerial(2018, i, 1)
            finMois = DateSerial(2018, i + 1, 0)
        End If
        If i = Month(Ncollab.embauche) Then
            opendays = Application.WorksheetFunction.NetworkDays(Ncollab.embauche, finMois, jferies)
        ElseIf i = Ncollab.Iteration Then
            opendays = Application.WorksheetFunction.NetworkDays(debutMois, Ncollab.finmission, jferies)
        Else
            opendays = Application.WorksheetFunction.NetworkDays(debutMois, finMois, jferies)
        End If
        Call IC_report(Ncollab, i, opendays)
        Cells(Ncollab.row, i * 5 + 19).Value = (opendays - Cells(Ncollab.row, i * 5 + 17).Value) * Ncollab.TJM
        Cells(Ncollab.row, i * 5 + 20).Value = (opendays - Cells(Ncollab.row, i * 5 + 17).Value) * Ncollab.TJM * Ncollab.margin
        Ncollab.CA = Ncollab.CA + ((opendays - Cells(Ncollab.row, i * 5 + 17).Value) * Ncollab.TJM)
    Next i
End Function

Function IC_report18(ByRef Ncollab As Variant, i As Integer, opendays As Integer)
    Select Case Ncollab.IC
        Case 0 '----> PASS'
        Case Is < opendays, Is = opendays
            Cells(Ncollab.row, i * 5 + 18).Value = Ncollab.IC
            Ncollab.IC = 0
        Case Is > opendays
            Cells(Ncollab.row, i * 5 + 18).Value = opendays
            Ncollab.IC = Ncollab.IC - opendays
    End Select
End Function
 
Function eval_CA18(ByRef Ncollab As Variant)
    Dim jferies As Range
    Set jferies = Getjferies
    If Ncollab.CA <> Application.WorksheetFunction.WorkDay(Ncollab.embauche, Ncollab.finmission, jferies) * Ncollab.TJM Then
        Debug.Print Ncollab.CA / (Application.WorksheetFunction.WorkDay(Ncollab.embauche, Ncollab.finmission,  feries) * Ncollab.TJM); Ncollab.row; ""
    Else
        'Debug.Print "1; "; Ncollab.row
    End If
End Function
 
Function report_WD18()
    Dim gen As Worksheet
    Set gen = ActiveWorkbook.Sheets("Planning 2018")
    Dim i As Integer
    Dim row As Integer
    Dim tpWD As Double
    For i = 1 To 12
        For row = 20 To 1000
            If Not IsEmpty(Cells(row, 5 * i + 19)) Then
                If Not Cells(row, 17).Value = 0 Then
                tpWD = tpWD + Cells(row, 5 * i + 19).Value / Cells(row, 17).Value
                End If
            Else
            End If
        Next row
        Cells(16, i * 5 + 19).Value = tpWD
        tpWD = 0
    Next i
End Function
