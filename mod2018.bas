Option Explicit

'---------------------------------------------------------------------------------------
' File   : mod2016
' Author : p0intz3r0
' Date   : 28/01/2016
' Purpose: macro de report sur 2017
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Method : report16
' Author : p0intz3r0
' Date   : 28/01/2016
' Purpose: Fonction principale
'---------------------------------------------------------------------------------------
Sub report18()
    Application.Calculation = xlCalculationManual
    Call clean18 'A update '
    Dim wsPlanning As Worksheet
    Set wsPlanning = ActiveWorkbook.Sheets("Planning 2017")
    Dim startRng As Range
    Dim cell As Range
    Set startRng = wsPlanning.Range("G20:G1500")
    Dim rowId As Integer, colId As Integer
    Dim Ncollab As Variant
    Set Ncollab = New Collab
    Dim tmpIteration As Integer
    Dim tmpIC As Double
    For Each cell In startRng
        If Not IsEmpty(cell) Then
            Ncollab.row = cell.row '-- Affectation des variables selon la table -- '
            Ncollab.embauche = wsPlanning.Cells(cell.row, cell.Column + 4).Value
            Ncollab.demarrage = wsPlanning.Cells(cell.row, cell.Column + 5).Value
            Ncollab.finmission = wsPlanning.Cells(cell.row, cell.Column + 6).Value
            Ncollab.TJM = wsPlanning.Cells(cell.row, cell.Column + 10).Value
            Ncollab.margin = wsPlanning.Cells(cell.row, cell.Column + 11).Value
            Ncollab.SBA = wsPlanning.Cells(cell.row, cell.Column + 9).Value
            If reject(Ncollab) = False Then '-- Fonction reject pour check dates s/ 2017 -- '
                Call year_switch(Ncollab) '-- Fonction
                tmpIteration = get_iterations(Ncollab)
                Ncollab.Iteration = tmpIteration
                tmpIC = get_Ic(Ncollab)
                Ncollab.IC = tmpIC
                Cells(Ncollab.row, 14).Value = Ncollab.IC
                Cells(Ncollab.row, 15).Value = DateDiff("m", Ncollab.embauche, Ncollab.finmission)
                Call funct17.report(Ncollab)
                Cells(Ncollab.row, 19).Value = Ncollab.CA
                Cells(Ncollab.row, 20).Value = Ncollab.CA * Ncollab.margin
                Call eval_CA(Ncollab)
                Ncollab.CA = 0
            ElseIf reject(Ncollab) = True Then
            Cells(Ncollab.row, 14).Value = 0
            End If
        End If
    Next cell
    Call report_WD
    Application.Calculation = xlCalculationAutomatic
wsPlanning.Range("R20").Select
End Sub

