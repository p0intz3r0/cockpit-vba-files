Attribute VB_Name = "getdata"
Option Explicit
Sub bddfact()
    Application.StatusBar = "Téléchargement des données en cours...."
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim sql_command As String
    Dim bres As Boolean
    Dim str_sourcefile As String
    Dim xlWs As Object
    Dim i As Range
    Dim factpath As String
    Dim n As String
    factpath = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\FACTURES 2016\"
    Dim rangeUrl As Range
    Set xlWs = ActiveWorkbook.Sheets("BDD")
    Set rangeUrl = xlWs.Range("P8:P1000")
    Dim shortn As String
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    shortn = "Voir la facture"
    shortn = Chr(34) & shortn & Chr(34)
    xlWs.UsedRange.ClearContents
    xlWs.Cells(7, 3).Value = "type"
    xlWs.Cells(7, 4).Value = "ID_collab"
    xlWs.Cells(7, 5).Value = "num_facture"
    xlWs.Cells(7, 6).Value = "date_facture"
    xlWs.Cells(7, 7).Value = "mois_facture"
    xlWs.Cells(7, 8).Value = "facture_libelle"
    xlWs.Cells(7, 9).Value = "TJM"
    xlWs.Cells(7, 10).Value = "nb_jours_factu"
    xlWs.Cells(7, 11).Value = "montant_ht"
    xlWs.Cells(7, 12).Value = "montant_ttc"
    xlWs.Cells(7, 13).Value = "is_factor"
    xlWs.Cells(7, 14).Value = "voir"
    xlWs.Cells(7, 15).Value = "client_name"
    Application.Calculation = xlCalculationManual
    'str_sourcefile = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\DB_FACTURES_2016.accdb" 'Absolute path de la BDD '
    str_sourcefile = "C:\Users\p0intz3r0\Desktop\DB_FACTURES_2016.accdb"
    Set access_db = DAO.DBEngine.OpenDatabase(str_sourcefile) 'Connection à la BDD '
    sql_command = "SELECT NUMFACTURE, TYPE, COLLAB, CLIENT, DATEFAC, PERIODE, LIBELLE, TJM, NBJOURS, MONTANTHT, MONTANTTTC, REGLEMENT, REFCLIENT, CLTNOM FROM [FACT] INNER JOIN [CLT] " & _
    "ON [FACT].CLIENT = [CLT].REFCLIENT;" 'String SQL de commande : Tout sélectionner de la table factu ATO'
    Set access_rs = access_db.OpenRecordset(sql_command, dbReadOnly)
    xlWs.Cells(8, 2).CopyFromRecordset access_rs
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    access_db.Close
    For Each i In rangeUrl
        If Not IsEmpty(Cells(i.Row, i.Column - 11)) Then
        n = Cells(i.Row, i.Column - 11).Value
        n = Chr(34) & factpath & n & ".pdf" & Chr(34)
        i.Formula = "=HYPERLINK(" & n & "," & shortn & ")"
        Else
        End If
        Next i
         xlWs.Range("C7:O5000").AutoFilter
            Application.ScreenUpdating = True
            
    Application.Calculation = xlCalculationAutomatic
End Sub
'Sub bddclt()
 'Dim access_db As DAO.Database
   ' Dim access_rs As DAO.Recordset
   ' Dim sql_command As String
  '  Dim bres As Boolean
  '  Dim str_sourcefile As String
  '  Dim xlWs As Object
  '  Set xlWs = ActiveWorkbook.Sheets("BDD Clients")
  '  xlWs.UsedRange.ClearContents
  '  Application.Calculation = xlCalculationManual
   ' str_sourcefile = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\DB_FACTURES_2016.accdb" 'Absolute path de la BDD '
  ''  Set access_db = DAO.DBEngine.OpenDatabase(str_sourcefile) 'Connection à la BDD '
  '  sql_command = "SELECT * FROM [CLT];" 'String SQL de commande : Tout sélectionner de la table factu Client'
   ' Set access_rs = access_db.OpenRecordset(sql_command, dbReadOnly)
   ' Application.StatusBar = "Téléchargement des données en cours...."
   ' xlWs.Cells(5, 1).CopyFromRecordset access_rs
   ' Application.Calculation = xlCalculationAutomatic
   ' Application.StatusBar = False
   ' access_db.Close
'End Sub
