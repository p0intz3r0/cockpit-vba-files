Attribute VB_Name = "fonctionsIT"
Public Function miseenpageIT(TJM, client, datefact, joursfactu, collabname)
    Dim isfactor As Byte, facturnumber As Double, typefact As String, ato As String, totalHT As Single, totalTTC As Single, delairglt As Integer, gen As Worksheet
    ato = "ATO 08/16"
    Sheets("GENERATEUR ATO").Activate
    ActiveSheet.Range("I19") = client
    ActiveSheet.Range("J14") = num
    ActiveSheet.Range("J13") = datefact
    ActiveSheet.Range("D34") = collabname
    ActiveSheet.Range("A34") = ato
    ActiveSheet.Range("F34") = joursfactu
    ActiveSheet.Range("C48") = collabname
    ActiveSheet.Range("H34") = TJM
    ActiveSheet.Range("J15") = "IT"
    facturnumber = Sheets("BDD VBA").Range("K5")
    facturnumber = facturnumber + 1
    Sheets("BDD VBA").Range("K5") = facturnumber
    ActiveSheet.Range("J14") = facturnumber
'-- RechercheV pour coordonnées client --'
    ActiveSheet.Range("I20") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 2, 0)
    ActiveSheet.Range("I21") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 3, 0)
    ActiveSheet.Range("I22") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 4, 0)
    ActiveSheet.Range("I23") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 5, 0)
    ActiveSheet.Range("I24") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 6, 0)
    ActiveSheet.Range("I25") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 9, 0)
    ActiveSheet.Range("C48") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 10, 0)
    ActiveSheet.Range("J16") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:N100"), 8, 0)
    sngclient = Application.VLookup(client, Sheets("BDD Clients").Range("B5:N100"), 8, 0)
    isfactor = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 11, 0)
    totalHT = ActiveSheet.Range("J40").Value
    totalTTC = ActiveSheet.Range("J44").Value
    delairglt = ActiveSheet.Range("C48").Value
   If isfactor = 1 Then
    'ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("A5")
   ' typefact = "Facture Factor CIC"
    ' ActiveSheet.Range("A53") = ""
    ElseIf isfactor = 2 Then
    ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("K1")
   'ActiveSheet.Range("A53") = "ATTENTION : changement de coordonnées bancaires à compter du 30/04"
    'ActiveSheet.Range("A53").Font.Color = vbRed
   ' ActiveSheet.Range("A53").Font.Bold = True
    ' ActiveSheet.Range("A53").Font.Size = 20
     typefact = "Facture Factor NATIXIS"
    '-- '
    Call fonctionsIT.csvnatixis(totalHT, totalTTC, datefact, facturnumber, delairglt, client, typefact)
       Call fonctionsIT.reportsurDBIT(TJM, sngclient, datefact, joursfactu, collabname, facturnumber, typefact, totalHT, totalTTC, delairglt, ato, isfactor)
       Call fonctionsIT.exportEnPdfIT(facturnumber)
    Else
        ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("A1")
        typefact = "Facture Directe"
         ActiveSheet.Range("A53") = ""
            Call fonctionsIT.reportsurDBIT(TJM, client, datefact, joursfactu, collabname, facturnumber, typefact, totalHT, totalTTC, delairglt, ato, isfactor)
       Call fonctionsIT.exportEnPdfIT(facturnumber)
End If
End Function
 Public Function reportsurDBIT(TJM, sngclient, datefact, joursfactu, collabname, facturnumber, typefact, totalHT, totalTTC, delairglt, ato, isfactor)
 Dim access_db As Object
 Const db_path As String = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\DB_FACTURES_2016.accdb"
 Set access_db = DAO.DBEngine.OpenDatabase(db_path)
 Dim sql_insert As String
 Dim sql_log As String
 Dim us3r As String
 Dim strIsFactor As String
 Dim sglClient As Single
 Dim sglCollab As Single
 If isfactor = 1 Or 2 Then
 strIsFactor = "F"
 Else
 strIsFactor = "N"
 End If
 typefact = Left(typefact, 1)
 us3r = Environ("Username") 'User = nom d'utilisateur de session Windows '
 'MAJ 01-08-16 : Au lieu de reporter les informations sur une feuille, elles sont envoyées dans la BDD '
 'MAJ 31-08-16 : TODO : Remplacer client et collab par leurs ID respectifs '
    sql_insert = "INSERT INTO [FACT] (NUMFACTURE, TYPE, COLLAB,CLIENT, DATEFAC, PERIODE, TJM,  LIBELLE, NBJOURS, MONTANTHT, MONTANTTTC, REGLEMENT) " & _
    "VALUES (" & facturnumber & "," & Chr(39) & typefact & Chr(39) & "," & Chr(39) & collabname & Chr(39) & ", " & Chr(39) & sngclient & Chr(39) & "," & Chr(39) & datefact & Chr(39) & "," & _
     Chr(39) & Month(datefact) & Chr(39) & "," & Chr(39) & TJM & Chr(39) & ", " & Chr(39) & ato & Chr(39) & "," & Chr(39) & joursfactu & Chr(39) & "," & Chr(39) & totalHT & _
     Chr(39) & "," & Chr(39) & totalTTC & Chr(39) & ", " & Chr(39) & strIsFactor & Chr(39) & ");"
    Debug.Print sql_insert
access_db.Execute (sql_insert)
'La deuxième commande sert à logger les actions dans une table separée, pour plus de tracabilité '
    sql_log = "INSERT INTO [LOG] (username, timest, command, num) VALUES" & _
    "(" & Chr(39) & us3r & Chr(39) & "," & Chr(39) & Now() & Chr(39) & ", " & Chr(39) & Left(sql_insert, 12) & facturnumber & Chr(39) & "," & facturnumber & ");"
    Debug.Print sql_log
access_db.Execute (sql_log)
access_db.Close
End Function
Public Function exportEnPdfIT(facturnumber)
    Sheets("GENERATEUR ATO").Activate
    facturnumber = CStr(facturnumber)
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\FACTURES 2016\" & facturnumber & ".pdf"
        ',Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False, DisplayAlerts:=True
    End With
End Function
Public Function csvnatixis(totalHT, totalTTC, datefact, facturnumber, delairglt, client, typefact)
Dim btypefact As String
btypefact = Left(typefact, 1)
Dim iNumeroClient As Long
iNumeroClient = Application.WorksheetFunction.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 8)
Dim lastrow As Long
Dim WsCsv As Worksheet
Set WsCsv = Sheets("CSVNATIXIS")
WsCsv.Activate
ptrlastrow = WsCsv.Cells(Rows.Count, "B").End(xlUp).Row
ptrlastrow = ptrlastrow + 1
WsCsv.Range("A" & ptrlastrow).Value = btypefact
WsCsv.Range("B" & ptrlastrow).Value = facturnumber
WsCsv.Range("C" & ptrlastrow).Value = datefact
WsCsv.Range("D" & ptrlastrow).Value = iNumeroClient
WsCsv.Range("E" & ptrlastrow).Value = totalHT
WsCsv.Range("F" & ptrlastrow).Value = totalTTC
WsCsv.Range("G" & ptrlastrow) = delairglt
Dim dEcheance As Date
dEcheance = datefact + delairglt
WsCsv.Range("H" & ptrlastrow) = dEcheance
WsCsv.Range("I" & ptrlastrow) = "VIR"
End Function
