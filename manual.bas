Attribute VB_Name = "manual"
Sub manuelIT()
    Dim isfactor As Byte
    Dim TJM As Double
    Dim client, collabname, typefact As String
    Dim datefact As Date
    Dim joursfactu, facturnumber As Double
    Dim totalHT, totalTTC, delairglt As Double
    Dim secteur As String
    Dim numtyp As Byte
    Dim typefact2 As String
    Dim typefactIndex As String
    Sheets("GENERATEUR MANUEL").Activate
    secteur = ActiveSheet.Range("J16")
    client = ActiveSheet.Range("I19").Value
    datefact = ActiveSheet.Range("J13")
    If ActiveSheet.Range("C13").Value = 2 Or ActiveSheet.Range("C13").Value = 3 Then
        collabname = ActiveSheet.Range("D34")
        motif = ActiveSheet.Range("A34")
        nbjours = ActiveSheet.Range("F34")
        TJM = ActiveSheet.Range("H34")
        datefact = ActiveSheet.Range("J14")
        joursfactu = ActiveSheet.Range("F34")
        facturnumber = Sheets("BDD VBA").Range("A12")
        facturnumber = facturnumber + 1
        numtyp = Sheets("GENERATEUR MANUEL").Range("C13").Value
       Sheets("BDD VBA").Range("A12") = facturnumber
        ActiveSheet.Range("J15") = facturnumber
        isfactor = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 11, 0)
       If isfactor = 1 Then
    ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("A5")
    typefact = "Facture Factor CIC"
    ElseIf isfactor = 2 Then
    ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("K1")
    ActiveSheet.Range("A54") = "ATTENTION : changement de coordonnées bancaires "
    ActiveSheet.Range("A54").Font.Color = vbRed
    ActiveSheet.Range("A54").Font.Bold = True
     ActiveSheet.Range("A54").Font.Size = 20
    Else
        ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("A1")
        typefact = "Facture Directe"
         ActiveSheet.Range("A54") = ""
End If
       
        typefact2 = Application.WorksheetFunction.Index(Sheets("BDD VBA").Range("A17:B21"), numtyp, 2)
        typefact = typefact2 & " " & typefact
        If numtyp = 3 Or numtyp = 5 Or numtyp = 4 Then
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 2
            delairglt = 0
        Else
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 1
            Sheets("GENERATEUR MANUEL").Range("C50").Font.ColorIndex = 3
            Sheets("GENERATEUR MANUEL").Range("H50").Font.ColorIndex = 3
            delairglt = ActiveSheet.Range("C50").Value
        End If
        totalHT = ActiveSheet.Range("J40").Value
        totalTTC = ActiveSheet.Range("J44").Value
        totalHT = ActiveSheet.Range("J42").Value
        totalTTC = ActiveSheet.Range("J46").Value
        ato = ActiveSheet.Range("A35").Value
        typefactIndex = Application.WorksheetFunction.Index(Sheets("BDD VBA").Range("A17:B21"), numtyp, 1)
        Call fonctionsIT.reportsurDBIT(TJM, client, datefact, joursfactu, collabname, facturnumber, typefact, totalHT, totalTTC, delairglt, ato, isfactor)
        With ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\FACTURES 2016\" & facturnumber & ".pdf" _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=False
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 1
        End With
    Else
        motif = ActiveSheet.Range("A35")
        rclient = ActiveSheet.Range("I19")
        rinvoicedate = ActiveSheet.Range("J14")
        sommefacture = ActiveSheet.Range("J46")
        metier = ActiveSheet.Range("D35")
        candidat = ActiveSheet.Range("D39")
        facturnumber = Sheets("Facturation ATO 2016").Range("B2")
        facturnumber = facturnumber + 1
        ActiveSheet.Range("J15") = facturnumber
        numtyp = Sheets("GENERATEUR MANUEL").Range("C13").Value
        Sheets("Facturation ATO 2016").Range("B2") = facturnumber
        isfactor = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 11, 0)
       If isfactor = 1 Then
    ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("A5")
    ActiveSheet.Range("A54") = ""
    typefact = "Facture Factor CIC"
    ElseIf isfactor = 2 Then
    
    ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("K1")
    ActiveSheet.Range("A54") = "ATTENTION : changement de coordonnées bancaires à compter du 30/04"
    ActiveSheet.Range("A54").Font.Color = vbRed
    ActiveSheet.Range("A54").Font.Bold = True
     ActiveSheet.Range("A54").Font.Size = 20
    typefact = "Facture Factor NATIXIS"
    Else
        ActiveSheet.Range("A57") = Sheets("BDD VBA").Range("A1")
        ActiveSheet.Range("A54") = ""
        typefact = "Facture Directe"
End If
        typefact2 = Application.WorksheetFunction.Index(Sheets("BDD VBA").Range("A17:B21"), numtyp, 2)
        typefact = typefact2 & " " & typefact
        totalHT = ActiveSheet.Range("J42").Value
        totalTTC = ActiveSheet.Range("J46").Value
        If numtyp = 3 Or numtyp = 5 Or numtyp = 4 Then
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 2
            delairglt = 0
        Else
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 1
            Sheets("GENERATEUR MANUEL").Range("C50").Font.ColorIndex = 3
            Sheets("GENERATEUR MANUEL").Range("H50").Font.ColorIndex = 3
            delairglt = ActiveSheet.Range("C50").Value
        End If
        Call rfonctions.reportsurDBRCT(motif, facturnumber, rinvoicedate, sommefacture, typefact, rclient, complt, totalHT, totalTTC, delairglt, metier, candidat)
        With ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\temp\" & facturnumber & ".pdf"
            Sheets("GENERATEUR MANUEL").Range("A50:H52").Font.ColorIndex = 1
        End With
        
    End If
End Sub
