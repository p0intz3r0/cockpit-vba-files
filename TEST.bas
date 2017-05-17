Attribute VB_Name = "TEST"
 Sub test()
 Dim rangeUrl As Range
 Dim i As Range
 Dim factpath As String
 Dim n As String
 Set rangeUrl = ActiveSheet.Range("B1:B200")
     factpath = "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\"
 For Each i In rangeUrl
        Debug.Print (factpath)
        n = Cells(i.Row, i.Column - 1).Value
        n = Chr(34) & factpath & n & Chr(34)
        Debug.Print n
        i.Formula = "=HYPERLINK(" & n & ")"
        Next i
End Sub
