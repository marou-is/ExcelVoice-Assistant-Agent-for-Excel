' Les remplir avec des membres aléatoires des lignes.

Public Sub VoiceMacro()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim targetRows As Long
    Dim i As Long
    Dim startIdx As Long
    Dim colID As Long, colAge As Long
    
    Set ws = ActiveWorkbook.Worksheets("Feuil1")
    Set tbl = ws.ListObjects("tblCustomers")
    
    ' Identifier les colonnes
    colID = tbl.ListColumns("Customer ID").Index
    colAge = tbl.ListColumns("Customer Age").Index
    
    Randomize
    
    ' Vérifier si la première ligne de données est vide
    If tbl.ListRows.Count = 1 Then
        If IsEmpty(tbl.DataBodyRange.Cells(1, colID).Value) And IsEmpty(tbl.DataBodyRange.Cells(1, colAge).Value) Then
            tbl.DataBodyRange.Cells(1, colID).Value = CLng(Int((9999 - 1000 + 1) * Rnd + 1000))
            tbl.DataBodyRange.Cells(1, colAge).Value = CLng(Int((70 - 20 + 1) * Rnd + 20))
            startIdx = 2
        Else
            startIdx = 1
        End If
    Else
        startIdx = 1
    End If
    
    ' Ajouter des lignes supplémentaires avec des nombres aléatoires
    targetRows = 10
    For i = startIdx To targetRows
        Set lr = tbl.ListRows.Add
        lr.Range.Cells(1, colID).Value = CLng(Int((9999 - 1000 + 1) * Rnd + 1000))
        lr.Range.Cells(1, colAge).Value = CLng(Int((70 - 20 + 1) * Rnd + 20))
    Next i
    
    ' Format des colonnes
    tbl.ListColumns(colID).DataBodyRange.NumberFormat = "0"
    tbl.ListColumns(colAge).DataBodyRange.NumberFormat = "0"
    
    ' Ajustement de la largeur des colonnes
    tbl.Range.Columns.AutoFit
End Sub
