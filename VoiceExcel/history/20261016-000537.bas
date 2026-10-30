' Crée un tableau qui contient l'id du client, et l'âge du client, et l'âge du client, Asha, et le prix par producteur, et filtre-les avec des valeurs randomes.

Public Sub VoiceMacro()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim tblName As String
    Dim tblIndex As Long
    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim sampleRows As Long
    Dim rngTable As Range
    
    Set ws = ActiveSheet
    
    ' Determine placement: 2 rows below existing data
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        startRow = 1
    Else
        startRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 1
    End If
    
    headers = Array("ID Client", "Âge client", "Montant achat", "Prix producteur")
    sampleRows = 8
    
    ' Write headers
    For i = LBound(headers) To UBound(headers)
        ws.Cells(startRow, i + 1).Value = headers(i)
    Next i
    
    ' Generate realistic random sample data
    Randomize
    For i = 1 To sampleRows
        ws.Cells(startRow + i, 1).Value = 1000 + i                                ' ID Client
        ws.Cells(startRow + i, 2).Value = Application.RandBetween(22, 65)         ' Âge client
        ws.Cells(startRow + i, 3).Value = Round(Application.RandBetween(150, 4500) + Rnd(), 2) ' Montant achat
        ws.Cells(startRow + i, 4).Value = Round(ws.Cells(startRow + i, 3).Value * (0.55 + Rnd() * 0.25), 2) ' Prix producteur
    Next i
    
    ' Define table range
    Set rngTable = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + sampleRows, UBound(headers) + 1))
    
    ' Generate unique table name
    tblName = "tblAchatsClients"
    tblIndex = 1
    Do While TableExists(ws, tblName)
        tblIndex = tblIndex + 1
        tblName = "tblAchatsClients" & tblIndex
    Loop
    
    ' Create Excel Table (ListObject)
    Set tbl = ws.ListObjects.Add(xlSrcRange, rngTable, , xlYes)
    tbl.Name = tblName
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Apply appropriate number formatting
    ws.Range(tbl.ListColumns("ID Client").DataBodyRange.Address).NumberFormat = "0"
    ws.Range(tbl.ListColumns("Âge client").DataBodyRange.Address).NumberFormat = "0"
    ws.Range(tbl.ListColumns("Montant achat").DataBodyRange.Address).NumberFormat = "#,##0.00 €"
    ws.Range(tbl.ListColumns("Prix producteur").DataBodyRange.Address).NumberFormat = "#,##0.00 €"
    
    ' Autofit columns
    tbl.Range.Columns.AutoFit
End Sub

Private Function TableExists(ws As Worksheet, tblName As String) As Boolean
    Dim t As ListObject
    For Each t In ws.ListObjects
        If StrComp(t.Name, tblName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next t
    TableExists = False
End Function
