' C'est un table qui contient deux colonnes, première c'est customer ID, deuxième c'est customer age.

Public Sub VoiceMacro()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim tblRange As Range
    Dim lo As ListObject
    Dim baseName As String, tblName As String
    Dim counter As Long

    Set ws = ActiveSheet

    ' Determine start row: empty sheet starts at A1, otherwise 2 rows below used range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        startRow = 1
    Else
        startRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 1
    End If

    ' Set headers
    ws.Cells(startRow, 1).Value = "Customer ID"
    ws.Cells(startRow, 2).Value = "Customer Age"

    ' Set initial range (header + 1 empty row for Excel table creation)
    Set tblRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 1, 2))

    ' Generate unique table name
    baseName = "tblCustomers"
    tblName = baseName
    counter = 1
    Do While TableExists(ws.Parent, tblName)
        counter = counter + 1
        tblName = baseName & counter
    Loop

    ' Create ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium2"

    ' Apply sensible number formatting
    lo.ListColumns(1).DataBodyRange.NumberFormat = "0"
    lo.ListColumns(2).DataBodyRange.NumberFormat = "0"

    ' Autofit columns
    tblRange.EntireColumn.AutoFit
End Sub

Private Function TableExists(wb As Workbook, tName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tName, vbTextCompare) = 0 Then
                TableExists = True
                Exit Function
            End If
        Next lo
    Next ws
    TableExists = False
End Function
