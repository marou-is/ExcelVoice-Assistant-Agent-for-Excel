' create a table that contains 3 rows and 4 lines

Public Sub VoiceMacro()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim tblRange As Range
    Dim lo As ListObject
    Dim tblName As String
    Dim i As Long
    Dim numCols As Long
    Dim numDataRows As Long

    Set ws = ActiveSheet

    ' 3 colonnes et 4 lignes de données
    numCols = 3
    numDataRows = 4

    ' Déterminer la position de départ (2 lignes après la plage utilisée)
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        startRow = 1
    Else
        startRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 1
    End If

    ' Définir la plage du tableau (1 ligne d'en-tête + 4 lignes de données)
    Set tblRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + numDataRows, numCols))

    ' Définir les en-têtes
    For i = 1 To numCols
        tblRange.Cells(1, i).Value = "Colonne " & i
    Next i

    ' Générer un nom unique pour la table
    tblName = GetUniqueTableName(ws, "tblDonnees")

    ' Créer le ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium2"

    tblRange.Columns.AutoFit
End Sub

Private Function GetUniqueTableName(ws As Worksheet, baseName As String) As String
    Dim candidate As String
    Dim index As Long
    Dim exists As Boolean
    Dim t As ListObject

    candidate = baseName
    index = 1

    Do
        exists = False
        For Each t In ws.ListObjects
            If StrComp(t.Name, candidate, vbTextCompare) = 0 Then
                exists = True
                Exit For
            End If
        Next t

        If exists Then
            index = index + 1
            candidate = baseName & index
        End If
    Loop While exists

    GetUniqueTableName = candidate
End Function
