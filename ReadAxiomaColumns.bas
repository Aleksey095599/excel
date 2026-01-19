Option Explicit

' === Public test entry point ===
Public Sub Test_ReadAxiomaColumns()
    Dim cols As Variant
    cols = ReadAxiomaColumnsFromParam()

    Debug.Print "Axioma columns (" & (UBound(cols) - LBound(cols) + 1) & "): " & Join(cols, vbTab)
End Sub

' === Core logic ===
Public Function ReadAxiomaColumnsFromParam() As Variant
    Dim ws As Worksheet
    Set ws = GetWorksheetOrFail("Param")

    Dim row_axioma_col As Long
    row_axioma_col = FindRowInColumnA_TrimmedExact(ws, "axioma_columns:")
    If row_axioma_col = 0 Then
        Err.Raise vbObjectError + 1001, "ReadAxiomaColumnsFromParam", _
                  "Не найдена метка 'axioma_columns:' в колонке A листа 'Param'."
    End If

    Dim c As Long
    c = 2 ' column B

    If Len(Trim$(CStr(ws.Cells(row_axioma_col, c).Value))) = 0 Then
        Err.Raise vbObjectError + 1002, "ReadAxiomaColumnsFromParam", _
                  "После 'axioma_columns:' нет ни одного значения (ячейка B" & row_axioma_col & " пустая)."
    End If

    Dim arr() As String
    Dim n As Long
    n = 0

    Do While Len(Trim$(CStr(ws.Cells(row_axioma_col, c).Value))) > 0
        ReDim Preserve arr(0 To n)
        arr(n) = Trim$(CStr(ws.Cells(row_axioma_col, c).Value))
        n = n + 1
        c = c + 1
    Loop

    ReadAxiomaColumnsFromParam = arr
End Function

' === Helpers ===
Private Function GetWorksheetOrFail(ByVal sheetName As String) As Worksheet
    On Error GoTo EH
    Set GetWorksheetOrFail = ThisWorkbook.Worksheets(sheetName)
    Exit Function
EH:
    Err.Raise vbObjectError + 1000, "GetWorksheetOrFail", _
              "Лист '" & sheetName & "' не найден в книге '" & ThisWorkbook.Name & "'."
End Function

Private Function FindRowInColumnA_TrimmedExact(ByVal ws As Worksheet, ByVal needle As String) As Long
    ' Надёжный поиск: сравниваем Trim(value) = needle по используемому диапазону колонки A
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim r As Long
    For r = 1 To lastRow
        If Trim$(CStr(ws.Cells(r, "A").Value)) = needle Then
            FindRowInColumnA_TrimmedExact = r
            Exit Function
        End If
    Next r

    FindRowInColumnA_TrimmedExact = 0
End Function
