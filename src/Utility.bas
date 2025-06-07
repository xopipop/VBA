Option Explicit

'      True/False
Public Function SheetExists(ByVal sheetName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'      ""  B1
Public Function GetProtectionPassword() As String
    On Error Resume Next
    GetProtectionPassword = ThisWorkbook.Sheets("").Range("B1").Value
    On Error GoTo 0
End Function

'       
Public Function EnsureSheetExists(ByVal sheetName As String, headers As Variant, Optional veryHidden As Boolean = False) As Worksheet
    Dim ws As Worksheet
    If SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Sheets(sheetName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
        If veryHidden Then ws.Visible = xlSheetVeryHidden
    End If
    Set EnsureSheetExists = ws
End Function

'      
Public Sub InitializeGlobals()
    Call EnsureSheetExists("", Array("Timestamp", "User", "Result", "Comment"))
    Call EnsureSheetExists("", Array("Timestamp", "User", "Sheet", "Address", "OldValue", "NewValue"))
    Call EnsureSheetExists("LockedCells", Array("Sheet", "Address", "Timestamp"), True)
End Sub

