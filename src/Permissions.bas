Option Explicit

' Снятие защиты со всех листов (и раскрытие их)
Public Sub UnprotectAllSheets(Optional wb As Workbook)
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
        On Error Resume Next
        ws.Unprotect Password:=GetProtectionPassword()
        On Error GoTo 0
    Next ws
End Sub

' Получение данных пользователя из листа "ПраваДоступа"
Public Function GetUserInfo(ByVal username As String) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim result(0 To 3) As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ПраваДоступа")
    On Error GoTo 0
    If ws Is Nothing Then
        result(0) = ""
        result(1) = "user"
        result(2) = "*"
        result(3) = ""
        GetUserInfo = result
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If LCase(Trim(ws.Cells(i, "A").Value)) = LCase(Trim(username)) Then
            result(0) = Trim(ws.Cells(i, "B").Value)
            result(1) = Trim(ws.Cells(i, "C").Value)
            result(2) = Trim(ws.Cells(i, "D").Value)
            result(3) = Trim(ws.Cells(i, "E").Value)
            GetUserInfo = result
            Exit Function
        End If
    Next i
    GetUserInfo = Null
End Function

' Проверка попадания ячейки в разрешённые диапазоны
Public Function IsCellInAllowedRange(ByVal cell As Range, ByVal allowedRanges As String) As Boolean
    Dim arr() As String, part As Variant, rngAllowed As Range
    Dim wsName As String, rngAddress As String
    If Trim(allowedRanges) = "*" Then
        IsCellInAllowedRange = True
        Exit Function
    End If
    arr = Split(allowedRanges, ";")
    For Each part In arr
        part = Trim(part)
        If part <> "" Then
            If InStr(part, "!") > 0 Then
                wsName = Split(part, "!")(0)
                rngAddress = Split(part, "!")(1)
                If LCase(cell.Parent.Name) = LCase(wsName) Then
                    On Error Resume Next
                    Set rngAllowed = cell.Parent.Range(rngAddress)
                    On Error GoTo 0
                    If Not rngAllowed Is Nothing Then
                        If Not Intersect(cell, rngAllowed) Is Nothing Then
                            IsCellInAllowedRange = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next part
    IsCellInAllowedRange = False
End Function

' Запись информации о заблокированной ячейке
Public Sub RecordLockedCell(ByVal targetCell As Range)
    Dim wsLocked As Worksheet
    Dim lastRow As Long, i As Long
    Dim found As Boolean
    Set wsLocked = EnsureSheetExists("LockedCells", Array("Sheet", "Address", "Timestamp"), True)
    found = False
    lastRow = wsLocked.Cells(wsLocked.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If wsLocked.Cells(i, "A").Value = targetCell.Parent.Name And _
           wsLocked.Cells(i, "B").Value = targetCell.Address Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        lastRow = lastRow + 1
        wsLocked.Cells(lastRow, "A").Value = targetCell.Parent.Name
        wsLocked.Cells(lastRow, "B").Value = targetCell.Address
        wsLocked.Cells(lastRow, "C").Value = Now
    End If
End Sub

' Восстановление блокировок при открытии книги
Public Sub ReapplyLockedCells()
    Dim wsLocked As Worksheet
    Dim lastRow As Long, i As Long
    Dim Sh As Worksheet, cell As Range
    Set wsLocked = EnsureSheetExists("LockedCells", Array("Sheet", "Address", "Timestamp"), True)
    lastRow = wsLocked.Cells(wsLocked.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        Dim sheetName As String, cellAddress As String
        sheetName = wsLocked.Cells(i, "A").Value
        cellAddress = wsLocked.Cells(i, "B").Value
        On Error Resume Next
        Set Sh = ThisWorkbook.Sheets(sheetName)
        If Not Sh Is Nothing Then
            Set cell = Sh.Range(cellAddress)
            If Not cell Is Nothing Then
                cell.Locked = True
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

' Разблокировка ячейки (для администратора)
Public Sub AdminUnlockCell(ByVal sheetName As String, ByVal cellAddress As String)
    Dim ws As Worksheet, wsLocked As Worksheet
    Dim lastRow As Long, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Set wsLocked = EnsureSheetExists("LockedCells", Array("Sheet", "Address", "Timestamp"), True)
    lastRow = wsLocked.Cells(wsLocked.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If wsLocked.Cells(i, "A").Value = sheetName And wsLocked.Cells(i, "B").Value = cellAddress Then
            wsLocked.Rows(i).Delete
            Exit For
        End If
    Next i
    ws.Unprotect Password:=GetProtectionPassword()
    ws.Range(cellAddress).Locked = False
    ws.Protect Password:=GetProtectionPassword(), UserInterfaceOnly:=True
End Sub

' Основное применение прав доступа
Public Sub ApplyPermissions(username As String, role As String, accessSheets As String, editRanges As String)
    Dim ws As Worksheet
    Dim firstVisibleSheet As Worksheet
    Dim adminSheets As Variant
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Сначала делаем листы видимыми и снимаем защиту
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
        On Error Resume Next
        ws.Unprotect Password:=GetProtectionPassword()
        On Error GoTo 0
    Next ws

    If SheetExists("СТАРТ", wb) Then
        If LCase(role) = "admin" Then
            wb.Sheets("СТАРТ").Visible = xlSheetVisible
        Else
            wb.Sheets("СТАРТ").Visible = xlSheetVeryHidden
        End If
    End If

    If LCase(role) = "admin" Then
        For Each ws In wb.Worksheets
            ws.Visible = xlSheetVisible
            If firstVisibleSheet Is Nothing Then Set firstVisibleSheet = ws
        Next ws
    Else
        adminSheets = Array("ПраваДоступа", "ЛогВхода", "ЛогИзменений", "LockedCells")
        Dim i As Long
        For i = LBound(adminSheets) To UBound(adminSheets)
            If SheetExists(adminSheets(i), wb) Then
                wb.Sheets(adminSheets(i)).Visible = xlSheetVeryHidden
            End If
        Next i

        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible And IsError(Application.Match(ws.Name, adminSheets, 0)) And ws.Name <> "СТАРТ" Then
                Set firstVisibleSheet = ws
                Exit For
            End If
        Next ws

        If firstVisibleSheet Is Nothing Then
            MsgBox "Нет доступных листов для отображения! Проверьте права доступа.", vbCritical
            Exit Sub
        End If

        firstVisibleSheet.Activate

        For Each ws In wb.Worksheets
            If IsError(Application.Match(ws.Name, adminSheets, 0)) And ws.Name <> "СТАРТ" Then
                If Not ws Is wb.ActiveSheet Then
                    ws.Visible = xlSheetVeryHidden
                End If
            End If
        Next ws

        If Trim(accessSheets) = "*" Then
            For Each ws In wb.Worksheets
                If IsError(Application.Match(ws.Name, adminSheets, 0)) And ws.Name <> "СТАРТ" Then
                    ws.Visible = xlSheetVisible
                End If
            Next ws
        Else
            Dim sheetName As Variant
            For Each sheetName In Split(accessSheets, ";")
                sheetName = Trim(sheetName)
                If SheetExists(sheetName, wb) Then
                    wb.Sheets(sheetName).Visible = xlSheetVisible
                End If
            Next sheetName
        End If
    End If

    If Not firstVisibleSheet Is Nothing Then
        firstVisibleSheet.Activate
    Else
        MsgBox "Нет доступных листов для отображения после применения прав. Проверьте настройки.", vbCritical
    End If

    For Each ws In wb.Worksheets
        ws.Cells.Locked = True
    Next ws

    If Trim(editRanges) = "*" Then
        For Each ws In wb.Worksheets
            ws.Cells.Locked = False
        Next ws
    ElseIf Trim(editRanges) <> "" Then
        Dim sheetPart As String, rangePart As String
        Dim rng As Variant
        For Each rng In Split(editRanges, ";")
            Dim parts() As String
            parts = Split(rng, "!")
            If UBound(parts) >= 1 Then
                sheetPart = Trim(parts(0))
                rangePart = Trim(parts(1))
                If SheetExists(sheetPart, wb) Then
                    On Error Resume Next
                    wb.Sheets(sheetPart).Range(rangePart).Locked = False
                    On Error GoTo 0
                End If
            End If
        Next rng
    End If

    If LCase(role) <> "admin" Then
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible Then
                ws.Protect Password:=GetProtectionPassword(), UserInterfaceOnly:=True
            End If
        Next ws
    End If

    Application.Visible = True

    If LCase(role) <> "admin" Then
        Call ReapplyLockedCells
    End If
End Sub

