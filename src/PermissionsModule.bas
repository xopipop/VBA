Option Explicit

' Глобальные переменные для проекта
Public g_UserRole As String
Public g_EditRanges As String
Public OldValues As Object  ' Будет хранить пары "Лист!Адрес" -> старое значение
Public Const Password As String = "mypassword"

' Инициализация словаря, если он ещё не создан
Public Sub InitializeGlobals()
    If OldValues Is Nothing Then
        Set OldValues = CreateObject("Scripting.Dictionary")
    End If
End Sub

' Снимает защиту со всех листов книги
Public Sub UnprotectAllSheets(Optional wb As Workbook)
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=Password
        On Error GoTo 0
    Next ws
End Sub

' Проверяет, существует ли лист с заданным именем
Public Function SheetExists(ByVal sheetName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Получает данные пользователя из листа "ПраваДоступа"
' Результат: массив из 4 элементов:
'   (0) требуемый пароль (если пусто, пароль не нужен);
'   (1) роль (например, "admin" или "user");
'   (2) перечень доступных листов (например, "*" или "Лист1;Лист2");
'   (3) перечень диапазонов для редактирования (например, "*" или "Лист1!A1:B10;Лист2!C1:D10")
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
    
    lastRow = ws.Cells(ws.ROWS.Count, "A").End(xlUp).Row
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

' Логирование попыток входа в систему
Public Sub LogLogin(username As String, result As String, comment As String)
    Dim logSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    If Not SheetExists("ЛогВхода", wb) Then
        MsgBox "Лист 'ЛогВхода' не найден!", vbCritical
        Exit Sub
    End If
    Set logSheet = wb.Sheets("ЛогВхода")
    logSheet.Unprotect Password:=Password
    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.ROWS.Count, "A").End(xlUp).Row + 1
    With logSheet
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = username
        .Cells(nextRow, 3).Value = result
        .Cells(nextRow, 4).Value = comment
    End With
    logSheet.Protect Password:=Password, UserInterfaceOnly:=True
End Sub

' Логирование изменения конкретной ячейки
Public Sub LogChangeForCell(ByVal cell As Range, ByVal oldVal As Variant)
    Dim logSheet As Worksheet
    Dim wb As Workbook
    Dim currentUser As String
    Dim loginSheet As Worksheet
    Dim nextRow As Long
    Set wb = ThisWorkbook
    If Not SheetExists("ЛогИзменений", wb) Then Exit Sub
    Set logSheet = wb.Sheets("ЛогИзменений")
    If Not SheetExists("ЛогВхода", wb) Then Exit Sub
    Set loginSheet = wb.Sheets("ЛогВхода")
    currentUser = loginSheet.Cells(loginSheet.ROWS.Count, "B").End(xlUp).Value
    ' Не логируем для администратора
    If LCase(currentUser) = "admin" Then Exit Sub
    logSheet.Unprotect Password:=Password
    nextRow = logSheet.Cells(logSheet.ROWS.Count, "A").End(xlUp).Row + 1
    With logSheet
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = currentUser
        .Cells(nextRow, 3).Value = cell.Parent.Name
        .Cells(nextRow, 4).Value = cell.Address
        .Cells(nextRow, 5).Value = oldVal
        .Cells(nextRow, 6).Value = cell.Value
    End With
    logSheet.Protect Password:=Password, UserInterfaceOnly:=True
End Sub

' Проверка, входит ли ячейка в разрешённый диапазон
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

' Запись информации о заблокированной ячейке в лист "LockedCells"
Public Sub RecordLockedCell(ByVal targetCell As Range)
    Dim wsLocked As Worksheet
    Dim lastRow As Long, i As Long
    Dim found As Boolean
    On Error Resume Next
    Set wsLocked = ThisWorkbook.Sheets("LockedCells")
    On Error GoTo 0
    If wsLocked Is Nothing Then
        Set wsLocked = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLocked.Name = "LockedCells"
        wsLocked.Visible = xlSheetVeryHidden
        wsLocked.Range("A1").Value = "Sheet"
        wsLocked.Range("B1").Value = "Address"
        wsLocked.Range("C1").Value = "Timestamp"
    End If
    found = False
    lastRow = wsLocked.Cells(wsLocked.ROWS.Count, "A").End(xlUp).Row
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

' Повторное применение блокировки ячеек при открытии файла
Public Sub ReapplyLockedCells()
    Dim wsLocked As Worksheet
    Dim lastRow As Long, i As Long
    Dim Sh As Worksheet, cell As Range
    On Error Resume Next
    Set wsLocked = ThisWorkbook.Sheets("LockedCells")
    On Error GoTo 0
    If wsLocked Is Nothing Then Exit Sub
    lastRow = wsLocked.Cells(wsLocked.ROWS.Count, "A").End(xlUp).Row
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
    Set wsLocked = ThisWorkbook.Sheets("LockedCells")
    On Error GoTo 0
    If ws Is Nothing Or wsLocked Is Nothing Then Exit Sub
    lastRow = wsLocked.Cells(wsLocked.ROWS.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If wsLocked.Cells(i, "A").Value = sheetName And wsLocked.Cells(i, "B").Value = cellAddress Then
            wsLocked.ROWS(i).Delete
            Exit For
        End If
    Next i
    ws.Unprotect Password:=Password
    ws.Range(cellAddress).Locked = False
    ws.Protect Password:=Password, UserInterfaceOnly:=True
End Sub

Option Explicit

Public Sub ApplyPermissions(username As String, role As String, accessSheets As String, editRanges As String)
    Dim ws As Worksheet
    Dim firstVisibleSheet As Worksheet
    Dim adminSheets As Variant
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Сначала делаем все листы видимыми и снимаем с них защиту
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    UnprotectAllSheets wb
    
    ' Пример управления видимостью листа "СТАРТ"
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
            Dim sheetName As Variant  ' Variant, так как Split возвращает массив Variants
            For Each sheetName In Split(accessSheets, ";")
                sheetName = Trim(sheetName)
                If SheetExists(sheetName, wb) Then
                    wb.Sheets(sheetName).Visible = xlSheetVisible
                End If
            Next sheetName
        End If
    End If  ' --- Завершаем блок If LCase(role)="admin" ... Else
    
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
                ws.Protect Password:=Password, UserInterfaceOnly:=True
            End If
        Next ws
    End If
    
    Application.Visible = True
    
    If LCase(g_UserRole) <> "admin" Then
        Call ReapplyLockedCells
    End If
End Sub

