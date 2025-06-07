Option Explicit

' Логирование попыток входа
Public Sub LogLogin(username As String, result As String, comment As String)
    Dim logSheet As Worksheet
    Set logSheet = EnsureSheetExists("ЛогВхода", Array("Timestamp", "User", "Result", "Comment"))
    logSheet.Unprotect Password:=GetProtectionPassword()
    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    With logSheet
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = username
        .Cells(nextRow, 3).Value = result
        .Cells(nextRow, 4).Value = comment
    End With
    logSheet.Protect Password:=GetProtectionPassword(), UserInterfaceOnly:=True
End Sub

' Логирование изменения ячейки
Public Sub LogChangeForCell(ByVal cell As Range, ByVal oldVal As Variant)
    Dim logSheet As Worksheet
    Dim loginSheet As Worksheet
    Dim currentUser As String
    Set logSheet = EnsureSheetExists("ЛогИзменений", Array("Timestamp", "User", "Sheet", "Address", "OldValue", "NewValue"))
    Set loginSheet = EnsureSheetExists("ЛогВхода", Array("Timestamp", "User", "Result", "Comment"))
    currentUser = loginSheet.Cells(loginSheet.Rows.Count, "B").End(xlUp).Value
    If LCase(currentUser) = "admin" Then Exit Sub
    logSheet.Unprotect Password:=GetProtectionPassword()
    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    With logSheet
        .Cells(nextRow, 1).Value = Now
        .Cells(nextRow, 2).Value = currentUser
        .Cells(nextRow, 3).Value = cell.Parent.Name
        .Cells(nextRow, 4).Value = cell.Address
        .Cells(nextRow, 5).Value = oldVal
        .Cells(nextRow, 6).Value = cell.Value
    End With
    logSheet.Protect Password:=GetProtectionPassword(), UserInterfaceOnly:=True
End Sub
