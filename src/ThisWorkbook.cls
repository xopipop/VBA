VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Workbook_Open()
    InitializeGlobals
    frmLogin.Show
    If frmLogin.IsAuthenticated Then
        Dim info As Variant
        info = frmLogin.UserInfo
        ApplyPermissions frmLogin.Username, info(1), info(2), info(3)
        LogLogin frmLogin.Username, "open", ""
    Else
        Application.DisplayAlerts = False
        ThisWorkbook.Close False
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LogChangeForCell Target, ""
End Sub
