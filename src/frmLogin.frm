VERSION 5.00
Begin VB.UserForm frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2400
   End
   Begin VB.TextBox txtPassword 
      PasswordChar    =   "*"
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2400
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   840
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   840
   End
   Begin VB.CommandButton btnManageUsers 
      Caption         =   "Manage Users"
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username:"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   960
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loginOK As Boolean
Private userInfo As Variant

Public Property Get IsAuthenticated() As Boolean
    IsAuthenticated = loginOK
End Property

Public Property Get Username() As String
    Username = Trim(Me.txtUsername.Value)
End Property

Public Property Get UserInfo() As Variant
    UserInfo = userInfo
End Property

Private Sub UserForm_Initialize()
    loginOK = False
    Set userInfo = Nothing
End Sub

Private Sub btnOK_Click()
    Dim info As Variant
    info = GetUserInfo(Username)
    If IsArray(info) Then
        If Trim(info(0)) = Trim(Me.txtPassword.Value) Then
            loginOK = True
            userInfo = info
            Me.Hide
            Exit Sub
        End If
    End If
    loginOK = False
    LogLogin Username, "fail", "invalid credentials"
    MsgBox "Неверное имя пользователя или пароль", vbCritical
End Sub

Private Sub btnCancel_Click()
    loginOK = False
    Me.Hide
End Sub

Private Sub btnManageUsers_Click()
    Dim info As Variant
    info = GetUserInfo(Username)
    If IsArray(info) Then
        If Trim(info(0)) = Trim(Me.txtPassword.Value) And LCase(info(1)) = "admin" Then
            Dim ws As Worksheet
            Set ws = EnsureSheetExists("ПраваДоступа", Array("Username", "Password", "Role", "Sheets", "Ranges"))
            ws.Visible = xlSheetVisible
            ws.Activate
        Else
            MsgBox "Только администратор может управлять пользователями", vbInformation
        End If
    Else
        MsgBox "Пользователь не найден", vbCritical
    End If
End Sub
