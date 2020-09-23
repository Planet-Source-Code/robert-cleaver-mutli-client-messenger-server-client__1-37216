VERSION 5.00
Begin VB.Form frmVerify 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrator Verification"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameVerify 
      Height          =   1845
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify"
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   1440
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "Password"
         Top             =   1050
         Width           =   1875
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Username"
         Top             =   420
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdVerify_Click()

    Dim srvUsername As String
    Dim srvPassword As String
    
    srvUsername = (txtUsername.Text)
    srvPassword = (txtPassword.Text)
    srvUsername = LCase(srvUsername)
    srvPassword = LCase(srvPassword)
    
    If srvUsername = ("1") And srvPassword = ("1") Then
        Call AdminVerified
        LoadUsers (App.Path & "/users.dat") 'remove when done
        frmConsole.lstClients.Nodes(1).ExpandedImage = 5
        frmConsole.lstClients.Nodes(1).Image = 6
        frmConsole.lstClients.Nodes(1).SelectedImage = 6
        Exit Sub
    End If
    
    MsgBox "Username or Password Incorrect.", vbCritical, "Login Incorrect"
    txtUsername.Text = ("Username")
    txtPassword.Text = ("Password")
    txtPassword.PasswordChar = ("")

End Sub

Private Sub Form_Load()

    frmConsole.Show

End Sub

Private Sub txtPassword_Change()

    txtPassword.PasswordChar = ("*")

End Sub

Function LoadUsers(userfile As String)
    
    Dim UserFreeFile
    Dim UserLoad As String
    Dim StringDelimiter
    Dim UserOnly As String
    Dim PassOnly As String
    
    StringDelimiter = ":"
    
    Usercount = 0
    frmConsole.lstClients.Nodes.Add , "online", "users", "Users", 1, 1
    UserFreeFile = FreeFile
    
    Open userfile For Input As UserFreeFile
    Do
    DoEvents
    Input #UserFreeFile, UserLoad
    If Len(UserLoad) > 1 Then
    If frmConsole.lstClients.Nodes.Count = 1 Then Usercount = 2 Else Usercount = Usercount + 1
    UserOnly = Left(UserLoad, InStr(UserLoad, StringDelimiter))
    UserOnly = Replace(UserOnly, ":", "")
    PassOnly = Mid(UserLoad, Len(UserOnly) + 2, Len(UserLoad) - Len(UserOnly) - 1)
    PassOnly = Replace(PassOnly, ":", "")
    frmConsole.lstClients.Nodes.Add "users", tvwChild, UserOnly, UserOnly, 1, 2
    UserArray(Usercount) = UserOnly
    PassArray(Usercount) = PassOnly
    End If
    Loop While EOF(UserFreeFile) = False
    Close #UserFreeFile
    
End Function
