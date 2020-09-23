VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2265
   LinkTopic       =   "Form2"
   ScaleHeight     =   2565
   ScaleWidth      =   2265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameLogin 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2145
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Text            =   "127.0.0.1"
         Top             =   480
         Width           =   1845
      End
      Begin MSWinsockLib.Winsock sockLogin 
         Left            =   1620
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3474
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   2040
         Width           =   945
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Text            =   "Dreams In Digital"
         Top             =   1020
         Width           =   1875
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Text            =   "blea?"
         Top             =   1650
         Width           =   1875
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label txtServerc 
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   780
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1410
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()

    sockLogin.Connect txtServer.Text, 3474
    TimeOut (0.3)
    Call Login(txtUsername.Text, txtPassword.Text, frmLogin.sockLogin)
    
End Sub

Private Sub Form_Load()

    Load frmClient
    frmClient.Visible = False

End Sub

Private Sub sockLogin_ConnectionRequest(ByVal requestID As Long)
    
    If sockLogin.State <> sckClosed Then sockLogin.Close
    
    sockLogin.Accept (requestID)
    
End Sub

Private Sub sockLogin_DataArrival(ByVal bytesTotal As Long)
    
    Dim IncomingString As String
    sockLogin.GetData IncomingString, vbString
    If InStr(IncomingString, "Login Successful") Then
    IncomingString = Replace(IncomingString, "Login Successful", "")
    UserIDNumber = IncomingString
    sockLogin.Close
    frmClient.sockClient.LocalPort = 3475
    frmClient.sockClient.Listen
    Me.Visible = False
    frmClient.Visible = True
    End If

End Sub

