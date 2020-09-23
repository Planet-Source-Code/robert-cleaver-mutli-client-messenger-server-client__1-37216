VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessage 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
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
   ScaleHeight     =   3150
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framePM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5175
      Begin RichTextLib.RichTextBox txtData 
         Height          =   2355
         Left            =   45
         TabIndex        =   4
         Top             =   150
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4154
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmClientMessage.frx":0000
      End
      Begin VB.TextBox txtUserTo 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtUserToID 
         Height          =   300
         Left            =   1170
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   810
      End
   End
   Begin VB.Frame frameSendPM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   2430
      Width           =   5175
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   345
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox txtMessage 
         Height          =   345
         Left            =   75
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4080
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()

End Sub

Public Sub cmdSend_Click()

    Call SendPM(Me.txtUserTo.Text, Me.txtUserToID.Text, Me.txtMessage.Text)
    Call AddMessage(frmLogin.txtUsername.Text, UserIDNumber, Me.txtMessage.Text)
    Me.txtMessage.Text = ("")
    
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdSend_Click
    End If
    
End Sub
