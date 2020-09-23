VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMessage 
   Caption         =   "-- Instant Message"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   StartUpPosition =   3  '____
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   2910
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3600
      TabIndex        =   3
      Top             =   2340
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   90
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2340
      Width           =   3465
   End
   Begin VB.PictureBox Picture1 
      Height          =   75
      Left            =   60
      ScaleHeight     =   15
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   2190
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   90
      Width           =   4575
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
