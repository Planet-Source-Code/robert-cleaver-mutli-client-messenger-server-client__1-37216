VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Support Client"
   ClientHeight    =   4155
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
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
   Begin VB.Frame frameConsole 
      Height          =   3825
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   3015
      Begin MSWinsockLib.Winsock sockClient 
         Left            =   2355
         Top             =   660
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3475
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H80000013&
         Caption         =   "?"
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   180
         Width           =   405
      End
      Begin VB.CommandButton cmdMessage 
         BackColor       =   &H80000013&
         Caption         =   "Message"
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H80000013&
         Caption         =   "Login"
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   1005
      End
      Begin ComctlLib.TreeView lstUsers 
         Height          =   3165
         Left            =   90
         TabIndex        =   1
         Top             =   525
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   5583
         _Version        =   327682
         Indentation     =   354
         LabelEdit       =   1
         Style           =   5
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   2460
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   6
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0322
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0644
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0966
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0C88
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":0EEA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSignOut 
         Caption         =   "Sign Out"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVerify_Click()

End Sub

Private Sub cmdHelp_Click()

    sockClient.SendData ("Ping!")

End Sub

Private Sub cmdLogin_Click()

    frmLogin.Show

End Sub

Private Sub cmdMessage_Click()

        lstUsers.Nodes.Item(3).Image = 3

End Sub

Private Sub lstUsers_Click()

    On Error Resume Next

    If lstUsers.SelectedItem.Index = 1 Then
        If lstUsers.SelectedItem.Expanded = True Then
            lstUsers.SelectedItem.Expanded = False
        ElseIf lstUsers.SelectedItem.Expanded = False Then
            lstUsers.SelectedItem.Expanded = True
        End If
    End If

End Sub

Private Sub lstUsers_DblClick()
    
    Dim FromUserID As String
    Dim FromUsername As String
    FromUsername = lstUsers.SelectedItem.Text
    FromUserID = lstUsers.SelectedItem.Index
    If lstUsers.SelectedItem.Index = 1 Then: Exit Sub
    frmMessageArr(FromUserID).Show
    frmMessageArr(FromUserID).Caption = FromUsername & " -- Instant Message"
    frmMessageArr(FromUserID).txtUserTo.Text = FromUsername
    frmMessageArr(FromUserID).txtUserToID.Text = FromUserID

End Sub

Private Sub lstUsers_Expand(ByVal Node As ComctlLib.Node)

    lstUsers.SelectedItem.Selected = False

End Sub

Private Sub mnuClose_Click()

    End

End Sub

Private Sub mnuMinimize_Click()

    frmClient.WindowState = 1

End Sub

Private Sub mnuSignOut_Click()

    sockClient.SendData ("LoggedOut")
    TimeOut (1)
    sockClient.Close
    frmClient.lstUsers.Nodes.Clear
    
End Sub

Private Sub sockClient_ConnectionRequest(ByVal requestID As Long)

    If sockClient.State <> sckClosed Then sockClient.Close
    sockClient.Accept requestID
    TimeOut (1)
    sockClient.SendData ("Connection Established")

End Sub

Private Sub sockClient_DataArrival(ByVal bytesTotal As Long)

    Dim IncomingData As String
    
    sockClient.GetData IncomingData, vbString
    
    If InStr(IncomingData, "Userlist:") Then
        IncomingData = Replace(IncomingData, "Userlist:", "")
        Call LoadUserlist(IncomingData)
        sockClient.SendData ("UsersLoaded")

        Exit Sub
    End If
    
    If InStr(IncomingData, "NewUserOnline:") Then
        IncomingData = Replace(IncomingData, "NewUserOnline:", "")
        MsgBox (IncomingData)
        lstUsers.Nodes(CInt(IncomingData)).Image = 3
        lstUsers.Nodes(CInt(IncomingData)).SelectedImage = 4
        Exit Sub
    End If
    If InStr(IncomingData, "Offline:") Then
        IncomingData = Replace(IncomingData, "Offline:", "")
        lstUsers.Nodes(CInt(IncomingData)).Image = 1
        lstUsers.Nodes(CInt(IncomingData)).SelectedImage = 2
        Exit Sub
    End If
    
    If InStr(IncomingData, "OnlineUsers:") Then
        IncomingData = Replace(IncomingData, "OnlineUsers:", "")
        Call AddOnlineUsers(IncomingData)
        Exit Sub
    End If
    
    If InStr(IncomingData, "PM:") Then
        IncomingData = Replace(IncomingData, "PM:", "")
        Call HandlePM(IncomingData)
    End If
End Sub

