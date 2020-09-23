VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmConsole 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command Console"
   ClientHeight    =   4125
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar frmStatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sockClients 
      Index           =   2
      Left            =   2550
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   3475
   End
   Begin VB.Frame frameConsole 
      Height          =   3825
      Left            =   90
      TabIndex        =   0
      Top             =   -30
      Width           =   3075
      Begin MSWinsockLib.Winsock sockListen 
         Left            =   2040
         Top             =   660
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   3474
      End
      Begin ComctlLib.TreeView lstClients 
         Height          =   3165
         Left            =   90
         TabIndex        =   5
         Top             =   570
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   5583
         _Version        =   327682
         Indentation     =   354
         LabelEdit       =   1
         Style           =   5
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H80000013&
         Caption         =   "?"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton cmdKickUser 
         BackColor       =   &H80000013&
         Caption         =   "Kick User"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdMessage 
         BackColor       =   &H80000013&
         Caption         =   "Message"
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton cmdVerify 
         BackColor       =   &H80000013&
         Caption         =   "Verify"
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   765
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   0
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   1
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   3
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   4
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   5
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   6
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   7
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   8
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   9
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   10
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   11
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   12
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   13
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   14
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   15
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   16
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   17
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   18
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   19
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   20
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   21
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   22
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   23
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   24
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   25
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   26
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   27
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   28
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   29
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   30
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   31
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   32
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   33
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   34
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   35
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   36
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   37
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   38
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   39
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   40
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   41
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   42
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   43
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   44
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   45
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   46
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   47
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   48
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   49
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
      End
      Begin MSWinsockLib.Winsock sockClients 
         Index           =   50
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   3475
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
               Picture         =   "frmConsole.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConsole.frx":0322
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConsole.frx":0644
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConsole.frx":0966
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConsole.frx":0C88
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConsole.frx":0EEA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAdminVerify 
         Caption         =   "Admin Verify"
      End
      Begin VB.Menu mnuUnverifyStop 
         Caption         =   "Unverify and Stop Server"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuChangeAdminInf 
         Caption         =   "Change Server Admin Info"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuUserOptions 
         Caption         =   "User Options"
         Begin VB.Menu mnuAddNewUser 
            Caption         =   "Add New User"
         End
      End
      Begin VB.Menu mnuServerOptions 
         Caption         =   "Server Options"
      End
   End
   Begin VB.Menu mnuServerControl 
      Caption         =   "Server Control"
      Begin VB.Menu mnuStartServer 
         Caption         =   "Start Server"
      End
      Begin VB.Menu mnuStopServer 
         Caption         =   "Stop Server"
      End
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Verified As Boolean
Dim ConnectedToClient As Boolean
Dim Userlist As String

Private Sub Command1_Click()

    frmVerify.Show vbModal

End Sub

Private Sub cmdHelp_Click()
    
    BroadCastMessage ("Offline:3")

End Sub

Private Sub cmdVerify_Click()

    frmVerify.Show vbModal

End Sub

Private Sub Form_Load()

    frmStatus.Panels(1).Width = (frmConsole.ScaleWidth)
    frmStatus.Panels(1).Text = ("Not Verified")
    frmVerify.Show vbModal
    
End Sub

Private Sub imgOffline_Click()

End Sub

Private Sub lstClients_Click()

    On Error Resume Next

    If lstClients.SelectedItem.Index = 1 Then
        If lstClients.SelectedItem.Expanded = True Then
            lstClients.SelectedItem.Expanded = False
        ElseIf lstClients.SelectedItem.Expanded = False Then
            lstClients.SelectedItem.Expanded = True
        End If
    End If

End Sub

Private Sub lstClients_DblClick()

    If lstClients.SelectedItem.Key = "users" Then: Exit Sub
    MsgBox PassArray(lstClients.SelectedItem.Index)
   'ServerToClientMsg (lstClients.SelectedItem.Text)

End Sub

Private Sub lstClients_Expand(ByVal Node As ComctlLib.Node)

    lstClients.SelectedItem.Selected = False

End Sub

Private Sub mnuAddNewUser_Click()

    Dim UsernameF As String
    Dim PasswordF As String
    
    UsernameF = InputBox("Enter the New Username:", "New User")
    If UsernameF = ("") Then: MsgBox "You must Have a Username", vbCritical, "Error Creating Username": Exit Sub
    PasswordF = InputBox("Enter the Password:", "Enter the User's Password")
    If PasswordF = ("") Then: MsgBox "You Must Have a Password.", vbCritical, "Error Creating Password": Exit Sub
    
    AddUser UsernameF, PasswordF

End Sub

Private Sub mnuAdminVerify_Click()

    frmVerify.Show vbModal

End Sub

Private Sub mnuClose_Click()

    End

End Sub

Private Sub mnuMinimize_Click()

    Me.WindowState = (1)

End Sub

Private Sub mnuStartServer_Click()
    
    sockListen.Close
    sockListen.Listen

End Sub

Private Sub mnuUnverifyStop_Click()

    Call AdminUnVerified

End Sub



Private Sub sockClients_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim UserRequest As String
    
    sockClients(Index).GetData UserRequest, vbString
    If InStr(UserRequest, "Ping!") Then
        MsgBox "Ping!"
    End If
    
    If InStr(UserRequest, "Connection Established") Then
        sockClients(Index).SendData ("Userlist:" + GetUserList)
        TimeOut (1)
        Call SignifyUsersOnline(sockClients(Index).Index)
        Exit Sub
        End If
    If InStr(UserRequest, "UsersLoaded") Then
        sockClients(Index).SendData ("OnlineUsers:" + GetOnlineUsers)
        Exit Sub
    End If
    If InStr(UserRequest, "LoggedOut") Then
        lstClients.Nodes(Index).SelectedImage = 2
        lstClients.Nodes(Index).Image = 1
        BroadCastMessage ("NewUserOffline:" & sockClients(Index))
        Exit Sub
    End If
    
    If InStr(UserRequest, "PM:") Then
        UserRequest = Replace(UserRequest, "PM:", "")
        HandleClientPMs (UserRequest)
        Exit Sub
    End If
    
End Sub

Private Sub sockListen_ConnectionRequest(ByVal requestID As Long)

    If sockListen.State <> sckClosed Then sockListen.Close
    sockListen.Accept (requestID)
    
End Sub

Private Sub sockListen_DataArrival(ByVal bytesTotal As Long)
    
    Dim UserString As String
    Dim PassString As String
    Dim IPString As String
    Dim OrigString As String
    Dim ParseCount As Long
    Dim OutgoingUsers As String
    Dim UserCheck As Long
    IPString = sockListen.RemoteHostIP
    
    sockListen.GetData OrigString, vbString
    
    UserString = Left(OrigString, InStr(OrigString, ":"))
    OrigString = Replace(OrigString, UserString, "")
    UserString = Replace(UserString, ":", "")
    PassString = Left(OrigString, InStr(OrigString, ";"))
    PassString = Replace(PassString, ";", "")
    OrigString = Replace(OrigString, PassString, "")
    

    
    IPString = sockListen.RemoteHostIP

    For UserCheck = 2 To lstClients.Nodes.Count - 1
        If UserArray(UserCheck) = UserString And LCase(PassArray(UserCheck)) = LCase(PassString) Then
        lstClients.Nodes.Item(UserCheck).Image = 3
        lstClients.Nodes.Item(UserCheck).SelectedImage = 4
        sockListen.SendData "Login Successful " & UserCheck
        TimeOut (0.5)
        sockListen.Close
        sockListen.Listen
        TimeOut (0.5)
        sockClients(UserCheck).Close
        ConnectedToClient = False
        sockClients(UserCheck).Connect IPString, 3475
        Exit Sub
        End If
    Next UserCheck
    
    sockListen.Close
    sockListen.Listen

End Sub

Function LoadUsers(userfile As String)
    
    Dim UserFreeFile
    Dim UserLoad As String
    Dim StringDelimiter
    Dim UserOnly As String
    Dim PassOnly As String
    
    StringDelimiter = ":"
    
    Usercount = 1
    lstClients.Nodes.Add , "online", "users", "Users", 1, 1
    UserFreeFile = FreeFile
    
    Open userfile For Input As UserFreeFile
    Do
    DoEvents
    Input #UserFreeFile, UserLoad
    If Len(UserLoad) > 1 Then
    If lstClients.Nodes.Count = 1 Then Usercount = 1 Else Usercount = Usercount + 1
    UserOnly = Left(UserLoad, InStr(UserLoad, StringDelimiter))
    UserOnly = Replace(UserOnly, ":", "")
    PassOnly = Mid(UserLoad, Len(UserOnly) + 2, Len(UserLoad) - Len(UserOnly) - 1)
    PassOnly = Replace(PassOnly, ":", "")
    lstClients.Nodes.Add "users", tvwChild, UserOnly, UserOnly, 1, 2
    UserArray(Usercount) = UserOnly
    PassArray(Usercount) = PassOnly
    End If
    Loop While EOF(UserFreeFile) = False
    Close #UserFreeFile
    
End Function

Function BroadCastMessage(MessageToBroadcast)

    Dim UserKeyLoop As Long
    
    For UserKeyLoop = 2 To Usercount
    If frmConsole.lstClients.Nodes(UserKeyLoop).Image = 3 Then
        sockClients(UserKeyLoop).SendData (MessageToBroadcast)
    End If
    
    Next UserKeyLoop
    
End Function

Function HandleClientPMs(ClientPMString As String)

    Dim MsgData As String
    Dim FromUsername As String
    Dim FromUserID As String
    Dim FromMessage As String
    Dim MessageDelimiter As String
    Dim UserDelimiter As String
    
    MsgData = ClientPMString
    UserDelimiter = InStr(MsgData, ":")
    FromUsername = Mid(MsgData, 1, UserDelimiter)
    MsgData = Replace(MsgData, FromUsername, "")
    FromUsername = Replace(FromUsername, ":", "")
    MessageDelimiter = InStr(MsgData, "?")
    FromUserID = Mid(MsgData, 1, MessageDelimiter)
    MsgData = Replace(MsgData, FromUserID, "")
    FromUserID = Replace(FromUserID, "?", "")
    FromMessage = MsgData
    
    sockClients(FromUserID).SendData ("PM:" + ClientPMString)

End Function

Function SignifyUsersOnline(UserIDtoSignify)

    BroadCastMessage ("NewUserOnline:" & UserIDtoSignify)
    
End Function
