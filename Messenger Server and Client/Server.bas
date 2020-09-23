Attribute VB_Name = "Module1"
Option Explicit
Global Usercount As Long
Global UserArray(2 To 50) As String
Global PassArray(2 To 50) As String

Global PMArray(2 To 50) As New frmMessage

Dim UserAccountInfo() As String




Function TimeOut(interval)

    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
    
End Function

Function AdminVerified()

    Unload frmVerify
    frmConsole.mnuAdminVerify.Enabled = False
    frmConsole.cmdVerify.Enabled = False
    frmConsole.mnuChangeAdminInf.Enabled = True
    frmConsole.mnuUnverifyStop.Enabled = True
    frmConsole.mnuOptions.Enabled = True
    frmConsole.mnuServerControl.Enabled = True
    frmConsole.cmdMessage.Enabled = True
    frmConsole.cmdKickUser.Enabled = True
    frmConsole.lstClients.Enabled = True
    frmConsole.frmStatus.Panels(1).Text = ("Admin Verified")
        
End Function

Function AdminUnVerified()

    Unload frmVerify
    frmConsole.mnuAdminVerify.Enabled = True
    frmConsole.cmdVerify.Enabled = True
    frmConsole.mnuChangeAdminInf.Enabled = False
    frmConsole.mnuUnverifyStop.Enabled = False
    frmConsole.mnuOptions.Enabled = False
    frmConsole.mnuServerControl.Enabled = False
    frmConsole.cmdMessage.Enabled = False
    frmConsole.cmdKickUser.Enabled = False
    frmConsole.lstClients.Enabled = False
    frmConsole.cmdVerify.Enabled = True
    frmConsole.frmStatus.Panels(1).Text = ("Not Verified")

End Function

Function AddUser(ByRef UsernameToAdd As String, ByRef PasswordToAdd As String)
    
    Dim OrigList As String
    Dim OpFile
    Dim OrigList2 As String
    OpFile = FreeFile
    Open App.Path & "/users.dat" For Input As #OpFile
        Do While Not EOF(OpFile)
        DoEvents
        Input #OpFile, OrigList2
        OrigList = OrigList + vbNewLine + OrigList2
        Loop
    Close OpFile
    
    OpFile = FreeFile
    
    Open App.Path & "/users.dat" For Output As #OpFile
        Print #1, UsernameToAdd + ":" + PasswordToAdd + ";" + vbNewLine + OrigList
    Close OpFile
    
End Function

Function PrivateMessage(UserToMessage As String)

    MsgBox (UserToMessage)

End Function

Function GetUserList() As String

    Dim UserKeyLoop As Long
    
    For UserKeyLoop = 2 To Usercount
        GetUserList = GetUserList + frmConsole.lstClients.Nodes(UserKeyLoop).Text + ","
    Next UserKeyLoop
    
        GetUserList = Mid(GetUserList, 1, Len(GetUserList) - 1)
        
End Function

Function GetOnlineUsers() As String

    Dim UserKeyLoop As Long
    
    For UserKeyLoop = 2 To Usercount
    If frmConsole.lstClients.Nodes(UserKeyLoop).Image = 3 Then
        GetOnlineUsers = GetOnlineUsers & frmConsole.lstClients.Nodes(UserKeyLoop).Index & ","
    End If
    Next UserKeyLoop
    
        GetOnlineUsers = Mid(GetOnlineUsers, 1, Len(GetOnlineUsers) - 1)
        
End Function
