Attribute VB_Name = "Client"
Global frmMessageArr(0 To 50) As New frmMessage

Global UserIDNumber As String



Function TimeOut(interval)

    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
    
End Function

Function Login(Username As String, Password As String, Sock As Winsock)

    Sock.SendData (Username & ":" & Password & ";" & Sock.LocalIP)

End Function

Function LoadUserlist(UserlistToLoad As String)

    Dim SpltUser
    Dim SplitCount As Long
    Dim StringDelimiter As String

    StringDelimiter = (",")
    SpltUser = Split(UserlistToLoad, StringDelimiter)
    frmClient.lstUsers.Nodes.Add , "online", "Users", "Users", 1, 1
    frmClient.lstUsers.Nodes(1).Image = 6
    frmClient.lstUsers.Nodes(1).ExpandedImage = 5
    For SplitCount = 0 To UBound(SpltUser)
        frmClient.lstUsers.Nodes.Add "Users", tvwChild, SpltUser(SplitCount), SpltUser(SplitCount), 1, 2
    Next SplitCount
    
End Function

Function AddOnlineUsers(OnlineUsersToAdd As String)

    Dim SpltUser
    Dim SplitCount As Long
    Dim StringDelimiter As String

    StringDelimiter = (",")
    SpltUser = Split(OnlineUsersToAdd, StringDelimiter)
    
    For SplitCount = 0 To UBound(SpltUser)
        frmClient.lstUsers.Nodes(Val(SpltUser(SplitCount))).Image = 3
        frmClient.lstUsers.Nodes(Val(SpltUser(SplitCount))).SelectedImage = 4
    Next SplitCount
    
End Function

Function HandlePM(DataToHandle As String)

    Dim MsgData As String
    Dim FromUsername As String
    Dim FromUserID As String
    Dim FromMessage As String
    Dim MessageDelimiter As String
    Dim UserDelimiter As String
    
    MsgData = DataToHandle
    UserDelimiter = InStr(MsgData, ":")
    FromUsername = Mid(MsgData, 1, UserDelimiter)
    MsgData = Replace(MsgData, FromUsername, "")
    FromUsername = Replace(FromUsername, ":", "")
    MessageDelimiter = InStr(MsgData, "?")
    FromUserID = Mid(MsgData, 1, MessageDelimiter)
    MsgData = Replace(MsgData, FromUserID, "")
    FromUserID = Replace(FromUserID, "?", "")
    FromMessage = MsgData
    
    'MsgBox FromUsername
    'MsgBox FromUserID
    'MsgBox FromMessage
    
    frmMessageArr(FromUserID).Show
    frmMessageArr(FromUserID).Caption = FromUsername & " -- Instant Message"
    frmMessageArr(FromUserID).txtUserTo.Text = FromUsername
    frmMessageArr(FromUserID).txtUserToID.Text = FromUserID
    
    Call AddMessage(FromUsername, FromUserID, FromMessage)

End Function

Function AddMessage(UserFrom As String, UserFromID As String, UserMessageFrom As String)

    With frmMessageArr(UserFromID).txtData
       .SelStart = Len(.Text)
        .SelColor = &H745456
        .SelBold = True
        .SelText = UserFrom & ": "
        .SelBold = False
        .SelColor = &H0&
        .SelText = UserMessageFrom + vbNewLine
    End With
    
End Function

Function SendPM(ToUsername As String, ToID As String, toMsg As String)
    
    frmClient.sockClient.SendData ("PM:" + ToUsername + ":" + ToID + "?" + toMsg)

End Function
