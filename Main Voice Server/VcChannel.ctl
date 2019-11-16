VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl VcChannel 
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Ls2 
      Left            =   2880
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws2 
      Index           =   0
      Left            =   960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ls 
      Left            =   2400
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   480
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "VcChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Limit As Integer
Private VCID(1 To 1000) As String
Private VCIndex(1 To 1000) As Integer
Private VCIgnores(0 To 1000) As String
Private VoiceUserList As String
Private MainTalker As String
Private MainIndex As Integer
Private SecondTalker As String
Private SecondIndex As Integer
Private SendListenerCounter As Long
Private SendListenerCounter2 As Long
Private ListenCount As Long
Private IgnoreCount As Long
Private PingCount As Long
Private RoomName As String
Private Port1 As String
Private Port2 As String

Private Sub UserControl_Initialize()
VoiceUserList = "~"
End Sub

Public Function VcList() As String
VcList = VoiceUserList
End Function


Public Sub StartChannel(LimitNum As Integer, RoomN As String, VcPort As String)
On Error Resume Next
MainTalker = ""
SendListenerCounter = 0
Dim i As Integer
Limit = LimitNum
RoomName = RoomN
Port1 = VcPort
Port2 = VcPort + 1
For i = 1 To Limit
VCID(i) = ""
VCIndex(i) = 0
VCIgnores(i) = ""
Load Ws(i)
Load Ws2(i)
Next i
DoEvents
Timer1 = False
VoiceUserList = "~"
Ls.Close
Ls2.Close
Ls.LocalPort = Port1
Ls2.LocalPort = Port2
Ls.Listen
PingCount = 0
Timer2 = True
End Sub

Public Sub StopChannel()
On Error Resume Next
Timer2 = False
PingCount = 0
Ls.Close
Ls2.Close
Dim i As Integer
For i = 1 To Limit
VCID(i) = ""
VCIndex(i) = 0
VCIgnores(i) = ""
Ws(i).Close
Ws2(i).Close
Unload Ws(i)
Unload Ws2(i)
Next i
DoEvents
Timer1 = False
VoiceUserList = "~"
MainTalker = ""
SecondTalker = ""
MainIndex = 0
SecondIndex = 0
RoomName = ""
Port1 = 0
Port2 = 0
End Sub

'''''''''''''''''''''''''''''''''Voice SOcket Stuff''''''''''''''''''''''''''''''''''''''''''

Private Sub Ls_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To Limit
If VCID(i) = "" And Ws(i).State <> 7 Then
Ws(i).Close
Ws(i).Accept requestID
Debug.Print "Accepted New Connection"
Ls.Close
Ls.LocalPort = Port1
Ls.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Server Full"
Ls.Close
Ls.LocalPort = Port1
Ls.Listen
End Sub

Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As Integer, TmpData As String, HeaderLength As Integer
HeaderLength = 12
With Ws(Index)
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
DataLength = Asc(Mid(Data, 2, 1))
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
Debug.Print "VOICE: " & TmpData
ProcessVoice TmpData, Index
DoEvents
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Public Function ProcessVoice(VCDATA As String, Index As Integer)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, SData() As String, TotalCount As Long
Casee = Mid(VCDATA, 6, 4)

Select Case Casee

Case "AUTH"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & VoiceUserList, "~" & Who & "~") > 0 Or Who = "" Or LCase(Who) = "admin" Then
Ws(Index).Close
Call Status(Who & " already in Voice, Rejected!")
Else
Call Status(Who & " Joined Voice!")
VCID(Index) = Who
VCIgnores(Index) = ""
VoiceUserList = VoiceUserList & "~" & Who & "~"
VoiceUserList = Replace(VoiceUserList, "~~", "~")
Debug.Print "Users:: " & VoiceUserList
'Prepare Audio Socket for User... To Listen!
Pause3 "0.4"
Ls2.Close
Ls2.LocalPort = Port2
Ls2.Listen
Pck = "AUDI|||" & Who & "|||" & Port2 & "|||"
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
ForwardUserJoin Who, Index 'Yet to code this where it tells everyone new user joined and then send whoel list to the new user!
End If

Case "TALK"
Who = Split(VCDATA, "|||")(1)
If Who = MainTalker Then MainTalker = "": Exit Function
If Who = SecondTalker Then SecondTalker = "": Exit Function
If MainTalker = "" Then
Call Status("<-" & Who & "->")
MainTalker = Who
MainIndex = 0
IgnoreCount = 0
SendListenerCounter = 20
Pck = "TALK|||" & Who
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Debug.Print "New MainTalker: " & MainTalker
Timer1 = False
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
ElseIf SecondTalker = "" Then
'Call Status2("<-" & Who & "->")
SData() = Split(VoiceUserList, "~")
TotalCount = UBound(SData) - 2
If IgnoreCount >= Int(TotalCount / 2) Then
Debug.Print "New MainTalker Swapped With SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & Who
Timer1 = False
IgnoreCount = 0
SecondIndex = MainIndex
SecondTalker = MainTalker
MainTalker = Who
MainIndex = 0
Call Status("<-" & MainTalker & "->")
SendListenerCounter = 20
Pck = "TALK|||" & Who
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
End If
SecondTalker = Who
SecondIndex = 0
SendListenerCounter2 = 20
Pck = "TALK|||" & Who
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Debug.Print "New SecondTalker: " & SecondTalker
Timer1 = False
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Else
SData() = Split(VoiceUserList, "~")
TotalCount = UBound(SData) - 2
If IgnoreCount >= Int(TotalCount / 2) Then
Debug.Print "New MainTalker Took Mic From Current MainTalker!! " & MainTalker & " - " & MainIndex & " <---> " & Who
Timer1 = False
IgnoreCount = 0
Pck = "STOP|||" & MainTalker
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(VCIndex(MainIndex)).SendData Pck 'Sent packet to tell user there is New Main Talker and they cant transmit voice right now!
MainTalker = Who
MainIndex = 0
Call Status("<-" & MainTalker & "->")
SendListenerCounter = 20
Pck = "TALK|||" & Who
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
End If
Pck = "STOP|||" & Who
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'Sent packet to tell user there is allready Main Talker and they cant transmit voice right now!
End If

Case "STOP"
Who = Split(VCDATA, "|||")(1)
If Who = MainTalker Then
If SecondTalker = "" Or SecondIndex = 0 Then
Timer1 = False
Call Status("<----->")
Debug.Print "MainTalker Stopped!! " & MainTalker & " - " & MainIndex
ForwardMainTalkerStopped MainTalker, Index 'Call Forward to all That User Stopped Talking yet to do this
MainTalker = ""
MainIndex = 0
Else 'Shift Second Talker into MainTalker position as Maintalker has timed out on audio data incoming
Debug.Print "New MainTalker Moved From SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & SecondTalker & " - " & SecondIndex
Timer1 = False
IgnoreCount = 0
MainIndex = SecondIndex
SecondIndex = 0
MainTalker = SecondTalker
SecondTalker = ""
Call Status("<-" & MainTalker & "->")
Timer1 = True
'Call Status2("<----->")
End If
ElseIf Who = SecondTalker Then
'Call Status2("<----->")
Debug.Print "SecondTalker Stopped!! " & SecondTalker & " - " & SecondIndex
ForwardSecondTalkerStopped SecondTalker, SecondIndex 'Call Forward to Main only That Second Stopped Talking
SecondTalker = ""
SecondIndex = 0
End If

Case "IGGY"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & VCIgnores(Index) & "~", "~" & Who & "~") > 0 Then GoTo Skip
VCIgnores(Index) = VCIgnores(Index) & "~" & Who & "~"
VCIgnores(Index) = Replace(VCIgnores(Index), "~~", "~")
Skip:

Case "UNIG"
Who = Split(VCDATA, "|||")(1)
VCIgnores(Index) = Replace(VCIgnores(Index), "~" & Who & "~", "~")
VCIgnores(Index) = Replace(VCIgnores(Index), "~~", "~")

Case Else

End Select
End Function

Private Sub Ws_Close(Index As Integer)
On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
VCIgnores(Index) = ""
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & VoiceUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Left Voice!!") 'Status
Dim i As Integer
Dim Pck As String
Pck = "LEFT|||" & TmpID
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If Ws(i).State = 7 Then Ws(i).SendData Pck: DoEvents
Next i
End If
'Update VoiceUserList to remove user who left from the string!
VoiceUserList = Replace(VoiceUserList, "~" & TmpID & "~", "~")
VoiceUserList = Replace(VoiceUserList, "~~", "~")
Debug.Print "Users:: " & VoiceUserList
End Sub

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
VCIgnores(Index) = ""
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & VoiceUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Left Voice!!") 'Status
Dim i As Integer
Dim Pck As String
Pck = "LEFT|||" & TmpID
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If Ws(i).State = 7 Then Ws(i).SendData Pck: DoEvents
Next i
End If
'Update VoiceUserList to remove user who left from the string!
VoiceUserList = Replace(VoiceUserList, "~" & TmpID & "~", "~")
VoiceUserList = Replace(VoiceUserList, "~~", "~")
Debug.Print "Users:: " & VoiceUserList
End Sub

''''''''''''''''''''''''''''AUDIO SOCKETS STUFF''''''''''''''''''''''''''''''''''''''

Private Sub Ls2_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To Limit
If Ws2(i).State <> 7 Then
Ws2(i).Close
Ws2(i).Accept requestID
Ls2.Close
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Ls2.Close
End Sub

Private Sub Ws2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String, PacketType As String, SData() As String, i As Long, TmpDat As String, TmpLen As Long
Ws2(Index).GetData Data, vbString, bytesTotal
'Debug.Print "AUDIO: " & Data
If Len(Data) < 200 Then

PacketType = Mid(Data, 6, 4)
If PacketType = "VOIP" Then
ProcessAudio Data, Index
ElseIf PacketType = "NAME" Then
Data = Split(Data, "|||")(1)
ProcessName Data, Index
End If

Else

SData = Split(Data, Chr(0) & Chr(0) & Chr(128))
For i = 1 To UBound(SData)
TmpLen = Asc(Right(SData(i - 1), 1))
TmpDat = Chr(0) & Chr(TmpLen) & Chr(0) & Chr(0) & Chr(128) & Mid(SData(i), 1, TmpLen - 5)
PacketType = Mid(TmpDat, 6, 4)
If PacketType = "VOIP" Then
ProcessAudio TmpDat, Index
DoEvents
ElseIf PacketType = "NAME" Then
TmpDat = Split(TmpDat, "|||")(1)
ProcessName TmpDat, Index
End If
Next i
DoEvents

End If
Exit Sub
Error:
Debug.Print "Audio Packet Error: (Data Arrival)  Data: " & Data
End Sub

Public Function ProcessName(Name As String, Index As Integer) 'This Gives the Audio Socket Index.. Knowledge of the Socket Index for a Users Voice Socket! Needed to process Ignores at Audio Arrival!
Dim i As Integer
For i = 1 To Limit
If LCase(VCID(i)) = LCase(Name) Then
VCIndex(Index) = i
Exit Function
End If
Next i
End Function

Private Function ProcessAudio(VCDATA As String, Index As Integer)
'On Error Resume Next
Dim WhoIsIt As String
WhoIsIt = Split(VCDATA, "|||")(1)
'Debug.Print "AUDIO: " & WhoIsIt & " - " & VCDATA
If WhoIsIt = MainTalker Then
Timer1 = False 'reset timeout for next packet to income from maintalker
ForwardMainTalker VCDATA, Index
Timer1 = True
ElseIf WhoIsIt = SecondTalker Then
ForwardSecondTalker VCDATA, Index
Else

End If
End Function

'''''''''''''''''''''''''''''''''''' Subs n Functions n Timers to be called an used '''''''''''''''''''''''''''''''

Private Function ForwardMainTalker(ThePck As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Listeners As Long
Dim Iggys As Long
MainIndex = Indy
Listeners = 0
Iggys = 0
For i = 1 To Limit
If i = Indy Then GoTo Skip
If InStr(1, "~" & VCIgnores(VCIndex(i)) & "~", "~" & MainTalker & "~") > 0 Then Iggys = Iggys + 1: GoTo Skip
If Ws2(i).State = 7 Then Listeners = Listeners + 1: Ws2(i).SendData ThePck: DoEvents
'Pause2 "0.001"
Skip:
Next i
DoEvents
ListenCount = Listeners
IgnoreCount = Iggys
SendListenerCounter = SendListenerCounter + 1
If SendListenerCounter >= 21 Then 'approx every 1 second
SendListenerCounter = 0 'reset count
SendListeners ThePck, ListenCount 'Send the count
End If
End Function

Private Function ForwardSecondTalker(ThePck As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Listeners As Long
SecondIndex = Indy
Listeners = 0
If MainTalker = "" Or MainIndex = 0 Then

Else
If Ws2(MainIndex).State = 7 Then
If InStr(1, "~" & VCIgnores(VCIndex(MainIndex)) & "~", "~" & SecondTalker & "~") > 0 Then GoTo Skip
Listeners = Listeners + 1: Ws2(MainIndex).SendData ThePck: DoEvents
Skip:
SendListenerCounter2 = SendListenerCounter2 + 1
If SendListenerCounter2 >= 21 Then 'approx every 1 second
SendListenerCounter2 = 0 'reset count
SendSecondListeners ThePck, Listeners 'Send the count
End If
End If
End If
End Function

Private Function SendListeners(Whom As String, HowMany As Long) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
Dim i As Integer
Dim Pck As String
Whom = Split(Whom, "|||")(1)
Pck = "NUM#|||" & Whom & "|||" & HowMany
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then
Ws(i).SendData Pck
Exit Function
End If
Next i
End Function

Private Function SendSecondListeners(Whom As String, HowMany As Long) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
Dim i As Integer
Dim Pck As String
Whom = Split(Whom, "|||")(1)
Pck = "NUM#|||" & Whom & "|||" & HowMany
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then
Ws(i).SendData Pck
Exit Function
End If
Next i
End Function

Private Function ForwardMainTalkerStopped(Whom As String, Indy As Integer)
On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = "FREE|||" & Whom
IgnoreCount = 0
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then GoTo Skip
If Ws(i).State = 7 Then Ws(i).SendData Pck
Skip:
Next i
DoEvents
End Function

Private Function ForwardSecondTalkerStopped(Whom As String, Indy As Integer)
On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = "FREE|||" & Whom
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = MainTalker Then
If Ws(i).State = 7 Then Ws(i).SendData Pck
End If
Next i
DoEvents
End Function

Private Function ForwardUserJoin(Whom As String, Indy As Integer)
On Error Resume Next
Dim i As Integer
Dim Pck As String
'Forward new user to all in voice.
Pck = "JOIN|||" & Whom
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If i = Indy Then GoTo Skip
If Ws(i).State = 7 Then Ws(i).SendData Pck
Pause4 "0.005"
Skip:
Next i
DoEvents
'Forward room list to new user
Pck = "LIST|||" & VoiceUserList
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Indy).State = 7 Then
Ws(Indy).SendData Pck
End If
End Function

Private Sub Timer1_Timer() 'Important leave this, it fixs and clear Maintalker, in the event they DC from socket while on AIR making other able to talk again
On Error Resume Next
If Timer1 = False Then Exit Sub
If MainTalker = "" Then

ElseIf SecondTalker = "" Or SecondIndex = 0 Then
Call Status("<----->")
Debug.Print "MainTalker Stopped By Timeout Timer!! " & MainTalker & " - " & MainIndex
ForwardMainTalkerStopped MainTalker, 0 'Call Forward to all That User Stopped Talking yet to do this
MainTalker = ""
MainIndex = 0
ListenerCount = 0
IgnoreCount = 0
Else 'Shift Second Talker into MainTalker position as Maintalker has timed out on audio data incoming
Debug.Print "New MainTalker Moved From SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & SecondTalker & " - " & SecondIndex
IgnoreCount = 0
MainIndex = SecondIndex
SecondIndex = 0
MainTalker = SecondTalker
SecondTalker = ""
Call Status("<-" & MainTalker & "->")
'Call Status2("<----->")
Exit Sub 'keep timer running
End If
Timer1 = False
End Sub

Private Sub Timer2_Timer()
Timer2 = False
PingCount = PingCount + 1
If PingCount = 120 Then
PingCount = 0
PingEveryone
End If
Timer2 = True
End Sub

Private Function PingEveryone()
On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = "PING|||KEEPALIVE"
Pck = Chr(0) & Chr(Len(Pck) + 5) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If Ws(i).State = 7 Then
Ws(i).SendData Pck
Pause2 "0.01"
End If
Next i
DoEvents
End Function

Private Sub Pause2(interval)
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Private Sub Pause3(interval)
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Private Sub Pause4(interval)
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

