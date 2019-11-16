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
Private VCIggyPoints(0 To 1000) As Long
Private VoiceUserList As String
Private MainTalker As String
Private MainIndex As Integer
Private SecondTalker As String
Private SecondIndex As Integer
Private SendListenerCounter As Long
Private SendListenerCounter2 As Long
Private SendListenerCounter3 As Long
Private ListenCount As Long
Private IgnoreCount As Long
Private IgnoreCount2 As Long
Private MainTalkerDomination As Long
Private SecondTalkerDomination As Long
Private DetectSecondDC As Long
Private TotalCount As Long
Private PingCount As Long
Private RoomName As String
Private Port1 As String
Private Port2 As String

Private Sub UserControl_Initialize()
On Error Resume Next
VoiceUserList = "~"
End Sub

Public Function VcList() As String
On Error Resume Next
VcList = VoiceUserList
End Function


Public Sub StartChannel(LimitNum As Integer, RoomN As String, VcPort As String)
On Error Resume Next
MainTalker = ""
SendListenerCounter = 0
SendListenerCounter2 = 0
SendListenerCounter3 = 0
MainTalkerDomination = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
IgnoreCount2 = 0
IgnoreCount = 0
TotalCount = 70
Dim i As Integer
Limit = LimitNum
RoomName = RoomN
Port1 = VcPort
Port2 = VcPort + 1
For i = 1 To Limit
Unload Ws(i)
Unload Ws2(i)
VCID(i) = ""
VCIndex(i) = 0
VCIgnores(i) = ""
VCIggyPoints(i) = 0
DoEvents
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
VCIggyPoints(i) = 0
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
'On Error Resume Next
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
'On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws(Index)
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
TmpData = Dee(Mid(TmpData, 11, Len(TmpData) - 10))
TmpData = "R4R4" & Chr(0) & Chr$(Int(Len(TmpData) / 256)) & Chr$(Len(TmpData) Mod 256) & Chr(0) & Chr(0) & Chr(128) & TmpData
Debug.Print "VOICE: " & TmpData
ProcessVoice TmpData, Index
DoEvents
Else
Exit Sub
End If
DoEvents
Else
GoTo Error
End If
Wend
End With
Exit Sub
Error:
'On Error Resume Next
Ws(Index).GetData TmpData
End Sub

Public Function ProcessVoice(VCDATA As String, Index As Integer)
'On Error Resume Next
Dim Who As String, Pck As String, Casee As String, SData() As String, TotalC As Long 'TotalCount As Long
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "AUTH"
Who = Split(VCDATA, "|||")(1)
TotalC = Split(VCDATA, "|||")(2)
If TotalC = 0 Then TotalC = 1
If InStr(1, "~" & VoiceUserList, "~" & Who & "~") > 0 Or Who = "" Or LCase(Who) = "admin" Then
Ws(Index).Close
Call Status(Who & " already in Voice, Rejected!")
Else
Call Status(Who & " Joined Voice!")
VCID(Index) = Who
VCIgnores(Index) = ""
VCIggyPoints(Index) = TotalC
VoiceUserList = VoiceUserList & "~" & Who & "~"
VoiceUserList = Replace(VoiceUserList, "~~", "~")
Debug.Print "Users:: " & VoiceUserList
'Prepare Audio Socket for User... To Listen!
'Pause3 "0.4"
Ls2.Close
Ls2.LocalPort = Port2
Ls2.Listen
Pck = Enn("AUDI|||" & Who & "|||" & Port2 & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
ForwardUserJoin Who, Index 'Yet to code this where it tells everyone new user joined and then send whoel list to the new user!
Exit Function
End If

Case "TALK"
Who = Split(VCDATA, "|||")(1)
If Who = MainTalker Then MainTalker = "": Exit Function
If Who = SecondTalker Then SecondTalker = "": Exit Function
If MainTalker = "" Then
''''''No MaintTalker Events, Assigns Talker As Main''''''''''''
'Call Status("<-" & Who & "->")
MainTalkerDomination = 0
MainTalker = Who
MainIndex = 0
IgnoreCount = 0
SendListenerCounter = 20
Pck = Enn("TALK|||" & Who)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Debug.Print "New MainTalker: " & MainTalker
Timer1 = False
DetectSecondDC = 0
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
ElseIf SecondTalker = "" Then
''''''''''No Second Talker Events, Checks Iggys For Main and Assigns Second Talker Or Main''''''''''''
'TotalCount = ListenCount + IgnoreCount
If TotalCount > 70 Then
TotalC = Int(TotalCount / 3)
Else
TotalC = Int(TotalCount / 2)
End If
If IgnoreCount >= TotalC Or MainTalkerDomination > 3899 Then
''''Second Talker Become Main Talker If Iggy Count high Enough''''
Debug.Print "New MainTalker Swapped With SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & Who
Timer1 = False
IgnoreCount2 = IgnoreCount
IgnoreCount = 0
SecondTalkerDomination = 0
MainTalkerDomination = 0
DetectSecondDC = 0
SecondIndex = MainIndex
SecondTalker = MainTalker
MainTalker = Who
MainIndex = 0
Call Status("<-" & MainTalker & "->")
SendListenerCounter = 20
Pck = Enn("TALK|||" & Who)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
End If
'''''Second Talker Is Assigned Here Making There Now 2 Talkers''''''
IgnoreCount2 = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
SecondTalker = Who
SecondIndex = 0
SendListenerCounter2 = 20
Pck = Enn("TALK|||" & Who)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Debug.Print "New SecondTalker: " & SecondTalker
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
Else
'''''''Already Main & Second Talker Events, Check Main Talker Iggycount and Take Mic If To Many''''''
If TotalCount > 70 Then
TotalC = Int(TotalCount / 3)
Else
TotalC = Int(TotalCount / 2)
End If
If IgnoreCount >= TotalC Or MainTalkerDomination = 3899 Then
''''''Takes Mic From Main Talker For High Iggy Count, Giving To New Talker!!'''''
Debug.Print "New MainTalker Took Mic From Current MainTalker!! " & MainTalker & " - " & MainIndex & " <---> " & Who
Timer1 = False
IgnoreCount = 0
MainTalkerDomination = 0
Pck = Enn("STOP|||" & MainTalker & "|||0/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(VCIndex(MainIndex)).SendData Pck 'Sent packet to tell user there is New Main Talker and they cant transmit voice right now!
MainTalker = Who
MainIndex = 0
'Call Status("<-" & MainTalker & "->")
SendListenerCounter = 20
Pck = Enn("TALK|||" & Who)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Timer1 = True
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
End If
If IgnoreCount2 >= TotalC Or SecondTalkerDomination > 399 Then
''''''Takes Mic From Second Talker For High Iggy Count, Giving To New Talker!!'''''
Debug.Print "New SecondTalker Took Mic From Current SecondTalker!! " & SecondTalker & " - " & SecondIndex & " <---> " & Who
IgnoreCount2 = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
Pck = Enn("STOP|||" & SecondTalker & "|||0/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(VCIndex(SecondIndex)).SendData Pck 'Sent packet to tell user there is New Main Talker and they cant transmit voice right now!
SecondTalker = Who
SecondIndex = 0
'Call Status("<-" & MainTalker & "->")
SendListenerCounter2 = 20
Pck = Enn("TALK|||" & Who)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'Sent packet to tell user he is Main Talker and can start transmitting audio data to server!
Exit Function
End If
''''Main & Second Talkers Assigned, No Position Available, So Told To Stop Talking'''''
Pck = Enn("STOP|||" & Who & "|||0/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'Sent packet to tell user there is allready Main Talker and they cant transmit voice right now!
Exit Function
End If

Case "STOP"
Who = Split(VCDATA, "|||")(1)
If Who = MainTalker Then
If SecondTalker = "" Or SecondIndex = 0 Then
Timer1 = False
'Call Status("<----->")
Debug.Print "MainTalker Stopped!! " & MainTalker & " - " & MainIndex
ForwardMainTalkerStopped MainTalker, Index 'Call Forward to all That User Stopped Talking yet to do this
MainTalkerDomination = 0
MainTalker = ""
MainIndex = 0
Else 'Shift Second Talker into MainTalker position as Maintalker has timed out on audio data incoming
Debug.Print "New MainTalker Moved From SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & SecondTalker & " - " & SecondIndex
Timer1 = False
IgnoreCount = 0
IgnoreCount2 = 0
MainTalkerDomination = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
MainIndex = SecondIndex
SecondIndex = 0
MainTalker = SecondTalker
SecondTalker = ""
'Call Status("<-" & MainTalker & "->")
Timer1 = True
'Call Status2("<----->")
End If
ElseIf Who = SecondTalker Then
'Call Status2("<----->")
Debug.Print "SecondTalker Stopped!! " & SecondTalker & " - " & SecondIndex
ForwardSecondTalkerStopped SecondTalker, SecondIndex 'Call Forward to Main only That Second Stopped Talking
IgnoreCount2 = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
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
'On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
VCIgnores(Index) = ""
VCIggyPoints(Index) = 0
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & VoiceUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Left Voice!!") 'Status
Dim i As Integer
Dim Pck As String
Pck = Enn("LEFT|||" & TmpID)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
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
'On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
VCIgnores(Index) = ""
VCIggyPoints(Index) = 0
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & VoiceUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Left Voice!!") 'Status
Dim i As Integer
Dim Pck As String
Pck = Enn("LEFT|||" & TmpID)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
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
'On Error Resume Next
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
'On Error GoTo Error
Dim PData As String
Ws2(Index).GetData PData, vbString, bytesTotal
'Debug.Print "AUDIO: " & PData
ProcessAudioArrival PData, Index
Exit Sub
Error:
'Debug.Print "Audio Packet Error: (Data Arrival)  Data: " & PData
End Sub

Private Function ProcessAudioArrival(Data As String, Index As Integer)
'On Error GoTo Error
Dim PacketType As String, SData() As String, i As Long, TmpDat As String, TmpLen As String
'Debug.Print "AUDIO: " & Data
If Len(Data) < 200 Then
PacketType = Mid(Data, 11, 4)
If PacketType = "VOIP" Then
ProcessAudio Data, Index
Exit Function
ElseIf PacketType = "PING" Then
Exit Function
Else
Data = Dee(Mid(Data, 11, Len(Data) - 10))
Data = "R4R4" & Chr(0) & Chr$(Int(Len(Data) / 256)) & Chr$(Len(Data) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Data
PacketType = Mid(Data, 11, 4)
If PacketType = "NAME" Then
Data = Split(Data, "|||")(1)
ProcessName Data, Index
Exit Function
Else
'Ws2(Index).Close
'Close Socket Because this should never happen
End If
End If
Exit Function
Else

SData = Split(Data, "R4R4" & Chr(0))
For i = 1 To UBound(SData)
'On Error GoTo Error
'TmpLen = Trim((256 * Asc(Mid(SData(i - 1), Len(SData(i - 1)) - 1, 1)) + Asc(Mid(SData(i - 1), Len(SData(i - 1)), 1))))
'TmpDat = Mid(SData(i), 1, TmpLen)
'TmpDat = "R4R4" & Chr(0) & Chr$(Int(Len(TmpDat) / 256)) & Chr$(Len(TmpDat) Mod 256) & Chr(0) & Chr(0) & Chr(128) & TmpDat
TmpDat = "R4R4" & Chr(0) & SData(i)
PacketType = Mid(TmpDat, 11, 4)
'Debug.Print "AUDIO: " & PacketType & " " & TmpDat
If PacketType = "VOIP" Then
'If Len(TmpDat) > 115 And Right(TmpDat, 3) = "|||" Then ProcessAudio TmpDat, Index
ProcessAudio TmpDat, Index
'DoEvents
ElseIf PacketType = "PING" Then
Else
TmpDat = Dee(Mid(TmpDat, 11, Len(TmpDat) - 10))
TmpDat = "R4R4" & Chr(0) & Chr$(Int(Len(TmpDat) / 256)) & Chr$(Len(TmpDat) Mod 256) & Chr(0) & Chr(0) & Chr(128) & TmpDat
PacketType = Mid(Data, 11, 4)
If PacketType = "NAME" Then
TmpDat = Split(TmpDat, "|||")(1)
ProcessName TmpDat, Index
Else
'Ws2(Index).Close
'Close Socket Because this should never happen
End If
End If
Next i
'DoEvents
Exit Function
End If
Exit Function
Error:
'Debug.Print "Audio Packet Error: (Data Arrival)  Data: " & Data
End Function

Public Function ProcessName(Name As String, Index As Integer) 'This Gives the Audio Socket Index.. Knowledge of the Socket Index for a Users Voice Socket! Needed to process Ignores at Audio Arrival!
'On Error Resume Next
Dim i As Integer
For i = 1 To Limit
If LCase(VCID(i)) = LCase(Name) Then
VCIndex(Index) = i
Exit Function
End If
Next i
'DoEvents
Ws2(Index).Close
End Function

Private Function ProcessAudio(VCDATA As String, Index As Integer)
'On Error Resume Next
Dim WhoIsIt As String
WhoIsIt = Split(VCDATA, "|||")(1)
'Debug.Print "AUDIO: " & WhoIsIt & " - " & VCDATA
If WhoIsIt = MainTalker Then
Timer1 = False 'reset timeout for next packet to income from maintalker
Timer1 = True
ForwardMainTalker VCDATA, Index
Exit Function
'Timer1 = True
ElseIf WhoIsIt = SecondTalker Then
DetectSecondDC = 0
ForwardSecondTalker VCDATA, Index
Else
SendOtherListeners VCDATA, 0
End If
End Function

'''''''''''''''''''''''''''''''''''' Subs n Functions n Timers to be called an used '''''''''''''''''''''''''''''''

Private Function ForwardMainTalker(ThePck As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Listeners As Long
Dim iggys As Long
MainIndex = Indy
Listeners = 0
iggys = 0
For i = 1 To Limit
If i = Indy Then GoTo Skip
If InStr(1, "~" & VCIgnores(VCIndex(i)) & "~", "~" & MainTalker & "~") > 0 Then iggys = iggys + VCIggyPoints(VCIndex(i)): GoTo Skip
If Ws2(i).State = 7 Then Listeners = Listeners + VCIggyPoints(VCIndex(i)): Ws2(i).SendData ThePck: DoEvents
Pause3 "0.002"
Skip:
Next i
DoEvents
ListenCount = Listeners
IgnoreCount = iggys
TotalCount = ListenCount + IgnoreCount
SendListenerCounter = SendListenerCounter + 1
MainTalkerDomination = MainTalkerDomination + 1
'Debug.Print "AUDIO: " & SendListenerCounter & "  " & ListenCount & "/" & TotalCount & " - " & ThePck
DetectSecondDC = DetectSecondDC + 1
If MainTalkerDomination > 4999 Then MainTalkerDomination = 0
If DetectSecondDC > 25 Then
ForwardSecondTalkerStopped SecondTalker, SecondIndex 'Call Forward to Main only That Second Stopped Talking
End If
If SendListenerCounter >= 21 Then 'approx every 1 second
SendListenerCounter = 0 'reset count
SendListeners ThePck, ListenCount 'Send the count
End If
End Function

Private Function ForwardSecondTalker(ThePck As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Listeners As Long
Dim iggys As Long
SecondIndex = Indy
Listeners = 0
If MainTalker = "" Or MainIndex = 0 Then

Else

For i = 1 To Limit
If i = Indy Then GoTo Skip
If InStr(1, "~" & VCIgnores(VCIndex(i)) & "~", "~" & SecondTalker & "~") > 0 Then iggys = iggys + VCIggyPoints(VCIndex(i)): GoTo Skip
If MainIndex = i Then
If Ws2(i).State = 7 Then Listeners = Listeners + VCIggyPoints(VCIndex(i)): Ws2(i).SendData ThePck: DoEvents
Else
If InStr(1, "~" & VCIgnores(VCIndex(i)) & "~", "~" & MainTalker & "~") > 0 Then
If Ws2(i).State = 7 Then Listeners = Listeners + VCIggyPoints(VCIndex(i)): Ws2(i).SendData ThePck: DoEvents
Pause2 "0.002"
End If
End If
Skip:
Next i
DoEvents

IgnoreCount2 = iggys
SendListenerCounter2 = SendListenerCounter2 + 1
SecondTalkerDomination = SecondTalkerDomination + 1
If SecondTalkerDomination > 1499 Then SecondTalkerDomination = 0
If SendListenerCounter2 >= 21 Then 'approx every 1 second
SendListenerCounter2 = 0 'reset count
SendSecondListeners ThePck, Listeners 'Send the count
CheckMainIggied
End If

End If
End Function

Private Function CheckMainIggied()
'On Error Resume Next
Dim Pck As String
Dim TotalC As Long
Dim TmpM As String
Dim TmpMI As Long
'TotalCount = ListenCount + IgnoreCount
If TotalCount > 70 Then
TotalC = Int(TotalCount / 3)
Else
TotalC = Int(TotalCount / 2)
End If
If IgnoreCount >= TotalC Or MainTalkerDomination > 3899 Then
If SecondTalker = "" Or SecondIndex = 0 Then

Else
Debug.Print "New MainTalker Swapped With SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & SecondTalker & " - " & SecondIndex
Timer1 = False
'Pck = Enn("STOP|||" & MainTalker)
'Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
'Ws(VCIndex(MainIndex)).SendData Pck 'Sent packet to tell user there is New Main Talker and they cant transmit voice right now!
IgnoreCount2 = IgnoreCount
IgnoreCount = 0
SecondTalkerDomination = 0
MainTalkerDomination = 0
DetectSecondDC = 0
SendListenerCounter = 20
TmpM = MainTalker
TmpMI = MainIndex
MainIndex = SecondIndex
SecondIndex = TmpMI
MainTalker = SecondTalker
SecondTalker = TmpM
Timer1 = True
End If
End If
End Function

Private Function SendListeners(Whom As String, HowMany As Long) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
'On Error GoTo Error
Dim i As Integer
Dim Pck As String
'TotalCount = ListenCount + IgnoreCount
Whom = Split(Whom, "|||")(1)
Pck = Enn("NUM#|||" & Whom & "|||" & HowMany & "/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then
Ws(i).SendData Pck
Pause4 "0.002"
Exit Function
End If
Next i
Error:
End Function

Private Function SendSecondListeners(Whom As String, HowMany As Long) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
'On Error GoTo Error
Dim i As Integer
Dim Pck As String
Whom = Split(Whom, "|||")(1)
Pck = Enn("NUM#|||" & Whom & "|||" & HowMany & "/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then
Ws(i).SendData Pck
Pause4 "0.002"
Exit Function
End If
Next i
Error:
End Function

Private Function SendOtherListeners(Whom As String, HowMany As Long) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
'On Error GoTo Error
Dim i As Integer
Dim Pck As String
SendListenerCounter3 = SendListenerCounter3 + 1
If SendListenerCounter3 >= 21 Then
SendListenerCounter3 = 0
Whom = Split(Whom, "|||")(1)
Pck = Enn("STOP|||" & Whom & "|||" & HowMany & "/" & TotalCount)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then
Ws(i).SendData Pck
Pause4 "0.002"
Exit Function
End If
Next i
End If
Error:
End Function

Private Function ForwardMainTalkerStopped(Whom As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = Enn("FREE|||" & Whom)
IgnoreCount = 0
MainTalkerDomination = 0
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = Whom Then GoTo Skip
If Ws(i).State = 7 Then Ws(i).SendData Pck
Skip:
Next i
DoEvents
End Function

Private Function ForwardSecondTalkerStopped(Whom As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = Enn("FREE|||" & Whom)
IgnoreCount2 = 0
SecondTalkerDomination = 0
DetectSecondDC = 0
SecondTalker = ""
SecondIndex = 0
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If VCID(i) = MainTalker Then
If Ws(i).State = 7 Then Ws(i).SendData Pck
End If
Next i
DoEvents
End Function

Private Function ForwardUserJoin(Whom As String, Indy As Integer)
'On Error Resume Next
Dim i As Integer
Dim Pck As String
'Forward new user to all in voice.
Pck = Enn("JOIN|||" & Whom)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If i = Indy Then GoTo Skip
If Ws(i).State = 7 Then Ws(i).SendData Pck
Pause4 "0.005"
Skip:
Next i
DoEvents
'Forward room list to new user
Pck = Enn("LIST|||" & VoiceUserList)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Indy).State = 7 Then
Ws(Indy).SendData Pck
End If
End Function

Private Sub Timer1_Timer() 'Important leave this, it fixs and clear Maintalker, in the event they DC from socket while on AIR making other able to talk again
'On Error Resume Next
If Timer1 = False Then Exit Sub
Timer1.interval = 4000
If MainTalker = "" Then

ElseIf SecondTalker = "" Or SecondIndex = 0 Then
'Call Status("<----->")
Debug.Print "MainTalker Stopped By Timeout Timer!! " & MainTalker & " - " & MainIndex
ForwardMainTalkerStopped MainTalker, 0 'Call Forward to all That User Stopped Talking yet to do this
MainTalker = ""
MainIndex = 0
ListenerCount = 0
IgnoreCount = 0
IgnoreCount2 = 0
MainTalkerDomination = 0
SecondTalkerDomination = 0
Else 'Shift Second Talker into MainTalker position as Maintalker has timed out on audio data incoming
Debug.Print "New MainTalker Moved From SecondTalker!! " & MainTalker & " - " & MainIndex & " <---> " & SecondTalker & " - " & SecondIndex
IgnoreCount = 0
MainTalkerDomination = 0
DetectSecondDC = 0
MainIndex = SecondIndex
SecondIndex = 0
MainTalker = SecondTalker
IgnoreCount2 = 0
SecondTalkerDomination = 0
SecondTalker = ""
'Call Status("<-" & MainTalker & "->")
'Call Status2("<----->")
Exit Sub 'keep timer running
End If
Timer1 = False
End Sub

Private Sub Timer2_Timer()
'On Error Resume Next
Timer2 = False
PingCount = PingCount + 1
If PingCount = 120 Then
PingCount = 0
PingEveryone
End If
Timer2 = True
End Sub

Private Function PingEveryone()
'On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = Enn("PING|||KEEPALIVE")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If Ws(i).State = 7 Then
Ws(i).SendData Pck
Pause2 "0.02"
End If
Next i
DoEvents
End Function

Private Sub Pause2(interval)
'On Error Resume Next
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Private Sub Pause3(interval)
'On Error Resume Next
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Private Sub Pause4(interval)
'On Error Resume Next
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

