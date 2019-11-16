VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl LoginChannel 
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Ls 
      Left            =   1440
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
Attribute VB_Name = "LoginChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Limit As Integer
Private VCID(1 To 9999) As String
Private RoomName(1 To 9999) As String
'Private VCIndex(1 To 10000) As Integer
'Private VCIgnores(0 To 10000) As String
Private LoginUserList As String
'Private MainTalker As String
'Private MainIndex As Integer
'Private SecondTalker As String
'Private SecondIndex As Integer
'Private SendListenerCounter As Long
'Private SendListenerCounter2 As Long
'Private ListenCount As Long
'Private IgnoreCount As Long
Private PingCount As Long
Private ServeIP As String
Private Port1 As String
'Private Port2 As String

Private Sub UserControl_Initialize()
LoginUserList = "~"
End Sub

Public Function VcList() As String
VcList = LoginUserList
End Function


Public Sub StartChannel(LimitNum As Integer, RoomN As String, VcPort As String)
On Error Resume Next
'MainTalker = ""
'SendListenerCounter = 0
Dim i As Integer
Limit = LimitNum
ServeIP = RoomN
Port1 = VcPort
'Port2 = VcPort + 1
For i = 1 To Limit
VCID(i) = ""
RoomName(i) = ""
'VCIndex(i) = 0
'VCIgnores(i) = ""
Load Ws(i)
Next i
DoEvents
Timer1 = False
LoginUserList = "~"
Ls.Close
Ls.LocalPort = Port1
Ls.Listen
PingCount = 0
Timer2 = True
End Sub

Public Sub StopChannel()
On Error Resume Next
Timer2 = False
PingCount = 0
Ls.Close
Timer1 = False
Dim i As Integer
For i = 1 To Limit
VCID(i) = ""
RoomName(i) = ""
'VCIndex(i) = 0
'VCIgnores(i) = ""
Ws(i).Close
Unload Ws(i)
Next i
DoEvents
LoginUserList = "~"
'MainTalker = ""
'SecondTalker = ""
'MainIndex = 0
'SecondIndex = 0
ServeIP = ""
Port1 = 0
'Port2 = 0
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
On Error GoTo Error
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
Debug.Print "User Socket: " & TmpData
ProcessUser TmpData, Index
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
On Error Resume Next
If Ws(Index).State = 7 Then Ws(Index).GetData TmpData
End Sub

Public Function ProcessUser(VCDATA As String, Index As Integer)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, SData() As String, TotalCount As Long
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "AUTH"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & LoginUserList, "~" & Who & "~") > 0 Or Who = "" Then
BADDY:
Ws(Index).Close
Call Status(Who & " already in Server, Rejected!")
Else
If LCase(Who) = LCase(NextLog) Then
NextLog = ""
VCID(Index) = Who
Pck = Enn("AUTH|||" & VCID(Index) & "|||" & NextBud & "|||" & Ws(Index).RemoteHostIP & "|||")
NextBud = ""
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Index).State = 7 Then Ws(Index).SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
Pck = Enn("IGYS|||" & VCID(Index) & "|||" & NextIgs & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Index).State = 7 Then Ws(Index).SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
NextIgs = ""
Pck = Enn("ONLS|||" & VCID(Index) & "|||" & NextOns & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Index).State = 7 Then Ws(Index).SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
NextOns = ""
Call Status(VCID(Index) & " Logged In!")
RoomName(Index) = ""
LoginUserList = LoginUserList & "~" & Who & "~"
LoginUserList = Replace(LoginUserList, "~~", "~")
Debug.Print "Users:: " & LoginUserList
Pck = "LOGN|||" & VCID(Index) & "|||" & Ws(Index).RemoteHostIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck 'sent packet telling new user they are in voice and need to connect audio socket now!
DoEvents
'ForwardUserLogin Who ', Index 'Yet to code this where it tells everyone new user joined and then send whoel list to the new user!
Else ' Bad User Not Meant TO Be Logging In
GoTo BADDY
End If
End If

Case "PING"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "ADDD"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "DENY"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "ACPT"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "FILE"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "XFIL"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "SFIL"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

Case "PROF"
If VCID(Index) = "" Then Exit Function
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA

            Case "CAMV"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case "DESK"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case "CALL"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case "CALC"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case "CALA"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case "CALD"
                If VCID(Index) = "" Then Exit Function
                If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
                
            Case Else
If VCID(Index) = "" Then Exit Function
Who = Split(VCDATA, "|||")(1)
If Who = VCID(Index) Then
If Form1.Ws.State = 7 Then Form1.Ws.SendData VCDATA
End If


End Select
End Function

Private Sub Ws_Close(Index As Integer)
On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
Dim Pck As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
RoomName(Index) = ""
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & LoginUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Logged Out!!") 'Status
'ForwardUserLogout TmpID
Pck = "EXIT|||" & TmpID & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck
End If
'Update VoiceUserList to remove user who left from the string!
LoginUserList = Replace(LoginUserList, "~" & TmpID & "~", "~")
LoginUserList = Replace(LoginUserList, "~~", "~")
Debug.Print "Users:: " & LoginUserList
End Sub

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If VCID(Index) = "" Then Exit Sub 'If VCID Blank then its not user leaving voice!
Dim TmpID As String
Dim Pck As String
TmpID = VCID(Index)
VCID(Index) = "" 'clear variable storing username for this index of it
RoomName(Index) = ""
'VCIgnores(Index) = ""
'forward user left voice here if user name still in VoiceUserList String variable!
If InStr(1, "~" & LoginUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " Logged Out!!") 'Status
'ForwardUserLogout TmpID
Pck = "EXIT|||" & TmpID & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck
End If
'Update VoiceUserList to remove user who left from the string!
LoginUserList = Replace(LoginUserList, "~" & TmpID & "~", "~")
LoginUserList = Replace(LoginUserList, "~~", "~")
Debug.Print "Users:: " & LoginUserList
End Sub


'''''''''''''''''''''''''''''''''''' Subs n Functions n Timers to be called an used '''''''''''''''''''''''''''''''

Public Function SendToAll(Packet As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
'If VCID(i) = Whom Then
If Ws(i).State = 7 Then
Ws(i).SendData Packet
Pause2 "0.001"
'Exit Function
End If
Next i
End Function

Public Function SendToAllJoin(Packet As String, Whom As String, Roomy As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
If VCID(i) = Whom Then
If Ws(i).State = 7 Then
RoomName(i) = Roomy
Ws(i).SendData Packet
Pause4 "0.001"
End If
Else
If LCase(RoomName(i)) = LCase(Roomy) Then
If Ws(i).State = 7 Then
Ws(i).SendData Packet
Pause3 "0.004"
End If
End If
End If
Next i
DoEvents
End Function

Public Function SendToAllLeft(Packet As String, Whom As String, Roomy As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
If VCID(i) = Whom Then
'If Ws(i).State = 7 Then
RoomName(i) = ""
'Ws(i).SendData Packet
'End If
Else
If LCase(RoomName(i)) = LCase(Roomy) Then
If Ws(i).State = 7 Then
Ws(i).SendData Packet
Pause4 "0.001"
End If
End If
End If
Next i
DoEvents
End Function

Public Function SendToAllInRoom(Packet As String, Whom As String, Roomy As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
If LCase(RoomName(i)) = LCase(Roomy) Then
If Ws(i).State = 7 Then
Ws(i).SendData Packet
Pause3 "0.001"
End If
End If
Next i
DoEvents
End Function

Public Function ForwardStat(Whom As String, Roomy As String, Packet As String)
On Error Resume Next
Dim i As Integer
Dim Pck As String
Pck = Enn("STAT|||" & Whom & "|||" & Roomy & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If InStr(1, "~" & LCase(Packet) & "~", "~" & LCase(VCID(i)) & "~") > 0 Then
If Ws(i).State = 7 Then
Ws(i).SendData Pck
Pause3 "0.001"
End If
End If
Next i
DoEvents
End Function

Public Function SendToAllExit(Packet As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
If RoomName(i) = "" Then

Else
RoomName(i) = ""
If Ws(i).State = 7 Then
Ws(i).SendData Packet
Pause3 "0.001"
End If
End If
Next i
DoEvents
End Function

Public Function KickUser(Whom As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Dim TmpID As String
Dim Pck As String
Pck = Enn("FAIL|||" & Whom & "|||Disconnected Because Your ID Logged In Somewhere Else!|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To Limit
If LCase(VCID(i)) = LCase(Whom) Then
If Ws(i).State = 7 Then
Ws(i).SendData Pck
TmpID = VCID(i)
VCID(i) = ""
RoomName(i) = ""
If InStr(1, "~" & LoginUserList, "~" & TmpID & "~") > 0 Then
Call Status("Status: " & TmpID & " D/C!!") 'Status
'ForwardUserLogout TmpID
Pck = "EXIT|||" & TmpID & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck
End If
'Update VoiceUserList to remove user who left from the string!
LoginUserList = Replace(LoginUserList, "~" & TmpID & "~", "~")
LoginUserList = Replace(LoginUserList, "~~", "~")
Debug.Print "Users:: " & LoginUserList
Pause "0.5"
Ws(i).Close
End If
Exit Function
End If
Next i
DoEvents
End Function

Public Function SendToOne(Packet As String, Whom As String) 'Find user in Ws Sockets via VCID indexing checking and send it listner count
On Error Resume Next
Dim i As Integer
Packet = Enn(Mid(Packet, 11, Len(Packet) - 10))
Packet = "R4R4" & Chr(0) & Chr$(Int(Len(Packet) / 256)) & Chr$(Len(Packet) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Packet
For i = 1 To Limit
If LCase(VCID(i)) = LCase(Whom) Then
If Ws(i).State = 7 Then
Ws(i).SendData Packet
End If
Exit Function
End If
Next i
DoEvents
End Function

Private Function ForwardUserLogin(Whom As String)
On Error Resume Next
Dim Pck As String
'Forward new user to all in.
Pck = "ONLN|||" & Whom & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck
End Function

Private Function ForwardUserLogout(Whom As String)
On Error Resume Next
Dim Pck As String
'Forward new user to all in.
Pck = "LGOF|||" & Whom & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Form1.Ws.State = 7 Then Form1.Ws.SendData Pck
End Function

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
Pck = Enn("PING|||KEEPALIVE")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
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

Public Function Ubounds() As Integer
On Error GoTo Error
Ubounds = Ws().Count
Exit Function
Error:
Ubounds = 0
End Function
