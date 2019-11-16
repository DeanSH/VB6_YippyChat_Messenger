VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Sub Login Server - Login Channels"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "R"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      ToolTipText     =   "Reconnect Now"
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox Text5 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5400
      List            =   "Form1.frx":000D
      TabIndex        =   11
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      ItemData        =   "Form1.frx":003B
      Left            =   3120
      List            =   "Form1.frx":0045
      TabIndex        =   10
      Text            =   "192.161.59.152"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<Set"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "4052"
      Top             =   120
      Width           =   615
   End
   Begin Project1.LoginChannel LoginChannel1 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update Users"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame F1 
      Caption         =   "Users: 0"
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6495
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "5000"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   2040
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




''''''''''''''''''''Buttons.............''''''''''''''


Private Sub Command1_Click()
On Error Resume Next
Timer1 = False
ServerON = False
Ws.Close
DoEvents
ServerIP = ""
ServerPort = 0
NextLog = ""
Ws.Connect Text5.Text, 4998
Timer1 = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
ServerON = False
Timer1 = False
Command2.Enabled = False
Ws.Close
DoEvents
ServerIP = ""
ServerPort = 0
NextLog = ""
List1.Clear
F1.Caption = "Users: " & List1.ListCount
LoginChannel1.StopChannel
DoEvents
Call Status("Status: Sub Server Channels Closed!!")
Text1.Enabled = True
Text2.Enabled = True
Command1.Enabled = True
Debug.Print "Sockets: " & LoginChannel1.Ubounds
End Sub


Private Sub Command3_Click()
On Error Resume Next
If Command2.Enabled = False Then Exit Sub
Dim TmpList As String
List1.Clear
Dim SData() As String
Dim i As Integer
TmpList = LoginChannel1.VcList
SData = Split(TmpList, "~")
For i = 0 To UBound(SData)
If Len(SData(i)) > 0 Then
List1.AddItem SData(i)
End If
Next i
DoEvents
F1.Caption = "Users: " & List1.ListCount
End Sub

Private Sub Command4_Click()
On Error Resume Next
Timer1.Enabled = False
Ws.Close
Ws.Connect Text5.Text, 4998
Timer1.Enabled = True
End Sub

Private Sub Command8_Click()
'LoginChannel1.SendToNone "lol", "Deano"
ServerIP = Text4.Text
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Timer1 = False Then Exit Sub
If Ws.State <> 7 Then
Ws.Close
Ws.Connect Text5.Text, 4998
End If
DoEvents
If Timer1 = False Then Exit Sub
Timer1.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Ws_Connect()
On Error Resume Next
Dim Pck As String
If ServerON = True Then
Call Status("Status: Sub-Server Re-Connected With Main Server!")
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws.State = 7 Then Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
Exit Sub
End If
ServerON = True
Command1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
ServerIP = Text4.Text
ServerPort = Text2
Pck = "PING|||" & ServerIP & "|||" & ServerPort & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Timer1 = True
LoginChannel1.StartChannel Text1, ServerIP, ServerPort
DoEvents
Command2.Enabled = True
Call Status("Status: Sub-Server Started!! && Connected With Main Server!")
Debug.Print "Sockets: " & LoginChannel1.Ubounds
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
Debug.Print "Main Server: " & TmpData
ProcessMain TmpData
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
If Ws.State = 7 Then Ws.GetData TmpData
End Sub

Public Function ProcessMain(VCDATA As String)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee
Case "PING"
Pck = "PING|||" & ServerIP & "|||" & ServerPort & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws.State = 7 Then Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents

Case "GOOD"
Who = Split(VCDATA, "|||")(1)
Indy = CountUsers(LoginChannel1.VcList)
Room = Split(VCDATA, "|||")(2)

If Who = "1" Then GoTo Skit
If Who = "2" Then GoTo Skit
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Who) & "~") > 0 Then
LoginChannel1.KickUser Who
End If
Skit:
If Who = "2" Then
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Room) & "~") > 0 Then
Pck = "GOOD|||" & Room & "|||" & Indy & "|||" & ServerIP & "|||" & ServerPort & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws.State = 7 Then Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
End If
End If
Pck = "GOOD|||" & Who & "|||" & Indy & "|||" & ServerIP & "|||" & ServerPort & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws.State = 7 Then Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
'End If
'End If

Case "LOGG"
NextLog = Split(VCDATA, "|||")(1)
NextBud = Split(VCDATA, "|||")(2)
NextIgs = Split(VCDATA, "|||")(3)
NextOns = Split(VCDATA, "|||")(4)

Case "STAT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Pck = Split(VCDATA, "|||")(3)
If Who = "" Then Exit Function
If Room = "" Then Exit Function
If Pck = "" Then Exit Function
LoginChannel1.ForwardStat Who, Room, Pck

Case "PMIM"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
SubIP = Split(VCDATA, "|||")(3)
If SubIP = "" Then Exit Function
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Room) & "~") > 0 Then
LoginChannel1.SendToOne VCDATA, Room
End If

Case "FULL"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Who) & "~") > 0 Then
LoginChannel1.SendToOne VCDATA, Who
End If

Case "ROMS"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Who) & "~") > 0 Then
LoginChannel1.SendToOne VCDATA, Who
End If

Case "BADD"
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Who) & "~") > 0 Then
LoginChannel1.SendToOne VCDATA, Who
End If

Case "EXIT"
LoginChannel1.SendToAllExit VCDATA
DoEvents

Case "JOIN"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
LoginChannel1.SendToAllJoin VCDATA, Who, Room
DoEvents

Case "CHAT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
SubIP = Split(VCDATA, "|||")(3)
If SubIP = "" Then Exit Function
LoginChannel1.SendToAllInRoom VCDATA, Who, Room
DoEvents

Case "COLR"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
SubIP = Split(VCDATA, "|||")(3)
If SubIP = "" Then Exit Function
LoginChannel1.SendToAllInRoom VCDATA, Who, Room
DoEvents

Case "LEFT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
LoginChannel1.SendToAllLeft VCDATA, Who, Room
DoEvents

Case Else
Who = Split(VCDATA, "|||")(1)
If InStr(1, "~" & LCase(LoginChannel1.VcList), "~" & LCase(Who) & "~") > 0 Then
LoginChannel1.SendToOne VCDATA, Who
'Else
'LoginChannel1.SendToAll VCDATA
End If

End Select
End Function

Private Function CountUsers(Allnames As String) As String
On Error Resume Next
Dim SData() As String
Dim ii As Integer
ii = 0
SData = Split(Allnames, "~")
ii = UBound(SData) - 2
CountUsers = ii
End Function

Private Sub Ws_Close()
On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
Timer1.Enabled = False
Ws.Close
Ws.Connect Text5.Text, 4998
Timer1.Enabled = True
End If
End Sub

Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
Timer1.Enabled = False
Ws.Close
Ws.Connect Text5.Text, 4998
Timer1.Enabled = True
End If
End Sub
