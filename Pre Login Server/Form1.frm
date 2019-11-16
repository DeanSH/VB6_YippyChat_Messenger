VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre-Login Server"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "R"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      ToolTipText     =   "Reconnect"
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3000
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   1560
      Top             =   480
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Text            =   "4000"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox Text5 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3360
      List            =   "Form1.frx":000D
      TabIndex        =   4
      Text            =   "ychat1.dyndns.tv"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "3001"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "20"
      Top             =   120
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Ls 
      Left            =   2520
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   2040
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws2 
      Left            =   3480
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
      TabIndex        =   6
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
ServerON = False
'If List2.ListCount > 200 Then Exit Sub
Ws2.Close
Timer2 = False
DoEvents
Ws2.Connect Text5.Text, Text3.Text
Timer2 = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
ServerON = False
Command2.Enabled = False
Timer2 = False
Dim i As Integer
For i = 1 To Text1
Ws(i).Close
DoEvents
Unload Ws(i)
Next i
DoEvents
Ws2.Close
Ls.Close
DoEvents
Timer2 = False
DoEvents
Call Status("Status: Sub Server Channels Closed!!")
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
Timer2.Enabled = False
Ws2.Close
Ws2.Connect Text5.Text, Text3.Text
Timer2.Enabled = True
End Sub

Private Sub Ls_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To Text1
If Ws(i).State <> 7 Then
Ws(i).Close
Ws(i).Accept requestID
Timer1(i).Enabled = True
Debug.Print "Accepted New Connection"
Ls.Close
Ls.LocalPort = Text2.Text
Ls.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Server Full"
Ls.Close
Ls.LocalPort = Text2.Text
Ls.Listen
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Error Resume Next
Ws(Index).Close
Timer1(Index).Enabled = False
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
Debug.Print "PRE-Login: " & TmpData
ProcessLogin TmpData, Index
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

Public Function ProcessLogin(VCDATA As String, Index As Integer)
On Error Resume Next
If Ws2.State <> 7 Then Exit Function
Dim Who As String, Pck As String, Casee As String, Room As String, PassW As String, Buds As String, Ignors As String, Onlines As String, OffData As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "LOGG"
Dim i As Integer
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2) 'Password
OffData = Split(VCDATA, "|||")(3) 'Password
Pck = "LOGG|||" & Who & "|||" & Room & "|||" & OffData & "|||" & Index & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents

Case Else
Timer1(Index).Enabled = False
Ws(Index).Close

End Select
End Function

Private Sub Ws_Close(Index As Integer)
On Error Resume Next
Timer1(Index).Enabled = False
End Sub

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Timer1(Index).Enabled = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Ws2_Connect()
On Error Resume Next
Dim Pck As String
If ServerON = True Then
Call Status("Status: Pre-Server Re-Connected With Main Server!")
Pck = "PING|||" & Text5.Text & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws2.State = 7 Then Ws2.SendData Pck 'sent packet telling new user the port for the room to connect to!
Exit Sub
End If
ServerON = True
Command1.Enabled = False
Dim i As Integer
For i = 1 To Text1
Load Ws(i)
Next i
DoEvents
Ls.Close
Ls.LocalPort = Text2.Text
Ls.Listen
Command2.Enabled = True
Call Status("Status: Pre-Server Started!! && Connected With Main Server!")
Pck = "PING|||" & Text5.Text & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws2.State = 7 Then Ws2.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Timer2 = True
End Sub

Private Sub Ws2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws2
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
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
Wend
End With
End Sub

Public Function ProcessMain(VCDATA As String)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee
Case "PING"
Pck = "PING|||" & Text5.Text & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws2.State = 7 Then Ws2.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents

Case "LOGG"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Indy = Split(VCDATA, "|||")(3)
SubIP = Split(VCDATA, "|||")(4)
SubPort = Split(VCDATA, "|||")(5)
'Call Status(Who & " Joining Channel " & Room & "!")
If Ws(Indy).State = 7 Then
Pck = Enn("LOGG|||" & Who & "|||" & Room & "|||" & Indy & "|||" & SubIP & "|||" & SubPort & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Indy).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If

Case "FAIL"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Indy = Split(VCDATA, "|||")(3)
If Ws(Indy).State = 7 Then
Pck = Enn("FAIL|||" & Who & "|||" & Room & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Indy).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If

Case "NEWV"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Indy = Split(VCDATA, "|||")(3)
If Ws(Indy).State = 7 Then
Pck = "NEWV|||" & Who & "|||" & Room & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Indy).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If

Case Else

End Select
End Function

Private Sub Ws2_Close()
On Error Resume Next
If Timer2.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
Timer2.Enabled = False
Ws2.Close
Ws2.Connect Text5.Text, Text3.Text
Timer2.Enabled = True
End If
End Sub

Private Sub Ws2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Timer2.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
Timer2.Enabled = False
Ws2.Close
Ws2.Connect Text5.Text, Text3.Text
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Timer2 = False Then Exit Sub
If Ws2.State <> 7 Then
Ws2.Close
Ws2.Connect Text5.Text, Text3.Text
End If
DoEvents
If Timer2 = False Then Exit Sub
Timer2.Enabled = False
Timer2.Enabled = True
End Sub
