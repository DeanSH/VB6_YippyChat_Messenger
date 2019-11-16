VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Main Vc Server"
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4440
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   3960
      Top             =   480
   End
   Begin MSWinsockLib.Winsock Ws2 
      Index           =   0
      Left            =   3480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ls2 
      Left            =   3000
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find?"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   "Public Chat:1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   1560
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "50"
      Top             =   120
      Width           =   495
   End
   Begin MSWinsockLib.Winsock Ls 
      Left            =   2520
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   2040
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
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
Command1.Enabled = False
Command3.Enabled = False
Text1.Enabled = False
Label2.Caption = 0
Dim i As Integer
For i = 1 To Text1
Load Timer1(i)
Load Timer2(i)
Load Ws(i)
Load Ws2(i)
Next i
DoEvents
Ls.Close
Ls.LocalPort = 5000
Ls2.Close
Ls2.LocalPort = 4999
Ls.Listen
Ls2.Listen
Command2.Enabled = True
Command3.Enabled = True
Timer3 = True
Call Status("Status: Voice Server Started!!")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
Dim i As Integer
Ls.Close
For i = 1 To Text1
Unload Timer1(i)
Unload Timer2(i)
Unload Ws(i)
Unload Ws2(i)
Next i
DoEvents
Call Status("Status: Voice Server Closed!!")
Text1.Enabled = True
Timer3 = False
Label2.Caption = 0
Command3.Enabled = False
Command1.Enabled = True
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim i As Integer
For i = 1 To Text1
If Ws2(i).State = 7 Then
Call Status(Who & " Joining Channel " & Room & "!")
Pck = "PORT|||Admin|||" & Text3.Text & "|||0|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(i).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
Next i
DoEvents
End Sub

Private Sub Form_Load()
'
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
Ls.LocalPort = 5000
Ls.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Server Full"
Ls.Close
Ls.LocalPort = 5000
Ls.Listen
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Error Resume Next
Ws(Index).Close
Timer1(Index).Enabled = False
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim i As Integer
Dim ii As Integer
ii = 0
For i = 1 To Text1
If Ws2(i).State = 7 Then
ii = ii + 1
End If
Next i
DoEvents
If Timer3 = False Then Exit Sub
Timer3 = False
Label2.Caption = ii
Timer3 = True
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
Debug.Print "PRE-VOICE: " & TmpData
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
On Error Resume Next
Ws(Index).GetData TmpData
End Sub

Public Function ProcessVoice(VCDATA As String, Index As Integer)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "P0RT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If Who = "" Or LCase(Who) = "admin" Then
Ws(Index).Close
Call Status(Who & " Rejected!")
Else
Dim i As Integer
For i = 1 To Text1
If Ws2(i).State = 7 Then
Call Status(Who & " Joining Channel " & Room & "!")
Pck = "PORT|||" & Who & "|||" & Room & "|||" & Index & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(i).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
Next i
DoEvents
End If

Case Else
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

Private Sub Ls2_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To Text1
If Ws2(i).State <> 7 Then
Timer2(i).Enabled = False
Ws2(i).Close
Ws2(i).Accept requestID
PingC(i) = 0
Timer2(i).interval = 5000
Timer2(i).Enabled = True
Debug.Print "Accepted New Sub Server Connection"
Ls2.Close
Ls2.LocalPort = 4999
Ls2.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Sub Servers Full"
Ls2.Close
Ls2.LocalPort = 4999
Ls2.Listen
End Sub

Private Sub Ws2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws2(Index)
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
Debug.Print "SubServer: " & TmpData
ProcessSubServer TmpData, Index
DoEvents
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Public Function ProcessSubServer(VCDATA As String, Index As Integer)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "PING"
PingC(Index) = 1

Case "PORT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Indy = Split(VCDATA, "|||")(3)
SubIP = Split(VCDATA, "|||")(4)
SubPort = Split(VCDATA, "|||")(5)
If Who = "Admin" Then
Call Status(Who & " Found Room! " & SubIP & ":" & SubPort)
Else
If Ws(Indy).State = 7 Then
Call Status(Who & " Joining Channel " & Room & "! " & SubIP & ":" & SubPort)
Pck = Enn("PORT|||" & Who & "|||" & Room & "|||" & SubIP & "|||" & SubPort)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Indy).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
End If

Case Else

End Select
End Function

Private Sub Ws2_Close(Index As Integer)
On Error Resume Next
'
End Sub

Private Sub Ws2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
'
End Sub

Private Sub Timer2_Timer(Index As Integer) 'Ping Timer
On Error Resume Next
Timer2(Index).Enabled = False
If PingC(Index) = 0 Then

Ws2(Index).Close
Exit Sub

Else

PingC(Index) = PingC(Index) + 1
If PingC(Index) >= 13 Then

PingC(Index) = 0
Pck = "PING|||STAYALIVE|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws2(Index).State = 7 Then
Ws2(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Else
Ws2(Index).Close
Exit Sub
End If

End If

End If
Timer2(Index).Enabled = True
End Sub
