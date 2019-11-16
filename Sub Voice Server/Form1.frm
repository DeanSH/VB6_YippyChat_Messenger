VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sub Voice Server - Vc Channels"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog C 
      Left            =   2640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Text5 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5400
      List            =   "Form1.frx":000D
      TabIndex        =   18
      Text            =   "ychat1.dyndns.tv"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      ItemData        =   "Form1.frx":003E
      Left            =   3120
      List            =   "Form1.frx":0048
      TabIndex        =   17
      Text            =   "192.161.59.152"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<Set"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame F2 
      Caption         =   "Rooms: 0"
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   3615
      Begin VB.CommandButton Command7 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "Public Chat:"
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "5000"
      Top             =   120
      Width           =   735
   End
   Begin Project1.VcChannel VcChannel1 
      Height          =   375
      Index           =   0
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
      Caption         =   "Get Room List"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame F1 
      Caption         =   "Users: 0"
      Height          =   3015
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   2775
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "60"
      Top             =   120
      Width           =   495
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
If List2.ListCount = 0 Then Exit Sub
ServerON = False
'If List2.ListCount > 200 Then Exit Sub
Ws.Close
Timer1 = False
DoEvents
Ws.Connect Text5.Text, 4999
End Sub

Private Sub Command2_Click()
On Error Resume Next
ServerON = False
Command2.Enabled = False
Dim i As Integer
For i = 0 To List2.ListCount - 1
RoomNames(i) = ""
RoomPort1(i) = 0
VcChannel1(i).StopChannel
DoEvents
Unload VcChannel1(i)
Next i
DoEvents
Ws.Close
DoEvents
Timer1 = False
DoEvents
Call Status("Status: Sub Server Channels Closed!!")
List1.Clear
F1.Caption = "Users: " & List1.ListCount
Text1.Enabled = True
Text2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command1.Enabled = True
End Sub


Private Sub Command3_Click()
On Error Resume Next
If List2.ListCount = 0 Then Exit Sub
If Command2.Enabled = False Then Exit Sub
Dim TmpList As String
List1.Clear
Dim SData() As String
Dim i As Integer
For i = 0 To List2.ListCount - 1
If LCase(RoomNames(i)) = LCase(Text3.Text) Then
TmpList = VcChannel1(i).VcList
GoTo Skip
End If
Next i
DoEvents
Skip:
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
List2.Clear
LoadList C, List2
DoEvents
'Dim i As Integer
'For i = Text6 To (Text6 + 29) '1600
'List2.AddItem Text3.Text & i
'Next i
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command5_Click()
On Error Resume Next
List2.Clear
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command6_Click()
On Error Resume Next
SaveList C, List2
End Sub

Private Sub Command7_Click()
On Error Resume Next
List2.AddItem Text3.Text
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command8_Click()
On Error Resume Next
ServerIP = Text4.Text
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub List2_Click()
On Error Resume Next
If List2.ListCount = 0 Then Exit Sub
Text3.Text = List2.Text
End Sub

Private Sub List2_DblClick()
On Error Resume Next
If List2.ListCount = 0 Then Exit Sub
If Command1.Enabled = False Then
Command3_Click
Else
List2.RemoveItem List2.ListIndex
F2.Caption = "Rooms: " & List2.ListCount
End If
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
If Timer1 = False Then Exit Sub
If Ws.State <> 7 Then
Ws.Close
Ws.Connect Text5.Text, 4999
End If
DoEvents
If Timer1 = False Then Exit Sub
Timer1.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Ws_Connect()
'On Error Resume Next
Dim Pck As String
If ServerON = True Then
Call Status("Status: Sub-Server Re-Connected With Main Server!")
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
Exit Sub
End If
ServerON = True
Command1.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
ServerIP = Text4.Text
Dim i As Integer
For i = 0 To List2.ListCount - 1
RoomNames(i) = List2.List(i)
RoomPort1(i) = (i * 3) + (Text2 + 3)
On Error Resume Next
Load VcChannel1(i)
DoEvents
VcChannel1(i).StartChannel Text1, RoomNames(i), RoomPort1(i)
Pause "0.01"
Next i
DoEvents
Command2.Enabled = True
Call Status("Status: Sub-Server Started!! && Connected With Main Server!")
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Timer1 = True
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
Debug.Print "PRE-VOICE: " & TmpData
ProcessVoice TmpData
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
Ws.GetData TmpData
End Sub

Public Function ProcessVoice(VCDATA As String)
'On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee
Case "PING"
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents

Case "PORT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
'If Who = "" Or LCase(Who) = "admin" Then
'Call Status(Who & " Rejected!")
'Else
Indy = Split(VCDATA, "|||")(3)
Dim i As Integer
For i = 0 To List2.ListCount - 1
If LCase(RoomNames(i)) = LCase(Room) Then
Call Status(Who & " Joining Channel " & Room & "!")
Pck = "PORT|||" & Who & "|||" & RoomNames(i) & "|||" & Indy & "|||" & ServerIP & "|||" & RoomPort1(i)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
End If
Next i
DoEvents
'End If

Case Else

End Select
End Function

Private Sub Ws_Close()
'On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
End If
End Sub

Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
End If
End Sub
