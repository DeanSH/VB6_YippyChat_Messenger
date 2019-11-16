VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Chatrooms Manager - The Beast"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog C 
      Left            =   2520
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Text5 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":000D
      TabIndex        =   17
      Text            =   "ychat1.ath.cx"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "?"
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4920
      TabIndex        =   15
      Text            =   "FindWho"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame F2 
      Caption         =   "Rooms: 0"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   3615
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Text            =   "10"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Text            =   "1"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Load"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Count"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Public Chat"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Gen"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   2680
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Text            =   "4051"
      Top             =   120
      Width           =   615
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
         Sorted          =   -1  'True
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
      Text            =   "50"
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
      Width           =   4695
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
ServerIP = Text5.Text
Ws.Connect ServerIP, Text2
End Sub

Private Sub Command10_Click()
On Error Resume Next
List2.Clear
F2.Caption = "Rooms: " & List2.ListCount
LoadList C, List2
DoEvents
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command2_Click()
On Error Resume Next
ServerON = False
Command2.Enabled = False
Dim i As Integer
For i = 0 To List2.ListCount - 1
If InStr(1, List2.List(i), ":") > 0 Then
RoomNames(i) = ""
UserList(i) = "~"
End If
Next i
DoEvents
Ws.Close
DoEvents
Timer1 = False
DoEvents
Call Status("Status: Rooms Server Closed!!")
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
Dim Sdata() As String
Dim i As Integer
For i = 0 To List2.ListCount - 1
If LCase(RoomNames(i)) = LCase(Text3.Text) Then
TmpList = LCase(UserList(i))
GoTo Skip
End If
Next i
DoEvents
Skip:
Sdata = Split(TmpList, "~")
For i = 0 To UBound(Sdata)
If Len(Sdata(i)) > 0 Then
List1.AddItem Sdata(i)
End If
Next i
DoEvents
F1.Caption = "Users: " & List1.ListCount
End Sub

Private Sub Command4_Click()
Dim i As Integer
For i = Text6.Text To Text7.Text '1600
List2.AddItem Text3.Text & ":" & i
Next i
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command5_Click()
List2.Clear
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command6_Click()
On Error Resume Next
SaveList C, List2
DoEvents
End Sub

Private Sub Command7_Click()
List2.AddItem Text3.Text
F2.Caption = "Rooms: " & List2.ListCount
End Sub

Private Sub Command8_Click()
On Error Resume Next
If List2.ListCount = 0 Then Exit Sub
If Command2.Enabled = False Then Exit Sub
Dim TmpList As String
'List1.Clear
'Dim SData() As String
Dim i As Integer
For i = 0 To List2.ListCount - 1
If InStr(1, "~" & LCase(UserList(i)), "~" & LCase(Text4.Text) & "~") > 0 Then
Call Status(Text4.Text & " found in " & RoomNames(i))
Text3.Text = RoomNames(i)
GoTo Skip
End If
Next i
DoEvents
Call Status(Text4.Text & " Not Found!")
Exit Sub
Skip:
Command3_Click
End Sub

Private Sub Command9_Click()
If Command1.Enabled = True Then Exit Sub
Dim i As Integer
Dim TmpList As Long
TmpList = 0
For i = 0 To List2.ListCount - 1
If InStr(1, List2.List(i), ":") > 0 Then
TmpList = TmpList + GetCount(UserList(i))
End If
Next i
DoEvents
Label2.Caption = TmpList
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub List2_Click()
If List2.ListCount = 0 Then Exit Sub
Text3.Text = List2.Text
End Sub

Private Sub List2_DblClick()
If List2.ListCount = 0 Then Exit Sub
If Command1.Enabled = False Then
Command3_Click
Else
List2.RemoveItem List2.ListIndex
F2.Caption = "Rooms: " & List2.ListCount
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Timer1 = False Then Exit Sub
If Ws.State <> 7 Then
Ws.Close
Ws.Connect ServerIP, Text2
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
Call Status("Status: Room Manager Re-Connected With Main Server!")
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
Limit = Text1
Dim i As Integer
For i = 0 To List2.ListCount - 1
If InStr(1, List2.List(i), ":") > 0 Then
RoomNames(i) = List2.List(i)
UserList(i) = "~"
Else

End If
Next i
DoEvents
Command2.Enabled = True
Call Status("Status: Rooms-Server Started!! && Connected With Main Server!")
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Timer1 = True
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
Debug.Print "Main Server Arrival: " & TmpData
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
Ws.GetData TmpData
End Sub

Public Function ProcessMain(VCDATA As String)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Colar As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)
Dim i As Integer

Select Case Casee
Case "PING"
Pck = "PING|||" & ServerIP & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents

Case "JOIN"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
Colar = Split(VCDATA, "|||")(3)
If Who = "" Then
Call Status(Who & " Rejected!")
GoTo OhNo
Else
'Indy = Split(VCDATA, "|||")(3)
For i = 0 To List2.ListCount - 1
If LCase(RoomNames(i)) = LCase(Room) Then
If InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "~") > 0 Or InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "<") > 0 Then
Call Status(Who & " Allready in " & RoomNames(i))
OhNo:
Pck = "BADD|||" & Who & "|||" & Room & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
Else
If GetCount(UserList(i)) >= Limit Then
Pck = "FULL|||" & Who & "|||" & Room & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
Else
Call Status(Who & " Joined " & RoomNames(i))
If Colar = "" Then
UserList(i) = UserList(i) & "~" & Who & "~"
Else
UserList(i) = UserList(i) & "~" & Who & "<" & Colar & ">~"
End If
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "JOIN|||" & Who & "|||" & RoomNames(i) & "|||" & UserList(i) & "|||" & Colar & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
End If
End If
Exit Function
End If
Next i
DoEvents
GoTo OhNo
End If

Case "EXIT"
Who = Split(VCDATA, "|||")(1)
For i = 0 To List2.ListCount - 1
If InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "~") > 0 Then
Call Status(Who & " Left " & RoomNames(i))
UserList(i) = Replace(UserList(i), "~" & Who & "~", "~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "LEFT|||" & Who & "|||" & RoomNames(i) & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
ElseIf InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "<") > 0 Then
Call Status(Who & " Left " & RoomNames(i))
Room = Split(UserList(i), "~" & Who & "<")(1)
Room = Split(Room, ">")(0)
UserList(i) = Replace(UserList(i), "~" & Who & "<" & Room & ">~", "~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "LEFT|||" & Who & "|||" & RoomNames(i) & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
End If
Next i
DoEvents

Case "LEFT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
For i = 0 To List2.ListCount - 1
If InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "~") > 0 Then
Call Status(Who & " Left " & RoomNames(i))
UserList(i) = Replace(UserList(i), "~" & Who & "~", "~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "LEFT|||" & Who & "|||" & RoomNames(i) & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
ElseIf InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "<") > 0 Then
Call Status(Who & " Left " & RoomNames(i))
Room = Split(UserList(i), "~" & Who & "<")(1)
Room = Split(Room, ">")(0)
UserList(i) = Replace(UserList(i), "~" & Who & "<" & Room & ">~", "~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "LEFT|||" & Who & "|||" & RoomNames(i) & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
End If
Next i
DoEvents

Case "COLR"
Who = Split(VCDATA, "|||")(1)
Colar = Split(VCDATA, "|||")(3)
If Colar = "" Then Exit Function
For i = 0 To List2.ListCount - 1
If InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "~") > 0 Then
'Call Status(Who & " Left " & RoomNames(i))
UserList(i) = Replace(UserList(i), "~" & Who & "~", "~" & Who & "<" & Colar & ">~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "COLR|||" & Who & "|||" & RoomNames(i) & "|||" & Colar & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
ElseIf InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "<") > 0 Then
'Call Status(Who & " Left " & RoomNames(i))
Room = Split(UserList(i), "~" & Who & "<")(1)
Room = Split(Room, ">")(0)
UserList(i) = Replace(UserList(i), "~" & Who & "<" & Room & ">~", "~" & Who & "<" & Colar & ">~")
UserList(i) = Replace(UserList(i), "~~", "~")
Pck = "COLR|||" & Who & "|||" & RoomNames(i) & "|||" & Colar & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
DoEvents
End If
Next i
DoEvents

Case "ROMS"
Who = Split(VCDATA, "|||")(1)
Indy = Split(VCDATA, "|||")(2)
SendRooms Who, Indy

Case "CHAT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If Who = "" Or LCase(Who) = "admin" Then
'Call Status(Who & " Rejected chat!")
Else
For i = 0 To List2.ListCount - 1
If LCase(RoomNames(i)) = LCase(Room) Then
If InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "~") > 0 Or InStr(1, LCase("~" & UserList(i)), "~" & LCase(Who) & "<") > 0 Then
'Call Status(Who & " Allready in " & RoomNames(i))
Ws.SendData VCDATA
DoEvents
End If
Exit Function
End If
Next i
DoEvents
End If

Case Else

End Select
End Function

Private Sub SendRooms(Whom As String, Index As Integer)
'On Error Resume Next
Dim i As Integer
Dim TmpList As String
Dim TmpList2 As Long
Dim Pck As String
TmpList = "~"
TmpList2 = 0
For i = 0 To List2.ListCount - 1
If InStr(1, List2.List(i), ":") > 0 Then
TmpList = TmpList & RoomNames(i) & "|" & GetCount(UserList(i)) & "~"
TmpList2 = TmpList2 + GetCount(UserList(i))
Else
TmpList = TmpList & List2.List(i) & "~"
End If
Next i
DoEvents
Pck = "ROMS|||" & Whom & "|||" & Index & "|||" & TmpList & "|||" & TmpList2 & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws.SendData Pck
End Sub

Private Function GetCount(TheList) As Integer
'On Error Resume Next
Dim Sdata() As String
Dim TmpCount As Integer
Sdata() = Split(TheList, "~")
'TmpCount = UBound(Sdata()) - 1
'If TmpCount < 0 Then
'TmpCount = 0
'End If
GetCount = UBound(Sdata()) - 1
End Function

Private Sub Ws_Close()
On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
End If
End Sub

Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Timer1.Enabled = False Then
Call Status("Failed Connect To Main Server!")
Else
Call Status("Disconnected From Main Server!")
End If
End Sub
