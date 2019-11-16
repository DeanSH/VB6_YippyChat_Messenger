VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Main Login Server"
   ClientHeight    =   975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   240
      Top             =   600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   2160
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "1.0.0"
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   480
   End
   Begin MSWinsockLib.Winsock Ls3 
      Left            =   1200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find?"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4080
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   3600
      Top             =   960
   End
   Begin MSWinsockLib.Winsock Ws2 
      Index           =   0
      Left            =   4080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ls2 
      Left            =   3600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "UserName"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "20"
      Top             =   120
      Width           =   495
   End
   Begin MSWinsockLib.Winsock Ls 
      Left            =   3120
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
      Left            =   2640
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
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
Option Explicit


''''''''''''''''''''Buttons.............''''''''''''''


Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
Command3.Enabled = False
Text1.Enabled = False
Timer3 = False
Timer4 = False
Label2.Caption = 0
Label3.Caption = 0
TCount = 0
OpenDataBase
Dim i As Integer
For i = 1 To Text1
Load Timer1(i)
If i < 21 Then Load Timer6(i)
If i < 21 Then Load Timer2(i)
LogIP(i) = ""
LogPort(i) = 0
LogCount(i) = 0
Load Ws(i)
If i < 21 Then Load Ws2(i)
Next i
DoEvents
Ls.Close
Ls.LocalPort = 4000
Ls2.Close
Ls2.LocalPort = 4998
Ls3.Close
Ls3.LocalPort = 4051
Ls3.Listen
Ls.Listen
Ls2.Listen
Command2.Enabled = True
Command3.Enabled = True
Timer3 = True
Call Status("Status: Main Login Server Started!!")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
Dim i As Integer
Ls.Close
Ls3.Close
For i = 1 To Text1
Unload Timer1(i)
If i < 21 Then Unload Timer2(i)
If i < 21 Then
Timer6(i) = False
TString(i) = ""
TIndex(i) = 0
Unload Timer6(i)
End If
LogIP(i) = ""
LogPort(i) = 0
LogCount(i) = 0
Unload Ws(i)
If i < 21 Then Unload Ws2(i)
Next i
DoEvents
Call Status("Status: Main Login Server Closed!!")
Text1.Enabled = True
Timer3 = False
Timer5 = False
Timer4 = False
Label2.Caption = 0
Label3.Caption = 0
Command3.Enabled = False
Command1.Enabled = True
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim i As Integer
Dim Pck As String
Call Status(Text3.Text & " Locating Connection If Any....")
For i = 1 To 20
If Ws2(i).State = 7 Then
Pck = "GOOD|||2|||" & Text3.Text & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(i).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
Next i
DoEvents
End Sub

Private Sub Ls_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To Text1
If Ws(i).State <> 7 Then
Ws(i).Close
Ws(i).Accept requestID
PingC2(i) = 0
Timer1(i).interval = 5000
Timer1(i).Enabled = True
Debug.Print "Accepted New Connection"
Ls.Close
Ls.LocalPort = 4000
Ls.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Server Full"
Ls.Close
Ls.LocalPort = 4000
Ls.Listen
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Error Resume Next
Dim Pck As String
Timer1(Index).Enabled = False
If PingC2(Index) = 0 Then

Ws(Index).Close
Exit Sub

Else

PingC2(Index) = PingC2(Index) + 1
If PingC2(Index) >= 13 Then

PingC2(Index) = 0
Pck = "PING|||STAYALIVE|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ws(Index).State = 7 Then
Ws(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Else
Ws(Index).Close
Exit Sub
End If

End If

End If
Timer1(Index).Enabled = True
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
iii = 0
ii = 0
For i = 1 To 20
If Ws2(i).State = 7 Then ii = ii + 1
If Ws(i).State = 7 Then iii = iii + 1
Next i
DoEvents
If Timer3 = False Then Exit Sub
Timer3 = False
Label2.Caption = ii
Label3.Caption = iii
Timer3 = True
End Sub

Private Sub Timer5_Timer()
Timer5 = False
Dim Pck As String
Pck = "GOOD|||1|||2|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0
End Sub

Private Sub Timer6_Timer(Index As Integer)
On Error Resume Next
Timer6(Index) = False
Dim SData() As String
Dim i As Long
Dim Pck As String
Dim Who As String, Room As String, PMTXT As String, PMC As Long, India As Integer
PMC = 0
SData = Split(TString(Index), "|~|")
TString(Index) = ""
India = TIndex(Index)
TIndex(Index) = 0
For i = 1 To UBound(SData) - 1
If Ws2(India).State = 7 Then
If Mid(SData(i), 5, 3) = "|||" Then
If Left(SData(i), 4) = "PMIM" Then
If PMC >= 30 Then GoTo Skipit
PMC = PMC + 1
Who = Split(SData(i), "|||")(1)
Room = Split(SData(i), "|||")(2)
PMTXT = Split(SData(i), "|||")(3)
PMTXT = Replace(PMTXT, "~|*~|*", "~|*")
PMTXT = Replace(PMTXT, "~|*", Chr(0))
Pck = "PMIM|||" & Who & "|||" & Room & "|||" & PMTXT & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(India).SendData Pck
DoEvents
Else
Pck = SData(i)
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(India).SendData Pck
DoEvents
End If
End If
End If
Skipit:
Next i
''
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
'TmpData = Dee(Mid(TmpData, 11, Len(TmpData) - 10))
'TmpData = "R4R4" & Chr(0) & Chr$(Int(Len(TmpData) / 256)) & Chr$(Len(TmpData) Mod 256) & Chr(0) & Chr(0) & Chr(128) & TmpData
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
Dim Who As String, Pck As String, Casee As String, Indy As Integer, Room As String, PassW As String, Buds As String, Ignors As String, Onlines As String, OffData As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "PING"
PingC2(Index) = 1

Case "LOGG"
Dim i As Integer
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2) 'Password
OffData = Split(VCDATA, "|||")(3) 'Password
Indy = Split(VCDATA, "|||")(4) 'Password
If Who = "" Then
Reject:
'Ws(Index).Close
Exit Function
Else
If IsAvail(Who) = True Then
Debug.Print Who & " account doesnt exist!"
If Ws(Index).State = 7 Then
Pck = "FAIL|||" & Who & "|||Bad Login ID!|||" & Indy & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
End If
End If
PassW = Get_Name_Info(Who, "Password")
If PassW = "" Then
Debug.Print Who & " failed get password " & PassW
GoTo Reject
End If

If Text2.Text = OffData Then

Else
If Ws(Index).State = 7 Then
Debug.Print Who & " Old Version " & OffData
Pck = "NEWV|||" & Who & "|||" & Text2.Text & "|||" & Indy & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
End If
End If

If Room = PassW Then
'Timer5(Index).Enabled = True
Pck = "GOOD|||" & Who & "|||" & Room & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
For i = 1 To 20
If Ws2(i).State = 7 Then
Ws2(i).SendData Pck 'sent packet telling sub login servers that this ID has logged in, DC any ID on sub server if matching this name
DoEvents
End If
Next i
DoEvents

Buds = Get_Name_Info(Who, "Buddys")
Ignors = Get_Name_Info(Who, "Ignores")
Onlines = GetOnlines(Buds)
DoEvents
If Buds = "" Then Buds = "~"
If Ignors = "" Then Ignors = "~"
If Onlines = "" Then Onlines = "~"

Dim ii As Integer
Dim iii As Integer
ii = 31111
iii = 0
For i = 1 To 20
If Ws2(i).State = 7 Then
If LogCount(i) < ii Then
ii = LogCount(i)
iii = i
End If
End If
Next i
DoEvents
If iii = 0 Then

Else
If Ws(Index).State = 7 Then
If Ws2(iii).State = 7 Then
'Call Status(Who & " Joining Channel " & Room & "!")
Pck = "LOGG|||" & Who & "|||" & Buds & "|||" & Ignors & "|||" & Onlines & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws2(iii).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Pck = "LOGG|||" & Who & "|||" & Room & "|||" & Indy & "|||" & LogIP(iii) & "|||" & LogPort(iii) & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Timer5 = True
Exit Function
End If
End If
'Ws(Index).Close
Debug.Print Who & " Closed For Having No SubServer Long in iii"
End If
'End If
Else
Debug.Print Who & " wrong Password " & Room
If Ws(Index).State = 7 Then
Pck = "FAIL|||" & Who & "|||Wrong Password!|||" & Indy & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws(Index).SendData Pck 'sent packet telling new user the port for the room to connect to!
DoEvents
Exit Function
End If
End If
End If

Case Else

End Select
End Function

Private Function GetOnlines(Budds As String) As String
On Error Resume Next
Dim i As Integer
Dim SDat() As String
Dim TmpD As String
Dim OffData As String
SDat = Split(Budds, "~")
TmpD = "~"
For i = 0 To UBound(SDat)
If SDat(i) = "" Then GoTo Skipy
OffData = Get_Name_Info(SDat(i), "Offline")
If OffData = "" Or OffData = "~" Then GoTo Skipy
If Len(OffData) <= 10 Or Mid(OffData, 11, 3) = "|~|" Then
OffData = ""
Else
If InStr(1, OffData, "|||") > 0 Then
OffData = Split(OffData, "|||")(1)
If Len(OffData) > 0 Then TmpD = TmpD & SDat(i) & "|" & OffData & "~"
End If
End If
Skipy:
Next i
GetOnlines = TmpD
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
For i = 1 To 20
If Ws2(i).State <> 7 Then
Timer2(i).Enabled = False
Ws2(i).Close
Ws2(i).Accept requestID
PingC(i) = 0
Timer2(i).interval = 5000
Timer2(i).Enabled = True
Debug.Print "Accepted New Sub Server Connection"
Ls2.Close
Ls2.LocalPort = 4998
Ls2.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Sub Servers Full"
Ls2.Close
Ls2.LocalPort = 4998
Ls2.Listen
End Sub

Private Sub Ws2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws2(Index)
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
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
Else
GoTo Error
End If
Wend
End With
Exit Sub
Error:
On Error Resume Next
If Ws2(Index).State = 7 Then Ws2(Index).GetData TmpData
End Sub

Public Function ProcessSubServer(VCDATA As String, Index As Integer)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, PMTXT As String, Indy As Integer, SubIP As String, SubPort As String, OffData As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "PING"
PingC(Index) = 1

Case "STAT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
OffData = Get_Name_Info(Who, "Buddys")
Set_Name_Info Who, "Offline", "1~Online~|||" & Room & "|||"
If OffData = "~" Then Exit Function
Pck = "STAT|||" & Who & "|||" & Room & "|||" & OffData & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0

Case "LOGN"
Who = Split(VCDATA, "|||")(1)
OffData = Get_Name_Info(Who, "Offline")
If Len(OffData) <= 10 Then
OffData = ""
Else
OffData = Mid(OffData, 11, Len(OffData) - 10)
If Mid(OffData, 8, 3) = "|||" Then
SendOfflines OffData, Index
End If
End If
Set_Name_Info Who, "Offline", "1~Online~|||Online|||"
OffData = Get_Name_Info(Who, "Buddys")
If OffData = "~" Then Exit Function
Pck = "STAT|||" & Who & "|||Online|||" & OffData & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0

Case "GOOD"
Who = Split(VCDATA, "|||")(1)
LogCount(Index) = Split(VCDATA, "|||")(2)
'Indy = Split(VCDATA, "|||")(3)
LogIP(Index) = Split(VCDATA, "|||")(3)
LogPort(Index) = Split(VCDATA, "|||")(4)
'Timer5 = False
If Who = Text3.Text Then
Call Status(Text3.Text & " Found Them! " & LogIP(Index) & ":" & LogPort(Index))
'Else
'Timer5 = True
End If

'Case "BADD"
'Who = Split(VCDATA, "|||")(1)
'LogCount(Index) = Split(VCDATA, "|||")(2)
'Indy = Split(VCDATA, "|||")(3)
'LogIP(Index) = Split(VCDATA, "|||")(4)
'LogPort(Index) = Split(VCDATA, "|||")(5)
'Timer5(Indy).Enabled = False
'If Who = "Admin" Then
'Call Status(Who & " Found Them! " & SubIP & ":" & SubPort)
'Else
'Ws(Indy).Close
'End If

Case "CHAT"
If Ls3.State = 7 Then
Ls3.SendData VCDATA
End If

Case "JOIN"
If Ls3.State = 7 Then
Ls3.SendData VCDATA
End If

Case "COLR"
If Ls3.State = 7 Then
Ls3.SendData VCDATA
End If

Case "LEFT"
If Ls3.State = 7 Then
Ls3.SendData VCDATA
End If

Case "EXIT"
Who = Split(VCDATA, "|||")(1)
If Ls3.State = 7 Then
Ls3.SendData VCDATA
End If
Set_Name_Info Who, "Offline", "0~Offline~"
OffData = Get_Name_Info(Who, "Buddys")
If OffData = "~" Then Exit Function
Pck = "STAT|||" & Who & "|||Offline|||" & OffData & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0
Timer5 = True

Case "ROMS"
Who = Split(VCDATA, "|||")(1)
Pck = "ROMS|||" & Who & "|||" & Index & "|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ls3.State = 7 Then
Ls3.SendData Pck
End If

Case "PMIM"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
PMTXT = Split(VCDATA, "|||")(3)
OffData = Get_Name_Info(Room, "Offline")
If Left(OffData, 8) = "1~Online" Then
ForwardRoomDataAll VCDATA, 0
Else
PMTXT = Replace(PMTXT, Chr(0), "~|*")
Pck = "PMIM|||" & Who & "|||" & Room & "|||" & PMTXT & "|||"
OffData = OffData & "|~|" & Pck & "|~|"
OffData = Replace(OffData, "|~||~|", "|~|")
Set_Name_Info Room, "Offline", OffData
End If

Case "ADDD"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If IsAvail(Who) = True Then Exit Function
OffData = Get_Name_Info(Who, "Offline")
If Left(OffData, 8) = "1~Online" Then
ForwardRoomDataAll VCDATA, 0
Else
OffData = OffData & "|~|" & Mid(VCDATA, 11, Len(VCDATA) - 10) & "|~|"
OffData = Replace(OffData, "|~||~|", "|~|")
Set_Name_Info Who, "Offline", OffData
End If

Case "DENY"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If IsAvail(Who) = True Then Exit Function
OffData = Get_Name_Info(Who, "Buddys")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Room & "~", "~")
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Who, "Buddys", OffData
DoEvents

OffData = Get_Name_Info(Room, "Buddys")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Who & "~", "~")
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Room, "Buddys", OffData
DoEvents

OffData = Get_Name_Info(Who, "Offline")
DoEvents
If Left(OffData, 8) = "1~Online" Then
ForwardRoomDataAll VCDATA, 0
Else
'OffData = OffData & VCDATA
'Set_Name_Info Who, "Offline", OffData
End If

Case "ACPT"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If IsAvail(Who) = True Then Exit Function
OffData = Get_Name_Info(Who, "Buddys")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Room & "~", "~")
OffData = Replace(OffData, "~~", "~")
OffData = OffData & "~" & Room & "~"
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Who, "Buddys", OffData
DoEvents
OffData = Get_Name_Info(Room, "Buddys")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Who & "~", "~")
OffData = Replace(OffData, "~~", "~")
OffData = OffData & "~" & Who & "~"
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Room, "Buddys", OffData
DoEvents
OffData = Get_Name_Info(Who, "Offline")
DoEvents
If Left(OffData, 8) = "1~Online" Then
ForwardRoomDataAll VCDATA, 0
Else
'OffData = OffData & VCDATA
'Set_Name_Info Who, "Offline", OffData
End If

Case "IGGY"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
If IsAvail(Room) = True Then Exit Function
OffData = Get_Name_Info(Who, "Ignores")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Room & "~", "~")
OffData = Replace(OffData, "~~", "~")
OffData = OffData & "~" & Room & "~"
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Who, "Ignores", OffData

Case "UNIG"
Who = Split(VCDATA, "|||")(1)
Room = Split(VCDATA, "|||")(2)
OffData = Get_Name_Info(Who, "Ignores")
DoEvents
If OffData = "" Then OffData = "~"
OffData = Replace(OffData, "~" & Room & "~", "~")
OffData = Replace(OffData, "~~", "~")
Set_Name_Info Who, "Ignores", OffData

Case "FILE"
ForwardRoomDataAll VCDATA, 0

Case "XFIL"
ForwardRoomDataAll VCDATA, 0

Case "SFIL"
ForwardRoomDataAll VCDATA, 0

Case Else
ForwardRoomDataAll VCDATA, 0

End Select
End Function

Private Sub SendOfflines(TheData As String, Index As Integer)
On Error Resume Next
TCount = TCount + 1
If TCount > 20 Then TCount = 1
TString(TCount) = TheData
TIndex(TCount) = Index
Timer6(TCount) = True
End Sub

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
Dim Pck As String
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


''''Chatrooms socket stuff''''

Private Sub Ls3_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Debug.Print "Accepted New Connection"
Timer4 = False
Ls3.Close
PingCC = 0
Call Status("Rooms Manager Server Connected!")
Ls3.Accept requestID
Timer4 = True
End Sub

Private Sub Ls3_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ls3
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
Debug.Print "RoomServer: " & TmpData
ProcessRoomServer TmpData
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
If Ls3.State = 7 Then Ls3.GetData TmpData
End Sub

Public Function ProcessRoomServer(VCDATA As String)
On Error Resume Next
Dim Who As String, Pck As String, Casee As String, Room As String, Indy As Integer, SubIP As String, SubPort As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "PING"
PingCC = 1

Case "ROMS"
Indy = Split(VCDATA, "|||")(2)
ForwardRoomDataSome VCDATA, Indy

Case Else
ForwardRoomDataAll VCDATA, 0

End Select
End Function

Public Sub ForwardRoomDataSome(Packet As String, Index As Integer)
On Error Resume Next
If Ws2(Index).State = 7 Then
Ws2(Index).SendData Packet 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
End Sub

Public Sub ForwardRoomDataAll(Packet As String, Index As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To 20
If Ws2(i).State = 7 Then
Ws2(i).SendData Packet 'sent packet telling new user the port for the room to connect to!
DoEvents
End If
Next i
DoEvents
End Sub

Private Sub Ls3_Close()
On Error Resume Next
If Command2.Enabled = True Then
If Timer4 = False Then Exit Sub
Timer4 = False
Dim Pck As String
Pck = "EXIT|||NOROOMS|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0
Call Status("Rooms Manager Server Disconnected!")
Ls3.Close
Ls3.LocalPort = 4051
Ls3.Listen
End If
End Sub



Private Sub Ls3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
If Command2.Enabled = True Then
If Timer4 = False Then Exit Sub
Timer4 = False
Dim Pck As String
Pck = "EXIT|||NOROOMS|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
ForwardRoomDataAll Pck, 0
Call Status("Rooms Manager Server Disconnected!")
Ls3.Close
Ls3.LocalPort = 4051
Ls3.Listen
End If
End Sub

Private Sub Timer4_Timer() 'Ping Timer 2
On Error Resume Next
Dim Pck As String
If PingCC = 0 Then

Ls3.Close
If Timer4.Enabled = False Then Exit Sub
Timer4.Enabled = False
If Command2.Enabled = True Then
Call Status("Rooms Manager Server Disconnected!")
Ls3.LocalPort = 4051
Ls3.Listen
End If
Exit Sub

Else
PingCC = PingCC + 1
If PingCC >= 13 Then

PingCC = 0
Pck = "PING|||STAYALIVE|||"
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
If Ls3.State = 7 Then
Timer4.Enabled = False
Ls3.SendData Pck 'sent ping
DoEvents
Else
Ls3.Close
If Timer4.Enabled = False Then Exit Sub
Timer4.Enabled = False
If Command2.Enabled = True Then
Call Status("Rooms Manager Server Disconnected!")
Ls3.LocalPort = 4051
Ls3.Listen
End If
Exit Sub
End If
Else
Timer4.Enabled = False
End If

End If
Timer4.Enabled = True
End Sub
