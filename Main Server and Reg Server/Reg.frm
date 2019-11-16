VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Register Server"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Test illegal"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Ls4 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   5000
      Left            =   2400
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Ws3 
      Index           =   0
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
Text1.Enabled = False
OpenDataBase
Dim i As Integer
For i = 1 To Text1
Load Timer6(i)
Load Ws3(i)
Next i
DoEvents
Ls4.Close
Ls4.LocalPort = 4049
Ls4.Listen
Command2.Enabled = True
Call Status("Status: Register Server Started!!")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
Dim i As Integer
Ls4.Close
For i = 1 To Text1
Unload Timer6(i)
Unload Ws3(i)
Next i
DoEvents
Call Status("Status: Register Server Closed!!")
Text1.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
If IsIllegalID(Text2.Text) = True Then
Call Status(Text2.Text & " Is Illegal!")
Else
Call Status(Text2.Text & " Is Allgood!")
End If
End Sub

'''''Acc Create Sockets''''

Private Sub Ls4_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 1 To 50
If Ws3(i).State <> 7 Then
Ws3(i).Close
Ws3(i).Accept requestID
Timer6(i).Enabled = True
Debug.Print "Accepted New Connection"
Ls4.Close
Ls4.LocalPort = 4049
Ls4.Listen
Exit Sub
End If
Next i
'If Reach Here, Server Full Limit Reached!
Debug.Print "Server Full"
Ls4.Close
Ls4.LocalPort = 4049
Ls4.Listen
End Sub

Private Sub Timer6_Timer(Index As Integer)
On Error Resume Next
Ws3(Index).Close
Timer6(Index).Enabled = False
End Sub

Private Sub Ws3_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String, DataLength As String, TmpData As String, HeaderLength As Integer
HeaderLength = 10
With Ws3(Index)
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
If Left(Data, 4) = "R4R4" Then
DataLength = Trim((256 * Asc(Mid(Data, 6, 1)) + Asc(Mid(Data, 7, 1))) + HeaderLength)
If DataLength < 16 Then GoTo Error
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
TmpData = Dee(Mid(TmpData, 11, Len(TmpData) - 10))
TmpData = "R4R4" & Chr(0) & Chr$(Int(Len(TmpData) / 256)) & Chr$(Len(TmpData) Mod 256) & Chr(0) & Chr(0) & Chr(128) & TmpData
'Debug.Print "Account Create: " & TmpData
ProcessAcc TmpData, Index
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
Timer6(Index).Enabled = False
Ws3(Index).Close
End Sub

Public Function ProcessAcc(VCDATA As String, Index As Integer)
On Error GoTo Error
Dim Who As String, Pck As String, Casee As String, Pass As String, Info As String
Casee = Mid(VCDATA, 11, 4)

Select Case Casee

Case "CHNG"
Who = Split(VCDATA, "|||")(1)
Pass = Split(VCDATA, "|||")(2) 'Password
Info = Split(VCDATA, "|||")(3) 'New Pass
If Who = "" Or LCase(Who) = "admin" Or Len(Who) < 2 Or Len(Who) > 25 Then GoTo Error
If IsAvail(Who) = True Then
Pck = Enn("FAIL|||" & Who & "|||Account was not Found in Database, Unable to Modify Password!|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Call Status(Who & " Not Found, Change Pass Rejected!")
Else
Call Status(Who & " Changing Passowrd!")
If ChangeUser(Who, Pass, Info) = True Then
Pck = Enn("CHGD|||" & Who & "|||" & Pass & "|||" & Info & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Else
Pck = Enn("FAIL|||" & Who & "|||The Password You Used Is Incorrect, Failed To Change Password|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Call Status(Who & " Change Pass Rejected, Wrong Pass!")
End If
End If

Case "MAKE"
Who = Split(VCDATA, "|||")(1)
Pass = Split(VCDATA, "|||")(2) 'Password
Info = Split(VCDATA, "|||")(3) 'Info
If IsIllegalID(Who) = True Then GoTo Baddy
If Who = "" Or LCase(Who) = "admin" Or Len(Who) < 2 Or Len(Who) > 25 Then
Baddy:
Pck = Enn("BADD|||" & Who & "|||Problem With User Name Length To Short Or Long, Or For ID Content Using Blocked Words Or Characters|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Call Status(Who & " Rejected from account creation!")
Else
Call Status(Who & " making account now!")
If AddNewUser(Who, Pass, Info) = True Then
Pck = Enn("MADE|||" & Who & "|||" & Pass & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Else
GoTo Baddy
End If
End If

Case "CHCK"
Who = Split(VCDATA, "|||")(1)
If IsIllegalID(Who) = True Then GoTo Baddy2
If Who = "" Or LCase(Who) = "admin" Or Len(Who) < 2 Or Len(Who) > 25 Then
Baddy2:
Pck = Enn("NOOO|||" & Who & "|||Problem With User Name Length To Short Or Long, Or For ID Content Using Blocked Words Or Characters|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Call Status(Who & " Rejected from account creation!")
Else
Call Status(Who & " checking available!!")
If IsAvail(Who) = True Then
Pck = Enn("YESS|||" & Who & "|||" & RandomGen("AaBbCcDdEeFfGgHhJjKkLMmNnPpQqRrSsTtUuVvWwXxYyZz", 7) & "|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
Else
Pck = Enn("NOOO|||" & Who & "|||Username Is Allready Taken!|||")
Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Ws3(Index).SendData Pck
End If
End If

Case Else

End Select
Exit Function
Error:
Timer6(Index).Enabled = False
Ws3(Index).Close
End Function

Private Sub Ws3_Close(Index As Integer)
On Error Resume Next
Timer6(Index).Enabled = False
End Sub

Private Sub Ws3_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Timer6(Index).Enabled = False
End Sub

