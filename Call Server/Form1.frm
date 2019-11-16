VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voice Call Server"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Ls 
      Left            =   2520
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   1920
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private KeyCode(0 To 1000) As String
Private KeyIndex(0 To 1000) As Integer

Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
Dim i As Integer
For i = 0 To 1000
KeyCode(i) = ""
KeyIndex(i) = 0
Next i
Label1.Caption = "0"
Label2.Caption = "0"
Ls.Close
Ls.LocalPort = "3440"
Ls.Listen
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
KeyCode(0) = ""
KeyIndex(0) = 0
Ws(0).Close
Ls.Close
Dim i As Integer
For i = 1 To 1000
KeyCode(i) = ""
KeyIndex(i) = 0
Ws(i).Close
Unload Ws(i)
Next i
DoEvents
Label1.Caption = "0"
Label2.Caption = "0"
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim i As Integer
If Command2.Enabled = False Then Exit Sub
If Label1.Caption = "0" Then Exit Sub
Label2.Caption = "0"
For i = 0 To Label1.Caption - 1
If Ws(i).State = 7 Then Label2.Caption = Label2.Caption + 1
Next i
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim i As Integer
If Command2.Enabled = False Then Exit Sub
If Label1.Caption = "0" Then Exit Sub
KeyCode(0) = ""
KeyIndex(0) = 0
Ws(0).Close
For i = 1 To Label1.Caption - 1
KeyCode(i) = ""
KeyIndex(i) = 0
Ws(i).Close
Unload Ws(i)
Next i
DoEvents
Label1.Caption = "0"
Label2.Caption = "0"
End Sub

Private Sub Ls_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
For i = 0 To 1000
If Ws(i).State <> 7 Then
Load Ws(i)
Ws(i).Close
KeyCode(i) = ""
KeyIndex(i) = 0
If (i + 1) > Label1.Caption Then Label1.Caption = (i + 1)
Ws(i).Accept requestID
GoTo ok
End If
Next i
ok:
Ls.Close
Ls.LocalPort = "3440"
Ls.Listen
End Sub

Private Sub Ws_Close(Index As Integer)
On Error Resume Next
Dim i As Integer
If Label1.Caption = "0" Then KeyCode(Index) = "": Exit Sub
For i = 0 To Label1.Caption - 1
If i = Index Then GoTo Skip
If Ws(i).State = 7 Then
If KeyCode(i) = KeyCode(Index) Then Ws(i).Close
End If
Skip:
Next i
DoEvents
KeyCode(Index) = ""
KeyIndex(Index) = 0
End Sub

Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Error
Dim Data As String
Ws(Index).GetData Data, vbString
'Debug.Print Index & " - " & Data
If Left(Data, 4) = "CODE" Then
KeyCode(Index) = Split(Data, "|||")(1)
Identify Data, Index
Exit Sub
End If
If Left(Data, 4) = "PING" Then
Exit Sub
End If
ForwardVoice Data, Index
Error:
End Sub

Private Function ForwardVoice(Dat As String, Indy As Integer)
On Error Resume Next
If Label1.Caption = "0" Then Exit Function
If KeyIndex(Indy) = Indy Then Exit Function
If KeyCode(KeyIndex(Indy)) = KeyCode(Indy) Then
If Ws(KeyIndex(Indy)).State = 7 Then Ws(KeyIndex(Indy)).SendData Dat
End If
End Function

Private Function Identify(Dat As String, Indy As Integer)
On Error Resume Next
Dim i As Integer
If Label1.Caption = "0" Then Exit Function
For i = 0 To Label1.Caption - 1
If i = Indy Then GoTo Skip
If Ws(i).State = 7 Then
If KeyCode(i) = KeyCode(Indy) Then
KeyIndex(i) = Indy
Ws(i).SendData Dat
DoEvents
KeyIndex(Indy) = i
Ws(Indy).SendData Dat
Exit Function
End If
End If
Skip:
Next i
'Ws(Indy).SendData Dat
End Function

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Dim i As Integer
If Label1.Caption = "0" Then KeyCode(Index) = "": Exit Sub
For i = 0 To Label1.Caption - 1
If i = Index Then GoTo Skip
If Ws(i).State = 7 Then
If KeyCode(i) = KeyCode(Index) Then Ws(i).Close
End If
Skip:
Next i
DoEvents
KeyCode(Index) = ""
KeyIndex(Index) = 0
End Sub
