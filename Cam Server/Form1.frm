VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Cam Server"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   9000
      Left            =   3120
      Top             =   840
   End
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
Private KeyCode(0 To 1000) As String ' Host Name
Private TempPic(0 To 1000) As String ' Host Temp Image Storage
Private PermPic(0 To 1000) As String ' Host Current Available Image Storage
Private SizePic(0 To 1000) As String ' Host Temp Image Size Expected
Private MyRatio(0 To 1000) As String ' Host Screen Size Ratio

Private KeyView(0 To 1000) As String ' Viewer Name
Private KeyWhom(0 To 1000) As String ' Viewing Who Saved With Viewer Index
Private KeyIndy(0 To 1000) As String ' Host Index Saved With Viewer Index



Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
Dim i As Integer
For i = 0 To 1000
KeyCode(i) = ""
KeyView(i) = ""
KeyWhom(i) = ""
KeyIndy(i) = ""
TempPic(i) = ""
PermPic(i) = ""
SizePic(i) = ""
MyRatio(i) = ""
Next i
Label1.Caption = "0"
Label2.Caption = "0"
Ls.Close
Ls.LocalPort = "3660"
Ls.Listen
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
Command2.Enabled = False
KeyCode(0) = ""
KeyView(0) = ""
KeyWhom(0) = ""
KeyIndy(0) = ""
TempPic(0) = ""
PermPic(0) = ""
SizePic(0) = ""
MyRatio(0) = ""
Timer1(0).Enabled = False
Ws(0).Close
Ls.Close
Dim i As Integer
For i = 1 To 1000
KeyCode(i) = ""
KeyView(i) = ""
KeyWhom(i) = ""
KeyIndy(i) = ""
TempPic(i) = ""
PermPic(i) = ""
SizePic(i) = ""
MyRatio(i) = ""
Timer1(i).Enabled = False
Ws(i).Close
Unload Ws(i)
Unload Timer1(i)
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
KeyView(0) = ""
KeyWhom(0) = ""
KeyIndy(0) = ""
TempPic(0) = ""
PermPic(0) = ""
SizePic(0) = ""
MyRatio(0) = ""
Timer1(0).Enabled = False
Ws(0).Close
For i = 1 To Label1.Caption - 1
KeyCode(i) = ""
KeyView(i) = ""
KeyWhom(i) = ""
KeyIndy(i) = ""
TempPic(i) = ""
PermPic(i) = ""
SizePic(i) = ""
MyRatio(i) = ""
Timer1(i).Enabled = False
Ws(i).Close
Unload Ws(i)
Unload Timer1(i)
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
Load Timer1(i)
Timer1(i).Enabled = False
Ws(i).Close
KeyCode(i) = ""
KeyView(i) = ""
KeyWhom(i) = ""
KeyIndy(i) = ""
TempPic(i) = ""
PermPic(i) = ""
SizePic(i) = ""
MyRatio(i) = ""
If (i + 1) > Label1.Caption Then Label1.Caption = (i + 1)
Timer1(i).Enabled = True
Ws(i).Accept requestID
GoTo ok
End If
Next i
ok:
Ls.Close
Ls.LocalPort = "3660"
Ls.Listen
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Error Resume Next
Timer1(Index).Enabled = False
CloseUser Index
End Sub

Private Sub Ws_Close(Index As Integer)
If Timer1(Index).Enabled = True Then CloseUser Index
End Sub

Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Timer1(Index).Enabled = False
Timer1(Index).Enabled = True
Ws(Index).GetData Data, vbString, bytesTotal
'Debug.Print Data
ProcessData Data, Index
End Sub

Private Function ProcessData(Dat As String, Indy As Integer)
On Error Resume Next
Dim i As Integer
If Label1.Caption = "0" Then CloseUser Indy: Exit Function

Select Case Left(Dat, 4)

Case "HOST"
If KeyCode(Indy) = "" Then
KeyCode(Indy) = Split(Dat, "|||")(1)
If InStr(1, Dat, "|||%") > 0 Then MyRatio(Indy) = Split(Dat, "|||%")(1)
TempPic(Indy) = ""
PermPic(Indy) = ""
SizePic(Indy) = ""
'Debug.Print "HOST:" & KeyCode(Indy)
Ws(Indy).SendData "SIZE|||" & KeyCode(Indy) & "|||"
End If
Exit Function

Case "SIZE"
If KeyCode(Indy) = "" Then

Else
TempPic(Indy) = ""
SizePic(Indy) = Split(Mid$(Dat, 5), "FILE")(0) 'SIZE456789FILEabcdefg.....
TempPic(Indy) = Split(Dat, SizePic(Indy) & "FILE")(1)
'Debug.Print SizePic(Indy)
If Len(TempPic(Indy)) >= SizePic(Indy) Then
PermPic(Indy) = TempPic(Indy)
'Debug.Print "DONEPIC:" & Len(TempPic(Indy))
TempPic(Indy) = ""
Ws(Indy).SendData "SIZE|||" & KeyCode(Indy) & "|||"
End If
End If
Exit Function

Case "KICK"
If KeyCode(Indy) = "" Then Exit Function
Dim KickWho As String
If LCase(KeyCode(Indy)) = LCase(Split(Dat, "|||")(1)) Then
KickWho = Split(Dat, "|||")(2)
'Debug.Print "KICK:" & KickWho
For i = 0 To Label1.Caption - 1
If i = Indy Then GoTo Skip
If Ws(i).State = 7 Then
If LCase(KeyView(i)) = LCase(KickWho) Then
CloseUser i
'Exit Function
End If
End If
Skip:
Next i
End If
'End If
Exit Function

Case "VIEW"
If KeyCode(Indy) = "" Then
If KeyView(Indy) = "" Then
KeyView(Indy) = Split(Dat, "|||")(1)
KeyWhom(Indy) = Split(Dat, "|||")(2)
For i = 0 To Label1.Caption - 1
If i = Indy Then GoTo Skip4
If Ws(i).State = 7 Then
If LCase(KeyCode(i)) = LCase(KeyWhom(Indy)) Then
KeyIndy(Indy) = i
Ws(i).SendData "JOIN|||" & KeyView(Indy) & "|||"
If MyRatio(i) = "" Then
Ws(Indy).SendData "PICL|||" & Len(PermPic(i)) & "|||"
'Debug.Print "VIEW:" & KeyView(Indy)
Else
Ws(Indy).SendData "PICL|||" & Len(PermPic(i)) & "|||%" & MyRatio(i)
'Debug.Print "VIEW:" & KeyView(Indy) & "::" & MyRatio(i)
End If
Exit Function
End If
End If
Skip4:
Next i
End If
End If
CloseUser Indy
Exit Function

Case "NEXT"
If KeyCode(Indy) = "" Then
If KeyView(Indy) = "" Then
CloseUser Indy
Else
'Debug.Print "NEXT:" & KeyView(Indy)
i = KeyIndy(Indy)
If LCase(KeyCode(i)) = LCase(KeyWhom(Indy)) Then
If KeyCode(i) = "" Then
CloseUser Indy
Else
If Ws(i).State = 7 Then
'If Split(Dat, "|||")(1) = KeyView(Indy) Then
If PermPic(i) = "" Then
'Debug.Print "NEXT:No PermPic"
Ws(Indy).SendData "PICL|||" & Len(PermPic(i)) & "|||"
Else
SendImage Indy, i
End If
'End If
End If
End If
Else
CloseUser Indy
End If
End If
End If
Exit Function

Case Else
If KeyCode(Indy) = "" Then

Else
TempPic(Indy) = TempPic(Indy) & Dat
'Debug.Print "PICDATA:" & Len(TempPic(Indy))
If Len(TempPic(Indy)) >= SizePic(Indy) Then
PermPic(Indy) = TempPic(Indy)
'Debug.Print "DONEPIC:" & Len(TempPic(Indy))
TempPic(Indy) = ""
Ws(Indy).SendData "SIZE|||" & KeyCode(Indy) & "|||"
End If
End If
Exit Function

End Select

End Function

Private Function SendImage(Index As Integer, Index2 As Integer)
On Error Resume Next
'Debug.Print "SEND:" & Len(PermPic(Index2))
Ws(Index).SendData "SIZE" & Len(PermPic(Index2)) & "FILE" & PermPic(Index2)
End Function

Private Sub Ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
CloseUser Index
End Sub

Public Sub CloseUser(Index As Integer)
On Error GoTo Error
Dim i As Integer
Dim Pck As String
Debug.Print Index & " Closed"
Timer1(Index).Enabled = False

If Label1.Caption = "0" Then GoTo Error
If KeyCode(Index) = "" Then
If KeyView(Index) = "" Then GoTo Error
If KeyWhom(Index) = "" Then GoTo Error
If KeyIndy(Index) = "" Then GoTo Error

i = KeyIndy(Index)
If i = Index Then GoTo Error
If Ws(i).State = 7 Then
If LCase(KeyCode(i)) = LCase(KeyWhom(Index)) Then
'Pck = "STOP|||" & KeyView(Index) & "|||" & KeyCode(i) & "|||"
'Pck = "R4R4" & Chr(0) & Chr$(Int(Len(Pck) / 256)) & Chr$(Len(Pck) Mod 256) & Chr(0) & Chr(0) & Chr(128) & Pck
Pck = "LEFT|||" & KeyView(Index) & "|||" & KeyCode(i) & "|||"
Ws(i).SendData Pck 'Send That Viewer Stopped Viewing To Host
End If
End If

Else

For i = 0 To Label1.Caption - 1
If i = Index Then GoTo Skip
If Ws(i).State = 7 Then
If LCase(KeyWhom(i)) = LCase(KeyCode(Index)) Then Ws(i).Close 'DC Viewers Because Host has DC'd
End If
Skip:
Next i
DoEvents

End If

Error:
KeyCode(Index) = ""
KeyView(Index) = ""
KeyWhom(Index) = ""
KeyIndy(Index) = ""
TempPic(Index) = ""
PermPic(Index) = ""
SizePic(Index) = ""
End Sub
