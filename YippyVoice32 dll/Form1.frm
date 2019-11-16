VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R4's Yahoo Audio Record/Play Example"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "Maximum Packs Per Second Recieved"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Text            =   "1440"
      ToolTipText     =   "Byte can Start from 480... and then 960... then 1440... etc! Just Increase/Decrease by amounts of 480 to keep Sound Good!"
      Top             =   360
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Real Time Playback"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "PlayBack Audio While Recording.. (May Cause Echo Effect If Not Using Headphones)"
      Top             =   840
      Width           =   1760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Packet Per Sec Counter"
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   72
      Left            =   2280
      Top             =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   1900
      X2              =   1900
      Y1              =   720
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<< Length/Bytes Of TrueSpeech Data !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Packets Per Sec:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   885
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents fRecorder As R4sRecorder
Attribute fRecorder.VB_VarHelpID = -1
Private WithEvents fPlayer As R4sPlayer
Attribute fPlayer.VB_VarHelpID = -1

Private RecData(0 To 99999) As String
Private RecIndex As Long
Private PlayIndex As Long

'NOTE Bytes for Chatroom = 1440 (480 x 3) making 96 bytes of Compressed Audio Data!!! And Confy = 960 (480 x 2) making 64 bytes of Compressed Audio Data!!!
'NOTE bytes increase or decrease in amount of +/- 480! So to modify the Bytes without effecting sound quality, adjust it in amounts of 480! not just any old increase or decrease!
'Examples...
'480 Bytes = 32 Bytes Of Compressed Audio Data Per Chunk
'960 Bytes = 64 Bytes Of Compressed Audio Data Per Chunk
'1440 Bytes = 96 Bytes Of Compressed Audio Data Per Chunk
'1880 Bytes = 128 Bytes Of Compressed Audio Data Per Chunk
'2400 Bytes = 160 Bytes Of Compressed Audio Data Per Chunk
'Etc

Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "Record" Then
Command1.Caption = "Stop"
Set fPlayer = Nothing
Timer2 = True
OverAllBytes = Text2(6).Text
RecIndex = 0
Text1.Text = "0"
Text4.Text = "0"
Set fPlayer = New R4sPlayer
fPlayer.Initalize
Set fRecorder = New R4sRecorder
fRecorder.Record
Else
Command1.Caption = "Record"
Timer2 = False
fRecorder.EndRecord
Set fRecorder = Nothing
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Play" Then
If RecIndex = 0 Then Exit Sub
Command2.Caption = "Stop"
PlayIndex = 0
Timer1.Interval = ((Text3 / 32) * 24)
Timer1 = True
Else
Timer1 = False
Command2.Caption = "Play"
End If
End Sub

Private Sub Form_Load()
RecIndex = 0
Load Form2
Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form2
Set fPlayer = Nothing
Set fRecorder = Nothing
End Sub

Private Sub fRecorder_onSoundCompressed(ByVal trueSpeechData As String, ByVal lBufferBytes As Long)
On Error Resume Next
Debug.Print "Compressed Data Chunk||" & lBufferBytes & "||" & trueSpeechData & "||" & Asc(Mid(trueSpeechData, Len(trueSpeechData), 1))
RecIndex = RecIndex + 1
Text1.Text = Text1.Text + 1
Me.Caption = "Recording String: " & RecIndex
Text3.Text = Len(trueSpeechData) - 1
RecData(RecIndex) = Mid(trueSpeechData, 1, Len(trueSpeechData) - 1)
If Check2.Value = 1 Then fPlayer.PlayWave RecData(RecIndex)
End Sub

Public Function ResolveMaxPacks(Nums As Long)
Dim Maxi As Long
Maxi = Text4
If Nums > Maxi Then Text4.Text = Nums
End Function

Private Sub Timer1_Timer()
If Timer1 = False Then Exit Sub
If PlayIndex >= RecIndex Then
Timer1 = False
Command2.Caption = "Play"
Exit Sub
End If
PlayIndex = PlayIndex + 1
Me.Caption = "Playing String: " & PlayIndex & "/" & RecIndex
fPlayer.PlayWave RecData(PlayIndex)
End Sub

Private Sub Timer2_Timer()
ResolveMaxPacks Text1
Text1.Text = "0"
End Sub
