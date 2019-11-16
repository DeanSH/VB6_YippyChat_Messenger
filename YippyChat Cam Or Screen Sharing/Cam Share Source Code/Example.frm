VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cam Usage Example"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Host Desktop Instead Of Cam (Host Only)"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open And View Thier Cam?"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open And Host Your Cam?"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "anthony"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "deano"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "yippychat.redirectme.net"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "yippychat.redirectme.net"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "<< Thier UserName"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Your UserName"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Thier IP Address"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Your IP Address"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Private Function GetINI(Key As String) As String
Dim Ret As String, NC As Long
  
  Ret = String(600, 0)
  NC = GetPrivateProfileString("P2PWebcam", Key, Key, Ret, 600, App.Path & "\Config.ini")
  If NC <> 0 Then Ret = Left$(Ret, NC)
  If Ret = Key Or Len(Ret) = 600 Then Ret = ""
  GetINI = Ret

End Function
'Read from INI

Private Sub WriteINI(ByVal Key As String, Value As String)
  
  WritePrivateProfileString "P2PWebcam", Key, Value, App.Path & "\Config.ini"

End Sub
'Write to INI
Private Sub Command1_Click()
On Error GoTo Error
WriteINI "Connect2", Text1
WriteINI "Connect1", Text2
WriteINI "ViewingPassword2", Text3
WriteINI "ViewingPassword", Text4
WriteINI "Driver", "0"
If Check1.Value = 0 Then WriteINI "Method2", "0"
If Check1.Value = 1 Then WriteINI "Method2", "1"
DoEvents
Shell App.Path & "\CamHost.exe", vbNormalFocus
Error:
End Sub

Private Sub Command2_Click()
On Error GoTo Error
WriteINI "Connect2", Text1
WriteINI "Connect1", Text2
WriteINI "ViewingPassword2", Text3
WriteINI "ViewingPassword", Text4
DoEvents
Shell App.Path & "\CamView.exe", vbNormalFocus
Error:
End Sub

Private Sub Form_Load()
Text1 = GetINI("Connect2")
Text2 = GetINI("Connect1")
Text3 = GetINI("ViewingPassword2")
Text4 = GetINI("ViewingPassword")
End Sub
