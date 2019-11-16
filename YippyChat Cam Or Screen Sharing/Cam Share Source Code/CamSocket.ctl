VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl CamSocket 
   BackColor       =   &H00FF0000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1200
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wsL 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Stats(Data As String)
Private sFile As String
Private sFile4 As String
Private WhoIsThis As String
Private IsSending As Boolean

Private Sub Timer1_Timer()
On Error Resume Next
Timer1 = False
If WhoIsThis = "" Then
RaiseEvent Stats("Host: Server Connection Timed Out!")
wsL.Close
End If
End Sub

Private Sub Timer2_Timer()
Timer2 = False
IsSending = False
ImageReady = False
If wsL.State = 7 Then SendPicture
End Sub

Private Sub wsL_Close()
Timer1 = False
If WhoIsThis = "" Then Exit Sub
RaiseEvent Stats("Host: " & WhoIsThis & " Server Disconnect!")
'RemoveList WhoIsThis, HostList.List1
'DoEvents
HostList.List1.Clear
WhoIsThis = ""
IsSending = False
ImageReady = False
If Desktop = True Then
Host.Caption = "Desktop"
Else
Host.Caption = "WebCam"
End If
End Sub

Public Sub ForceStop(StopWho As String)
On Error Resume Next
If WhoIsThis = "" Then
'Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
Else
wsL.SendData "KICK|||" & WhoIsThis & "|||" & StopWho & "|||" 'Send Kick User Command Packet
RemoveList StopWho, HostList.List1
DoEvents
'WhoIsThis = ""
If Desktop = True Then
Host.Caption = "Desktop (" & HostList.List1.ListCount & " Viewing)"
Else
Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
End If
End If
End Sub

Public Sub ClosingTime()
WhoIsThis = ""
Timer1 = False
Timer2 = False
IsSending = False
ImageReady = False
wsL.Close
End Sub

Public Function TheState() As Boolean
If wsL.State = 7 Then
TheState = True
Else
TheState = False
End If
End Function

Public Sub ConnectServer(ServaIP As String, Porta As String)
On Error Resume Next
RaiseEvent Stats("Host: Connecting to Server..")
wsL.Close
Timer1 = True
wsL.Connect ServaIP, Porta
End Sub

Private Sub wsL_Connect()
RaiseEvent Stats("Host: Connected and Broadcasting!")
WhoIsThis = WhoAmI
If Desktop = True Then
wsL.SendData "HOST|||" & WhoIsThis & "|||%" & Sratio 'Send Host Command
Else
wsL.SendData "HOST|||" & WhoIsThis & "|||" 'Send Host Command
End If
End Sub

Private Sub wsL_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Error
Dim sData As String
Dim sWho As String
  wsL.GetData sData
  'Debug.Print sData
      Select Case Left$(sData, 4)
      Case "JOIN"
        If InStr(1, sData, "SIZE|||") > 0 Then
        Timer1 = False
        IsSending = False
        ImageReady = False
        SendPicture
        End If
        sWho = Split(sData, "|||")(1)
        Joiner sWho
        Exit Sub

      Case "LEFT"
        If InStr(1, sData, "SIZE|||") > 0 Then
        Timer1 = False
        IsSending = False
        ImageReady = False
        SendPicture
        End If
        sWho = Split(sData, "|||")(1)
        Leaver sWho
        Exit Sub
        
      Case "SIZE"
        If InStr(1, sData, "LEFT|||") > 0 Then
        sWho = Split(sData, "LEFT|||")(1)
        sWho = Split(sWho, "|||")(0)
        Leaver sWho
        End If
        If InStr(1, sData, "JOIN|||") > 0 Then
        sWho = Split(sData, "JOIN|||")(1)
        sWho = Split(sWho, "|||")(0)
        Joiner sWho
        End If
        Timer1 = False
        IsSending = False
        ImageReady = False
        SendPicture

      End Select
Error:
End Sub
'parsing results

Private Sub Joiner(sWho As String)
 If InList(sWho, HostList.List1) = True Or LCase(sWho) = LCase(WhoAmI) Then
        RaiseEvent Stats("Host: Blocked Duplicate User Attempt!")
        ForceStop sWho
        Exit Sub
End If
        HostList.List1.AddItem sWho
        RaiseEvent Stats("Host: " & sWho & " Started Viewing!")
    If Desktop = True Then
        Host.Caption = "Desktop (" & HostList.List1.ListCount & " Viewing)"
            Else
        Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
    End If
End Sub

Private Sub Leaver(sWho As String)
If InList(sWho, HostList.List1) = False Then
        Exit Sub
        End If
        RemoveList sWho, HostList.List1
        DoEvents
        RaiseEvent Stats("Host: " & sWho & " Stopped Viewing!")
        If Desktop = True Then
        Host.Caption = "Desktop (" & HostList.List1.ListCount & " Viewing)"
        Else
        Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
        End If
End Sub

Private Sub wsL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Timer1 = False
ImageReady = False
IsSending = False
If WhoIsThis = "" Then Exit Sub
RaiseEvent Stats("Host: " & WhoIsThis & " Server Disconnect!")
HostList.List1.Clear
WhoIsThis = ""
If Desktop = True Then
Host.Caption = "Desktop"
Else
Host.Caption = "WebCam"
End If
End Sub
'duh

Private Sub SendPicture()
On Error GoTo Err
    'get jpg file
'wait:
    'If sFile3 = "" Then GoTo wait
    'If ImageReady = True Then Pause "0.001": GoTo wait
    If Host.CamPause.Checked = True Then
    Pause "0.3"
    Else
    If sFile4 = sFile3 Then Pause "0.02" ': GoTo wait
    End If
    Timer2 = False
    IsSending = True
    ImageReady = True
    sFile = sFile3
    sFile4 = sFile
    Timer2 = True
    wsL.SendData "SIZE" & Len(sFile) & "FILE" & sFile
Exit Sub
Err:
ImageReady = False
IsSending = False
Debug.Print "SendPic Error"
End Sub

Private Sub wsL_SendComplete()
'Timer2 = False
If IsSending = True Then
IsSending = False
ImageReady = False
End If
End Sub

Private Sub Pause(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub
