VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Host 
   BackColor       =   &H00000000&
   Caption         =   "WebCam (0 Viewers)"
   ClientHeight    =   4590
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5550
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000B&
   Icon            =   "Host.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin HostCam.CamSocket CamSocket 
      Height          =   375
      Left            =   10080
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4335
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9252
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   480
   End
   Begin VB.CommandButton CmdFormat 
      Caption         =   "Resolution"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Pic3 
      Height          =   1935
      Left            =   7800
      ScaleHeight     =   1875
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdHost 
      Caption         =   "Start Hosting"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdSettings 
      Caption         =   "Settings"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picCapture 
      Height          =   1245
      Left            =   0
      ScaleHeight     =   1185
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start WebCam Feed"
      Height          =   255
      Left            =   10080
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image picCapture2 
      BorderStyle     =   1  'Fixed Single
      Height          =   4095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
   Begin VB.Menu mn2 
      Caption         =   "---"
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu viewrs 
         Caption         =   "Show Viewers"
         Shortcut        =   ^L
      End
      Begin VB.Menu CamPause 
         Caption         =   "Pause"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu ops 
      Caption         =   "Settings"
      Begin VB.Menu advops 
         Caption         =   "Advanced Options"
      End
      Begin VB.Menu setupops 
         Caption         =   "Image Compression"
         Shortcut        =   ^Q
      End
      Begin VB.Menu scaling 
         Caption         =   "Quality Scale"
         Visible         =   0   'False
         Begin VB.Menu FQ 
            Caption         =   "Full Quality"
            Checked         =   -1  'True
         End
         Begin VB.Menu hq 
            Caption         =   "Half Quality"
         End
         Begin VB.Menu lq 
            Caption         =   "Low Quality"
         End
      End
      Begin VB.Menu vidops 
         Caption         =   "Video Format"
      End
   End
   Begin VB.Menu paus 
      Caption         =   "Paused!"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<=== R4's Notes ==>
'Used A Public P2P and FTP Cam Example, That didnt work due to no pause between sending size and File..
'Now i made it send the Size on the front of the file, and made it not split the file sending chunks,
'Instead made it send it whole, reducing overall image transferal times, now very smooth Streaming..
'Also Added Pausing Ability, Resizing Ability, Minimizing, And Fullscreening. After Editing The Way,
'That Images Are Displayed, Originally Using Picture Box's.. Now Displaying With ImageBox's Instead!
'Allow for the Benefit of Stretch to Fit Property. And Layout Completely Modified.
'Added Extra Option For Video Format, To Control Resolution If Wanted!
'Also Made it Into multi viewer hosting, not single p2p.. So now can stream to upto 50 Viewers.
'The Ability to Control Quality COmpression means users can control image sizes, therefore controlling bandwidth!
'</=== R4's Notes ==>

'<=== 2nd Original Notes ==>
'I used public coding to get a stream of the webcam.
'I resized it to max myself, added the winsock, auth, p2p, all that
'good stuff myself. sorry if the coding is not the best, it is easy
'for me to understand, and my style.
'-------
'openurl and bmp2jpg was also public code on pscode
'thanks to everyone which makes their code public, it is a great learning
'resource. and a quick easy way to make something better of something small :)
'-------------
'the upload script not is kind of sketchy. it would be better to find a way to
'get it to only refresh when a new image is up. save time and a few problems
'email: FinalCry@GMail.Com
'if you figure out a solution to this or make one please do let me know
'--Enjoy the code
'Also note, the lstdevices coding, is a hidden feature to save pictures
'for testing purposes(incomplete), but good enough if you want to add it on
'</== Second Original Notes ==>

'<=== Original Notes ==>
'For further information on this Feel Free
'to contact me at brtiwari@yahoo.com
'and do not forget to vote me on PSC
'</== Original Notes ==>

Const WM_CAP As Integer = &H400

Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP + 41
Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP + 42

Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM As Integer = 1

Dim iDevice As Long  ' Current device ID


Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "USER32" (ByVal hndw As Long) As Boolean
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

Private Sub advops_Click()
CmdSettings_Click
End Sub

Private Sub CamPause_Click()
If Dir$(App.Path & "\Files\cam.bmp") = "" Then Exit Sub
If cmdStart.Enabled = True Then Exit Sub
If CamPause.Checked = False Then
CamPause.Checked = True
Timer1 = False
paus.Visible = True
Else
CamPause.Checked = False
paus.Visible = False
Timer1 = True
End If
End Sub

Private Sub CamSocket_Stats(Data As String)
StatusBar1.Panels(1).Text = Data
End Sub

Private Sub CmdFormat_Click()
capDlgVideoFormat hHwnd
End Sub

Private Sub cmdHost_Click()
On Error GoTo Err
  If cmdHost.Caption = "Start Hosting" Then
    cmdHost.Caption = "Stop Hosting"
    'StatusBar1.Panels(1).Text = "Host: Broadcasting!!"
    'ListenUp
  End If
Err:
End Sub
'toggle hosting

Private Sub CmdSettings_Click()
  capDlgVideoSource hHwnd
End Sub
'show camera settings

Private Sub cmdStart_Click()
If Desktop = False Then iDevice = HostSetup.lstDevices.ListIndex
    cmdStart.Enabled = False
    OpenPreviewWindow
End Sub
'start feed

Private Sub Form_Load()
On Error GoTo Err
If App.PrevInstance = True Then
MsgBox "Error!! YippyChat CamHost.exe is allready Running!" & vbCrLf & "Running 2 Copies at once is Not Allowed, Please Close it First to Start it again Fresh! If unable to locate, open Task Manager and Force it Closed in there!", , "Error! CamHost.exe is allready Running!"
DoEvents
End
Exit Sub
End If
If Dir$("C:\Windows\System32\gdi32.dll") <> "" Then

Else
MsgBox "Error!! 'gdi32.dll' is Missing Or Not Registered!" & vbCrLf & "Please Redo The Run-Time Files To Fix This Error!", , "Error! Needed File Is Missing!"
DoEvents
End
Exit Sub
End If
If Dir$("C:\Windows\System32\VIC32.DLL") <> "" Then

Else
MsgBox "Error!! 'VIC32.DLL' is Missing Or Not Registered!" & vbCrLf & "Please Redo The Run-Time Files To Fix This Error!", , "Error! Needed File Is Missing!"
DoEvents
End
Exit Sub
End If
If Dir$("C:\Windows\System32\avicap32.dll") <> "" Then

Else
MsgBox "Error!! 'avicap32.dll' is Missing Or Not Registered!" & vbCrLf & "Please Redo The Run-Time Files To Fix This Error!", , "Error! Needed File Is Missing!"
DoEvents
End
Exit Sub
End If
If Dir$("C:\Windows\System32\wiaaut.dll") <> "" Then

Else
MsgBox "Error!! 'wiaaut.dll' is Missing Or Not Registered!" & vbCrLf & "Please Redo The Run-Time Files To Fix This Error!", , "Error! Needed File Is Missing!"
DoEvents
End
Exit Sub
End If
    picCapture.Height = "0"
    picCapture.Width = "0"
    picCapture.BorderStyle = 0
    sLastSendOut = "N/A"
    'duh
    Me.Height = "5715"
    Me.Width = "5805"
    HostSetup.Show
    HostList.Show
    HostSetup.Visible = False
    HostList.Visible = False
    ops.Enabled = False
    scaling.Enabled = False
    CamPause.Enabled = False
    viewrs.Enabled = False
    Select Case GetINI("Method")
    Case "0"
      HostSetup.Check1.Value = 0
    Case "1"
      HostSetup.Check1.Value = 1
    Case Else
      HostSetup.Check1.Value = 0
      WriteINI "Method", "0"
    End Select
    'password users must use to view ur webcam
    WhoAmI = GetINI("ViewingPassword")
    TheIP = GetINI("Connect1")
    HostList.Label1.Caption = WhoAmI
    mn2.Caption = WhoAmI
    ScaleDown = 1
    'password used to connect to others
    'jpeg quality
   'Select Case GetINI("Method2")
    'Case "0"
      Desktop = False
      HostSetup.sQuality.Value = Val(GetINI("Quality"))
      HostSetup.sQuality.ToolTipText = HostSetup.sQuality.Value
      Wsize = Screen.Width / 15
      Hsize = Screen.Height / 15
      Debug.Print Wsize
      Debug.Print Hsize
      Sratio = (100 / Wsize) * Hsize
      Host.Caption = "WebCam (0 Viewers)"
      Form_Resize
    'Case "1"
      'StatusBar1.Panels(1).Text = "Host: Loading Desktop..."
      'Desktop = True
      'HostSetup.sQuality.Value = 30
      'HostSetup.sQuality.ToolTipText = HostSetup.sQuality.Value
      'Me.Visible = False
      'Wsize = Screen.Width / 15
      'Hsize = Screen.Height / 15
      'Debug.Print Wsize
      'Debug.Print Hsize
      'Sratio = (100 / Wsize) * Hsize
      'Host.Caption = "Desktop (0 Viewers)"
      'Form_Resize
      'cmdStart_Click
      'Exit Sub
    'Case Else
      'WriteINI "Method2", "1"
      'StatusBar1.Panels(1).Text = "Host: Loading Desktop..."
      'Desktop = True
      'HostSetup.sQuality.Value = 30
      'HostSetup.sQuality.ToolTipText = HostSetup.sQuality.Value
      'Me.Visible = False
      'Wsize = Screen.Width / 15
      'Hsize = Screen.Height / 15
      'Debug.Print Wsize
      'Debug.Print Hsize
      'Sratio = (100 / Wsize) * Hsize
      'Host.Caption = "Desktop (0 Viewers)"
      'Form_Resize
      'cmdStart_Click
      'Exit Sub
    'End Select
    DoEvents
    LoadDeviceList
    DoEvents
    'loads cam devices
    If HostSetup.lstDevices.ListCount > 0 Then
      HostSetup.lstDevices.Selected(0) = True
      HostSetup.lstDevices.ListIndex = 0
      'sets index
      StatusBar1.Panels(1).Text = "Host: Loading Cam..."
      Me.Visible = False
      cmdStart_Click
    Else
      StatusBar1.Panels(1).Text = "Host: Failed To Load Cam!! Error!!"
      cmdStart.Enabled = False
    End If
    Me.Show
    DoEvents
    StayOnTop Host
    'DoEvents
    'if no devises disable hosting
    Exit Sub
Err:
      StatusBar1.Panels(1).Text = "Host: Failed To Load Cam!! Error!!"
      cmdStart.Enabled = False
End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer As String
    Dim iReturn As Boolean
    Dim x As Long
    
    x = 0
    strName = Space(100)
    strVer = Space(100)

    ' Load name of all available devices into lstDevices
    Do
        ' Get Driver name and version
        iReturn = capGetDriverDescriptionA(x, strName, 100, strVer, 100)
        ' If there was a device add device name to the list
        If iReturn Then HostSetup.lstDevices.AddItem Trim$(strName)
        x = x + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()

If Desktop = False Then
    ' Open Preview window in picturebox 352 288
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 352, 288, picCapture.hWnd, 0)
     
    ' Connect to device
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then

        'Set the preview scale
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0

        'Set the preview rate in milliseconds
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0

        'Start previewing the image from the camera
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0

        ' Resize window to fit in picturebox
        'SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, SWP_NOMOVE Or SWP_NOZORDER
    CmdSettings.Enabled = True
    CmdFormat.Enabled = True
    cmdHost.Enabled = True
    ops.Enabled = True
    scaling.Enabled = False
    CamPause.Enabled = True
    viewrs.Enabled = True
    On Error Resume Next
    'Me.Show
    DoJPG
    'DoEvents
    'StayOnTop Me
    CamSocket.ConnectServer TheIP, "3660"
    'CamSocket.ConnectServer "192.161.59.152", "3660"
    Timer1 = True
    HostSetup.lstDevices.Enabled = False
    cmdHost_Click
    
    Else

        ' Error connecting to device close window
        DestroyWindow hHwnd
        Timer1 = False
        cmdStart.Enabled = True
        StatusBar1.Panels(1).Text = "Host: Failed To Load Cam!! Error!!"
    End If
    
    Else
    
    CmdSettings.Enabled = False
    CmdFormat.Enabled = False
    cmdHost.Enabled = True
    ops.Enabled = True
    scaling.Enabled = True
    advops.Enabled = False
    vidops.Enabled = False
    CamPause.Enabled = True
    viewrs.Enabled = True
    On Error Resume Next
    Me.Show
    DoEvents
    DoJPG
    DoEvents
    CamSocket.ConnectServer TheIP, "3660"
    'CamSocket.ConnectServer "192.161.59.152", "3660"
    Timer1 = True
    cmdHost_Click
    
    End If
 End Sub

Private Sub ClosePreviewWindow()
On Error Resume Next
If Desktop = False Then
    ' Disconnect from device
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0
    Timer1 = False
    ' close window
    DestroyWindow hHwnd
    Else
    Timer1 = False
    End If
End Sub
'disc driver


Private Sub Form_Resize()
On Error GoTo Error
If Sratio = "" Then Exit Sub
If Desktop = True Then
If Me.Width <= "4399" Then Me.Width = "4400"
Me.Height = (Me.Width / 100) * Sratio
picCapture2.Width = (Me.Width - 240)
picCapture2.Height = (Me.Height - StatusBar1.Height - 870)
If Me.Width > "4399" And picCapture2.Left > "0" Then picCapture2.Left = "0"
Exit Sub
End If
If Me.Width <= "4399" Then Me.Width = "4400": Me.Height = "4200": Exit Sub
If Me.Height <= "7499" Then Me.Height = Me.Width / 100 * 96: GoTo Skip
If Me.Height <= "9999" Then Me.Height = Me.Width / 100 * 92: GoTo Skip
If Me.Height <= "12499" Then Me.Height = Me.Width / 100 * 88: GoTo Skip
If Me.Height > "12499" Then Me.Height = Me.Width / 100 * 86
Skip:
If Me.Height <= "4199" Then Me.Width = "4400": Me.Height = "4200": Exit Sub
picCapture2.Height = Me.Height - StatusBar1.Height - 870
picCapture2.Width = Me.Width - 240
If Me.Height > "4199" And Me.Width > "4399" And picCapture2.Left > "0" Then picCapture2.Left = "0"
Exit Sub
Error:
Form_Resize2
End Sub

Private Sub Form_Resize2()
On Error Resume Next
If Desktop = True Then
picCapture2.Height = (Me.Height - StatusBar1.Height - 870)
picCapture2.Width = (picCapture2.Height / Sratio) * 100
If ((Me.Width - picCapture2.Width - 200) / 2) < 0 Then
picCapture2.Left = "0"
Else
picCapture2.Left = (Me.Width - picCapture2.Width - 200) / 2
End If
Exit Sub
End If
picCapture2.Height = Me.Height - StatusBar1.Height - 870
picCapture2.Width = Me.Height / 85 * 100
If ((Me.Width - picCapture2.Width - 200) / 2) < 0 Then
picCapture2.Left = "0"
Else
picCapture2.Left = (Me.Width - picCapture2.Width - 200) / 2
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  If cmdStart.Enabled = False Then ClosePreviewWindow   'if camera is on
  'disables camera
  DoEvents
  CamSocket.ClosingTime
  DoEvents
  'Unload CamSocket
  'DoEvents
  Unload HostSetup
  DoEvents
  If Dir$(App.Path & "\Files\cam.bmp") <> "" Then Kill App.Path & "\Files\cam.bmp"
  If Dir$(App.Path & "\Files\cam.jpg") <> "" Then Kill App.Path & "\Files\cam.jpg"
  'deletes sending
  End
End Sub

Function capDlgVideoFormat(ByVal lwnd As Long) As Boolean
   capDlgVideoFormat = SendMessage(lwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
End Function

Function capDlgVideoSource(ByVal lwnd As Long) As Boolean
   capDlgVideoSource = SendMessage(lwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
End Function
'for settings such as source, color, hue, tilt bla bla

Private Sub FQ_Click()
ScaleDown = 1
hq.Checked = False
FQ.Checked = True
lq.Checked = False
End Sub

Private Sub hq_Click()
ScaleDown = 2
hq.Checked = True
FQ.Checked = False
lq.Checked = False
End Sub

Private Sub lq_Click()
ScaleDown = 3
hq.Checked = False
FQ.Checked = False
lq.Checked = True
End Sub

Private Sub setupops_Click()
HostSetup.Visible = True
HostSetup.Left = Me.Left + 500
StayOnTop HostSetup
End Sub

Private Sub Timer1_Timer()
Timer1 = False
If CamSocket.TheState = True Then
If Desktop = False Then Me.WindowState = 0 ': StayOnTop Host
  If cmdStart.Enabled = False Then
  If CamPause.Checked = True Then Exit Sub
  'If CmdSettings.Enabled = True Then
    'If CmdView.Caption = "Start Viewing" Then
      DoJPG
    'End If
  'End If
  Exit Sub
  Else
  Exit Sub
  End If
End If
  'shows a prevue of jpg quality if viewing mode is off
Timer1 = True
End Sub

Sub DoJPG()
On Error Resume Next
Dim s As String

  If Desktop = False Then
  s = Clipboard.GetText
  SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
  Pic3.Picture = Clipboard.GetData
  Clipboard.SetText s
  If HostSetup.Check1.Value = 0 Then
  picCapture2.Picture = Pic3.Picture
  End If
  Else
  Set Pic3.Picture = CaptureScreen()
  Pause2 "0.01"
  If HostSetup.Check1.Value = 0 Then
  picCapture2.Picture = Pic3.Picture
  End If
  If Dir$(App.Path & "\Files\cam.bmp") <> "" Then Kill App.Path & "\Files\cam.bmp"
  SavePicture Pic3.Picture, App.Path & "\Files\cam.bmp"
  Dim imgPhoto As WIA.ImageFile
  Set imgPhoto = New WIA.ImageFile
  'imgPhoto.FileData.Picture = Pic3.Picture
  imgPhoto.LoadFile App.Path & "\Files\cam.bmp"
  Set imgPhoto = ResizeImage(imgPhoto, Trim(Wsize / ScaleDown), Trim(Hsize / ScaleDown))
  Set Pic3.Picture = imgPhoto.FileData.Picture
  End If
  'restores text
  'copies to clipboard, then to our picturebox
  
  If Dir$(App.Path & "\Files\cam.bmp") <> "" Then Kill App.Path & "\Files\cam.bmp"
  If Dir$(App.Path & "\Files\cam.jpg") <> "" Then Kill App.Path & "\Files\cam.jpg"
  'deletes old pics
  'Pause2 "0.01"
  
  SavePicture Pic3.Picture, App.Path & "\Files\cam.bmp"
  Pause2 "0.01"
  
  'kills old bmp, saves new
  BMPtoJPG App.Path & "\Files\cam.bmp", App.Path & "\Files\cam.jpg", HostSetup.sQuality.Value
  
  If HostSetup.Check1.Value = 1 Then
  Pause2 "0.01"
  picCapture2.Picture = LoadPicture(App.Path & "\Files\cam.jpg")
  Else
  'picCapture2.Picture = Pic3.Picture
  End If
  'converts bmp to jpg, shows jpg
  Open App.Path & "\Files\cam.jpg" For Binary As #1
  sFile2 = Space(LOF(1))
  Get #1, , sFile2
  Close #1
  'DoEvents
  'Debug.Print Len(sFile2)
  'Debug.Print FileLen(App.Path & "\Files\cam.jpg")
    'stores file
  sFile3 = sFile2
  Pause2 "0.01"
  Timer1 = True
End Sub
'makes a jpeg image and displays it

Private Sub vidops_Click()
CmdFormat_Click
End Sub

Private Sub viewrs_Click()
If viewrs.Checked = True Then
viewrs.Checked = False
HostList.Visible = False
StayOnTop Me
Else
HostList.Left = Me.Left + 500
viewrs.Checked = True
HostList.Visible = True
StayOnTop HostList
End If
End Sub

