VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form HostSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quality Control"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1130
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin MSComctlLib.Slider sQuality 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Min             =   30
      Max             =   85
      SelStart        =   85
      TickFrequency   =   2
      Value           =   85
   End
   Begin VB.Frame Frame1 
      Caption         =   "Webcam Drivers"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
      Begin VB.TextBox txtVPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtYourIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtTheirIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ListBox lstDevices 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "<-- Password"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Your IP"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Their IP"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Display Your Cam At Set Compressed Quality?"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Min 30% <-- Image Compression Quality --> 85% Max"
      Height          =   235
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "The Lower The Compression % The Less Clear Images Become!! The Higher The Value The More Clear Images Become!!"
      Top             =   830
      Width           =   4095
   End
End
Attribute VB_Name = "HostSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error Resume Next
Command3.Enabled = True
If Check1.Value = 1 Then
WriteINI "Method", "1"
Else
WriteINI "Method", "0"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Command3.Enabled = False
Me.Visible = False
End Sub

Private Sub Command2_Click()
On Error Resume Next
StayOnTop Host
Command3.Enabled = False
Me.Visible = False
End Sub

Private Sub Command3_Click()
On Error Resume Next
Command3.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = Host.Icon
'StayOnTop Me
End Sub

Private Sub sQuality_Change()
On Error Resume Next
Dim s As String
Command3.Enabled = True
  s = sQuality.Value
  sQuality.ToolTipText = sQuality.Value
  If Desktop = False Then WriteINI "Quality", s
End Sub
'stores new quality

'Private Sub txtTheirIP_Change()
'On Error Resume Next
  'WriteINI "ConnectTo", txtTheirIP.Text
'End Sub
'saves new viewing ip

'Private Sub txtVPassword_Change()
'On Error Resume Next
  'WriteINI "ViewingPassword", txtVPassword.Text
'End Sub
'saves new viewing pass
