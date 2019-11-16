VERSION 5.00
Begin VB.Form HostList 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewers List"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Double Click To Kick A Viewer"
      Top             =   340
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   45
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   70
      Width           =   2535
   End
End
Attribute VB_Name = "HostList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.Icon = Host.Icon
End Sub

Private Sub Label2_Click()
Host.viewrs.Checked = False
HostList.Visible = False
End Sub

Private Sub List1_DblClick()
On Error Resume Next
If List1.ListCount = 0 Then Exit Sub
'Dim I As Integer
Dim WhoKick As String
WhoKick = List1
'For I = 1 To 50
'If Host.CamSocket(I).TheName = WhoKick Then
Host.CamSocket.ForceStop WhoKick
'GoTo Done
'End If
'Next I
'DoEvents
Done:
List1.RemoveItem List1.ListIndex
DoEvents
Host.Caption = "WebCam (" & List1.ListCount & " Viewing)"
End Sub
