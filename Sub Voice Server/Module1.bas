Attribute VB_Name = "Module1"
Public RoomNames(0 To 50) As String
Public RoomPort1(0 To 50) As String
Public ServerIP As String
Public ServerON As Boolean


Public Function Status(Stat As String) As String
On Error Resume Next
Form1.Label1.Caption = Stat
End Function

Public Sub Pause(interval)
On Error Resume Next
Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Public Function Enn(Data As String) As String
On Error Resume Next
Dim en As New ModR4
Dim AntiFreeze As Long
Dim TmpStr As String
AntiFreeze = 0
TmpStr = ""
Start:
If AntiFreeze > 199 Then Enn = "Error": Exit Function
AntiFreeze = AntiFreeze + 1
TmpStr = en.EncryptString(Data, Chr(&H7B - &H4D) + Chr(&H81 - &H53) + Chr(&H95 - &H48) + Chr(&H5F - &H6) + Chr(&H4C - &H2B) + Chr(&H9A - &H57) + Chr(&H98 - &H50) + Chr(&H68 - &H27) + Chr(&HB5 - &H61) + Chr(&H86 - &H58) + Chr(&H36 - &H8) + Chr(&H92 - &H60) + Chr(&H56 - &H26) + Chr(&H67 - &H36) + Chr(&H81 - &H4E))
DoEvents
If Len(TmpStr) <> Len(Data) + 5 Then GoTo Start
If Dee(TmpStr) = Data Then
Enn = TmpStr
Else
GoTo Start
End If
End Function

Public Function Dee(Data As String) As String
On Error Resume Next
Dim De2 As New ModR4
Dee = De2.DecryptString(Data, Chr(&H7B - &H4D) + Chr(&H81 - &H53) + Chr(&H95 - &H48) + Chr(&H5F - &H6) + Chr(&H4C - &H2B) + Chr(&H9A - &H57) + Chr(&H98 - &H50) + Chr(&H68 - &H27) + Chr(&HB5 - &H61) + Chr(&H86 - &H58) + Chr(&H36 - &H8) + Chr(&H92 - &H60) + Chr(&H56 - &H26) + Chr(&H67 - &H36) + Chr(&H81 - &H4E))
End Function

Public Sub LoadList(dialogCommon As CommonDialog, List As ListBox)
Dim lstInput As String
On Error GoTo Error_Killer
    With dialogCommon
    .DialogTitle = "Load List"
    .Filter = "*.txt"
    .ShowOpen
On Error Resume Next
    Open .FileName For Input As #1
    While Not EOF(1)
        Input #1, lstInput$
    If lstInput$ = "" Then Exit Sub
        List.AddItem lstInput$
    Wend
    Close #1
 End With
Exit Sub
Error_Killer:
End Sub

Public Function SaveList(dialogCommon As CommonDialog, list45 As ListBox)
On Error GoTo Error_Killer
    With dialogCommon
    .DialogTitle = "Save List"
    .Filter = "*.txt"
    .ShowSave
    Dim Nbr As Long
On Error Resume Next
    Open .FileName For Output As #1
    For Nbr = 0 To list45.ListCount - 1
    Print #1, list45.List(Nbr)
    Next Nbr
    Close #1
    End With
Exit Function
Error_Killer:
End Function
