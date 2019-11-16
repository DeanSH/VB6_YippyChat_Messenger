Attribute VB_Name = "Module1"
Public RoomNames(0 To 3000) As String
Public ServerIP As String
Public ServerON As Boolean
Public Limit As Integer
Public UserList(0 To 3000) As String

Public Function Status(Stat As String) As String
Form1.Label1.Caption = Stat
End Function

Public Sub Pause(interval)
 Dim x
 x = Timer
  Do While Timer - x < Val(interval)
  DoEvents
 Loop
End Sub

Public Sub LoadList(dialogCommon As CommonDialog, List As ListBox)
Dim lstInput As String
On Error GoTo Error_Killer
    With dialogCommon
    .DialogTitle = "Load List"
    .Filter = "*.txt"
    .ShowOpen
On Error Resume Next
    Open .fileName For Input As #1
    While Not EOF(1)
        Input #1, lstInput$
    If lstInput$ = "" Then Exit Sub
        List.AddItem lstInput$
    Wend
    Close #1
 End With
Exit Sub
Error_Killer:
Exit Sub
End Sub

Public Function SaveList(dialogCommon As CommonDialog, list45 As ListBox)
On Error GoTo Error_Killer
    With dialogCommon
    .DialogTitle = "Save List"
    .Filter = "*.txt"
    .ShowSave
    Dim Nbr As Long
On Error Resume Next
    Open .fileName For Output As #1
    For Nbr = 0 To list45.ListCount - 1
    Print #1, list45.List(Nbr)
    Next Nbr
    Close #1
    End With
Exit Function
Error_Killer:
Exit Function
End Function
