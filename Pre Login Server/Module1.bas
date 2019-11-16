Attribute VB_Name = "Module1"
Public ServerON As Boolean

Public Function Status(Stat As String) As String
Form1.Label1.Caption = Stat
End Function

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



