Attribute VB_Name = "Module2"
Option Explicit
'These set our variables to laod our databse tables into
'Remember to select Microsoft DAO Library from Reference


'Declaires variables for bits and bobs

'These set our variables to laod our databse tables into
'Remember to select Microsoft DAO Library from Reference
Dim db As Database
Dim rs As Recordset
Dim Wss As Workspace

Dim max As Long
Dim jid As Long
Dim entry_date As String

''''''''''''''''''''''''''''''''''''''''''

Public Sub OpenDataBase()
On Error Resume Next
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists(App.Path & "\Activity Logs") Then

Else
  fs.CreateFolder App.Path & "\Activity Logs"
End If
End Sub

Public Function IsAvail(Name As String) As Boolean
On Error Resume Next
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists(App.Path & "\Activity Logs\" & LCase(Name)) Then
If InStr(1, " " & LCase(Name), "/") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), "\") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), "?") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), "|") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), Chr(34)) > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), "*") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), "<") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), ">") > 0 Then GoTo Badi
If InStr(1, " " & LCase(Name), ":") > 0 Then GoTo Badi
IsAvail = False
Else
Badi:
IsAvail = True
'fs.CreateFolder App.Path & "\Activity Logs\" & LCase(Name)
End If
End Function

Public Function AddNewUser(Name As String, Pass As String, Info As String) As Boolean
On Error GoTo Exist
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists(App.Path & "\Activity Logs\" & LCase(Name)) Then
AddNewUser = False
Exit Function
Else
fs.CreateFolder App.Path & "\Activity Logs\" & LCase(Name)
End If

If Set_Name_Info2(Name, Pass, Info) = True Then
AddNewUser = True
Else
AddNewUser = False
End If
Exit Function
Exist:
AddNewUser = False
End Function

Private Function Set_Name_Info2(Name As String, Pass As String, Info As String) As Boolean
On Error GoTo Error
Dim DataType As String
DataType = "Password"
SaveFile Name, DataType, Pass
DoEvents

DataType = "Details"
SaveFile Name, DataType, Info
DoEvents

DataType = "Buddys"
SaveFile Name, DataType, "~"
DoEvents

DataType = "Ignores"
SaveFile Name, DataType, "~"
DoEvents

DataType = "Offline"
SaveFile Name, DataType, "0~Offline~"
DoEvents

Set_Name_Info2 = True
Exit Function

Error:
Set_Name_Info2 = False
End Function

Private Sub SaveFile(User As String, FileType As String, TheData As String)
On Error Resume Next
Dim intE As Integer
Dim TheFile As String

TheFile = App.Path & "\Activity Logs\" & LCase(User) & "\" & FileType & ".txt"
intE = FreeFile

Open TheFile For Output As #intE
Print #intE, TheData
Close #intE
DoEvents
End Sub

Public Sub Set_Name_Info(Name As String, DataType As String, Info As String)
SaveFile Name, DataType, Info
End Sub

Public Function Get_Name_Info(Name As String, DataType As String) As String
On Error GoTo Error
Dim TmpDat As String

TmpDat = get_entries(Name, DataType)
TmpDat = Replace(TmpDat, Chr(10), "")
TmpDat = Replace(TmpDat, Chr(13), "")
If DataType = "Details" Then GoTo Skip
If DataType = "Offline" Then GoTo Skip
TmpDat = Replace(TmpDat, " ", "")
Skip:
Get_Name_Info = TmpDat
Exit Function

Error:
Get_Name_Info = ""
End Function

Private Function get_entries(User As String, DataType As String) As String
On Error Resume Next
Dim intE As Integer
Dim TmpData As String
Dim TheFile As String
Dim lstInput As String
TheFile = App.Path & "\Activity Logs\" & LCase(User) & "\" & DataType & ".txt"
intE = FreeFile
TmpData = ""

    Open TheFile For Input As #intE
    While Not EOF(1)
        Input #intE, lstInput$
        'If lstInput$ = "" Then Exit Sub
        TmpData = TmpData & lstInput$
    Wend
    Close #intE
 DoEvents
get_entries = TmpData
End Function

Public Function IsIllegalID(Name) As Boolean
On Error Resume Next

Dim LettersAllowed As String
Dim LetterArray As Long
Dim UndyCount() As String

UndyCount = Split(Name, "_")
If UBound(UndyCount) > 10 Then GoTo Done

If InStr(1, Name, "_____") > 0 Then GoTo Done

If Len(Name) < 2 Or Len(Name) > 25 Then GoTo Done

LettersAllowed = "zABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
If InStr(1, LettersAllowed, Left(Name, 1)) > 0 Then

LettersAllowed = "0_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
If InStr(1, LettersAllowed, Right(Name, 1)) > 0 Then

LettersAllowed = "_ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_"
For LetterArray = 2 To Len(Name) - 1
If InStr(1, LettersAllowed, Mid(Name, LetterArray, 1)) > 0 Then
'LetterFound
Else
GoTo Done
End If
Next LetterArray
IsIllegalID = False
Exit Function

End If
End If
DoEvents
Done:
IsIllegalID = True
End Function
