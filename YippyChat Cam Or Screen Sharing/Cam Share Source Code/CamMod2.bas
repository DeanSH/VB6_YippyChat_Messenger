Attribute VB_Name = "CamMod2"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Wsize As String
Public Hsize As String
Public Sratio As String

Public Function GetINI(Key As String) As String
Dim Ret As String, NC As Long
  
  Ret = String(600, 0)
  NC = GetPrivateProfileString("P2PWebcam", Key, Key, Ret, 600, App.Path & "\Config.ini")
  If NC <> 0 Then Ret = Left$(Ret, NC)
  If Ret = Key Or Len(Ret) = 600 Then Ret = ""
  GetINI = Ret

End Function
'Read from INI

Public Sub WriteINI(ByVal Key As String, Value As String)
  
  WritePrivateProfileString "P2PWebcam", Key, Value, App.Path & "\Config.ini"

End Sub
'Write to INI

Public Function RandomGen2(rChars As String, rCount As Integer) As String
On Error Resume Next
  Dim tmpStr As String, x As Integer
    Randomize
      Do Until Len(tmpStr) = rCount
        x = Len(rChars) * Rnd + 1
        tmpStr = tmpStr & (Mid$(rChars, x, 1))
      Loop
        RandomGen2 = tmpStr
End Function

Public Sub Pause(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub

  
