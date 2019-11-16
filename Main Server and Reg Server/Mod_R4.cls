VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModR4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Event Progress(Percent As Integer)
Const Rounds = 8
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Private AddTbl(255, 255) As Byte
Private XTbl(255, 255) As Byte
Private LsTbl(255, 255) As Byte
Private RsTbl(255, 255) As Byte

Private Sub Append(ByRef StringData As String, Optional Length As Long)
    On Error Resume Next
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 4096
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function
Public Function EnHex(Data As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function EncryptString(Text As String, Optional Key As String) As String
    EncryptString = StrConv(EncryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
    'EncryptString = EnHex(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String) As String
    On Error Resume Next
    'Text = DeHex(Text)
    DecryptString = StrConv(DecryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 4096
    ReDim byteArray(hiBound)
End Sub
Public Function EncryptByte(DS() As Byte, Key() As Byte)
    On Error Resume Next
    Dim tmp2() As Byte, p As Integer, i As Long, Bound As Integer, r As Integer

    If UBound(Key) < 0 Then
        ReDim Key(2): Key(0) = 1: Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0): tmp2(0) = Key(0): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 1 Then
        ReDim tmp2(1): tmp2(0) = Key(0): tmp2(1) = Key(1): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = tmp2(1): Key(2) = 1
    End If

    ReDim tmp2((UBound(DS)) + 5)

    Randomize Timer
    tmp2(0) = Int((Rnd * 254) + 1)
    tmp2(1) = Int((Rnd * 254) + 1)
    tmp2(2) = Int((Rnd * 254) + 1)
    tmp2(UBound(tmp2)) = Int((Rnd * 254) + 1)
    tmp2(UBound(tmp2) - 1) = Int((Rnd * 254) + 1)

    Call CopyMem(tmp2(3), DS(0), UBound(DS) + 1)
    ReDim DS(UBound(tmp2)) As Byte
    DS() = tmp2()
    ReDim tmp2(0)

    For r = 1 To Rounds
        Bound = (UBound(Key))
        p = 0
        For i = 0 To UBound(DS) - 1
            If p = Bound Then p = 0
            DS(i) = XTbl(DS(i), AddTbl(DS(i + 1), Key(p)))
            DS(i + 1) = XTbl(DS(i), DS(i + 1))
            DS(i) = XTbl(DS(i), AddTbl(DS(i + 1), Key(p + 1)))
            p = p + 1
        Next
        RaiseEvent Progress((r / Rounds) * 100)
    Next
    
    EncryptByte = DS()
End Function
Public Function DecryptByte(DS() As Byte, Key() As Byte)
    On Error Resume Next
    Dim tmp2() As Byte, p As Long, i As Long, Bound As Integer, r As Integer
    If UBound(Key) < 0 Then
        ReDim Key(2): Key(0) = 1: Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0): tmp2(0) = Key(0): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = 1: Key(2) = 1
    ElseIf UBound(Key) = 1 Then
        ReDim tmp2(1): tmp2(0) = Key(0): tmp2(1) = Key(1): ReDim Key(2)
        Key(0) = tmp2(0): Key(1) = tmp2(1): Key(2) = 1
    End If

    For r = 1 To Rounds
        Bound = (UBound(Key))
        p = (UBound(DS)) Mod (UBound(Key))
        
        For i = (UBound(DS)) To 1 Step -1
            If p = 0 Then p = Bound
            DS(i - 1) = XTbl(DS(i - 1), AddTbl(DS(i), Key(p)))
            DS(i) = XTbl(DS(i - 1), DS(i))
            DS(i - 1) = XTbl(DS(i - 1), AddTbl(DS(i), Key(p - 1)))
            p = p - 1
        Next
        RaiseEvent Progress((r / Rounds) * 100)
    Next

    tmp2() = DS()
    ReDim DS((UBound(tmp2)) - 4) As Byte
    Call CopyMem(DS(0), tmp2(3), UBound(DS))
    ReDim Preserve DS(UBound(DS) - 1) As Byte
    
    DecryptByte = DS()
End Function
Private Sub Class_Initialize()
Dim i As Integer, j As Integer

For i = 0 To 255
    For j = 0 To 255
        XTbl(i, j) = CByte(i Xor j)
        AddTbl(i, j) = CByte((i + j) Mod 255)
    Next
Next
End Sub


