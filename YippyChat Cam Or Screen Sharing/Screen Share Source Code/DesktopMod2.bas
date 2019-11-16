Attribute VB_Name = "DesktopMod2"
Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Type RGBTRIPLE
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbReserved As Byte
End Type

Private Type BITMAP '14 bytes
bmType As Long
bmWidth As Long
bmHeight As Long
bmWidthBytes As Long
bmPlanes As Integer
bmBitsPixel As Integer
bmBits As Long
End Type

Private Type BITMAPFILEHEADER
bfType As Integer
bfSize As Long
bfReserved1 As Integer
bfReserved2 As Integer
bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
biSize As Long
biWidth As Long
biHeight As Long
biPlanes As Integer
biBitCount As Integer
biCompression As Long
biSizeImage As Long
biXPelsPerMeter As Long
biYPelsPerMeter As Long
biClrUsed As Long
biClrImportant As Long
End Type

Private Type BITMAPINFO
bmHeader As BITMAPINFOHEADER
'bmColors(0 To 255) As RGBTRIPLE
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Const SPI_SETDESKWALLPAPER = 20
'key constants
Private Const HKEY_CURRENT_USER = &H80000001
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_SUCCESS = 0&
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const READ_WRITE = 2
Private Const READAPI = 0
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Private Const REG_NONE = 0                       ' No value type
Private Const KEY_WOW64_64KEY As Long = &H100& '32 bit app to access 64 bit hive
Private Const KEY_WOW64_32KEY As Long = &H200& '64 bit app to access 32 bit hive


Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Type Rect
left As Integer
top As Integer
right As Integer
bottom As Integer
End Type


Private Function CreateFromPicture(ByVal sPic As IPicture, ByRef pixels() As Byte)

Dim tB As BITMAP
Dim hbm As Long
Dim oldbmp As Long
hbm = GetObject(sPic.Handle, Len(tB), tB)
picHeight = tB.bmHeight
picWidth = tB.bmWidth

Dim hdc As Long
Dim memdc As Long
Dim old
hdc = GetWindowDC(0)
hbm = CreateCompatibleBitmap(hdc, picWidth, picHeight)
memdc = CreateCompatibleDC(hdc)
oldbmp = SelectObject(memdc, hbm)
Dim rc As Rect
rc.left = 0
rc.top = 0
rc.right = picWidth
rc.bottom = picHeight

Call sPic.Render(memdc, 0, 0, picWidth, picHeight, 0, sPic.Height, sPic.Width, -sPic.Height, rc)
Dim bitmap_info As BITMAPINFO
Dim bytes_per_scanline As Integer
Dim pad_per_scanline As Integer

With bitmap_info.bmHeader 'start load picture data
.biSize = Len(bitmap_info.bmHeader)
.biHeight = picHeight
.biWidth = picWidth
.biPlanes = 1
.biBitCount = 32
.biCompression = 0 'BI_RGB
bytes_per_scanline = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) 'get bytes
pad_per_scanline = bytes_per_scanline - (((.biWidth * .biBitCount) + 7) \ 8) 'get pad
.biSizeImage = bytes_per_scanline * Abs(.biHeight)
End With

ReDim pixels(1 To 4, 1 To picWidth, 1 To picHeight)

GetDIBits memdc, hbm, 0, picHeight, pixels(1, 1, 1), bitmap_info, 0 'DIB_RGB_COLORS

' Fill in the BITMAPFILEHEADER.
Dim bitmap_file_header As BITMAPFILEHEADER
With bitmap_file_header
.bfType = &H4D42 ' "BM"
.bfOffBits = Len(bitmap_file_header) + _
Len(bitmap_info.bmHeader)
.bfSize = .bfOffBits + _
bitmap_info.bmHeader.biSizeImage
End With

Kill App.Path & "\temp.bmp"
file_name = App.Path & "\temp.bmp"
' Open the output bitmap file.
fnum = FreeFile
Open file_name For Binary As fnum

' Write the BITMAPFILEHEADER.
Put #fnum, , bitmap_file_header

' Write the BITMAPINFOHEADER.
' (Note that memory_bitmap.bitmap_info.bmiHeader.biHeight
' must be positive for this.)
Put #fnum, , bitmap_info

' Write the DIB bits.
Put #fnum, , pixels

' Close the file.
Close fnum

CreateFromPicture = file_name
SetWallPaper App.Path & "\black.bmp"

Call SelectObject(memdc, oldbmp)
Call DeleteDC(memdc)

End Function

Public Function PictureFromByteStream(ByRef b() As Byte) As IPicture
Dim LowerBound As Long
Dim ByteCount As Long
Dim hMem As Long
Dim lpMem As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown

On Error GoTo Err_Init
If UBound(b, 1) < 0 Then
Exit Function
End If

LowerBound = LBound(b)
ByteCount = (UBound(b) - LowerBound) + 1
hMem = GlobalAlloc(&H2, ByteCount)
If hMem <> 0 Then
lpMem = GlobalLock(hMem)
If lpMem <> 0 Then
MoveMemory ByVal lpMem, b(LowerBound), ByteCount
Call GlobalUnlock(hMem)
If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture) = 0 Then
Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture, PictureFromByteStream)
End If
End If
End If
End If

Exit Function

Err_Init:
If Err.Number = 9 Then
'Uninitialized array
'MsgBox "???????????!"
Else
'MsgBox Err.Number & " - " & Err.Description
End If
End Function

 

'''''''''''''''''''''''''''' TEST '''''''''''''''''''''''''''

Public Sub TestConvertToBMP(filepath As String)

Dim iFile As Integer
iFile = FreeFile()
'Dim filepath As String
' filepath = "jpg, jpeg, gif, png, bmp, ... any image file path"
Dim bindata() As Byte
Open filepath For Binary As iFile
lByteLen = LOF(iFile)
ReDim bindata(1 To lByteLen)
Get iFile, 1, bindata
Close iFile

On Error Resume Next
Set p = PictureFromByteStream(bindata)

If Err.Number <> 0 Or VarType(p) <> vbDataObject Then
If Err.Number <> 0 Then
'MsgBox Err.Description & ", We will use normal avatar!"
End If
'Image26.Picture = App.Path & "\normal.png"
Else
Dim bin() As Byte
ReDim bin(0)
'On Error Resume Next
bmppath = CreateFromPicture(p, bin)
End If

End Sub

Public Function GetWallPaperPathAnConvert() As String
Dim RegVal As Long: Dim StrData As String: Dim LenVal As Long: Dim hKey As Long
RegVal = RegOpenKeyEx(HKEY_CURRENT_USER, "Control Panel\Desktop", 0, KEY_READ Or KEY_WOW64_64KEY, hKey)
If RegVal = 0 Then
RegVal = RegQueryValueEx(hKey, "Wallpaper", 0, REG_SZ, ByVal 0, LenVal)
If RegVal = 0 Then
StrData = Space(LenVal)
RegVal = RegQueryValueEx(hKey, "Wallpaper", 0, REG_SZ, ByVal StrData, LenVal)
If RegVal = 0 Then
StrData = left(StrData, InStr(StrData, vbNullChar) - 1)
'Picture11.Picture = LoadPicture(StrData)
TestConvertToBMP StrData
GetWallPaperPathAnConvert = StrData
Exit Function
End If
End If
End If
GetWallPaperPathAnConvert = "Failed"
End Function

Public Sub SetWallPaper(StrData As String)
'On Error Resume Next
SystemParametersInfo SPI_SETDESKWALLPAPER, 0, StrData, &H1 Or &H2
End Sub
