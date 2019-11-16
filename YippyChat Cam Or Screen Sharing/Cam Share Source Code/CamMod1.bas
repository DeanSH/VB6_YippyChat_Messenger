Attribute VB_Name = "CamMod1"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Wsize As String
Public Hsize As String
Public Sratio As String
Public ScaleDown As Long

Public Desktop As Boolean
Public WhoAmI As String
Public TheIP As String
Public sFile2 As String
Public sFile3 As String
Public ImageReady As Boolean
Public hHwnd As Long ' Handle to preview window

Type imgdes
    ibuff As Long
    stx As Long
    sty As Long
    endx As Long
    endy As Long
    buffwidth As Long
    palette As Long
    colors As Long
    imgtype As Long
    bmh As Long
    hBitmap As Long
    End Type


Type BITMAPINFOHEADER
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
    
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function bmpinfo Lib "VIC32.DLL" (ByVal Fname As String, bdat As BITMAPINFOHEADER) As Long
Declare Function allocimage Lib "VIC32.DLL" (Image As imgdes, ByVal wid As Long, ByVal leng As Long, ByVal BPPixel As Long) As Long
Declare Function loadbmp Lib "VIC32.DLL" (ByVal Fname As String, desimg As imgdes) As Long
Declare Sub freeimage Lib "VIC32.DLL" (Image As imgdes)
Declare Function convert1bitto8bit Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes) As Long
Declare Sub copyimgdes Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes)
Declare Function savejpg Lib "VIC32.DLL" (ByVal Fname As String, srcimg As imgdes, ByVal Quality As Long) As Long
    'end declarations
    'the sub

Public Sub StayOnTop(frm As Form)
Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub BMPtoJPG(Thebmp As String, Thejpg As String, Quality As Long)
    Dim tmpimage As imgdes ' Image descriptors
    Dim tmp2image As imgdes
    Dim rcode As Long
    Dim vbitcount As Long
    Dim bdat As BITMAPINFOHEADER ' Reserve space For BMP struct
    Dim bmp_fname As String
    Dim jpg_fname As String
    bmp_fname = Thebmp
    jpg_fname = Thejpg
    ' Get info on the file we're to load
    rcode = bmpinfo(bmp_fname, bdat)


    If (rcode <> NO_ERROR) Then
        'cannot find file!
        Exit Sub
    End If
    vbitcount = bdat.biBitCount


    If (vbitcount >= 16) Then ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
        vbitcount = 24
    End If
    ' Allocate space for an image
    rcode = allocimage(tmpimage, bdat.biWidth, bdat.biHeight, vbitcount)


    If (rcode <> NO_ERROR) Then
        'not enuf memory!
        Exit Sub
    End If
    ' Load image
    rcode = loadbmp(bmp_fname, tmpimage)


    If (rcode <> NO_ERROR) Then
        freeimage tmpimage ' Free image On Error
        'cannot load file
        Exit Sub
    End If


    If (vbitcount = 1) Then ' If we loaded a 1-bit image, convert To 8-bit grayscale
        ' because jpeg only supports 8-bit grays
        '     cale or 24-bit color images
        rcode = allocimage(tmp2image, bdat.biWidth, bdat.biHeight, 8)


        If (rcode = NO_ERROR) Then
            rcode = convert1bitto8bit(tmpimage, tmp2image)
            freeimage tmpimage ' Replace 1-bit image With grayscale image
            copyimgdes tmp2image, tmpimage
        End If
    End If
    ' Save image
    rcode = savejpg(jpg_fname, tmpimage, Quality)
    freeimage tmpimage
End Sub

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

Public Function InList(Data As String, List45 As ListBox) As Boolean
On Error Resume Next
If List45.ListCount = 0 Then InList = False: Exit Function
Dim I As Integer
For I = 0 To List45.ListCount - 1
If UCase(Data) = UCase(List45.List(I)) Then InList = True: Exit Function
Next I
InList = False
End Function

Public Function RemoveList(Data As String, List45 As ListBox)
On Error Resume Next
If List45.ListCount = 0 Then Exit Function
Dim I As Integer
For I = 0 To List45.ListCount - 1
If UCase(Data) = UCase(List45.List(I)) Then List45.RemoveItem I: Exit Function
Next I
End Function

Public Sub Pause2(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub

Public Function ResizeImage( _
    ByVal Original As WIA.ImageFile, _
    ByVal WidthPixels As Long, _
    ByVal HeightPixels As Long) As WIA.ImageFile

    'Scale the photo to fit supplied dimensions w/o distortion.
    With New WIA.ImageProcess
        .Filters.Add .FilterInfos!Scale.FilterID
        With .Filters(1).Properties
            '!PreserveAspectRatio = True by default, so just:
            !MaximumWidth = WidthPixels
            !MaximumHeight = HeightPixels
        End With
        Set ResizeImage = .Apply(Original)
    End With
End Function

