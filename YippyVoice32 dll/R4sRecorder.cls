VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "R4sRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements StreamEvent

Private Const TRUESPEECH = 34
Private wfx             As New Wavefx

Private fCaptureFormat    As WAVEFORMATEX
Private fTrueSpeechFormat As WAVEFORMATEX

Private fCaptureStream    As CaptureStream
Private fStreamConverter  As StreamConverter

Event onSoundCompressed(ByVal trueSpeechData As String, ByVal lBufferBytes As Long)


Private Sub Class_Initialize()
  
  With fTrueSpeechFormat
     .FormatTag = TRUESPEECH
     .channels = 1        'Mono.
     .SamplesPerSec = 8000 '8khz.
     .BitsPerSample = 1    'bit rate.
     .BlockAlign = 32
     .AvgBytesPerSec = 1067
     .cbSize = 32
     
              .extraBytes(0) = &H1    'required by the truespeech codec.
              .extraBytes(2) = &HF0   'required by the truespeech codec.
  End With
  
End Sub

Sub Record()
    
  Set fCaptureStream = New CaptureStream
     
     fCaptureFormat = wfx.createFormat(1, 8000, 16)
        
     Set fStreamConverter = New StreamConverter
        fStreamConverter.streamOpen fCaptureFormat, fTrueSpeechFormat
        
        With fCaptureStream
            .setCaptureDescription fCaptureFormat, Me, OverAllBytes
                Call .startCapture((-1))
        End With
  
End Sub

Sub EndRecord()
  Call fCaptureStream.stopCapture
     Set fCaptureStream = Nothing
  
  Call fStreamConverter.streamClose
     Set fStreamConverter = Nothing
End Sub

Private Sub StreamEvent_onCapture(waveBuffer() As Byte, lBytesCaptured As Long)
 On Error Resume Next
  Dim length As Long
  Dim wavData() As Byte
  Dim Datt As String

    wavData = fStreamConverter.Convert(waveBuffer(), lBytesCaptured)
    Datt = StrConv(wavData, vbUnicode)
    length = Len(Datt)
    If (length > 0) Then RaiseEvent onSoundCompressed(Datt, length)

End Sub

Private Function StreamEvent_onWrite(waveBuffer() As Byte, lBufferBytes As Long) As Long
'' not implemented in this class.
End Function



