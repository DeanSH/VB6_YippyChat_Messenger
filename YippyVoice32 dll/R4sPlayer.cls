VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "R4sPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event onSoundComplete()

Implements StreamEvent

Private Const TRUESPEECH = 34

Private wfx             As New Wavefx
Private fSoundBuffer    As New StreamIO

Private fSoundFormat      As WAVEFORMATEX
Private fTrueSpeechFormat As WAVEFORMATEX

Private fSoundStream      As SoundStream
Private fStreamConverter  As StreamConverter


Private Sub Class_Initialize()
  
  'set the truespeech format attributes.check out DSP group's website for more information.
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

Public Sub Initalize()

 Set fSoundStream = New SoundStream
     fSoundFormat = wfx.createFormat(1, 8000, 16)   'typical pcm 16-bit,mono format.
     
     fSoundStream.setSoundDescription fSoundFormat, Me, OverAllBytes  '1440 bytes of 16-bit pcm data at a sample rate of 8khz can be de/compressed into 96 bytes of truespeech data.exactly what we need for yahoo voice.
     
     
  'the only thing ima explain in this example is how to use the StreamConverter to convert between wave formats.
  'in this example i have chosen dsp group truespeech format.
  Set fStreamConverter = New Wavefx32.StreamConverter   'create an instance of the converter.
     fStreamConverter.streamOpen fTrueSpeechFormat, fSoundFormat  'open the conversion stream,specifying the input and output formats.

End Sub

Public Sub PlayWave(ByVal trueSpeechData As String)
  Dim wavData() As Byte
     wavData = StrConv(trueSpeechData, vbFromUnicode)
     
     fSoundBuffer.Write_ wavData(), UBound(wavData) - LBound(wavData) + 1, 0
          If (fSoundBuffer.chunkSize = 4) Then fSoundStream.beginWrite
End Sub


Private Sub Class_Terminate()
  fSoundStream.closeSound
     Set fSoundStream = Nothing
End Sub

Private Sub StreamEvent_onCapture(waveBuffer() As Byte, lBytesCaptured As Long)
' not implemented in this class.
End Sub

Private Function StreamEvent_onWrite(waveBuffer() As Byte, lBufferBytes As Long) As Long
  Dim tsWavData() As Byte
  Dim waveLength As Long
     
     If (fSoundBuffer.chunkSize < 1) Then
        fSoundBuffer.Clear
        RaiseEvent onSoundComplete
        StreamEvent_onWrite = 0
        
     ElseIf (fSoundBuffer.chunkSize >= 1) Then
     
        waveLength = fSoundBuffer.Read_(tsWavData(), 0)  'read the next block of data in line.
        waveBuffer = fStreamConverter.Convert(tsWavData(), waveLength)  'convert this block back to pcm.
        
        lBufferBytes = UBound(waveBuffer)
        StreamEvent_onWrite = 1
        
     End If
End Function

