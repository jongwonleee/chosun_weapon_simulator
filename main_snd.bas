Attribute VB_Name = "main_snd"
Private DX As New DirectX7
Private DS As DirectSound

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCPAINT = &HEE0086
Public Const BLACKNESS = &H42
Public Const SRCAND = &H8800C6


Public bg_main As DirectSoundBuffer
Public eff_boing(6) As DirectSoundBuffer
Public bg_start As DirectSoundBuffer
Public chain As DirectSoundBuffer

Public Sub sound_init(Main_form As Form)
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel Main_form.hWnd, DSSCL_PRIORITY

    
    DSLoadWave bg_main, App.Path + "\sound\bg_main.wav" '메인 bg
    DSLoadWave bg_start, App.Path + "\sound\bg_start.wav" '스타트 bg
    DSLoadWave eff_boing(1), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave eff_boing(2), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave eff_boing(3), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave eff_boing(4), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave eff_boing(5), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave eff_boing(6), App.Path + "\sound\Spring-Boing.wav" '스타트
    DSLoadWave chain, App.Path + "\sound\Chain Dropped.wav" '체인
    
End Sub

Public Sub DSLoadWave(dsbuffer As DirectSoundBuffer, ByVal file As String)
    On Error GoTo s10
        Dim bufferdesc As DSBUFFERDESC
        Dim waveformat As WAVEFORMATEX
        bufferdesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        waveformat.nFormatTag = WAVE_FORMAT_PCM
        waveformat.nChannels = 2 ' 채널
        waveformat.lSamplesPerSec = 44100 ' 샘플수
        waveformat.nBitsPerSample = 16 ' 비트
        waveformat.nBlockAlign = waveformat.nBitsPerSample / 8 & waveformat.nChannels ' 샘플당 바이트
        waveformat.lAvgBytesPerSec = waveformat.lSamplesPerSec * waveformat.nBlockAlign ' 초당 샘플 바이트
        
        Set dsbuffer = DS.CreateSoundBufferFromFile(file, bufferdesc, waveformat)
s10:
End Sub

Public Sub playsound(dsbuffer As DirectSoundBuffer, Optional ByVal opt As Integer)
On Local Error Resume Next
    If opt = 1 Then
        dsbuffer.Play DSBPLAY_LOOPING
        dsbuffer.SetVolume 0
    ElseIf opt = 2 Then
        dsbuffer.Play DSBPLAY_DEFAULT
        dsbuffer.SetVolume 0
    ElseIf opt = 3 Then
        dsbuffer.Play DSBPLAY_DEFAULT
        dsbuffer.SetVolume -1000
    ElseIf opt = 4 Then
        dsbuffer.Play DSBPLAY_DEFAULT
        dsbuffer.SetVolume -2000
    ElseIf opt = 5 Then
        dsbuffer.Play DSBPLAY_LOOPING
        dsbuffer.SetVolume -1000
    End If
End Sub

Public Sub stopsound(dsbuffer As DirectSoundBuffer)
    dsbuffer.Stop
    dsbuffer.SetCurrentPosition 0
    End Sub


