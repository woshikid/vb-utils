Attribute VB_Name = "WaveIn"
Option Explicit
Public Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function WaveInStart Lib "winmm.dll" Alias "waveInStart" (ByVal hWaveIn As Long) As Long
Public Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public toend As Boolean

Public Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
End Type
Public Type WAVEHDR
        lpData As Long
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        dwLoops As Long
        lpNext As Long
        Reserved As Long
End Type
Public Const WAVE_FORMAT_PCM = 1
Public Const CHANNEL = 1
Public Const SAMPLES = 11025&
Public Const BITS = 16
Public Const WAVE_MAPPER = -1&
Public Const CALLBACK_FUNCTION = &H30000
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_SHARE = &H2000
Public Const GMEM_ZEROINIT = &H40
Public Const BUF_SIZE = 32768
Public Const MM_WIM_DATA = &H3C0
Public Const MM_WOM_DONE = &H3BD
Private WavInFmt As WAVEFORMAT
Private hWaveIn As Long
Private hMemIn As Long
Private inHdr As WAVEHDR

Private WavOutFmt As WAVEFORMAT
Private hWaveOut As Long
Private hMemOut As Long
Private outHdr As WAVEHDR

Public Const WIM_OPEN = &H1
Public Const WIM_CLOSE = &H2
Public Const WIM_DATA = &H3
Public Const WOM_OPEN = &H4
Public Const WOM_CLOSE = &H5
Public Const WOM_DONE = &H6

Public waveReady As Boolean
Public waveData() As Byte
Dim a() As Byte
Dim b() As Byte
Private prevWndProc As Long
Private DesthWnd As Long
Public Const GWL_WNDPROC = (-4)
Public Const CALLBACK_WINDOW = &H10000

Public Sub WaveInInit(ByVal nCh As Integer, ByVal Sample As Long, ByVal nBits As Integer)
    If waveReady = True Then Exit Sub
    prevWndProc = GetWindowLong(Form1.hWnd, GWL_WNDPROC)
    SetWindowLong Form1.hWnd, GWL_WNDPROC, AddressOf WaveInProc
    Dim ret As Long
    WavInFmt.wFormatTag = WAVE_FORMAT_PCM
    WavInFmt.nChannels = nCh
    WavInFmt.nSamplesPerSec = Sample
    WavInFmt.nBlockAlign = nBits * nCh / 8
    WavInFmt.wBitsPerSample = nBits
    WavInFmt.cbSize = 0
    WavInFmt.nAvgBytesPerSec = nBits * Sample * nCh / 8
    ret = waveInOpen(hWaveIn, WAVE_MAPPER, WavInFmt, Form1.hWnd, 0, CALLBACK_WINDOW)
    If ret <> 0 Then Exit Sub
    
    'hMemIn = GlobalAlloc(GMEM_MOVEABLE + GMEM_SHARE + GMEM_ZEROINIT, BUF_SIZE)
    ReDim a(BUF_SIZE - 1)
    inHdr.lpData = VarPtr(a(0))
    inHdr.dwBufferLength = BUF_SIZE
    inHdr.dwFlags = 0
    inHdr.dwLoops = 0
    inHdr.dwUser = 0
    ret = waveInPrepareHeader(hWaveIn, inHdr, Len(inHdr))
    If ret = 0 Then waveReady = True
End Sub

Public Function WaveInProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Select Case Msg
    Case MM_WIM_DATA
        MsgBox "Buffer full."
        outHdr.lpData = inHdr.lpData
    Case MM_WOM_DONE
        MsgBox "Playback finished."

End Select

WaveInProc = CallWindowProc(prevWndProc, hWnd, Msg, wParam, lParam)

End Function

Public Sub WaveInRecord()
    If waveReady = False Then Exit Sub
    Dim ret As Long
    ret = waveInAddBuffer(hWaveIn, inHdr, Len(inHdr))
    If ret <> 0 Then Exit Sub
    ret = WaveInStart(hWaveIn)
    If ret <> 0 Then
        MsgBox "err"
        Exit Sub
    End If
End Sub

Public Sub WaveOutInit(ByVal nCh As Integer, ByVal Sample As Long, ByVal nBits As Integer)
    Dim ret As Long
    WavOutFmt.wFormatTag = WAVE_FORMAT_PCM
    WavOutFmt.nChannels = nCh
    WavOutFmt.nSamplesPerSec = Sample
    WavOutFmt.nBlockAlign = nBits * nCh / 8
    WavOutFmt.wBitsPerSample = nBits
    WavOutFmt.cbSize = 0
    WavOutFmt.nAvgBytesPerSec = nBits * Sample * nCh / 8
    ret = waveOutOpen(hWaveOut, WAVE_MAPPER, WavOutFmt, AddressOf WaveOutProc, 0, CALLBACK_FUNCTION)
    If ret <> 0 Then Exit Sub
    'hMemOut = GlobalAlloc(GMEM_MOVEABLE + GMEM_SHARE + GMEM_ZEROINIT, BUF_SIZE)
    ReDim b(BUF_SIZE - 1)
    outHdr.lpData = VarPtr(b(0))
    outHdr.dwBufferLength = BUF_SIZE
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0
    outHdr.dwUser = 0
    ret = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If ret <> 0 Then
        MsgBox "waveOutPrepareHeader"
        Exit Sub
    End If
End Sub

Public Sub WaveOutPlayback()
    waveOutWrite hWaveOut, outHdr, Len(outHdr)
End Sub

Public Sub WaveInDeinit()
    Dim ret As Long
    ret = waveInStop(hWaveIn)
    If ret <> 0 Then Exit Sub
    ret = waveInReset(hWaveIn)
    If ret <> 0 Then Exit Sub
    ret = waveInUnprepareHeader(hWaveIn, inHdr, Len(inHdr))
    If ret <> 0 Then Exit Sub
    GlobalUnlock hMemIn
    GlobalFree hMemIn
    ret = waveInClose(hWaveIn)
    If ret <> 0 Then
            Exit Sub
    End If
    SetWindowLong Form1.hWnd, GWL_WNDPROC, prevWndProc
End Sub

Public Sub WaveOutDeinit()
    Dim ret As Long
    ret = waveOutReset(hWaveOut)
    If ret <> 0 Then Exit Sub
    ret = waveOutUnprepareHeader(hWaveOut, outHdr, Len(outHdr))
    If ret <> 0 Then Exit Sub
    GlobalUnlock hMemOut
    GlobalFree hMemOut
    ret = waveOutClose(hWaveOut)
    If ret <> 0 Then
            Exit Sub
    End If
End Sub

Public Function WaveOutProc(ByVal hwi As Long, ByVal Msg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long) As Long
Select Case Msg
    Case MM_WOM_DONE
        If toend = False Then WaveOutPlayback
End Select
End Function
