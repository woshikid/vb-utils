Attribute VB_Name = "WaveIO"
Option Explicit

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Public Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveInGetID Lib "winmm.dll" (ByVal hWaveIn As Long, lpuDeviceID As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveInGetPosition Lib "winmm.dll" (ByVal hWaveIn As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveInMessage Lib "winmm.dll" (ByVal hWaveIn As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function WaveInStart Lib "winmm.dll" Alias "waveInStart" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutBreakLoop Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveOutGetID Lib "winmm.dll" (ByVal hWaveOut As Long, lpuDeviceID As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveOutGetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwPitch As Long) As Long
Public Declare Function waveOutGetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, lpdwRate As Long) As Long
Public Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Public Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function waveOutMessage Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal Msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Public Declare Function waveOutSetPitch Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwPitch As Long) As Long
Public Declare Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Public Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long

Public Const MAXPNAMELEN = 32

Public Type MMTIME
        wType As Long
        u As Long
End Type
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
Public Type WAVEINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
End Type
Public Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

Public Const MM_WIM_CLOSE = &H3BF
Public Const MM_WIM_DATA = &H3C0
Public Const MM_WIM_OPEN = &H3BE
Public Const MM_WOM_CLOSE = &H3BC
Public Const MM_WOM_DONE = &H3BD
Public Const MM_WOM_OPEN = &H3BB
Public Const WAVE_MAPPER = -1&

Public Const CALLBACK_WINDOW = &H10000
Public Const CALLBACK_FUNCTION = &H30000

Public Const WAVE_FORMAT_PCM = 1

Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_SHARE = &H2000
Public Const GMEM_ZEROINIT = &H40

Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400


