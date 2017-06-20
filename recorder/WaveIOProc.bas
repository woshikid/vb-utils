Attribute VB_Name = "WaveIOProc"
Option Explicit

Public Const BUF_SIZE = 327680
Public Const SAMPLES = 22050&
Public Const BITS = 16
Public Const CHANNEL = 1

Public Const WIM_OPEN = &H1
Public Const WIM_CLOSE = &H2
Public Const WIM_DATA = &H3
Public Const WOM_OPEN = &H4
Public Const WOM_CLOSE = &H5
Public Const WOM_DONE = &H6


Private hWaveIn As Long
Private hWaveOut As Long
Private inHdr As WAVEHDR
Private outHdr As WAVEHDR
Private WavInFmt As WAVEFORMAT
Private WavOutFmt As WAVEFORMAT
Private hMemIn As Long
Private hMemOut As Long
Private prevWndProc As Long
Private DesthWnd As Long

Public Sub RegisterWinProc(ByVal hWnd As Long)
prevWndProc = GetWindowLong(hWnd, GWL_WNDPROC)
SetWindowLong hWnd, GWL_WNDPROC, AddressOf WndProc
DesthWnd = hWnd

End Sub

Public Sub UnRegisterWinProc(ByVal hWnd As Long)
SetWindowLong hWnd, GWL_WNDPROC, prevWndProc

End Sub

Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'If (Msg = WM_USER) Then
'    Select Case wParam
'        Case WIM_OPEN
'            MsgBox "WaveInOpen ok."
'        Case WIM_CLOSE
'            MsgBox "WaveInClose ok."
'        Case WIM_DATA
'            MsgBox "Buffer full."
'        Case WOM_OPEN
'            MsgBox "WaveOutOpen ok."
'        Case WOM_CLOSE
'            MsgBox "WaveOutClose ok."
'        Case WOM_DONE
'            MsgBox "Data sent."
'    End Select
'End If

Select Case Msg
    Case MM_WIM_DATA
        MsgBox "Buffer full."
        outHdr.lpData = inHdr.lpData
    Case MM_WIM_OPEN
'        MsgBox "WaveInOpen ok."
    Case MM_WIM_CLOSE
'        MsgBox "WaveInClose ok."
    Case MM_WOM_DONE
        MsgBox "Playback finished."
    Case MM_WOM_OPEN
'        MsgBox "WaveOutOpen ok."
    Case MM_WOM_CLOSE
'        MsgBox "WaveOutClose ok."
End Select

WndProc = CallWindowProc(prevWndProc, hWnd, Msg, wParam, lParam)
End Function

Public Function MywaveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long) As Long
Select Case (uMsg)
    Case MM_WIM_DATA
'        PostMessage DesthWnd, WM_USER, WIM_DATA, 0
    Case MM_WIM_OPEN
'        PostMessage DesthWnd, WM_USER, WIM_OPEN, 0
    Case MM_WIM_CLOSE
'        PostMessage DesthWnd, WM_USER, WIM_CLOSE, 0
End Select

End Function

Public Function MywaveOutProc(ByVal hwo As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long) As Long
Select Case (uMsg)
    Case MM_WOM_DONE
'        PostMessage DesthWnd, WM_USER, WOM_DONE, 0
    Case MM_WOM_OPEN
'        PostMessage DesthWnd, WM_USER, WOM_OPEN, 0
    Case MM_WOM_CLOSE
'        PostMessage DesthWnd, WM_USER, WOM_CLOSE, 0
End Select

End Function

Public Sub WaveInInit(ByVal nCh As Integer, ByVal Sample As Long, ByVal nBits As Integer)
Dim ret As Long
Dim Msg As String * 200
WavInFmt.wFormatTag = WAVE_FORMAT_PCM
WavInFmt.nChannels = nCh
WavInFmt.nSamplesPerSec = Sample
WavInFmt.nBlockAlign = nBits * nCh / 8
WavInFmt.wBitsPerSample = nBits
WavInFmt.cbSize = 0
WavInFmt.nAvgBytesPerSec = nBits * Sample * nCh / 8
ret = waveInOpen(hWaveIn, WAVE_MAPPER, WavInFmt, DesthWnd, 0, CALLBACK_WINDOW) 'AddressOf MywaveInProc, 0, CALLBACK_FUNCTION)
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInOpen:" & Msg
        Exit Sub
End If
hMemIn = GlobalAlloc(GMEM_MOVEABLE + GMEM_SHARE + GMEM_ZEROINIT, BUF_SIZE)
inHdr.lpData = GlobalLock(hMemIn)
inHdr.dwBufferLength = BUF_SIZE
inHdr.dwFlags = 0
inHdr.dwLoops = 0
inHdr.dwUser = 0
ret = waveInPrepareHeader(hWaveIn, inHdr, Len(inHdr))
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInPrepareHeader" & Msg
        Exit Sub
End If

End Sub

Public Sub WaveInRecord()
Dim ret As Long
Dim Msg As String * 200
ret = waveInAddBuffer(hWaveIn, inHdr, Len(inHdr))
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInAddBuffer" & Msg
        Exit Sub
End If
ret = WaveInStart(hWaveIn)
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInAddBuffer" & Msg
        Exit Sub
End If

End Sub
Public Sub WaveInDeinit()
Dim ret As Long
Dim Msg As String * 200
ret = waveInStop(hWaveIn)
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInStop" & Msg
        Exit Sub
End If
ret = waveInReset(hWaveIn)
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInReset" & Msg
        Exit Sub
End If
ret = waveInUnprepareHeader(hWaveIn, inHdr, Len(inHdr))
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInUnprepareHeader" & Msg
        Exit Sub
End If
GlobalUnlock hMemIn
GlobalFree hMemIn
ret = waveInClose(hWaveIn)
If ret <> 0 Then
        waveInGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveInClose" & Msg
        Exit Sub
End If

End Sub

Public Sub WaveOutInit(ByVal nCh As Integer, ByVal Sample As Long, ByVal nBits As Integer)
Dim ret As Long
Dim Msg As String * 200
WavOutFmt.wFormatTag = WAVE_FORMAT_PCM
WavOutFmt.nChannels = nCh
WavOutFmt.nSamplesPerSec = Sample
WavOutFmt.nBlockAlign = nBits * nCh / 8
WavOutFmt.wBitsPerSample = nBits
WavOutFmt.cbSize = 0
WavOutFmt.nAvgBytesPerSec = nBits * Sample * nCh / 8
ret = waveOutOpen(hWaveOut, WAVE_MAPPER, WavOutFmt, DesthWnd, 0, CALLBACK_WINDOW) ' AddressOf MywaveOutProc, 0, CALLBACK_FUNCTION)
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutOpen:" & Msg
        Exit Sub
End If
hMemOut = GlobalAlloc(GMEM_MOVEABLE + GMEM_SHARE + GMEM_ZEROINIT, BUF_SIZE)
outHdr.lpData = GlobalLock(hMemOut)
outHdr.dwBufferLength = BUF_SIZE
outHdr.dwFlags = 0
outHdr.dwLoops = 0
outHdr.dwUser = 0
ret = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutPrepareHeader" & Msg
        Exit Sub
End If

End Sub

Public Sub WaveOutPlayback()
Dim ret As Long
Dim Msg As String * 200

ret = waveOutWrite(hWaveOut, outHdr, Len(outHdr))
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutWrite" & Msg
        Exit Sub
End If

End Sub
Public Sub WaveOutDeinit()
Dim ret As Long
Dim Msg As String * 200
ret = waveOutReset(hWaveOut)
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutReset" & Msg
        Exit Sub
End If
ret = waveOutUnprepareHeader(hWaveOut, outHdr, Len(outHdr))
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutUnprepareHeader" & Msg
        Exit Sub
End If
GlobalUnlock hMemOut
GlobalFree hMemOut
ret = waveOutClose(hWaveOut)
If ret <> 0 Then
        waveOutGetErrorText ret, Msg, Len(Msg)
        MsgBox "waveOutClose" & Msg
        Exit Sub
End If

End Sub

