Attribute VB_Name = "Module1"
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumherOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
'定义API函数中用到的所有常量
'GetStdHandle函数的 nStdHandle参数的取值
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
'SetConsoleTextAttribute函数的wAttributes参数的取值（按RGB方式组合）
Private Const FOREGROUND_bLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80
Private Const FOREGROUND_YELLOW = FOREGROUND_RED Or FOREGROUND_GREEN
'SetConsoleMode的输入模式
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'SetConsoleMode的输出模式
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Private hConsoleIn As Long '控制台窗口的 input handle
Private hConsoleOut As Long '控制台窗口的output handle
Private hConsoleErr As Long '控制台窗口的error handle

Private Sub Main()
    Dim szUserInput As String
    AllocConsole '创建 console window
    SetConsoleTitle "VB控制台应用程序"
    '设置console window的标题
    '取得console window的三个句柄
    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
    SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    '前景：亮绿；背景：黑
    ConsolePrint "What's your name?"
    szUserInput = ConsoleRead()
    If Not szUserInput = vbNullString Then
      ConsolePrint "Hello, " & szUserInput & "!" & vbCrLf
    Else
      ConsolePrint "You don't have a name?" & vbCrLf
    End If
    ConsolePrint vbCrLf & "Press enter to exit!"
    Call ConsoleRead
    FreeConsole '销毁 console window
End Sub

Private Sub ConsolePrint(szOut As String)
    WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull
End Sub

Private Function ConsoleRead() As String
    Dim sUserInput As String * 256
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    '截掉字符串结尾的&H00和回车、换行符
    ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function
