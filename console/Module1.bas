Attribute VB_Name = "Module1"
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumherOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
'����API�������õ������г���
'GetStdHandle������ nStdHandle������ȡֵ
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
'SetConsoleTextAttribute������wAttributes������ȡֵ����RGB��ʽ��ϣ�
Private Const FOREGROUND_bLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80
Private Const FOREGROUND_YELLOW = FOREGROUND_RED Or FOREGROUND_GREEN
'SetConsoleMode������ģʽ
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'SetConsoleMode�����ģʽ
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
Private hConsoleIn As Long '����̨���ڵ� input handle
Private hConsoleOut As Long '����̨���ڵ�output handle
Private hConsoleErr As Long '����̨���ڵ�error handle

Private Sub Main()
    Dim szUserInput As String
    AllocConsole '���� console window
    SetConsoleTitle "VB����̨Ӧ�ó���"
    '����console window�ı���
    'ȡ��console window���������
    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
    SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or FOREGROUND_INTENSITY
    'ǰ�������̣���������
    ConsolePrint "What's your name?"
    szUserInput = ConsoleRead()
    If Not szUserInput = vbNullString Then
      ConsolePrint "Hello, " & szUserInput & "!" & vbCrLf
    Else
      ConsolePrint "You don't have a name?" & vbCrLf
    End If
    ConsolePrint vbCrLf & "Press enter to exit!"
    Call ConsoleRead
    FreeConsole '���� console window
End Sub

Private Sub ConsolePrint(szOut As String)
    WriteConsole hConsoleOut, szOut, Len(szOut), vbNull, vbNull
End Sub

Private Function ConsoleRead() As String
    Dim sUserInput As String * 256
    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
    '�ص��ַ�����β��&H00�ͻس������з�
    ConsoleRead = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)
End Function
