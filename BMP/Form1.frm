VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMPendecoder"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Decode"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "c:\out.cmp"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "c:\test.bmp"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "cmp file:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "bmp file:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
Dim d(2) As Byte
Dim scount As Integer
Dim sin As Boolean
Dim spos As Long
sin = False
Kill "c:\out.cmp"
Err.Clear
Open Text1.Text For Binary As #1
Open "c:\out.cmp" For Binary As #2
If Err Then
    MsgBox Err.Description
    Close #1
    Close #2
    Exit Sub
End If
Dim data() As Byte
ReDim data(LOF(1) - 1)
Dim data2() As Byte
ReDim data2(UBound(data))
Get #1, , data
Dim i As Long
Dim j As Integer
Dim k As Long
Do While i <= UBound(data)
    If i + 2 > UBound(data) Then
        data2(k) = data(i)
        k = k + 1
        If i + 1 = UBound(data) Then
            data2(k) = data(i + 1)
            k = k + 1
        End If
        Exit Do
    End If
    d(0) = data(i)
    d(1) = data(i + 1)
    d(2) = data(i + 2)
    i = i + 3
    For j = 1 To 254
        If i + 2 > UBound(data) Then Exit For
        If d(0) <> data(i) Or d(1) <> data(i + 1) Or d(2) <> data(i + 2) Then Exit For
        i = i + 3
    Next j
    If j = 1 Then
        If sin = False Then
            sin = True
            scount = 1
            data2(k) = 0
            data2(k + 1) = 1
            spos = k + 1
            k = k + 2
        Else
            scount = scount + 1
            If scount = 255 Then
                sin = False
                data2(spos) = scount
            End If
        End If
        data2(k) = d(0)
        data2(k + 1) = d(1)
        data2(k + 2) = d(2)
        k = k + 3
    Else
        If sin = True Then
            sin = False
            data2(spos) = scount
        End If
        data2(k) = j
        data2(k + 1) = d(0)
        data2(k + 2) = d(1)
        data2(k + 3) = d(2)
        k = k + 4
    End If
    If UBound(data2) - k < 1000 Then ReDim Preserve data2(UBound(data2) + 1000000)
Loop
If spos <> 0 Then data2(spos) = scount
ReDim Preserve data2(k - 1)
Put #2, , data2
Close #2
Close #1
MsgBox "output to C:\out.cmp", vbOKOnly, "done!"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Open Text2.Text For Binary As #1
Open "c:\out.org" For Binary As #2
If Err Then
    MsgBox Err.Description
    Close #1
    Close #2
    Exit Sub
End If
Dim data() As Byte
ReDim data(LOF(1) - 1)
Dim data2() As Byte
ReDim data2(UBound(data))
Get #1, , data
Dim i As Long
Dim j As Integer
Dim k As Integer
Dim m As Long
Do While i <= UBound(data)
    If i + 2 > UBound(data) Then
        data2(m) = data(i)
        m = m + 1
        If i + 1 = UBound(data) Then
            data2(m) = data(i + 1)
            m = m + 1
        End If
        Exit Do
    End If
    j = data(i)
    If j = 0 Then
        j = data(i + 1)
        i = i + 2
        For k = 1 To j
            data2(m) = data(i)
            data2(m + 1) = data(i + 1)
            data2(m + 2) = data(i + 2)
            i = i + 3
            m = m + 3
        Next k
    Else
        For k = 1 To j
            data2(m) = data(i + 1)
            data2(m + 1) = data(i + 2)
            data2(m + 2) = data(i + 3)
            m = m + 3
        Next k
        i = i + 4
    End If
    If UBound(data2) - m < 1000 Then ReDim Preserve data2(UBound(data2) + 1000000)
Loop
ReDim Preserve data2(m - 1)
Put #2, , data2
Close #2
Close #1
MsgBox "output to C:\out.org", vbOKOnly, "done!"
End Sub
