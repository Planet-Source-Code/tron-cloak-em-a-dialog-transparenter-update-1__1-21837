Attribute VB_Name = "Mod_Windows"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const BM_SETSTYLE = &HF4
Private Const BS_SOLID = 0

' I did not create the code all my self,
' What I did do was make the code readable
' instend of that crapy looking code style
' which I cannot see how people even update
' their projects when looking like that, anyway
' I also added some of my own code and style to it.
' I got most of code parts from the following
' URLs at PSC:
'             Transparent Window:  http://www.planet-source-code.com/xq/ASP/txtCodeId.13386/lngWId.1/qx/vb/scripts/ShowCode.htm
'   Close Window (partial title):  http://www.planet-source-code.com/xq/ASP/txtCodeId.8784/lngWId.1/qx/vb/scripts/ShowCode.htm
'               C++ Button Style:  http://www.planet-source-code.com/xq/ASP/txtCodeId.6122/lngWId.1/qx/vb/scripts/ShowCode.htm
'
' Date Code was edited [03/22/2001]

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_CLOSE = &H10
Public AppTitle As String
Public ApphWnd As Long

Declare Function EnumChildWindows Lib "user32.dll" (ByVal hwndParent As Long, ByVal lpenumfunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim sLength As Long, WinText As String, sLengthC As Long
Dim Buffer As String, ClassName As String, WinCap As String
Dim Retval As Long

Static WinNum As Integer
WinNum = WinNum + 1

    If Start = 1 Then
        Start = 0
    End If

    ClassName = Space(255)
    sLength = GetClassName(hwnd, ClassName, 255)
    ClassName = Left(ClassName, sLength)
    
    sLengthC = GetWindowTextLength(hwnd) + 1
    If sLengthC > 128 Then sLengthC = 128
    WinCap = ""
    
    If sLengthC > 1 Then
        Buffer = Space(sLengthC)
        Retval = GetWindowText(hwnd, Buffer, sLengthC)
        WinCap = Left(Buffer, sLengthC - 1)
    End If

    If OnlyCaps <> 1 Then
        frmMain.List_IDs.AddItem hwnd
        frmMain.List_Windows.AddItem WinCap
    End If
    
EnumChildProc = 1
End Function

Public Function CButton(Button As CommandButton) As Long
    SendMessage Button.hwnd, BM_SETSTYLE, BS_SOLID, 1
End Function
