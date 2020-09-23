Attribute VB_Name = "Mod_Cloak"
Option Explicit

'I did not make the transparent code and I have no idea who really did.
'I got the code for transparents from the mIRC DDE project that was posted
'on PCS at http://www.planet-source-code.com/xq/ASP/txtCodeId.11275/lngWId.1/qx/vb/scripts/ShowCode.htm
'

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32" (ByVal Hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type SIZE
    cX As Long
    cY As Long
End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Const WS_EX_LAYERED = &H80000
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H1
Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
Public Const AC_SRC_NO_ALPHA = &H2
Public Const AC_DST_NO_PREMULT_ALPHA = &H10
Public Const AC_DST_NO_ALPHA = &H20
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4

Public lRet As Long

Function CheckLayered(ByVal Hwnd As Long) As Boolean
    lRet = GetWindowLong(Hwnd, GWL_EXSTYLE)
    If (lRet And WS_EX_LAYERED) = WS_EX_LAYERED Then
        CheckLayered = True
    Else
        CheckLayered = False
    End If
End Function

Function SetLayered(ByVal Hwnd As Long, SetAs As Boolean, bAlpha As Byte)
    lRet = GetWindowLong(Hwnd, GWL_EXSTYLE)
    If SetAs = True Then
        lRet = lRet Or WS_EX_LAYERED
    Else
        lRet = lRet And Not WS_EX_LAYERED
    End If
    SetWindowLong Hwnd, GWL_EXSTYLE, lRet
    SetLayeredWindowAttributes Hwnd, 0, bAlpha, LWA_ALPHA
End Function


