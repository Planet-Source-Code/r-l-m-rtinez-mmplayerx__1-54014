Attribute VB_Name = "mMouseWheel"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbSrc As Long)

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WHEEL_DELTA = 120
Private Const MK_LBUTTON = &H1
Private Const MK_RBUTTON = &H2
Private Const MK_SHIFT = &H4
Private Const MK_CONTROL = &H8
Private Const MK_MBUTTON = &H10
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private lpPrevWndProc As Long

Function LoWord(ByVal dwDoubleWord As Long) As Integer
    Call CopyMemory(LoWord, dwDoubleWord, 2)
End Function

Function HiWord(ByVal dwDoubleWord As Long) As Integer
    Call CopyMemory(HiWord, ByVal VarPtr(dwDoubleWord) + 2, 2)
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim fwKeys As Integer, zDelta As Integer
    Select Case uMsg
    Case WM_MBUTTONDOWN
        '/* The Wheel button is down
    Case WM_MBUTTONUP
        '/* The Wheel button isn't down anymore
    Case WM_MBUTTONDBLCLK
        '/ The Wheel button has been double-clicked
    Case WM_MOUSEWHEEL
        fwKeys = LoWord(wParam)
        zDelta = HiWord(wParam) / WHEEL_DELTA
        '/* Wheel rotate
        '/* abs(zDelta) ---> Ticks,Points
        '/* zDelta > 0  ---> Rotate forward
        '/* zDelta < 0  ---> Rotate backward
                
        If zDelta > 0 Then ' forward
          MusicMp3.Ajust_Volume MusicMp3.imgNormal(16).Top - 2  'Mas volumen
        Else '/* backward
          MusicMp3.Ajust_Volume MusicMp3.imgNormal(16).Top + 2  'Menos volumen
        End If
       
        ' If (fwKeys And MK_LBUTTON) = MK_LBUTTON Then '/* The left button was down
        ' If (fwKeys And MK_RBUTTON) = MK_RBUTTON Then '/* The right button was down
        ' If (fwKeys And MK_SHIFT) = MK_SHIFT Then     '/* The shift key was down
        ' If (fwKeys And MK_CONTROL) = MK_CONTROL Then '/* The ctrl key was down
        ' If (fwKeys And MK_MBUTTON) = MK_MBUTTON Then '/* The Wheel button was down
    End Select
    
    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub Hook()
 On Error Resume Next
    lpPrevWndProc = SetWindowLong(MusicMp3.PicMusic.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
  On Error Resume Next
    Call SetWindowLong(MusicMp3.PicMusic.hwnd, GWL_WNDPROC, lpPrevWndProc)
End Sub
