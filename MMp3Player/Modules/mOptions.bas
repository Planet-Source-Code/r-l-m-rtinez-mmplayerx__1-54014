Attribute VB_Name = "mSysIco"
Option Explicit


Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'// constantes para la api Shell_NotifyIcon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'// constantes para capturar los eventos del formulario
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'// api para mostrar un icono en la barra de sistema con los para metros enviados
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
             (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub ColocarIcono(ByVal hwnd As Long, ByVal hIcon As Long, ByVal sToolTip As String)
'// colocar icono en la barra de sistema
'// parametros
'//  [hWnd] -> de algun objeto para identificarlo despues
'//  [hIcon] -> Icono a mostrar
'//  [sToolTip] -> Mensaje a mostrar como tooltiptext
Dim udtNOTIFYICONDATA As NOTIFYICONDATA
With udtNOTIFYICONDATA
   .cbSize = Len(udtNOTIFYICONDATA)
   .hwnd = hwnd
   '.UID = 1&
   .UID = vbNull
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallbackMessage = WM_MOUSEMOVE
   .hIcon = hIcon
   If IsEmpty(sToolTip) Then
      .szTip = "" & vbNullChar
   Else
      .szTip = sToolTip & vbNullChar
   End If
End With

'agregar a la Barra
Shell_NotifyIcon NIM_ADD, udtNOTIFYICONDATA
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub CambiarIcono(ByVal hwnd As Long, ByVal hIcon As Long, Optional ByVal sToolTip As String)
'// colocar icono en la barra de sistema
'// parametros
'//  [hWnd] -> de algun objeto para identificarlo despues
'//  [hIcon] -> Icono a mostrar
'//  [sToolTip] -> Mensaje a mostrar como tooltiptext

Dim udtNOTIFYICONDATA As NOTIFYICONDATA

With udtNOTIFYICONDATA
   .cbSize = Len(udtNOTIFYICONDATA)
   .hwnd = hwnd
   '.UID = 1&
   .UID = vbNull
   .hIcon = hIcon
   .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   .uCallbackMessage = WM_MOUSEMOVE
   If IsEmpty(sToolTip) Then
      .szTip = "" & vbNullChar
   Else
      .szTip = sToolTip & vbNullChar
   End If
End With

'agregar a la Barra
Shell_NotifyIcon NIM_MODIFY, udtNOTIFYICONDATA

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub QuitarIcono(ByVal hwnd As Long)
'// Quitar icono
'// parametros
'// [hWnd] -> Manejador de objeto indicado cuando agramos el icono para quitarlo
Dim udtNOTIFYICONDATA As NOTIFYICONDATA
With udtNOTIFYICONDATA
   .cbSize = Len(udtNOTIFYICONDATA)
   .hwnd = hwnd
   '.UID = 1&
   .UID = vbNull
End With

Shell_NotifyIcon NIM_DELETE, udtNOTIFYICONDATA
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

