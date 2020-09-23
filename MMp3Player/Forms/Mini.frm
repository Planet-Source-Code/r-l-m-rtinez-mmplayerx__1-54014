VERSION 5.00
Begin VB.Form frmMini 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Mini.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMini 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      Begin VB.PictureBox picFondo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   330
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   11
         Top             =   1410
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox picBotones 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2430
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   10
         Top             =   1125
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   6
         Left            =   3990
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   9
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   5
         Left            =   3810
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   8
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   4
         Left            =   3555
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   7
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   3
         Left            =   3360
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   6
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   3165
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   5
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   2955
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   4
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   2745
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   3
         Top             =   240
         Width           =   120
      End
      Begin VB.PictureBox picScroll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   705
         MousePointer    =   99  'Custom
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   124
         TabIndex        =   1
         Top             =   210
         Width           =   1860
      End
      Begin VB.Label lblTiempoT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   105
         TabIndex        =   2
         Top             =   195
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------- posicion minimascara ------------------------------------
Dim bolDragMini As Boolean
Dim StartDragX As Single, StartDragY As Single
Dim rWorkArea As RECT
Dim GraphicsHeight As Integer, desAlto As Integer, desAncho As Integer, orgX As Integer, orgAncho As Integer, orgAlto As Integer
Private mAttachedToRight As Boolean
Private mAttachedToLeft As Boolean
Private mAttachedToTop As Boolean
Private mAttachedToBottom As Boolean
Private mSnapDistance As Long
Public bolTimeAct As Boolean


Private Sub Form_KeyPress(KeyAscii As Integer)
With MusicMp3
 If KeyAscii = 45 Then .Ajust_Volume .imgNormal(16).Top + 2  '-
 If KeyAscii = 43 Then .Ajust_Volume .imgNormal(16).Top - 2  '+
 If KeyAscii = 65 Or KeyAscii = 97 Then .Five_Seg_Backward 'A Atras 5 seg
 If KeyAscii = 68 Or KeyAscii = 100 Then .Five_Seg_Forward 'D Adelante 5 seg
End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
With MusicMp3
  If KeyCode = 90 Then .Previous_Track 'Z
  If KeyCode = 88 Then .Play 'X
  If KeyCode = 67 Then .Pause_Play 'C
  If KeyCode = 86 Then .Stop_Player 'V
  If KeyCode = 66 Then .Next_Track 'B
  If Shift = vbShiftMask And KeyCode = 226 Then .Next_Album: Exit Sub ' > Siguiente Album
  If KeyCode = 226 Then .Previous_Album ' < Anterior Album
  If KeyCode = 76 Then .Front_Click 'L Cambiar caratula
  If KeyCode = 73 Then .Intro 'I Intro 10 seg
  If KeyCode = 82 Then .Player_Repeat 'R Repetir
  If KeyCode = 83 Then .Player_Mute 'S Silencio
If KeyCode = 81 Then 'Q Orden aleatorio Album
 frmPopUp.Menu_Aleatorio_Album
End If
If KeyCode = 87 Then 'W Orden aleatorio coleccion
 frmPopUp.Menu_Aleatorio_Coleccion
End If
  If KeyCode = 77 Then frmPopUp.MostaRCaratulA  'M Mostrar caratula
  If KeyCode = 70 Then frmPopUp.NuevABusQuEdA 'F Nueva busqueda
End With

End Sub

Private Sub Form_Load()
    mSnapDistance = 10 * Screen.TwipsPerPixelX
End Sub

Private Sub lblTiempoT_DblClick()
 '// show diferent curent time
 
 bolTimeAct = Not bolTimeAct
End Sub

Private Sub lblTiempoT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then MiniMaskDown X, Y
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal

End Sub

Private Sub lblTiempoT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then MiniMaskMove X, Y
End Sub

Private Sub lblTiempoT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bolDragMini = False
End Sub

Private Sub picMini_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err
 If Button = vbLeftButton Then MiniMaskDown X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
Exit Sub
err:
  bolDragMini = False
End Sub

Sub MiniMaskMove(X As Single, Y As Single)
 On Error Resume Next
  Dim DiffX As Long, DiffY As Long
  Dim NewX As Long, NewY As Long
  Dim ToLeftDistance As Long
  Dim ToRightDistance As Long
  Dim ToTopDistance As Long
  Dim ToBottomDistance As Long

 '// si estamos arrastrando
 If bolDragMini = True Then
    '// resta para mantener la posicion
    '// del cursor en la posicion inicial del objeto
    DiffX = X - StartDragX
    DiffY = Y - StartDragY
  
   If DiffX = 0 And DiffY = 0 Then Exit Sub
     '// obtener las coordenadas corectas
     NewX = Me.left + DiffX
     NewY = Me.Top + DiffY

    '// Enkontrar los bordes del escritorio
    
    
    ToRightDistance = rWorkArea.Right - (NewX + Me.Width)
    ToLeftDistance = NewX - rWorkArea.left
    ToBottomDistance = rWorkArea.Bottom - (NewY + Me.Height)
    ToTopDistance = NewY - rWorkArea.Top
    
    '// si no esta anklado
    If Not mAttachedToBottom Then
        '// si esta en el area minima para arrastrarse para abajo
        If Abs(ToBottomDistance) <= mSnapDistance Then
            '// anklar al borde de abajo
            NewY = rWorkArea.Bottom - Me.Height
            mAttachedToBottom = True
        End If
    Else
        
        If Abs(ToBottomDistance) > mSnapDistance Then
            '// Romper el anklado
            mAttachedToBottom = False
        Else
            '// mantener la actual posicion
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToTop Then
        If Abs(ToTopDistance) <= mSnapDistance Then
            NewY = rWorkArea.Top
            mAttachedToTop = True
        End If
    Else
        If Abs(ToTopDistance) > mSnapDistance Then
            mAttachedToTop = False
        Else
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToRight Then
        If Abs(ToRightDistance) <= mSnapDistance Then
            NewX = rWorkArea.Right - Me.Width
            mAttachedToRight = True
        End If
    Else
        If Abs(ToRightDistance) > mSnapDistance Then
            mAttachedToRight = False
        Else
            NewX = Me.left
        End If
    End If

    If Not mAttachedToLeft Then
        If Abs(ToLeftDistance) <= mSnapDistance Then
            NewX = rWorkArea.left
            mAttachedToLeft = True
        End If
    Else
        If Abs(ToLeftDistance) > mSnapDistance Then
            mAttachedToLeft = False
        Else
            NewX = Me.left
        End If
    End If
   
   '// mover a la actual posicion
   Me.Move NewX, NewY
'   SetWindowPos frmMini.hwnd, -2, NewX / Screen.TwipsPerPixelX, NewY / Screen.TwipsPerPixelY, _
'      Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 0
  
End If

End Sub

Private Sub picMini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MiniMaskMove X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
End Sub

Private Sub picMini_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 bolDragMini = False
End Sub

Sub MiniMaskDown(X As Single, Y As Single)
 On Error Resume Next
    '// Obtener el Area de trabajo en rWorkArea
    '// del escritorio sin kontar la taskbar
    
    'SystemParametersInfo SPI_GETWORKAREA, 0&, rWorkArea, 0&
    SystemGetWorkArea SPI_GETWORKAREA, 0&, rWorkArea, 0&
    
    '// Convretirlos de pixeles a twips
    rWorkArea.Top = rWorkArea.Top * Screen.TwipsPerPixelY
    rWorkArea.left = rWorkArea.left * Screen.TwipsPerPixelX
    rWorkArea.Bottom = rWorkArea.Bottom * Screen.TwipsPerPixelY
    rWorkArea.Right = rWorkArea.Right * Screen.TwipsPerPixelX
    
    '// variable para empezar a arrastrar
    bolDragMini = True
    '// almacenar las coordenadas iniciales
    StartDragX = X
    StartDragY = Y
     
End Sub


Private Sub picNormal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Index > 0 And Index < 4 Then Exit Sub
        Images_Buttons Index, True
End Sub

Sub Images_Buttons(IndexButton As Integer, Active As Boolean)
    
    desAncho = picBotones.ScaleWidth / 7
    desAlto = picBotones.ScaleHeight / 2
    orgAncho = picBotones.ScaleWidth / 7
    orgAlto = picBotones.ScaleHeight / 2
  
  orgX = (IndexButton) * (picBotones.ScaleWidth / 7)
 If Active = True Then
    GraphicsHeight = picBotones.ScaleHeight / 2
    picNormal(IndexButton).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
 Else
    GraphicsHeight = 0
    picNormal(IndexButton).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
 End If
End Sub

Private Sub picNormal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Index > 0 And Index < 4 Then GoTo etiqueta
        Images_Buttons Index, False
etiqueta:
 If Index = 0 Then MusicMp3.Previous_Track
 If Index = 1 Then MusicMp3.Play
 If Index = 2 Then MusicMp3.Pause_Play
 If Index = 3 Then MusicMp3.Stop_Player
 If Index = 4 Then MusicMp3.Next_Track
 If Index = 6 Then Unload MusicMp3
 If Index = 5 Then Change_Mask False
End Sub


Private Sub picScroll_DblClick()
If MusicMp3.bolToyBuscando = True Then MusicMp3.bolToyBuscando = False
 If (TextWidth(ScrollText) / 15) <= picScroll.ScaleWidth Then Exit Sub
 MusicMp3.Timer_Texto.Enabled = Not MusicMp3.Timer_Texto.Enabled
End Sub

Private Sub picScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then MiniMaskDown X, Y
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

Private Sub picScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 bolDragMini = False
End Sub
