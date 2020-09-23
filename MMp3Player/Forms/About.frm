VERSION 5.00
Begin VB.Form frmAcerca 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " About..."
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   4170
      Top             =   1380
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1965
      Left            =   15
      Picture         =   "About.frx":000C
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   0
      Top             =   1485
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.geocities.com/skoria_36"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   555
      MouseIcon       =   "About.frx":4FB0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   15
      Picture         =   "About.frx":5102
      Top             =   45
      Width           =   3360
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'El texto actual a correr. Tambien se puede leer desde un archivo de texto
Private ScrollText As String
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Punto izquierdo superior del PICSCROLL
Dim RectHeight As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub RunMain()

Const IntervalTime As Long = 60 '// Velocidad variable del scroll del texto
'Muestra la forma
frmAcerca.Refresh
'Obtiene el tamaño del PICSCROLL y lo reemplaza por la constante DT_CALRECT
rt = DrawText(picScroll.hDC, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'Si marca error
    'MsgBox "Error scrolling text", vbCritical
Else
    '// obtener un rectangulo segun el ancho del piccscroll y alto
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Arregla la altura del PICSCROLL
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
 On Error Resume Next
  Me.Caption = Trim(arryLanguage(35))
  Me.Icon = MusicMp3.Icon
  bolAcercaShow = True
  ScrollText = "MUSIC MP3 PLAYER X " & vbCrLf & _
                "VERSION 1.1." & vbCrLf & vbCrLf & _
                "DEVELOPED BY:" & vbCrLf & _
                "<< RAUL MARTINEZ HERNANDEZ >>" & vbCrLf & _
                "VALLE DE SANTIAGO" & vbCrLf & _
                "GUANAJUATO - MEXICO" & vbCrLf & vbCrLf & _
                "MAY 2004" & vbCrLf & vbCrLf & _
                "If you have any ideas," & vbCrLf & _
                "comments, doubts, suggestions," & vbCrLf & _
                "bugs, skins, languages, etc," & vbCrLf & _
                "please email me." & vbCrLf & vbCrLf & _
                "E-mail :" & vbCrLf & _
                "escorpio36@hotmail.com" & vbCrLf & vbCrLf & _
                "Web Site :" & vbCrLf & _
                "www.geocities.com/skoria_36" & vbCrLf & vbCrLf & _
                "  ,,,,,  " & vbCrLf & _
                ".|_   _|." & vbCrLf & _
                "(| O-O |)" & vbCrLf & _
                " |  >  | " & vbCrLf & _
                " | \_/ | " & vbCrLf & _
                "  \___/  " & vbCrLf
                                
  Timer1.Enabled = True
  Me.left = (Screen.Width - Me.Width) / 2   '// centrar formulario
  Me.Top = (Screen.Height - Me.Height) / 2
  RunMain '// empezar a mover el texto
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFFFF
Label1.FontUnderline = False

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Unload(Cancel As Integer)
    bolAcercaShow = False
    Timer1.Enabled = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFFFF
Label1.FontUnderline = False

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label1.Move Label1.left + 1, Label1.Top + 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFFC0
Label1.FontUnderline = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 Dim lngRETURN As Long
 Label1.ForeColor = &HFFFFFF
 Label1.FontUnderline = False
 Label1.Move Label1.left - 1, Label1.Top - 1
 lngRETURN = ShellExecute(Me.hwnd, "Open", Label1.Caption, "", "", vbNormalFocus)
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub picScroll_Click()
 Timer1.Enabled = Not Timer1.Enabled
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Timer1_Timer()
        picScroll.Cls  '// borrar imagen anterior
        
        DrawText picScroll.hDC, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        '// Actualiza las coordenadas del rectangulo
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        '// Controla el PICSCROLL y lo reinicia si se sale de su limite(si termina)
        If DrawingRect.Top < -(RectHeight) Then '// Tiempo de reinicio
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
    DoEvents
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
