VERSION 5.00
Begin VB.Form frmPopUp 
   Caption         =   " Opciones Generales"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Popup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBotonesMini 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   405
      Picture         =   "Popup.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   1455
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   1950
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   7
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picVol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   165
      Picture         =   "Popup.frx":0ED0
      ScaleHeight     =   1470
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      Picture         =   "Popup.frx":19CC
      ScaleHeight     =   360
      ScaleWidth      =   2400
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox picMini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   2580
      Picture         =   "Popup.frx":4F4F
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   5
      Top             =   105
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox PicMusic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3150
      Left            =   1620
      Picture         =   "Popup.frx":DB5B
      ScaleHeight     =   3150
      ScaleWidth      =   5595
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   5595
   End
   Begin VB.PictureBox picBotones 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   150
      Picture         =   "Popup.frx":4725F
      ScaleHeight     =   540
      ScaleWidth      =   1725
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.PictureBox picDiscos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   360
      Picture         =   "Popup.frx":4A298
      ScaleHeight     =   270
      ScaleWidth      =   405
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picRep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   135
      Picture         =   "Popup.frx":4A8C4
      ScaleHeight     =   135
      ScaleWidth      =   2085
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Menu mnuMenuPrincipal 
      Caption         =   "MenuPrincipal"
      Begin VB.Menu mnuNuevaBusqueda 
         Caption         =   "F   Nueva Busqueda ..."
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCFront 
         Caption         =   "   Cover Front"
         Begin VB.Menu mnuCambiarListaCaratula 
            Caption         =   "L   Cambiar Lista Rep/Caratula"
         End
         Begin VB.Menu mnuWallpapper 
            Caption         =   "    Colocar Caratula como Wallpaper"
         End
         Begin VB.Menu mnuMCaratula 
            Caption         =   "    Maximizar Caratula"
         End
      End
      Begin VB.Menu mnuB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowsers 
         Caption         =   "   Browsers"
         Begin VB.Menu mnuExplorar 
            Caption         =   "    Explorar ..."
         End
         Begin VB.Menu mnuExpAlbum 
            Caption         =   "    Explorar Album(s)"
         End
         Begin VB.Menu mnuZ 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTagEditor 
            Caption         =   "   Tag Editor"
         End
         Begin VB.Menu mnuLyrics 
            Caption         =   "    Lyrics"
         End
      End
      Begin VB.Menu mnuC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControles 
         Caption         =   "    Controles de Reproduccion"
         Begin VB.Menu mnuVolumen 
            Caption         =   "   Volumen"
            Begin VB.Menu mnuSubirVolumen 
               Caption         =   "+   Subir Volumen"
            End
            Begin VB.Menu mnuBajarVolumen 
               Caption         =   "-   Bajar Volumen"
            End
         End
         Begin VB.Menu mnuD 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrackAnterior 
            Caption         =   "Z   Track Anterior"
         End
         Begin VB.Menu mnuReproducir 
            Caption         =   "X   Reproducir"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "C   Pausa"
         End
         Begin VB.Menu mnuDetener 
            Caption         =   "V   Detener"
         End
         Begin VB.Menu mnuSigTrack 
            Caption         =   "B   Siguiente Track"
         End
         Begin VB.Menu mnuE 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSigAlbum 
            Caption         =   ">   Siguiente Album/Folder"
         End
         Begin VB.Menu mnuAnteriorAlbum 
            Caption         =   "<   Anterior Album/Folder"
         End
         Begin VB.Menu mnuF 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIntro 
            Caption         =   "I   Intro 10 Segundos"
         End
         Begin VB.Menu mnuRepetir 
            Caption         =   "R   Repetir Track"
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "S   Silencio"
         End
         Begin VB.Menu mnuj 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOrdenAleatorio 
            Caption         =   "  Orden Aleatorio"
            Begin VB.Menu mnuAleatorioActAlbum 
               Caption         =   "Q   Actual Album/Folder"
            End
            Begin VB.Menu mnuAleatorioTodaColec 
               Caption         =   "W   Toda la Colección"
            End
         End
         Begin VB.Menu mnug 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAtras5Seg 
            Caption         =   "A   Atras 5 Segundos"
         End
         Begin VB.Menu mnuAdelante5Seg 
            Caption         =   "D   Adelante 5 Segundos"
         End
      End
      Begin VB.Menu mnuh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "    Opciones ..."
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "    Skins"
         WindowList      =   -1  'True
         Begin VB.Menu mnuExpSkins 
            Caption         =   "<<  Explorador de Skins >>"
         End
         Begin VB.Menu mnuXXX 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSkinsAdd 
            Caption         =   " Default"
            Index           =   1
         End
      End
      Begin VB.Menu mnuWOpacity 
         Caption         =   "    Window Opacity"
         Begin VB.Menu mnuAlpha 
            Caption         =   "100%"
            Index           =   0
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "90%"
            Index           =   1
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "80%"
            Index           =   2
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "70%"
            Index           =   3
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "60%"
            Index           =   4
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "40%"
            Index           =   6
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "30%"
            Index           =   7
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "20%"
            Index           =   8
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "10%"
            Index           =   9
         End
         Begin VB.Menu mnuAlphaPer 
            Caption         =   "Personalizar..."
         End
      End
      Begin VB.Menu mnui 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "    Acerca de ..."
      End
      Begin VB.Menu mnuXX 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimizar 
         Caption         =   "    Minimizar"
      End
      Begin VB.Menu mnuCambiarMascaras 
         Caption         =   "    Cambiar Mascara"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "    Salir"
      End
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAcercaDe_Click()
 If bolAcercaShow = True Then
   frmAcerca.ZOrder 0
 Else
   frmAcerca.Show
 End If
End Sub


Private Sub mnuAdelante5Seg_Click()
 MusicMp3.Five_Seg_Forward
End Sub
Sub Menu_Aleatorio_Album()
 If TotalAlbumS = 0 Or MusicMp3.bolToyBuscando = True Then Exit Sub
 If frmPopUp.mnuAleatorioActAlbum.Checked = False Then
   MusicMp3.Images_Buttons 8, True: MusicMp3.Randomize_Order "Album"
   frmPopUp.mnuAleatorioActAlbum.Checked = True
   frmPopUp.mnuAleatorioTodaColec.Checked = False
 Else
  MusicMp3.Randomize_Click
  frmPopUp.mnuAleatorioActAlbum.Checked = False
 End If
End Sub

Private Sub mnuAleatorioActAlbum_Click()
If TotalAlbumS = 0 Or MusicMp3.bolToyBuscando = True Then Exit Sub

 If frmPopUp.mnuAleatorioActAlbum.Checked = False Then
   frmPopUp.mnuAleatorioActAlbum.Checked = True
   frmPopUp.mnuAleatorioTodaColec.Checked = False
   MusicMp3.Images_Buttons 8, True: MusicMp3.Randomize_Order "Album"

 Else
  MusicMp3.Randomize_Click
  frmPopUp.mnuAleatorioActAlbum.Checked = False
 End If
End Sub


Sub Menu_Aleatorio_Coleccion()
 If TotalAlbumS = 0 Or MusicMp3.bolToyBuscando = True Then Exit Sub
 If frmPopUp.mnuAleatorioTodaColec.Checked = False Then
   MusicMp3.Images_Buttons 8, True: MusicMp3.Randomize_Order "WholeColl"
   frmPopUp.mnuAleatorioTodaColec.Checked = True
   frmPopUp.mnuAleatorioActAlbum.Checked = False
 Else
   MusicMp3.Randomize_Click
   frmPopUp.mnuAleatorioTodaColec.Checked = False
 End If
End Sub
Private Sub mnuAleatorioTodaColec_Click()
 If TotalAlbumS = 0 Or MusicMp3.bolToyBuscando = True Then Exit Sub
 If frmPopUp.mnuAleatorioTodaColec.Checked = False Then
   frmPopUp.mnuAleatorioTodaColec.Checked = True
   frmPopUp.mnuAleatorioActAlbum.Checked = False
 Else
   MusicMp3.Randomize_Click
   frmPopUp.mnuAleatorioTodaColec.Checked = False
 End If
End Sub

Private Sub mnuAlpha_Click(Index As Integer)
On Error GoTo Hell
 Dim tAlpha
 Dim i As Integer
   tAlpha = mnuAlpha(Index).Caption
   tAlpha = left(tAlpha, Len(tAlpha) - 1)
  Call SetWindowLong(MusicMp3.hwnd, GWL_EXSTYLE, GetWindowLong(MusicMp3.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(MusicMp3.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
  mnuAlpha(Index).Checked = True
  OpcionesMusic.Alpha = tAlpha
    
    frmPopUp.mnuAlphaPer.Caption = Trim(arryLanguage(34))
    frmPopUp.mnuAlphaPer.Checked = False
  For i = 0 To 9
   If i <> Index Then mnuAlpha(i).Checked = False
  Next i
 Exit Sub
Hell:
End Sub

Private Sub mnuAlphaPer_Click()
frmOpciones.Show
frmOpciones.TabStrip1.Tabs(3).Selected = True

End Sub

Private Sub mnuAnteriorAlbum_Click()
 MusicMp3.Previous_Album
End Sub

Private Sub mnuAtras5Seg_Click()
 MusicMp3.Five_Seg_Backward
End Sub

Private Sub mnuBajarVolumen_Click()
MusicMp3.Ajust_Volume MusicMp3.imgNormal(16).Top + 2
End Sub

Private Sub mnuCambiarListaCaratula_Click()
MusicMp3.Front_Click
End Sub

Private Sub mnuCambiarMascaras_Click()
 If bolMiniMascara = True Then
  Change_Mask False
Else
  Change_Mask True
End If
End Sub



Private Sub mnuDetener_Click()
 MusicMp3.Stop_Player
End Sub

Private Sub mnuExpAlbum_Click()
If bolDirectoriosShow = False Then
 frmDirectorios.Show
Else
 frmDirectorios.ZOrder 0
End If
End Sub

Private Sub mnuExplorar_Click()
On Error Resume Next
Dim X As Long
Dim strPathExplore As String
 If TotalAlbumS = 0 Then
   strPathExplore = Path_Exe(PathExe)
 Else
   strPathExplore = MusicMp3.picAlbum(intActiveAlbum).ToolTipText
 End If
X = Shell("explorer.exe " & strPathExplore, vbMaximizedFocus)

End Sub

Private Sub mnuExpSkins_Click()
 frmOpciones.Show
frmOpciones.TabStrip1.Tabs(2).Selected = True
End Sub

Private Sub mnuIntro_Click()
  frmPopUp.mnuIntro.Checked = Not frmPopUp.mnuIntro.Checked
  MusicMp3.Intro
End Sub
Sub MostaRCaratulA()
 frmCaratula.Show
End Sub

Private Sub mnuLyrics_Click()
 If bolLyricsShow = True Then
   frmLyrics.ZOrder 0
 Else
   frmLyrics.Visible = True
 End If
End Sub

Private Sub mnuMCaratula_Click()
If bolCaratulaShow = False Then
 frmCaratula.Show
Else
 frmCaratula.ZOrder 0
End If
End Sub

Private Sub mnuMinimizar_Click()
 MusicMp3.Minimize_Me
End Sub

Private Sub mnuNuevaBusqueda_Click()
  NuevABusQuEdA
End Sub
Sub NuevABusQuEdA()
 On Error GoTo Hell
Dim strNuevaPath As String, ruta As String

 strNuevaPath = Explorador_Para_Directorios(Me.hwnd, arryLanguage(76))
If Trim(strNuevaPath) = "" Then Exit Sub
ruta = Path_Exe(PathSkin)

MusicMp3.Search_Mp3s strNuevaPath

Exit Sub
Hell:
If Dir(Path_Exe(PathSkin) & "curMain.cur") <> "" Then PicMusic.MouseIcon = LoadPicture(Path_Exe(PathSkin) & "curMain.cur")
End Sub
Private Sub mnuOpciones_Click()
 frmOpciones.Show
End Sub
Private Sub mnuPausa_Click()
 MusicMp3.Pause_Play
End Sub

Private Sub mnuRepetir_Click()
 frmPopUp.mnuRepetir.Checked = Not frmPopUp.mnuRepetir.Checked
 MusicMp3.Player_Repeat
End Sub

Private Sub mnuReproducir_Click()
 MusicMp3.Play
End Sub
Private Sub mnuSalir_Click()
 Unload MusicMp3
 End
End Sub

Private Sub mnuSigAlbum_Click()
 MusicMp3.Next_Album
End Sub

Private Sub mnuSigTrack_Click()
 MusicMp3.Next_Track
End Sub

Private Sub mnuSilencio_Click()
  frmPopUp.mnuSilencio.Checked = Not frmPopUp.mnuSilencio.Checked
  MusicMp3.Player_Mute
End Sub

Private Sub mnuSkinsAdd_Click(Index As Integer)
 On Error Resume Next
 Dim Skins As String, MiRuta As String
 Dim i As Integer
 Skins = Trim(mnuSkinsAdd(Index).Caption)
 '// si es el mismo skin salir
 If LCase(Skins) = LCase(strSkinSeleccionado) Then Exit Sub
 '// seleccionar el skin
 For i = 1 To mnuSkinsAdd.Count
  If i = Index Then
   mnuSkinsAdd(Index).Checked = True
  Else
   mnuSkinsAdd(i).Checked = False
  End If
 Next i
 
MiRuta = Path_Exe(PathExe) & "MMp3Player\Skins\"

If Skins = "Default" Then
   strSkinSeleccionado = "\" & Skins
    '// si esta la minimascara
    If bolMiniMascara = True Then
       frmMini.Visible = False
    Else
       MusicMp3.Visible = False
    End If
 
    '// cambiar el skin
    Change_Skin strSkinSeleccionado
    '// ajustar los bordes
    Form_Mini_Normal
    '// si esta la minimascara
    If bolMiniMascara = True Then
    frmMini.Visible = True
    Else
      MusicMp3.Visible = True
    End If
    Exit Sub
End If

'// chekar si existe la carpeta
If Dir(MiRuta & Skins, vbDirectory) <> "" Then
    strSkinSeleccionado = Skins
    '// si esta la minimascara
    If bolMiniMascara = True Then
       frmMini.Visible = False
    Else
       MusicMp3.Visible = False
    End If

    '// Cambiar el skin
    Change_Skin Skins
    '// ajustar los bordes
    Form_Mini_Normal
    
    If bolMiniMascara = True Then
      frmMini.Visible = True
    Else
      MusicMp3.Visible = True
    End If
End If

End Sub

Private Sub mnuSubirVolumen_Click()
MusicMp3.Ajust_Volume MusicMp3.imgNormal(16).Top - 2
End Sub

Private Sub mnuTagEditor_Click()
 If bolTagsShow = True Then
    frmTags.Load_Album_Tags
    frmTags.ZOrder 0
 Else
   frmTags.Show
 End If
End Sub

Private Sub mnuTrackAnterior_Click()
 MusicMp3.Previous_Track
End Sub

Private Sub mnuWallpapper_Click()
 If MusicMp3.ListaRep.ListCount = 0 Then Exit Sub
  mnuWallpapper.Checked = Not mnuWallpapper.Checked
   If mnuWallpapper.Checked = True Then
     ConfigurarWallpaper
   Else
     PoneRWallPapeROriginaL
   End If
  
End Sub

