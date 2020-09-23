VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MusicMp3 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "MusicMp3 1.0"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   ForeColor       =   &H00000000&
   Icon            =   "MMp3Player.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSTIcon 
      Height          =   330
      Index           =   1
      Left            =   1260
      TabIndex        =   45
      Text            =   "Text3"
      Top             =   4020
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   2
      Left            =   1035
      TabIndex        =   44
      Text            =   "Text3"
      Top             =   4050
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   43
      Text            =   "Text3"
      Top             =   4380
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   3
      Left            =   645
      TabIndex        =   42
      Text            =   "Text3"
      Top             =   4290
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   41
      Text            =   "Text3"
      Top             =   4440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ListBox LyricsRef 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1635
      Sorted          =   -1  'True
      TabIndex        =   40
      Top             =   3555
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.FileListBox fileCaratulas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   1500
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   22
      Top             =   3165
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.FileListBox FileSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Hidden          =   -1  'True
      Left            =   1455
      MousePointer    =   99  'Custom
      Pattern         =   "*.mp3"
      System          =   -1  'True
      TabIndex        =   39
      Top             =   3090
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.DirListBox DirSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   975
      TabIndex        =   38
      Top             =   3675
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picWallOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2355
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   37
      Top             =   3030
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picWallProp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2445
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   36
      Top             =   3090
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.FileListBox FileAleatorio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Hidden          =   -1  'True
      Left            =   1425
      MousePointer    =   99  'Custom
      Pattern         =   "*.mp3"
      System          =   -1  'True
      TabIndex        =   31
      Top             =   3030
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   4020
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer playTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   375
      Top             =   3735
   End
   Begin VB.PictureBox PicMusic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   17
      Top             =   0
      Width           =   5835
      Begin VB.PictureBox picSliderRep 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   240
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   139
         TabIndex        =   25
         Top             =   2250
         Width           =   2085
         Begin VB.PictureBox imgNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   15
            Left            =   15
            MousePointer    =   99  'Custom
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   27
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox picSliderVol 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   2535
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   24
         Top             =   720
         Width           =   135
         Begin VB.PictureBox imgNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   16
            Left            =   0
            MousePointer    =   99  'Custom
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   26
            Top             =   15
            Width           =   135
         End
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2025
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   35
         Top             =   3450
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.PictureBox picBotones 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2025
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   114
         TabIndex        =   34
         Top             =   3870
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.PictureBox picDiscos 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2040
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   27
         TabIndex        =   33
         Top             =   4410
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.PictureBox picFondo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   420
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1785
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   4
         ToolTipText     =   "Orden Aleatorio"
         Top             =   1350
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1365
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   3
         ToolTipText     =   "R  Repetir Track"
         Top             =   1350
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   930
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         ToolTipText     =   "S  Silencio"
         Top             =   1350
         Width           =   210
      End
      Begin VB.PictureBox picAlbum 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   315
         MousePointer    =   99  'Custom
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   0
         Top             =   555
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   5355
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   15
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   5130
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   14
         ToolTipText     =   "Mini Mascara"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   4905
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   13
         ToolTipText     =   "Minimizar"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   4335
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   12
         ToolTipText     =   ">  Siguiente Album/Folder"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   450
         MousePointer    =   99  'Custom
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   5
         ToolTipText     =   "Z  Anterior Track"
         Top             =   2460
         Width           =   342
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   795
         MousePointer    =   99  'Custom
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   6
         ToolTipText     =   "X  Reproducir"
         Top             =   2460
         Width           =   342
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1140
         MousePointer    =   99  'Custom
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   7
         ToolTipText     =   "C  Pausa"
         Top             =   2460
         Width           =   342
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1830
         MousePointer    =   99  'Custom
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   9
         ToolTipText     =   "B  Siguiente Track"
         Top             =   2460
         Width           =   342
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1485
         MousePointer    =   99  'Custom
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   23
         TabIndex        =   8
         ToolTipText     =   "V  Detener"
         Top             =   2460
         Width           =   342
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   495
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   1
         ToolTipText     =   "I  Intro 10 seg"
         Top             =   1350
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3660
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   10
         ToolTipText     =   "<  Anterior Album/Folder"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox imgNormal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   3990
         MousePointer    =   99  'Custom
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   11
         ToolTipText     =   "L  Lista Rep/Caratula"
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox picScroll 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         MousePointer    =   99  'Custom
         ScaleHeight     =   12
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   30
         Top             =   1860
         Width           =   2280
      End
      Begin VB.FileListBox ListaRep 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2505
         Hidden          =   -1  'True
         Left            =   2775
         MousePointer    =   99  'Custom
         Pattern         =   "*.mp3;*.wav;*.wma"
         System          =   -1  'True
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.Label lblBitrate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "000 kbps"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   105
         TabIndex        =   28
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label lblTiempoT 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   135
         TabIndex        =   18
         Top             =   2055
         Width           =   420
      End
      Begin VB.Label lblFreq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00.0 Hz"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   1635
         TabIndex        =   29
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label lblDuracion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   2040
         TabIndex        =   20
         Top             =   2040
         Width           =   420
      End
      Begin VB.Label lblTrackRuta 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscando..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   255
         Width           =   2370
      End
      Begin VB.Label lblTrackRep 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "                           "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   690
         TabIndex        =   21
         Top             =   2070
         Width           =   1230
      End
      Begin VB.Image ImagenCaratulA 
         Appearance      =   0  'Flat
         Height          =   2505
         Left            =   2865
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2865
      End
   End
   Begin VB.Timer Timer_Intro 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   3630
   End
   Begin VB.Timer Timer_Texto 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   75
      Top             =   3480
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   2205
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MMp3Player.frx":0D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MMp3Player.frx":10E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MMp3Player.frx":1436
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MMp3Player.frx":1788
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MMp3Player.frx":1ADA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MusicMp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|-> Proyect: Music Mp3 Player X                                                     |
'|-> Version: 1.1                                                                    |
'|-> Author: Raúl Martínez                                                           |
'|-> Email: escorpio36@hotmail.com                                                   |
'|-> Update: May 2004, Valle de Santiago, Guanajuato, México                         |
'|-> Note:                                                                           |
'|<>You do NOT have rights to redistribute this code, in whole or in part            |
'|   without my permission.  You also may not recompile the code and release         |
'|   it as another program without my permission.  If you would like to modify this  |
'|   code and distribute it in either as source code or as a compiled program please |
'|   contact me at [escorpio36@hotmail.com] before doing so.  I would appreciate     |
'|   being notified of any modifications even if you do not intend to redistribute it|
'|<>This proyect use internal player, if not run, check in proyect-references        |
'|   and active [Active Movie control type library] and run now :)                   |
'|<>If you like see "The XP Theme" then compile the code or make .EXE and see it.    |
'|<>This proyect make two file [.INI] and [.manifest] don't worry :D                 |
'|<>Sorry but the comments are in spanish :(                                         |
'|<>Any idea, comment, suggestions, doubts, bugs, more skins, etc.                   |
'|   please email me.                                                                |
'|                        is all. and thank you......                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Dim PlayerIntro As Boolean, TiempoIntro As Integer, PlayerLoop As Boolean, PlayerMute As Boolean
Dim GraphicsHeight As Integer, desAlto As Integer, desAncho As Integer, orgX As Integer, orgAncho As Integer, orgAlto As Integer
'-----------------------------------------------
Public PlayerIsPlaying As String
Public bolToyBuscando As Boolean
Public bolShowFront As Boolean
Dim Player As FilgraphManager   'Referencia el reproductor
Public PlayerPos As IMediaPosition 'Referencia para determinar la posicion
Dim PlayerAU As IBasicAudio     'Referencia para determinar el volumen
Dim X As Variant
Dim i As Integer
Public VolumeNActuaL As Integer
'------------------------------------------------
Dim bolAleatorioAlbum As Boolean
'--------------------------------------------------
Dim AleatorioCol() As String '// arreglo para aleatorio en la colleccion
Dim AleatorioRola() As Integer  '// Aleatorio en actual album
Dim stcAleatCol As Integer
Dim bolFirstAleatCol As Boolean

'-posiciones de los sliders volumen posicion ----------------------
Dim slidePos As Boolean, slideVol As Boolean
Dim DragX As Single, DragY As Single, PosVol As Integer, Pos As Integer
'El texto actual a correr
Dim rt As Long
Dim DrawingRect As RECT
Dim Izquierda As Boolean

Public strArtist As String, strAlbum As String
Public LyricsIndex As Integer

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Select_Album(Index As Integer)
   GraphicsHeight = 0
   desAncho = picDiscos.ScaleWidth / 3:  desAlto = picDiscos.ScaleHeight / 2
   orgX = (2) * (picDiscos.ScaleWidth / 3): orgAncho = picDiscos.ScaleWidth / 3
   orgAlto = picDiscos.ScaleHeight / 2
   If intActiveAlbum > 0 Then picAlbum(intActiveAlbum).PaintPicture picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
   GraphicsHeight = picDiscos.ScaleHeight / 2
   picAlbum(Index).PaintPicture picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Play_Album(Index As Integer)
Search_Caratula picAlbum(Index).ToolTipText
Select_Album Index
intActiveAlbum = Index
ListaRep.Path = picAlbum(Index).ToolTipText
 If frmPopUp.mnuAleatorioActAlbum.Checked = True And bolAleatorioAlbum = True Then
    bolAleatorioAlbum = False:    PlayerIsPlaying = "false": Randomize_Order "Album": Exit Sub
 End If
 If ListaRep.ListCount > 0 Then
    If frmPopUp.mnuAleatorioTodaColec.Checked = False Then ListaRep.Selected(0) = True: ListaRep.ListIndex = 0
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Search_Caratula(strPath As String)
 On Error Resume Next
 
 Dim bolEureka As Boolean, bolCaratula As Boolean
 fileCaratulas.Path = strPath
 
 '// Buskar caratula por todo el fileCaratulas hasta enkontrarlo
 '// search cover front
 For i = 0 To fileCaratulas.ListCount - 1
     fileCaratulas.ListIndex = i
     bolEureka = LCase(Trim(fileCaratulas.FileName)) Like "*caratula*"
      If bolEureka = False Then bolEureka = LCase(Trim(fileCaratulas.FileName)) Like "*portada*"
      If bolEureka = False Then bolEureka = LCase(Trim(fileCaratulas.FileName)) Like "*front*"
      If bolEureka = False Then bolEureka = LCase(Trim(fileCaratulas.FileName)) Like "*frt*"
      If bolEureka = True Then bolCaratula = True: Exit For
 Next i
 
'// si enkuentra alguna
'// I find one
If bolCaratula = True Then
    ImagenCaratulA.Stretch = True
    ImagenCaratulA.Picture = LoadPicture(fileCaratulas.Path & "\" & fileCaratulas.FileName)
    
    If bolCaratulaShow = True Then ' si esta cargado el frmcaratula mostrar la caratula
      frmCaratula.Picture1.Picture = LoadPicture(fileCaratulas.Path & "\" & fileCaratulas.FileName)
      frmCaratula.Mover_Form
    End If
    strRutaCaratula = fileCaratulas.Path & "\" & fileCaratulas.FileName
    If frmPopUp.mnuWallpapper.Checked = True Then ConfigurarWallpaper
    
Else
    If bolCaratulaShow = True Then 'si esta caragado y no tiene caratula mostrar la default
      frmCaratula.Picture1.Picture = frmCaratula.Picture2.Picture
      frmCaratula.Mover_Form
    End If
    strRutaCaratula = ""
    If frmPopUp.mnuWallpapper.Checked = True Then ConfigurarWallpaper
    ImagenCaratulA.Picture = LoadPicture("")
End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  BUSKEDA METODO UNO: MAS RAPIDO PERO UTILIZANDO OBJETOS DIR Y FILE :)                 |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Search_Mp3s(strPath As String)
 On Error GoTo Hell
 Dim strPathCur As String, strPathern As String
 
 '// Primero buscar en el directorio padre para buscar despues en subdirectorios
 '// first search in parent directory and after in subdirectories
 If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
 
 CopyMp3Totales = 0  '// resetear valores
 CopyTotalAlbums = 1
 
 '// set pather at files list box
 If OpcionesMusic.MP3FILE = True Then strPathern = "*.mp3"
 If OpcionesMusic.WMAFILE = True Then
    If strPathern = "" Then
      strPathern = "*.wma"
    Else
      strPathern = strPathern & ";*.wma"
    End If
 End If
 
 If OpcionesMusic.WAVFILE = True Then
    If strPathern = "" Then
      strPathern = "*.wav"
    Else
       strPathern = strPathern & ";*.wav"
    End If
 End If
 
 If strPathern = "" Then strPathern = "*.mp3"
   
  FileSearch.Pattern = strPathern
  
  FileSearch.Path = strPath
 If FileSearch.ListCount > 0 Then
   picAlbum(1).ToolTipText = strPath  '// poner el primer album
   CopyMp3Totales = CopyMp3Totales + FileSearch.ListCount
   CopyTotalAlbums = 2
 End If
 
 bolToyBuscando = True
 lblTrackRuta.Caption = arryLanguage(68)
 
 '// poner cursor de busqueda si hay del skin
 strPathCur = Path_Exe(PathSkin)
 If Dir(strPathCur & "curFind.cur") <> "" Then
   PicMusic.MouseIcon = LoadPicture(strPathCur & "curFind.cur")
   If bolSplashScreen = True Then
     frmSplash.MouseIcon = LoadPicture(strPathCur & "curFind.cur")
     frmSplash.lblSplash(0).MouseIcon = LoadPicture(strPathCur & "curFind.cur")
     frmSplash.lblSplash(1).MouseIcon = LoadPicture(strPathCur & "curFind.cur")
     frmSplash.lblSplash(2).MouseIcon = LoadPicture(strPathCur & "curFind.cur")
   End If
 End If
 
 If bolSplashScreen = True Then
   frmSplash.lblSplash(1).Caption = arryLanguage(75)
 End If
 
 '// si esta la minimacara mostrar en el picture el de buscar
 If bolMiniMascara = True Then Scroll_Text arryLanguage(68), frmMini.picScroll
 
 '---------------------------------------------------------------------------------
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Click
    frmPopUp.mnuAleatorioTodaColec.Checked = False
  End If
  If frmPopUp.mnuAleatorioActAlbum.Checked = True Then
    Randomize_Click
    frmPopUp.mnuAleatorioActAlbum.Checked = False
  End If
'---------------------------------------------------------------------------------
 
 '// Empezar ha buskar
  Call Start_Search(strPath)
 bolToyBuscando = False
 '// Akomodar los albums si enkuentra
  Call Process_Albums(True)
  
  ListaRep.Pattern = strPathern
  FileAleatorio.Pattern = strPathern
  
 '// variable para determinar la path en el form directorios
  If CopyMp3Totales > 0 Then
    strTraySearch = strPath
    If bolDirectoriosShow = True Then frmDirectorios.Load_Albums
  End If
  
 If Dir(strPathCur & "curMain.cur") <> "" Then PicMusic.MouseIcon = LoadPicture(strPathCur & "curMain.cur")
Exit Sub
Hell:
 If Dir(strPathCur & "curMain.cur") <> "" Then PicMusic.MouseIcon = LoadPicture(strPathCur & "curMain.cur")
End Sub

'// metod for search is very faster

Sub Start_Search(strPath As String)
 On Error Resume Next  '// manejador de error por si permisos de acceso a los directorios
 
 DoEvents '// para que deje trabajar el Windows
 Dim subdirs As Integer, k As Integer, intFolder As Integer
 ReDim subdirs_name(0 To 10) As String  '// arreglo para directorios
 subdirs = 0

If bolToyBuscando = False Then Exit Sub  '// para cancelar si keremos
  
 '// Poner el Dir en la direccion para iniciar busqueda y en subdirectorios
 DirSearch.Path = strPath
For intFolder = 0 To DirSearch.ListCount - 1  '// buskar en los elementos del dir
      '// Komo todos son directorios almacenarlos en el arreglo para despues buskar
      '// en subdirectorios
      subdirs_name(subdirs) = DirSearch.List(intFolder)
      subdirs = subdirs + 1
      '// si se pasan los directorios del maximo del arreglo
      '// aumentar otros 10
      If subdirs Mod 10 = 0 Then ReDim Preserve subdirs_name(0 To subdirs + 10)
      
      '// Verifikar si hay mp3s con el file
      FileSearch.Path = DirSearch.List(intFolder)
      If FileSearch.ListCount > 0 Then
        '// Ir kontando todos los mp3's
        CopyMp3Totales = CopyMp3Totales + FileSearch.ListCount
        '// Verifikar si no se han cargado ahun los picAlbums para que no marke error
        '// sino pus kargarlo
          If CopyTotalAlbums > picAlbum.Count Then Load picAlbum(CopyTotalAlbums)
          '// Almecenar la ruta en el tooltiptext para reproducirlos despues
            picAlbum(CopyTotalAlbums).ToolTipText = DirSearch.List(intFolder)
          
           If bolSplashScreen = True Then frmSplash.lblSplash(2).Caption = "Albums: [ " & CopyTotalAlbums & " ]   Files: [ " & CopyMp3Totales & " ]"
           'lblTrackRuta.Caption = arryLanguage(68)
           
          '// Ir contando los albums totales
           CopyTotalAlbums = CopyTotalAlbums + 1
          End If
Next intFolder

'//-----------Buscamos en subdirectorios ----------------------------------------
'// como es una procedimento que se llama a si mismo las variables anteriores
'// se siguen conserbando hasta que termine
For k = 0 To subdirs - 1
 '// mostramos los directorios de busqueda
 
 If bolSplashScreen = True Then frmSplash.lblSplash(0).Caption = subdirs_name(k)
 If bolMiniMascara = True Then
   frmMini.picScroll.ToolTipText = subdirs_name(k)
 Else
   lblTrackRuta.ToolTipText = subdirs_name(k)
 End If
 Start_Search subdirs_name(k)
Next

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Process_Albums(Normal As Boolean)
 On Error GoTo Hell
 lblTrackRuta.ToolTipText = ""
 frmMini.picScroll.ToolTipText = ""
 
 '// si no se encuentran mp3's
 If CopyMp3Totales = 0 Then
    If PlayerIsPlaying = "true" Or PlayerIsPlaying = "pause" Then
      Load_File_Tags
    Else
      lblTrackRuta.Caption = arryLanguage(69)
    End If
   Exit Sub
 End If
 
 CopyTotalAlbums = CopyTotalAlbums - 1

'// Okultar los albums anteriores
For i = TotalAlbumS To CopyTotalAlbums Step -1
  picAlbum(i).Visible = False
Next i

 PicMusic.Refresh
 TotalAlbumS = CopyTotalAlbums
 MP3totales = CopyMp3Totales
  

  GraphicsHeight = 0

'// cargar los albums
For i = 1 To TotalAlbumS
  'si es el primer album se queda en la misma posicion
 If i <= 48 Then  ' comparar los albums que se pueden ver maximo 48
  If i <> 1 And i < 13 Then '// primera linea de 12 elementos
    picAlbum(i).Top = picAlbum(1).Top
    picAlbum(i).left = picAlbum(i - 1).left + 13
  End If
  
  If i > 12 And i < 25 Then '// Segunda linea de 12 elementos
   picAlbum(i).Top = picAlbum(1).Top + 13
   picAlbum(i).left = picAlbum(i - 12).left
  End If
  
  If i > 24 And i < 37 Then '// Tercera linea de 12 elementos
   picAlbum(i).Top = picAlbum(13).Top + 13
   picAlbum(i).left = picAlbum(i - 24).left
  End If
  
  If i > 36 And i < 49 Then '// Cuarta linea de 12 elementos
   picAlbum(i).Top = picAlbum(25).Top + 13
   picAlbum(i).left = picAlbum(i - 36).left
  End If
  
 '// Poner la imagen ahora si
  desAncho = picDiscos.ScaleWidth / 3
  desAlto = picDiscos.ScaleHeight / 2
  orgX = (2) * (picDiscos.ScaleWidth / 3)
  orgAncho = picDiscos.ScaleWidth / 3
  orgAlto = picDiscos.ScaleHeight / 2
  
  picAlbum(i).PaintPicture picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
  picAlbum(i).Picture = picAlbum(i).Image
  picAlbum(i).Visible = True
 End If
Next i

If Normal = True Then
  If TotalAlbumS > 0 Then ListaRep.ListIndex = -1: Play_Album 1
End If
Exit Sub
Hell:
MsgBox err.Description
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Play()
 On Error Resume Next
  If ListaRep.ListCount = 0 Or TotalAlbumS = 0 Then Exit Sub
  
  If PlayerIntro = True Then Timer_Intro.Enabled = True: TiempoIntro = 0
  If PlayerIsPlaying = "pause" Then Pause_Play: Exit Sub
  'If PlayerIsPlaying = "true" Then Five_Seg_Forward: Exit Sub
  
  '// check if player int frmtags is playing
  If PlayerState = "true" Or PlayerState = "pause" Then frmTags.Stop_Player
  
  playTimer.Enabled = False
   Start_Play
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Image_State_Rep()
   GraphicsHeight = 0
   desAncho = picBotones.ScaleWidth / 5:   desAlto = picBotones.ScaleHeight / 2:   orgAncho = picBotones.ScaleWidth / 5
   orgAlto = picBotones.ScaleHeight / 2:   orgX = (3) * (picBotones.ScaleWidth / 5)
   orgX = (1) * (picBotones.ScaleWidth / 5)
   imgNormal(1).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
   orgX = (3) * (picBotones.ScaleWidth / 5)
   imgNormal(3).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
   orgX = (2) * (picBotones.ScaleWidth / 5)
   imgNormal(2).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
   GraphicsHeight = picBotones.ScaleHeight / 2
   frmMini.Images_Buttons 1, False
   frmMini.Images_Buttons 2, False
   frmMini.Images_Buttons 3, False
 Select Case PlayerIsPlaying
   Case "true"  'Reproduciendo
     orgX = (1) * (picBotones.ScaleWidth / 5)
     imgNormal(1).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
     frmMini.Images_Buttons 1, True
   Case "false" 'detenido
     orgX = (3) * (picBotones.ScaleWidth / 5)
     imgNormal(3).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
     frmMini.Images_Buttons 3, True
   Case "pause" 'Pausado
     orgX = (2) * (picBotones.ScaleWidth / 5)
     imgNormal(2).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
     frmMini.Images_Buttons 2, True
 End Select
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Images_Buttons(Button As Integer, Active As Boolean)
    desAncho = picMenu.ScaleWidth / 10
    desAlto = picMenu.ScaleHeight / 2
    orgAncho = picMenu.ScaleWidth / 10
    orgAlto = picMenu.ScaleHeight / 2
  
  orgX = (Button - 5) * (picMenu.ScaleWidth / 10)
 If Active = True Then
    GraphicsHeight = picMenu.ScaleHeight / 2
    imgNormal(Button).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
 Else
    GraphicsHeight = 0
    imgNormal(Button).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
 End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Start_Play()
On Error GoTo error
 Set Player = New FilgraphManager  '// ajustar los reproductores
 Set PlayerPos = Player
 Set PlayerAU = Player
     
   
   Player.RenderFile sFileMainPlaying 'cargar archivo
   Player.Run                         'executar player
   
    If PlayerMute = True Then
     PlayerAU.Volume = -10000
    Else
    PlayerAU.Volume = VolumeNActuaL
    End If
    
   
   Load_File_Tags
   
   Start_Lyrics
   
   playTimer.Enabled = True
   
   '//Calcular la duracion de la RoLa
   lblDuracion.Caption = Convert_Time(PlayerPos.Duration)
   '// cargar tags
   
   PlayerIsPlaying = "true"
   Image_State_Rep
Exit Sub
error:
   lblTrackRuta.Caption = arryLanguage(73)
   Stop_Player
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Stop_Player()
 On Error Resume Next
 If TotalAlbumS = 0 Then Exit Sub
  imgNormal(15).left = 0
  imgNormal(15).Picture = imgNormal(15).Image
  picSliderRep.Picture = picSliderRep.Image
  
  playTimer.Enabled = False
  
  If PlayerIntro = True Then Timer_Intro.Enabled = False
  
  lblTiempoT.Caption = "00:00"
  frmMini.lblTiempoT = "00:00"
  PlayerIsPlaying = "false"
  
  Image_State_Rep
 
  Player.Stop
 
   Set Player = Nothing
   Set PlayerPos = Nothing
   Set PlayerAU = Nothing

 If HasLyrics3Tag = True And bolLyricsShow = True Then frmLyrics.Reset_Values

End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Pause_Play()
 Dim CurState As Long
 
 If ListaRep.ListCount = 0 Or TotalAlbumS = 0 Then Exit Sub
 
  If PlayerIsPlaying = "false" Then Exit Sub
     Player.GetState X, CurState
 '------'Esta Reproduciendo, pausar-------------------------------------------
     If CurState = 2 Then
       PlayerIsPlaying = "pause"
       Image_State_Rep
       Player.Pause
       If PlayerIntro = True Then Timer_Intro.Enabled = False
     Else
'------'Si esta pausado, reproducir---------------------------------------------
       If PlayerMute = True Then PlayerAU.Volume = -10000
       If PlayerIntro = True Then Timer_Intro.Enabled = True
       PlayerIsPlaying = "true"
       Player.Run
       Image_State_Rep
     End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Function Convert_Time(ByVal LSec As Long) As String
 Dim HH As Long, MM As Long, SS As Long
 Dim tmp As String
 
 HH = LSec \ 3600  '// calkular horas
 MM = LSec \ 60 Mod 60 '// Calkular minutos
 SS = LSec Mod 60  '// calkular segundos
 
 If HH > 0 Then tmp = Format$(HH, "00:")
 Convert_Time = tmp & Format$(MM, "00:") & Format$(SS, "00")
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub ShowTime(ByVal iDur As Integer)
  Dim actual As Variant
  actual = iDur
    imgNormal(15).left = 1 + Int(((actual * 1000) / (PlayerPos.Duration * 1000)) * 128)
    If imgNormal(15).left > 129 Then imgNormal(15).left = 129
     imgNormal(15).Picture = imgNormal(15).Image
     picSliderRep.Picture = picSliderRep.Image
    DoEvents
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Forward()
 Dim CurPos
  If ListaRep.ListCount = 0 Or PlayerIsPlaying <> "true" Then Exit Sub
  
  CurPos = PlayerPos.CurrentPosition
  CurPos = CurPos + 5
  If CurPos > PlayerPos.Duration Then CurPos = PlayerPos.Duration
  PlayerPos.CurrentPosition = CurPos
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Backward()
 Dim CurPos
  If ListaRep.ListCount = 0 Or PlayerIsPlaying <> "true" Then Exit Sub
  CurPos = PlayerPos.CurrentPosition
  CurPos = CurPos - 5
  If CurPos < 0 Then CurPos = 0
  PlayerPos.CurrentPosition = CurPos
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Next_Album()
 If TotalAlbumS = 0 Or picAlbum.Count = 1 Then Exit Sub
 
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Click
    frmPopUp.mnuAleatorioTodaColec.Checked = False
  End If
  
  If intActiveAlbum >= TotalAlbumS Then
    Play_Album 1
     Exit Sub
  End If
  
  Play_Album intActiveAlbum + 1
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Previous_Album()
 If TotalAlbumS = 0 Or picAlbum.Count = 1 Then Exit Sub
 
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Click
    frmPopUp.mnuAleatorioTodaColec.Checked = False
  End If
  
  If intActiveAlbum = 1 Then
   Play_Album TotalAlbumS
   Exit Sub
  End If
  
   Play_Album intActiveAlbum - 1
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Next_Track()
 Dim a As Integer
  If ListaRep.ListCount = 0 Then Exit Sub
  
  If frmPopUp.mnuAleatorioActAlbum.Checked = True Then
    Randomize_Order "Album"
    Exit Sub
  End If
  
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Order "Coleccion"
    Exit Sub
  End If
  
  a = ListaRep.ListIndex
  a = a + 1
 If a < ListaRep.ListCount Then
  ListaRep.Selected(a) = True
 Else
  Next_Album
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Previous_Track()
 Dim a As Integer
  If ListaRep.ListCount = 0 Then Exit Sub
  
  If frmPopUp.mnuAleatorioActAlbum.Checked = True Then
    Randomize_Order "Album"
    Exit Sub
  End If
  
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Order "Coleccion"
    Exit Sub
  End If
 
 a = ListaRep.ListIndex
 If a = 0 Then Previous_Album
 If a <> 0 Then a = a - 1
 If a >= 0 Or a < ListaRep.ListCount Then ListaRep.Selected(a) = True
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Front_Click()
 On Error Resume Next
  ListaRep.Visible = Not ListaRep.Visible
  ImagenCaratulA.Visible = Not ImagenCaratulA.Visible
     
     desAncho = picMenu.ScaleWidth / 10
     desAlto = picMenu.ScaleHeight / 2
     orgX = (5) * (picMenu.ScaleWidth / 10)
     orgAncho = picMenu.ScaleWidth / 10
     orgAlto = picMenu.ScaleHeight / 2
     If ListaRep.Visible = False Then
      GraphicsHeight = picMenu.ScaleHeight / 2
      imgNormal(10).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
      bolShowFront = True
     Else
      GraphicsHeight = 0
      imgNormal(10).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
      bolShowFront = False
     End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 90 Then Previous_Track 'Z
  If KeyCode = 88 Then Play 'X
  If KeyCode = 67 Then Pause_Play 'C
  If KeyCode = 86 Then Stop_Player 'V
  If KeyCode = 66 Then Next_Track 'B
  If Shift = vbShiftMask And KeyCode = 226 Then Next_Album: Exit Sub ' > Siguiente Album
  If KeyCode = 226 Then Previous_Album ' < Anterior Album
  If KeyCode = 76 Then Front_Click 'L Cambiar caratula
  If KeyCode = 73 Then Intro 'I Intro 10 seg
  If KeyCode = 82 Then Player_Repeat 'R Repetir
  If KeyCode = 83 Then Player_Mute 'S Silencio
  If KeyCode = 81 Then 'Q Orden aleatorio Album
    frmPopUp.Menu_Aleatorio_Album
  End If
  If KeyCode = 87 Then 'W Orden aleatorio coleccion
    frmPopUp.Menu_Aleatorio_Coleccion
  End If
  If KeyCode = 77 Then frmPopUp.MostaRCaratulA  'M Mostrar caratula
  If KeyCode = 70 Then frmPopUp.NuevABusQuEdA 'F Nueva busqueda
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Randomize_Click()
 If TotalAlbumS = 0 Or bolToyBuscando = True Then Exit Sub
 
 If frmPopUp.mnuAleatorioActAlbum.Checked = False And frmPopUp.mnuAleatorioTodaColec.Checked = False Then
    If TotalAlbumS = 1 Then
       frmPopUp.mnuAleatorioTodaColec.Enabled = False
    Else
       frmPopUp.mnuAleatorioTodaColec.Enabled = True
    End If
    
  PopupMenu frmPopUp.mnuOrdenAleatorio
  
   If frmPopUp.mnuAleatorioActAlbum.Checked = True Then
     Images_Buttons 8, True
     Randomize_Order "Album"
   End If
   
   If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
     Images_Buttons 8, True
     Randomize_Order "WholeColl"
   End If
   
  Exit Sub
 End If
  
 If frmPopUp.mnuAleatorioActAlbum.Checked = True Or frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    bolAleatorioAlbum = False
    bolFirstAleatCol = False
    stcAleatCol = 0
    Images_Buttons 8, False
    
    frmPopUp.mnuAleatorioActAlbum.Checked = False
    frmPopUp.mnuAleatorioTodaColec.Checked = False
 End If
 
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ImagenCaratulA_DblClick()
 If bolCaratulaShow = True Then
   frmCaratula.ZOrder 0
 Else
   frmCaratula.Show
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ImagenCaratulA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Intro()
  If PlayerIntro = False Then
    'poner intro activado
    Images_Buttons 5, True
    '-------------------------------------------
    PlayerIntro = True
    TiempoIntro = 0
    Timer_Intro.Enabled = True
    frmPopUp.mnuIntro.Checked = True
  Else
  'poner intro desactivado
    Images_Buttons 5, False
   '----------------------------------------------
    PlayerIntro = False
    Timer_Intro.Enabled = False
    frmPopUp.mnuIntro.Checked = False
  End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Randomize_Order(MoDo As String)
 On Error Resume Next
  Dim aleatorio() As String
  Dim AleatAlbum As Integer
  Dim AleatRola As Integer
  Dim i As Integer, j As Integer
  Static stcRolaAleat As Integer

If MoDo = "Album" Then
'------- ALEATORIO DE ALBUMS -----------------------------------------------------------
  '// si es la perimera vez
  If bolAleatorioAlbum = False Then
     '// redimencionar arreglo con el numero de elementos de la lista de reprod.
     ReDim AleatorioRola(ListaRep.ListCount - 1)
     
     Randomize
          
     If PlayerIsPlaying = "false" Then
       AleatorioRola(0) = Int(ListaRep.ListCount * Rnd)
     Else
       AleatorioRola(0) = ListaRep.ListIndex
        If AleatorioRola(0) = -1 Then AleatorioRola(0) = Int(ListaRep.ListCount * Rnd)
     End If
     
   '// numero de aleatorios a kalkular
   For j = 1 To ListaRep.ListCount - 1
     DoEvents
      '// skar numero aleatorio
      Randomize
      AleatorioRola(j) = Int(ListaRep.ListCount * Rnd)
         '// compararlo con los aleatorios anteriores
         '// deskontando el anterior
         For i = 0 To j - 1
            If AleatorioRola(j) = AleatorioRola(i) Then
              j = j - 1
               If j < 1 Then j = 1
              Exit For
            End If
         Next i
    Next j
     bolAleatorioAlbum = True
     '// variable para apuntar al numero de arreglo
     stcRolaAleat = 0
     If PlayerIsPlaying = "false" Then
      ListaRep.ListIndex = -1
      ListaRep.ListIndex = AleatorioRola(stcRolaAleat)
      'ListaRep.Selected(AleatorioRola(stcRolaAleat)) = True
     End If
     
  '// si no es la primera vez
  Else
    stcRolaAleat = stcRolaAleat + 1
    If stcRolaAleat < ListaRep.ListCount Then
      ListaRep.ListIndex = AleatorioRola(stcRolaAleat)
      ListaRep.Selected(AleatorioRola(stcRolaAleat)) = True
    Else
      If TotalAlbumS = 1 Then Stop_Player: Randomize_Click: Exit Sub
      Next_Album
    End If
  End If

'// arden aleatorio entoda la coleccion
'--------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------
  '// si es la primera vez
  If stcAleatCol = 0 Then
     ReDim AleatorioCol(0)
     
     If bolFirstAleatCol = False And PlayerIsPlaying = "true" Or PlayerIsPlaying = "pause" Then
       '// kalkular aleatorio NUMERO_DE_ALBUM  ,  TRACK_NUMBER
       AleatorioCol(stcAleatCol) = intActiveAlbum & "," & ListaRep.ListIndex
       stcAleatCol = stcAleatCol + 1
       bolFirstAleatCol = True
       Exit Sub
     End If
     
      Randomize '// albums
       AleatAlbum = Int(Rnd * (TotalAlbumS) + 1)
       AleatorioCol(stcAleatCol) = AleatAlbum
       
      Randomize '// rolas albums
       FileAleatorio.Path = picAlbum(AleatAlbum).ToolTipText
       AleatRola = Int(FileAleatorio.ListCount * Rnd)
       
       AleatorioCol(stcAleatCol) = AleatAlbum & "," & AleatRola
       
       Play_Album AleatAlbum
       ListaRep.ListIndex = AleatRola
       
  Else '// si no es la primera vez
    '// redim al nuevo numero aleatorio
    ReDim Preserve AleatorioCol(stcAleatCol)
AleatorioNuevo:
     Randomize 'albums
      AleatAlbum = Int(Rnd * (TotalAlbumS) + 1)
      AleatorioCol(stcAleatCol) = AleatAlbum
      
    Randomize 'rolas albums
      FileAleatorio.Path = picAlbum(AleatAlbum).ToolTipText
      AleatRola = Int(FileAleatorio.ListCount * Rnd)
      
      '// almacenar aleatorio en arreglo
      AleatorioCol(stcAleatCol) = AleatAlbum & "," & AleatRola
    
   For j = 0 To UBound(AleatorioCol) - 1
     aleatorio() = Split(AleatorioCol(j), ",", , vbTextCompare)
      'compara si son iguales a los anteriores
     If aleatorio(0) = AleatAlbum And aleatorio(1) = AleatRola Then
      GoTo AleatorioNuevo
     End If
   Next j
  
   '// si ya se hicieron todos los mp3 comenzar de new
   If stcAleatCol = (MP3totales - 1) Then
    stcAleatCol = 0
   End If
   
   Play_Album AleatAlbum
   ListaRep.ListIndex = AleatRola
   
 End If
   '// aumentar a la siguiente aleatorio
    stcAleatCol = stcAleatCol + 1
    
End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub imgNormal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Index = 15 Then
  SliderReproduccioN_Move X, Y
 Else
  SliderVolumen_Move X, Y
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub imgNormal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then Exit Sub
   GraphicsHeight = 0
    
   '// botones de intro - mute - repeat - aleatorio
   If Index >= 5 And Index <= 7 Then Exit Sub
   
   '// botones de reproduccion
   If Index < 5 Then
      If Index > 0 And Index < 4 Then GoTo etiqueta
        desAncho = picBotones.ScaleWidth / 5
        desAlto = picBotones.ScaleHeight / 2
        orgX = (Index) * (picBotones.ScaleWidth / 5)
        orgAncho = picBotones.ScaleWidth / 5
        orgAlto = picBotones.ScaleHeight / 2
        imgNormal(Index).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
        
   ElseIf Index < 15 Then '// menus
         If Index = 10 Then GoTo etiqueta
            GraphicsHeight = 0
            desAncho = picMenu.ScaleWidth / 10
            desAlto = picMenu.ScaleHeight / 2
            orgX = (Index - 5) * (picMenu.ScaleWidth / 10)
            orgAncho = picMenu.ScaleWidth / 10
            orgAlto = picMenu.ScaleHeight / 2
            imgNormal(Index).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
            
       Else '// sliders de rep y vol
            desAncho = picDiscos.ScaleWidth / 3
            desAlto = picDiscos.ScaleHeight / 2
            orgX = (Index - 15) * (picDiscos.ScaleWidth / 3)
            orgAncho = picDiscos.ScaleWidth / 3
            orgAlto = picDiscos.ScaleHeight / 2
            imgNormal(Index).PaintPicture picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
            imgNormal(Index).Picture = imgNormal(Index).Image
            
               If Index = 15 Then
                 SliderReproduccioN_Up X, Y
               Else
                 slideVol = False
                  Scroll_Text ScrollText, picScroll
               End If
               
       End If
       
etiqueta:
 If Index = 0 Then Previous_Track
 If Index = 4 Then Next_Track
 If Index = 9 Then Previous_Album
 If Index = 10 Then Front_Click
 If Index = 11 Then Next_Album
 If Index = 12 Then Minimize_Me
 If Index = 13 Then Change_Mask True
 If Index = 14 Then Unload Me
 
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Minimize_Me()
   If bolCaratulaShow = True Then frmCaratula.Hide
   If bolDirectoriosShow = True Then frmDirectorios.Hide
   If bolOpcionesShow = True Then frmOpciones.Hide
   If bolAcercaShow = True Then frmAcerca.Hide
   If bolTagsShow = True Then frmTags.Hide
   If bolLyricsShow = True Then frmLyrics.Hide
   
   If bolMiniMascara = True Then
      frmMini.Hide
   Else
      Me.Hide
   End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Player_Repeat()
 If PlayerLoop = False Then
   '---Activar loop -----------------------------
    Images_Buttons 7, True
    PlayerLoop = True
    frmPopUp.mnuRepetir.Checked = True
  Else
   '--- Descativar el loop ---------------------------
    Images_Buttons 7, False
    PlayerLoop = False
    frmPopUp.mnuRepetir.Checked = False
  End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Player_Mute()
 On Error Resume Next
  If PlayerMute = False Then
    '--activar silencio --------------------
    Images_Buttons 6, True
    PlayerMute = True
    If PlayerAU Is Nothing Then
      VolumeNActuaL = 0: frmPopUp.mnuSilencio.Checked = True
    Else
      VolumeNActuaL = PlayerAU.Volume: PlayerAU.Volume = -10000
      frmPopUp.mnuSilencio.Checked = True
    End If
  Else
    'Desactivar el mute --------------------------------
    Images_Buttons 6, False
    PlayerMute = False
    If PlayerAU Is Nothing Then
      VolumeNActuaL = -10000: frmPopUp.mnuSilencio.Checked = False
    Else
      PlayerAU.Volume = VolumeNActuaL
      frmPopUp.mnuSilencio.Checked = False
    End If
  End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblBitrate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblDuracion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblFreq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblTiempoT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblTrackRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button = vbLeftButton Then FormDrag Me
 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub lblTrackRuta_DblClick()
 If bolToyBuscando = True Then
   bolToyBuscando = False
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ListaRep_Click()
 On Error Resume Next
 If ListaRep.ListCount = 0 Or ListaRep.List(ListaRep.ListIndex) = "" Then Exit Sub
 
 sFileMainPlaying = Trim(ListaRep.Path & "\" & ListaRep.FileName)

 lblTrackRep.Caption = "Track " & ListaRep.ListIndex + 1 & " of " & ListaRep.ListCount
 
 PlayerIsPlaying = "true"
 
  Play
 
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub imgNormal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  If Index = 8 Then Randomize_Click: Exit Sub
  If Index = 5 Then Intro: Exit Sub
  If Index = 6 Then Player_Mute: Exit Sub
  If Index = 7 Then Player_Repeat: Exit Sub
  If Index = 1 Then Play
  If Index = 2 Then Pause_Play
  If Index = 3 Then Stop_Player
  If Index = 10 Then Exit Sub

   If Index < 5 Then ' botones de reproduccion atras adela
      If Index > 0 And Index < 4 Then Exit Sub
        GraphicsHeight = picBotones.ScaleHeight / 2
        desAncho = picBotones.ScaleWidth / 5
        desAlto = picBotones.ScaleHeight / 2
        orgX = (Index) * (picBotones.ScaleWidth / 5)
        orgAncho = picBotones.ScaleWidth / 5
        orgAlto = picBotones.ScaleHeight / 2
        imgNormal(Index).PaintPicture picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
   ElseIf Index < 15 Then
            GraphicsHeight = picMenu.ScaleHeight / 2
            desAncho = picMenu.ScaleWidth / 10
            desAlto = picMenu.ScaleHeight / 2
            orgX = (Index - 5) * (picMenu.ScaleWidth / 10)
            orgAncho = picMenu.ScaleWidth / 10
            orgAlto = picMenu.ScaleHeight / 2
            imgNormal(Index).PaintPicture picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
        Else
            GraphicsHeight = picDiscos.ScaleHeight / 2
            desAncho = picDiscos.ScaleWidth / 3
            desAlto = picDiscos.ScaleHeight / 2
            orgX = (Index - 15) * (picDiscos.ScaleWidth / 3)
            orgAncho = picDiscos.ScaleWidth / 3
            orgAlto = picDiscos.ScaleHeight / 2
            imgNormal(Index).PaintPicture picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
            imgNormal(Index).Picture = imgNormal(Index).Image
            picSliderVol.Picture = picSliderVol.Image
        If Index = 15 Then
         SliderReproduccioN_Down X, Y
        Else
         SliderVolumen_Down X, Y
        End If
   End If
 End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 45 Then Ajust_Volume imgNormal(16).Top + 2  '-
 If KeyAscii = 43 Then Ajust_Volume imgNormal(16).Top - 2  '+
 If KeyAscii = 65 Or KeyAscii = 97 Then Five_Seg_Backward 'A Atras 5 seg
 If KeyAscii = 68 Or KeyAscii = 100 Then Five_Seg_Forward 'D Adelante 5 seg
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Ajust_Volume(Posicion As Integer)
 On Error Resume Next
 
   Dim intPorcentaje As Integer
   
      imgNormal(16).Top = Posicion
      
     If imgNormal(16).Top < 0 Then imgNormal(16).Top = 0
     If imgNormal(16).Top > 89 Then imgNormal(16).Top = 89
     
       imgNormal(16).Picture = imgNormal(16).Image
       picSliderVol.Picture = picSliderVol.Image
       
       intPorcentaje = CInt((imgNormal(16).Top * 100) / 89)
       
       frmPopUp.mnuVolumen.Caption = arryLanguage(12) & " [ " & 100 - intPorcentaje & " % ]"
       
        If slideVol = True Then
          picScroll.Cls
          picScroll.Print "          [ " & Trim(arryLanguage(12)) & " " & 100 - intPorcentaje & " % ]"
          picScroll.Refresh
        End If
      
    If PlayerAU Is Nothing Then
      VolumeNActuaL = -(intPorcentaje * 50)
    Else
      PlayerAU.Volume = -(intPorcentaje * 50)
      VolumeNActuaL = PlayerAU.Volume
    End If
    
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Load()
 On Error Resume Next
  LyricsIndex = 1
  PlayerIsPlaying = "false"
  '/* inizializar detectar la rueda de la rata :)
  Call Hook
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
   
   Stop_Player
   
   Save_Settings_INI
   
   If App.PrevInstance = False Then
     If frmPopUp.mnuWallpapper.Checked = True Then PoneRWallPapeROriginaL
   End If
     
     'Borrar el archivo de wallpaper creado si se hizo
   If Dir(DirectoriOWindowS & "MusicMp3.bmp") <> "" Then
     Kill DirectoriOWindowS & "MusicMp3.bmp"
   End If
        
   '/* eliminar monitorizar mause
   Call Unhook
   Set Player = Nothing
   Set PlayerPos = Nothing
   Set PlayerAU = Nothing
   Set MusicMp3 = Nothing
  End
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picAlbum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then Exit Sub
 
 If intActiveAlbum = Index Then Exit Sub  ' no reproducir de nuevo el disco activo
  
  If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
    Randomize_Click
    frmPopUp.mnuAleatorioTodaColec.Checked = False
  End If
 
 Play_Album Index

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub PicMusic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
   If Button = vbLeftButton Then FormDrag Me
   If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picScroll_Click()
  If (TextWidth(ScrollText) / 15) <= picScroll.ScaleWidth Then Exit Sub
  Timer_Texto.Enabled = Not Timer_Texto.Enabled
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub playTimer_Timer()
On Error Resume Next
 '//si esta reproduciendo
  If PlayerIsPlaying = "true" Then
    '// si se esta arrastrando el slider rep
    If slidePos = True Then Exit Sub
    
    '// si se akabo la rola
    If PlayerPos.CurrentPosition >= PlayerPos.Duration Then
      Stop_Player
      PlayerIsPlaying = "false"
     
     '// si esta seleccionada el check para el loop
      If PlayerLoop = True Then Play: Exit Sub
     
      If frmPopUp.mnuAleatorioTodaColec.Checked = True Then
        Randomize_Order ("TodosLosAlbums")
        Exit Sub
      End If
      
      If frmPopUp.mnuAleatorioActAlbum.Checked = True Then
        Randomize_Order ("Album")
        Exit Sub
      End If

      If ListaRep.ListIndex < ListaRep.ListCount - 1 Then
           ListaRep.ListIndex = ListaRep.ListIndex + 1
      Else
           Next_Album
      End If
     Exit Sub
   End If
  
  '// Si esta la minimaskara
   If bolMiniMascara = True Then
     If frmMini.bolTimeAct = True Then
       frmMini.lblTiempoT.Caption = Convert_Time(PlayerPos.CurrentPosition)
     Else
       frmMini.lblTiempoT.Caption = "-" & Convert_Time(PlayerPos.Duration - PlayerPos.CurrentPosition)
     End If
   Else
     lblTiempoT.Caption = Convert_Time(PlayerPos.CurrentPosition)
   End If
   
      ShowTime PlayerPos.CurrentPosition
   
  If HasLyrics3Tag = True And bolLyricsShow = True Then
    Update_Lyrics
  End If
 End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Load_File_Tags()
On Error Resume Next

Dim strFolder As String
Dim aryTitulo() As String
Dim accMP3Info As MPEGInfo

 
 
  accMP3Info = Load_MPEGInfo(sFileMainPlaying)
  
  lblBitrate.Caption = accMP3Info.BITRATE
  lblFreq.Caption = accMP3Info.FREQ
   
  If accMP3Info.Artist <> "" And accMP3Info.Album <> "" Then
     strArtist = accMP3Info.Artist
     strAlbum = accMP3Info.Album
  
      If bolToyBuscando = False Then
         lblTrackRuta.Caption = accMP3Info.Artist & " - " & accMP3Info.Album
           'por si no cabe en la che label
            If Len(lblTrackRuta.Caption) > 33 Then
              lblTrackRuta.Caption = left(lblTrackRuta.Caption, 30) & "..."
              lblTrackRuta.ToolTipText = accMP3Info.Artist & " - " & accMP3Info.Album
            Else
              lblTrackRuta.ToolTipText = ListaRep.Path
            End If

      End If
        
        '// title
        If accMP3Info.Title <> "" Then
           ScrollText = accMP3Info.Title
           picScroll.ToolTipText = sFileMainPlaying
           CambiarIcono Text1.hwnd, Me.Icon.Handle, "<< " & accMP3Info.Title & " >> - << " & accMP3Info.Artist & " >>"
        Else
           ScrollText = Trim(left(Trim(ListaRep.FileName), Len(Trim(ListaRep.FileName)) - 4))
           picScroll.ToolTipText = sFileMainPlaying
           CambiarIcono Text1.hwnd, Me.Icon.Handle, "<< " & ScrollText & " >>"
        End If


  Else
     '// si no tiene tags o estan bacios
     ScrollText = Trim(left(Trim(ListaRep.FileName), Len(Trim(ListaRep.FileName)) - 4))
     picScroll.ToolTipText = sFileMainPlaying

     CambiarIcono Text1.hwnd, Me.Icon.Handle, "<< " & ScrollText & " >>"
     aryTitulo = Split(ListaRep.Path, "\")
      i = UBound(aryTitulo)
    
    '// poner el nombre del folder de mp3s
     strFolder = Trim(aryTitulo(i))
   
     If bolToyBuscando = False Then
       lblTrackRuta.Caption = strFolder
       lblTrackRuta.ToolTipText = ListaRep.Path
       
       If Len(lblTrackRuta.Caption) > 23 Then
         lblTrackRuta.ToolTipText = lblTrackRuta.Caption
         lblTrackRuta.Caption = left(lblTrackRuta.Caption, 23) & "..."
       End If
      End If
     strArtist = strFolder
     strAlbum = strFolder
 End If

 '// si esta la mini maskara
 If bolMiniMascara = True Then
  If bolToyBuscando = False Then Scroll_Text ScrollText, frmMini.picScroll
 Else
   Scroll_Text ScrollText, picScroll
 End If

If HasLyrics3Tag = True And Trim(accMP3Info.LYRICS) <> "" Then
  Show_Lyrics Trim(accMP3Info.LYRICS)
Else
  LyricsRef.Clear
End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
 Sub SliderReproduccioN_Down(X As Single, Y As Single)
    If PlayerIsPlaying = "false" Then Exit Sub
    If slidePos = False Then
        DragX = X: DragY = imgNormal(15).left
        slidePos = True
    End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub SliderReproduccioN_Move(X As Single, Y As Single)
    If slidePos = True Then
        Pos = DragY + (X - DragX)
         If Pos < 0 Then Pos = 0
         If Pos > 129 Then Pos = 129
        DragY = Pos: imgNormal(15).left = Pos
        imgNormal(15).Picture = imgNormal(15).Image
        picSliderRep.Picture = picSliderRep.Image

     Dim P As Variant
       P = Int(((Pos - 1) / 128) * (PlayerPos.Duration * 1000))
     Dim CurPos
       CurPos = P / 1000
       If CurPos < 0 Then CurPos = 0
       lblTiempoT.Caption = Convert_Time(CurPos)
    End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub SliderReproduccioN_Up(X As Single, Y As Single)
 On Error GoTo Hell
 Dim P As Variant
   If PlayerIsPlaying = "false" Then Exit Sub
   P = Int(((Pos - 1) / 128) * (PlayerPos.Duration * 1000))
 Dim CurPos
   CurPos = P / 1000
   If CurPos < 0 Then CurPos = 0
   PlayerPos.CurrentPosition = CurPos
   slidePos = False
   If PlayerIsPlaying = "pause" Then Pause_Play
Exit Sub
Hell:
slidePos = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub SliderVolumen_Down(X As Single, Y As Single)
    If slideVol = False Then
      Timer_Texto.Enabled = False
      If PlayerMute = True Then Player_Mute
        DragY = Y: DragX = imgNormal(16).Top
        slideVol = True
    End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub SliderVolumen_Move(X As Single, Y As Single)
 On Error GoTo Hell
 Dim intPorcentaje As Integer
    If slideVol = True Then
        PosVol = DragX + (Y - DragY)
          If PosVol < 0 Then PosVol = 0
          If PosVol > 89 Then PosVol = 89
        DragX = PosVol: Ajust_Volume PosVol
    End If
Exit Sub
Hell:
slideVol = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static rec As Boolean, msg As Long
   msg = X / Screen.TwipsPerPixelX
   If rec = False Then
      rec = True
      ' Captura cada evento de botones del Raton
      Select Case msg
        Case WM_LBUTTONDBLCLK  ' Doble click Boton Izquierdo
           If bolAcercaShow = True Then frmAcerca.Show
           If bolCaratulaShow = True Then frmCaratula.Show
           If bolDirectoriosShow = True Then frmDirectorios.Show
           If bolOpcionesShow = True Then frmOpciones.Show
           If bolLyricsShow = True Then frmLyrics.Show
           If bolTagsShow = True Then frmTags.Show
                      
           If bolMiniMascara = True Then
               frmMini.Show
           Else
               Me.Show
           End If
       Case WM_LBUTTONDOWN  ' Boton Izquierdo pulsado
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
        Case WM_RBUTTONDBLCLK ' Doble Click Boton Derecho
        Case WM_RBUTTONDOWN ' Boton derecho pulsado
        Case WM_RBUTTONUP  ' Boton derecho Arriba
           Me.PopupMenu frmPopUp.mnuMenuPrincipal
     End Select
      rec = False
   End If
   DoEvents
End Sub



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Intro_Timer()
 TiempoIntro = TiempoIntro + 1
 If TiempoIntro = 10 Then
  If PlayerLoop = True Then
    Play
  Else
    Next_Track
  End If
  TiempoIntro = 0
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Scroll_Text(ScrollText As String, picScroll As PictureBox)
 If (TextWidth(ScrollText) / 15) > picScroll.ScaleWidth Then
  RunMain picScroll
  Timer_Texto.Enabled = True
 Else
  Timer_Texto.Enabled = False
   picScroll.Cls
   picScroll.CurrentX = (picScroll.ScaleWidth / 2) - ((TextWidth(ScrollText) / 2) / 15)
   picScroll.Print ScrollText
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Texto_Timer()
 On Error Resume Next
 Static Espera As Integer, ToyDerecha As Boolean
      If ToyDerecha = False Then
       
         '// si esta la mini maskara
         If bolMiniMascara = True Then
            frmMini.picScroll.Cls
            DrawText frmMini.picScroll.hDC, ScrollText, -1, DrawingRect, DT_SINGLELINE
         Else
            picScroll.Cls
            DrawText picScroll.hDC, ScrollText, -1, DrawingRect, DT_SINGLELINE
         End If
         
        'Actualiza las coordenadas del rectangulo
        If Izquierda = False Then
         DrawingRect.left = DrawingRect.left - 1
         DrawingRect.Right = DrawingRect.Right - 1
        Else
         DrawingRect.left = DrawingRect.left + 1
         DrawingRect.Right = DrawingRect.Right + 1
        End If
        
        '// si esta la mini maskara
         If bolMiniMascara = True Then
           If DrawingRect.Right < (frmMini.picScroll.ScaleWidth - 20) And Izquierda = False Then   'Tiempo de reinicio
              ToyDerecha = True
              Izquierda = True
           End If
         Else
           If DrawingRect.Right < (picScroll.ScaleWidth - 20) And Izquierda = False Then   'Tiempo de reinicio
              ToyDerecha = True
              Izquierda = True
           End If
         End If
         
         If DrawingRect.left > 20 And Izquierda = True Then   'Tiempo de reinicio
            Izquierda = False
            ToyDerecha = True
         End If
        
        '// si esta la mini maskara
         If bolMiniMascara = True Then
            frmMini.picScroll.Refresh
         Else
            picScroll.Refresh
         End If
      
      Else
        Espera = Espera + 1
        If Espera = 30 Then ToyDerecha = False: Espera = 0
      End If
    DoEvents
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub RunMain(picScroll As PictureBox)
  On Error Resume Next
  rt = DrawText(picScroll.hDC, ScrollText, -1, DrawingRect, DT_CALCRECT)
 If rt <> 0 Then 'Si marca error
    DrawingRect.Top = (picScroll.ScaleHeight / 2) - ((TextHeight(ScrollText) / 2) / 15)
    DrawingRect.Right = TextWidth(ScrollText) / 15
    DrawingRect.Bottom = DrawingRect.Bottom + (TextHeight(ScrollText) / 15)
 Else
    Timer_Texto.Enabled = False
 End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| fUNCTION FOR ORDER THE LYRICS IN THE LIST FOR SHOW
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Start_Lyrics()
  '// has lyrics
 LyricsIndex = 1
 
 If HasLyrics3Tag = True And LyricsRef.ListCount > 0 Then
   '// form lyrics showing
   If bolLyricsShow = True Then
      frmLyrics.lblAlbum.Caption = strAlbum
      frmLyrics.lblArtist.Caption = strArtist
      frmLyrics.lblTitle.Caption = ScrollText
      frmLyrics.picLyrics.Visible = True
      frmLyrics.lblNoLyrics.Visible = False
   End If
 Else
   If bolLyricsShow = True Then
     frmLyrics.lblAlbum.Caption = strAlbum
     frmLyrics.lblArtist.Caption = strArtist
     frmLyrics.lblTitle.Caption = ScrollText
     frmLyrics.picLyrics.Visible = False
     frmLyrics.lblNoLyrics.Visible = True
   End If
 End If

End Sub


Sub Show_Lyrics(strLyrics As String)
 On Error Resume Next
   Dim RawLyrics() As String
   Dim Start As Integer
   Dim i As Integer, L As Integer, j As Integer, fin As Integer
   Dim strTemp As String
   Dim strTemp2 As String
   Dim startLyrics As Integer
   LyricsRef.Clear
  
   If Trim(strLyrics) = "" Then Exit Sub
   'check for timestamps
   If InStr(strLyrics, "[") = 0 Then Exit Sub
   'ok, it has lyrics, now put them into an array
   
   RawLyrics = Split(strLyrics, vbCr)
   L = UBound(RawLyrics)
   
   For i = 0 To L - 1
      Start = 1
      RawLyrics(i) = Trim(RawLyrics(i))
      Do
         j = InStr(Start, RawLyrics(i), "[")
         If j > 0 Then
            fin = InStr(Start, RawLyrics(i), "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) < 9 Then
             '// extract time
              strTemp = Mid$(RawLyrics(i), j + 1, fin - j - 1)
              '// extract lyrics
               startLyrics = InStrRev(RawLyrics(i), "]", Len(RawLyrics(i)))
              strTemp2 = Right(RawLyrics(i), Len(RawLyrics(i)) - startLyrics)
              '// 00:00:00
              LyricsRef.AddItem strTemp & "    " & strTemp2
            End If
         End If
         Start = fin + 1
      Loop Until j = 0

   Next i
   
   If bolLyricsShow = True Then
      frmLyrics.Order_lblLyrics
   End If
End Sub

Private Sub Update_Lyrics()
   Dim NumberOfLines As Integer
   Dim sCurrentTime As String
   
   'now display the lyrics
   NumberOfLines = LyricsRef.ListCount
   On Error GoTo Hell
   
   sCurrentTime = Trim(lblTiempoT.Caption)
   
   'do we need to go forward?
   If sCurrentTime >= Trim(left$(LyricsRef.List(LyricsIndex), 9)) Then
      'yes.. how much??
      Do Until sCurrentTime <= Trim(left$(LyricsRef.List(LyricsIndex), 9)) Or LyricsIndex = NumberOfLines
         frmLyrics.Move_Next_Focus_Lyrics
       LyricsIndex = LyricsIndex + 1
      Loop
   
   'do we need to go backwards?
   Else
      'yes, how much?

      Do Until sCurrentTime >= Trim(left$(LyricsRef.List(LyricsIndex - 1), 9)) Or LyricsIndex = 1
         LyricsIndex = LyricsIndex - 1
         frmLyrics.Move_Previous_Focus_Lyrics
      Loop

   End If
   Exit Sub
Hell:

End Sub


Private Sub txtSTIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static rec As Boolean, msg As Long
   msg = X / Screen.TwipsPerPixelX
   If rec = False Then
      rec = True
      ' Captura cada evento de botones del Raton
      Select Case msg
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
           Select Case Index
               Case 0 '// Previous
                 Previous_Track
               Case 1 '// Play
                 Play
               Case 2 '// Pause
                 Pause_Play
               Case 3 '// Stop
                 Stop_Player
               Case 4 '// Next
                 Next_Track
           End Select
          End Select
      rec = False
   End If
   DoEvents

End Sub
