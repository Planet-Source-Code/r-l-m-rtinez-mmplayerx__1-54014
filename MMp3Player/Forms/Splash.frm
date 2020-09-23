VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "Splash.frx":000C
   ScaleHeight     =   2400
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSplash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   135
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   105
      Width           =   45
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   1
      Left            =   135
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1545
      Width           =   4740
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading... "
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
      Height          =   600
      Index           =   0
      Left            =   135
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1695
      Width           =   4665
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
If MusicMp3.bolToyBuscando = True Then MusicMp3.bolToyBuscando = False
End Sub

Private Sub Form_Load()
bolSplashScreen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
bolSplashScreen = False
If bolMiniMascara = True Then
  frmMini.Show
Else
  MusicMp3.Show
End If

End Sub

Private Sub lblSplash_DblClick(Index As Integer)
  If MusicMp3.bolToyBuscando = True Then MusicMp3.bolToyBuscando = False
End Sub

