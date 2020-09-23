VERSION 5.00
Begin VB.Form frmLyrics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lyrics"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lyrics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   2430
      Left            =   15
      ScaleHeight     =   2400
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   600
      Width           =   4365
      Begin VB.PictureBox picLyrics 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2025
         Left            =   15
         ScaleHeight     =   2025
         ScaleWidth      =   4320
         TabIndex        =   1
         Top             =   0
         Width           =   4320
         Begin VB.Shape shpFocus 
            BorderColor     =   &H00FF0000&
            Height          =   240
            Left            =   0
            Top             =   15
            Width           =   4305
         End
         Begin VB.Label lblLyrics 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Letras"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   15
            Width           =   4305
         End
      End
      Begin VB.Label lblNoLyrics 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ no lyrics found ]"
         Height          =   240
         Left            =   60
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guns and Roses"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   15
      TabIndex        =   4
      Top             =   360
      Width           =   4395
   End
   Begin VB.Label lblAlbum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guns and Roses"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   15
      TabIndex        =   3
      Top             =   165
      Width           =   4395
   End
   Begin VB.Label lblArtist 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guns and Roses"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   15
      TabIndex        =   2
      Top             =   -30
      Width           =   4395
   End
End
Attribute VB_Name = "frmLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iLinesLyrics As Integer
Dim iCurrentLine As Integer

Dim lblForeColor As Long
Dim iLinesMax As Integer

Public Sub Reset_Values()
On Error Resume Next
 iLinesMax = 0
 
 If iCurrentLine > 0 Then lblLyrics(iCurrentLine).Font.Bold = False
 
 iCurrentLine = 0
 iLinesLyrics = MusicMp3.LyricsRef.ListCount
 picLyrics.Top = 0
 lblLyrics(0).Font.Bold = True
 shpFocus.Top = lblLyrics(0).Top - 20
End Sub

Public Sub Move_Previous_Focus_Lyrics()
    
  iCurrentLine = iCurrentLine - 1
   If iCurrentLine < 0 Then
      iCurrentLine = 0
      iLinesMax = 0
      Exit Sub
   End If
 
   iLinesMax = iLinesMax - 1
   
   If iLinesMax < 0 Then
     iLinesMax = 9
     picLyrics.Top = (lblLyrics(0).Height * 10) - (lblLyrics(0).Height * (iCurrentLine + 1))
   End If
   
   shpFocus.Top = lblLyrics(iCurrentLine).Top - 20
 
   lblLyrics(iCurrentLine + 1).Font.Bold = False
   lblLyrics(iCurrentLine).Font.Bold = True
End Sub


Public Sub Move_Next_Focus_Lyrics()
  iCurrentLine = iCurrentLine + 1
   If iCurrentLine > iLinesLyrics Then
      iCurrentLine = iLinesLyrics
      Exit Sub
   End If
 
   iLinesMax = iLinesMax + 1
   
   If iLinesMax > 9 Then
     iLinesMax = 0
     picLyrics.Top = -(lblLyrics(0).Height * iCurrentLine)
   End If
   
   
    shpFocus.Top = lblLyrics(iCurrentLine).Top - 20
      
    lblLyrics(iCurrentLine - 1).Font.Bold = False
    lblLyrics(iCurrentLine).Font.Bold = True
End Sub

Public Sub Order_lblLyrics()
 Dim i As Integer
 Dim intHeight As Integer
 Dim strLyrics As String
 
 If MusicMp3.LyricsRef.ListCount = 0 Then Exit Sub
  lblForeColor = lblNoLyrics.ForeColor
  
   iLinesMax = 0
   iCurrentLine = 0
   iLinesLyrics = MusicMp3.LyricsRef.ListCount
   picLyrics.Top = 0
   shpFocus.Top = lblLyrics(0).Top - 20
   lblLyrics(0).Font.Bold = False

  
  For i = 0 To MusicMp3.LyricsRef.ListCount - 1
    If i >= lblLyrics.Count Then
      Load lblLyrics(i)
    End If
    
    If i > 0 Then lblLyrics(i).Top = lblLyrics(i - 1).Top + lblLyrics(i - 1).Height
    strLyrics = Right(MusicMp3.LyricsRef.List(i), Len(MusicMp3.LyricsRef.List(i)) - 9)
    lblLyrics(i).Caption = strLyrics
    lblLyrics(i).ForeColor = lblForeColor
    lblLyrics(i).Visible = True
    intHeight = intHeight + lblLyrics(i).Height
  Next i
 picLyrics.Height = intHeight
 lblLyrics(0).Font.Bold = True
End Sub


Private Sub Form_Load()

  bolLyricsShow = True
  
  Me.Caption = Trim(arryLanguage(10))
Me.Icon = MusicMp3.Icon
  frmLyrics.left = (Screen.Width - frmLyrics.Width) / 2
  frmLyrics.Top = (Screen.Height - frmLyrics.Height) / 2

     picLyrics.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
     picBody.BackColor = picLyrics.BackColor
     
     shpFocus.BorderColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
     lblNoLyrics.ForeColor = shpFocus.BorderColor

 MusicMp3.LyricsIndex = 1
 If MusicMp3.LyricsRef.ListCount > 0 Then
      lblAlbum.Caption = MusicMp3.strAlbum
      lblArtist.Caption = MusicMp3.strArtist
      lblTitle.Caption = ScrollText
      Order_lblLyrics
      picLyrics.Visible = True
      lblNoLyrics.Visible = False
 Else
      lblAlbum.Caption = MusicMp3.strAlbum
      lblArtist.Caption = MusicMp3.strArtist
      lblTitle.Caption = ScrollText
      picLyrics.Visible = False
      lblNoLyrics.Visible = True
      lblNoLyrics.Caption = arryLanguage(77)
 End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
 bolLyricsShow = False
End Sub
