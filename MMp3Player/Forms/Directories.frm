VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDirectorios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Albums Browser"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "Directories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView TreeAlbums 
      Height          =   2220
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   3916
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox FileSearch 
      Appearance      =   0  'Flat
      Height          =   615
      Hidden          =   -1  'True
      Left            =   2745
      Pattern         =   "*.mp3;*.wav;*.wma"
      System          =   -1  'True
      TabIndex        =   1
      Top             =   2670
      Visible         =   0   'False
      Width           =   1605
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1125
      Top             =   2775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Directories.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Directories.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Directories.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Directories.frx":0A02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPaths 
      Caption         =   "Paths"
      Visible         =   0   'False
      Begin VB.Menu mnuExpArchivos 
         Caption         =   "Explorar Archivos"
      End
      Begin VB.Menu mnuEditTags 
         Caption         =   "Editar Tags"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "Reproducir"
      End
   End
End
Attribute VB_Name = "frmDirectorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
On Error Resume Next
  bolDirectoriosShow = True
  Me.Icon = MusicMp3.Icon
  Me.Caption = Trim(arryLanguage(8)) & " [ " & TotalAlbumS & " Albums ]"

  frmDirectorios.left = (Screen.Width - frmDirectorios.Width) / 2
  frmDirectorios.Top = (Screen.Height - frmDirectorios.Height) / 2
  
  Me.mnuExpArchivos.Caption = "  " & Trim(arryLanguage(7))
  Me.mnuEditTags.Caption = "  " & Trim(arryLanguage(9))
  Me.mnuPlay.Caption = "  " & Trim(Right(arryLanguage(16), Len(arryLanguage(16)) - 1))

 Load_Albums
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Load_Albums()
 Dim iAlbum As Integer, iTrack As Integer
 Dim ruta As String, a As String
 Dim Discos As Integer, i As Integer
 
On Error Resume Next
 
 ruta = strTraySearch
FileSearch.Pattern = MusicMp3.ListaRep.Pattern
TreeAlbums.Nodes.Clear

 '// add albums folders
 For iAlbum = 1 To TotalAlbumS
   a = Mid(MusicMp3.picAlbum(iAlbum).ToolTipText, Len(ruta), Len(MusicMp3.picAlbum(iAlbum).ToolTipText))
    If Trim(a) <> "" Then
      TreeAlbums.Nodes.Add , , CStr(iAlbum & " \"), "[ " & a & " ]", 1
    Else
      a = Mid(MusicMp3.picAlbum(iAlbum).ToolTipText, InStrRev(MusicMp3.picAlbum(iAlbum).ToolTipText, "\") + 1, Len(MusicMp3.picAlbum(iAlbum).ToolTipText))
      TreeAlbums.Nodes.Add , , CStr(iAlbum & " \"), "[ " & a & " ]", 1
    End If
    
   '// add files in album
     FileSearch.Path = MusicMp3.picAlbum(iAlbum).ToolTipText
        For iTrack = 0 To FileSearch.ListCount - 1
          If LCase(Right(FileSearch.List(iTrack), 3)) = "mp3" Then
             TreeAlbums.Nodes.Add CStr(iAlbum & " \"), tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 2
          ElseIf LCase(Right(FileSearch.List(iTrack), 3)) = "wma" Then
                   TreeAlbums.Nodes.Add CStr(iAlbum & " \"), tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 3
              Else
                  TreeAlbums.Nodes.Add CStr(iAlbum & " \"), tvwChild, CStr(iAlbum & " \ " & iTrack), FileSearch.List(iTrack), 4
              End If
        Next iTrack
     a = ""
 Next iAlbum
  '// seleccionar el album reproduciendo
  TreeAlbums.Nodes(CStr(intActiveAlbum & " \")).Selected = True
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Resize()
  TreeAlbums.Width = Me.ScaleWidth + 50
  TreeAlbums.Height = Me.ScaleHeight + 50
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Unload(Cancel As Integer)
bolDirectoriosShow = False
End Sub

Private Sub mnuEditTags_Click()
  Dim nodAlbum As Node
  Dim arryTemp() As String
  
  On Error Resume Next
  
  Set nodAlbum = TreeAlbums.SelectedItem
   
  arryTemp = Split(nodAlbum.Key, "\")
  
  If UBound(arryTemp) = 1 Then '// is a file click
     If bolTagsShow = True Then
       frmTags.FileTags.Path = MusicMp3.picAlbum(CInt(arryTemp(0))).ToolTipText
       If Trim(arryTemp(1)) <> "" Then frmTags.FileTags.Selected(CInt(arryTemp(1))) = True
       frmTags.ZOrder 0
     Else
       frmTags.Show
       frmTags.FileTags.Path = MusicMp3.picAlbum(CInt(arryTemp(0))).ToolTipText
       If Trim(arryTemp(1)) <> "" Then frmTags.FileTags.Selected(CInt(arryTemp(1))) = True
     End If
  End If

 
 
End Sub

Private Sub mnuExpArchivos_Click()
  Dim nodAlbum As Node
  Dim arryTemp() As String
  Dim x As Long
  
  If TreeAlbums.Nodes.Count = 0 Then Exit Sub
   
  On Error Resume Next
  
  Set nodAlbum = TreeAlbums.SelectedItem
   
  arryTemp = Split(nodAlbum.Key, "\")
  
  If UBound(arryTemp) = 1 Then
       x = Shell("explorer.exe " & MusicMp3.picAlbum(CInt(arryTemp(0))).ToolTipText, vbMaximizedFocus)
  End If

End Sub

Private Sub mnuPlay_Click()
  Dim nodAlbum As Node
  Dim arryTemp() As String
  
  On Error Resume Next
  
  If TreeAlbums.Nodes.Count = 0 Then Exit Sub

  Set nodAlbum = TreeAlbums.SelectedItem
   
  arryTemp = Split(nodAlbum.Key, "\")
  
  If UBound(arryTemp) = 1 Then '// is a file click
     MusicMp3.Play_Album CInt(arryTemp(0))
     If Trim(arryTemp(1)) <> "" Then MusicMp3.ListaRep.Selected(CInt(arryTemp(1))) = True
  End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub TreeAlbums_DblClick()
  Dim nodAlbum As Node
  Dim arryTemp() As String
  
  On Error Resume Next
  
  If TreeAlbums.Nodes.Count = 0 Then Exit Sub

  Set nodAlbum = TreeAlbums.SelectedItem
   
  arryTemp = Split(nodAlbum.Key, "\")
  
  If UBound(arryTemp) = 1 Then '// is a file click
     MusicMp3.Play_Album CInt(arryTemp(0))
     If Trim(arryTemp(1)) <> "" Then MusicMp3.ListaRep.Selected(CInt(arryTemp(1))) = True
  End If
End Sub

Private Sub TreeAlbums_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If TreeAlbums.Nodes.Count = 0 Then Exit Sub
 If Button = vbRightButton Then PopupMenu Me.mnuPaths
End Sub


