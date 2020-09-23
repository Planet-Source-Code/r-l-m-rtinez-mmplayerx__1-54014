Attribute VB_Name = "mStart"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   VARIABLES UTILIZADAS PARA TODO EL PROGRAMA                                          |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public strRutaCaratula As String
Public CopyMp3Totales As Integer
Public CopyTotalAlbums As Integer
Public bolCaratulaShow As Boolean, bolDirectoriosShow As Boolean
Public bolAcercaShow As Boolean, bolOpcionesShow As Boolean
Public bolLyricsShow As Boolean, bolTagsShow As Boolean
Public OriginalWallpaperStyle As Integer
Public OriginalTileWallpaper As Integer
Public OriginalRutaWallpaper As String
Public bolCaratulaDefault As Boolean
Public bolSplashScreen As Boolean
Public strTraySearch As String
Public intActiveAlbum As Integer
Public TotalAlbumS As Integer
Public MP3totales As Integer
Public ScrollText As String
Public sFileMainPlaying As String
Public PlayerState As String

Public RegionMusic As Long
Public RegionMini As Long
Public Type Entry
    NoAlteraR As Boolean
    Mosaico As Boolean
    Centrar As Boolean
    Proporcional As Boolean
    Expander As Boolean
    Directorio As Boolean
    Language As String
    Ingles As Boolean
    Alpha As Integer
    SiempreTop As Boolean
    Splash As Boolean
    Instancias As Boolean
    MP3FILE As Boolean
    WAVFILE As Boolean
    WMAFILE As Boolean
End Type

Public Type TrayIcon
    Previous As Boolean
    Play As Boolean
    Pause As Boolean
    Stop As Boolean
    Next As Boolean
End Type

Public strSkinSeleccionado As String
Public bolMiniMascara As Boolean
Public OpcionesMusic As Entry
Public PlayerTrayIcon As TrayIcon


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  INICIO DE LA APLICATION                                                              |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Main()
  On Error Resume Next
  Dim strPath As String, args As String
  
 '----------------------------------------
   '// Optional Load XP Theme need the component [Microsoft Windows Common Controls 6.0]
   '// or load if no cheched
    InitCommonControls
    XPStyle False
 '----------------------------------------
   
 '// running right click in explorer or other
 '// HKCR\Directory\shell\Search Music Mp3 Player X\command\predeterminado
   args = Trim$(Command$)
 If args <> "" Then '// si lo executan desde el explorador solo buskar alli
    MusicMp3.bolShowFront = True
    Load_Settings_INI False
    MusicMp3.lblTrackRuta.Caption = arryLanguage(68)
    MusicMp3.Search_Mp3s (args)
 Else
    MusicMp3.bolShowFront = True
    Load_Settings_INI True
    
    strPath = Path_Exe(PathExe)
    'strPath = "C:\"
     If MP3totales = 0 Then
       MusicMp3.Search_Mp3s (strPath)
     End If
 End If
Unload frmSplash
If OpcionesMusic.SiempreTop = True Then Always_on_Top
If bolMiniMascara = True Then Change_Mask True
End Sub

