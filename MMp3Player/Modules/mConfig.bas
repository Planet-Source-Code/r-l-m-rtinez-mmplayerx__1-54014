Attribute VB_Name = "mConfig"
Option Explicit
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| OPERATING SYSTEM VERSIN INFORMATION                                                   |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| EXECUTAR APLICACIONES CON LOS PARAMETROS DADOS                                        |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| ARRASTRE DEL FORMULARIO                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS PARA PONER SIEMPRE ARRIBA EL FORMULARIO                                          |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  No usar el orden Z del propietario
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| MOVER EL TEXTO POR LOS PICTURES
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Const DT_BOTTOM As Long = &H8
Public Const DT_CALCRECT As Long = &H400
Public Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Public Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Public Const DT_WORDBREAK As Long = &H10

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|    Declaraciones para Layered Windows (sólo Windows 2000 y superior)                  |
'|    APIS PARA PONER TRASPARENTE EL FORM                                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_ALPHA As Long = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const RDW_INVALIDATE = &H1
Public Const RDW_ERASE = &H4
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_FRAME = &H400

'
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function RedrawWindow2 Lib "user32" Alias "RedrawWindow" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  APIS PARA LEER LAS CONFIGURACIONES DE LOS ARCHIVOS .INI O DEMAS
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Declare Function GetPrivateProfileString Lib "Kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long
    
 Public Enum Sel_Option
  PathExe = 0
  PathSkin = 1
End Enum

Public Function Read_INI(Section As String, Value As String, Default As Variant, Optional IsColor As Boolean = False, Optional ConfigurationMusic As Boolean = False) As Variant
 '// Funcion para leer configuraciones del INI
 '// Parametros
 '// [Section] -> Rama principal del .ini : ei:  [Configuration]
 '// [Value] -> Valor de la Seccion , ej: Intro = False
 '// [Default] -> Valor de retorno si no se encuantra el valor
 '// [IsColor] -> Opcional saber si es color para tratarlo diferente
 '// [ConfigurationMusic] -> opcional, Leer el el archivo principal del programa
 '// Valor de retorno el valor de la seccion si se encuantra
 
 Dim ColorArr As Variant
 Dim Str As String
    
  If ConfigurationMusic = True Then
    Str = String(255, Chr(0))
    Str = left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), Path_Exe(PathExe) & App.EXEName & ".ini"))
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Trim(Default)
    Else
       Read_INI = Trim(Str)
    End If
    Exit Function
  End If
      
  If IsColor = True Then ' is a color
    Str = String(255, Chr(0))
    Str = left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), Path_Exe(PathSkin) & "Skin.ini"))
    
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Default
    Else
      ColorArr = Split(Str, ",")
       If UBound(ColorArr) <> 2 Then ' si esta mal la che clave
         Read_INI = Default
       Else
         Read_INI = RGB(ColorArr(0), ColorArr(1), ColorArr(2))
       End If
    End If
  Else
    Str = String(255, Chr(0))
    Str = left(Str, GetPrivateProfileString(Section, ByVal Value, "NO_TA", Str, Len(Str), Path_Exe(PathSkin) & "Skin.ini"))
    If Str = "NO_TA" Then ' si no encuentra la clave
       Read_INI = Trim(Default)
    Else
       Read_INI = Trim(Str)
    End If
  End If
End Function

Public Function Write_INI(Section As String, KeyName As String, KeyValue As String) As Boolean
Dim Ret As Long
    Ret = WritePrivateProfileString(Section, KeyName, KeyValue, Path_Exe(PathExe) & "Skin.ini")
    If Ret = 0 Then
        Write_INI = True
    Else
        Write_INI = False
    End If
End Function

Sub Load_Settings_INI(bolNormal As Boolean)
 Dim strRes As Variant, strAlbums As String, strPathern As String
 Dim i As Integer, intMp3 As Integer
 Dim strKeyQuery As Variant
 Dim lngRootKey As Long

 On Error Resume Next
  strKeyQuery = vbNullString
  lngRootKey = HKEY_CURRENT_USER
 
  '// si existe el archivo de configuracion
  'If Dir(Path_Exe(PathExe) & App.EXEName & ".ini") = "" Then Exit Sub
 
  '// Multiples instancias
  strRes = Read_INI("Configuration", "MulInstances", "false", , True)
  
  If LCase(strRes) = "false" Then       '// Si este en falso y hay otra
    If App.PrevInstance = True Then     '// Instancia terminar
      'AppActivate "MusicMp3 v.1.0"
      End
      Exit Sub
    End If
  End If
  If LCase(strRes) = "true" Then OpcionesMusic.Instancias = True
    
  '// Mostrar Splash Screen
  strRes = Read_INI("Configuration", "SplashScreen", "false", , True)
  If LCase(strRes) = "true" Then
     frmSplash.lblSplash(0).Caption = "Loading configuration..."
     frmSplash.Show
     'MusicMp3.Hide
     OpcionesMusic.Splash = True
  End If
  
  '// Kargar Skin
  strRes = Read_INI("Configuration", "Skin", "/Default", , True)
  strSkinSeleccionado = Trim(strRes)
  Change_Skin strSkinSeleccionado '// cambiar skin, posicion de los controles
  Form_Mini_Normal '// si tiene zonas irregulares ajustar el form
  Load_Skins_Menu LCase(strSkinSeleccionado) '// kargar el menu de Skins y seleccionar el actual
  
  '// Estado de la maskara mini - normal
  strRes = Read_INI("Configuration", "Mini", "false", , True)
  If LCase(strRes) = "true" Then bolMiniMascara = True
  
  '// Mover los formularios
  strRes = Read_INI("Configuration", "MX", 0, , True)
  If IsNumeric(strRes) = False Then strRes = 0
     MusicMp3.left = CInt(strRes)

   
  strRes = Read_INI("Configuration", "MY", 0, , True)
  If IsNumeric(strRes) = False Then strRes = 0
     MusicMp3.Top = CInt(strRes)
    
  '// si no esta seleccionado el splash screen mostrar los form ahora
  If bolSplashScreen = False Then
   If bolMiniMascara = True Then
     frmMini.Show
   Else
     MusicMp3.Show
   End If
  End If
   
  '-----------------------------------------------------------------------
  'Guardar la ruta del Wallpaper al inicio que se tiene
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "Wallpaper")
  OriginalRutaWallpaper = strKeyQuery
  
  '-----------------------------------------------------------------------
  'Guardar el Estilo Wallpaper al inicio
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "WallpaperStyle")
  OriginalWallpaperStyle = strKeyQuery
  
  '-----------------------------------------------------------------------
  'Guardar el tileWallpaper al inicio
  strKeyQuery = regQuery_A_Key(lngRootKey, "Control panel\Desktop", "TileWallpaper")
  OriginalTileWallpaper = strKeyQuery
  
  
  '// Guardar los estilos de Walppaper al inicio
  strRes = Read_INI("Configuration", "Wallpaper", 0, , True)
  
  If CInt(strRes) < 0 Or CInt(strRes) > 3 Or IsNumeric(strRes) = False Then strRes = 0
  
  '//Poner valores correctos por si modifican el archivo
  If strRes = 0 Then
    OpcionesMusic.NoAlteraR = True
  ElseIf strRes = 1 Then
        OpcionesMusic.Mosaico = True
      ElseIf strRes = 2 Then
            OpcionesMusic.Centrar = True
          Else
            OpcionesMusic.Expander = True
          End If
          
  '// check proporcional
  strRes = Read_INI("Configuration", "Proportional", "false", , True)
  If LCase(strRes) = "true" Then OpcionesMusic.Proporcional = True
  
  '// check Directorio
  strRes = Read_INI("Configuration", "Directory", "false", , True)
  If LCase(strRes) = "true" Then OpcionesMusic.Directorio = True
 
   '// system tray icons
  ColocarIcono MusicMp3.Text1.hwnd, MusicMp3.Icon.Handle, "Music Mp3 Player X v.1.0 - by Raul Martinez"
  
  strRes = Read_INI("Configuration", "SysTrayNext", "false", , True)
  If LCase(strRes) = "true" Then
     PlayerTrayIcon.Next = True
     ColocarIcono MusicMp3.txtSTIcon(4).hwnd, MusicMp3.ImageList.ListImages(5).ExtractIcon.Handle, "Next Track - MMp3Player"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayStop", "false", , True)
  If LCase(strRes) = "true" Then
     PlayerTrayIcon.Stop = True
     ColocarIcono MusicMp3.txtSTIcon(3).hwnd, MusicMp3.ImageList.ListImages(4).ExtractIcon.Handle, "Stop - MMp3Player"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayPause", "false", , True)
  If LCase(strRes) = "true" Then
     PlayerTrayIcon.Pause = True
     ColocarIcono MusicMp3.txtSTIcon(2).hwnd, MusicMp3.ImageList.ListImages(3).ExtractIcon.Handle, "Pause  - MMp3Player"
  End If
  
  strRes = Read_INI("Configuration", "SysTrayPlay", "false", , True)
  If LCase(strRes) = "true" Then
     PlayerTrayIcon.Play = True
     ColocarIcono MusicMp3.txtSTIcon(1).hwnd, MusicMp3.ImageList.ListImages(2).ExtractIcon.Handle, "Play - MMp3Player"
  End If
    
    strRes = Read_INI("Configuration", "SysTrayPrevious", "false", , True)
  If LCase(strRes) = "true" Then
     PlayerTrayIcon.Previous = True
     ColocarIcono MusicMp3.txtSTIcon(0).hwnd, MusicMp3.ImageList.ListImages(1).ExtractIcon.Handle, "Previous - MMp3Player"
  End If
  
  '// load lenguaje y cambiarlo
  strRes = Read_INI("Configuration", "Language", "English", , True)
  OpcionesMusic.Language = strRes
  Load_Language OpcionesMusic.Language
  
 '//----------------------------------------------------------------------------------
  '// play .mp3 files format
  strRes = Read_INI("Configuration", "PlayMP3", "true", , True)
  If LCase(strRes) = "true" Then
     OpcionesMusic.MP3FILE = True
     strPathern = "*.mp3"
  End If
  
  '// play .wma files format
  strRes = Read_INI("Configuration", "PlayWMA", "true", , True)
  If LCase(strRes) = "true" Then
     OpcionesMusic.WMAFILE = True
     If strPathern = "" Then
       strPathern = "*.wma"
     Else
       strPathern = strPathern & ";*.wma"
     End If
  End If
  
  '// play wav files format
  strRes = Read_INI("Configuration", "PlayWAV", "true", , True)
  If LCase(strRes) = "true" Then
     OpcionesMusic.WAVFILE = True
      If strPathern = "" Then
        strPathern = "*.wav"
      Else
        strPathern = strPathern & ";*.wav"
      End If
  End If
   
  If strPathern = "" Then strPathern = "*.mp3"
  
  MusicMp3.ListaRep.Pattern = strPathern
  MusicMp3.FileSearch.Pattern = strPathern
  MusicMp3.FileAleatorio.Pattern = strPathern
 '//----------------------------------------------------------------------------------
  
  '// Trasparencia del form
  strRes = Read_INI("Configuration", "Alpha", 100, , True)
  If strRes < 10 Or strRes > 100 Then strRes = 100
  OpcionesMusic.Alpha = strRes
  If bolMiniMascara = True Then
    Make_Transparent frmMini.hwnd, OpcionesMusic.Alpha '// Poner Trasparente
  Else
    Make_Transparent MusicMp3.hwnd, OpcionesMusic.Alpha '// Poner Trasparente
  End If
      
      For i = 0 To 9
       If left(frmPopUp.mnuAlpha(i).Caption, Len(frmPopUp.mnuAlpha(i).Caption) - 1) = OpcionesMusic.Alpha Then
         frmPopUp.mnuAlpha(i).Checked = True
            frmPopUp.mnuAlphaPer.Caption = Trim(arryLanguage(34))
            frmPopUp.mnuAlphaPer.Checked = False
         Exit For
       Else
         frmPopUp.mnuAlphaPer.Caption = Trim(arryLanguage(34)) & " [ " & OpcionesMusic.Alpha & "% ]"
         frmPopUp.mnuAlphaPer.Checked = True
       End If
     Next i

  strRes = Read_INI("Configuration", "AlwaysTop", "false", , True)
  If LCase(strRes) = "true" Then OpcionesMusic.SiempreTop = True
   
 
  '// Ajustar Volumen
  strRes = Read_INI("Configuration", "Volume", 0, , True)
  If strRes < 0 Or strRes > 89 Then strRes = 0
   MusicMp3.Ajust_Volume CInt(strRes)
    
'// -------------------------------------------------------------------------------
If bolNormal = True Then '// si es cargado normalmente
 strRes = Read_INI("Configuration", "Intro", "false", , True)
  If LCase(strRes) = "true" Then MusicMp3.Intro
 
 strRes = Read_INI("Configuration", "Mute", "false", , True)
  If LCase(strRes) = "true" Then MusicMp3.Player_Mute
 
 strRes = Read_INI("Configuration", "Repeat", "false", , True)
  If LCase(strRes) = "true" Then MusicMp3.Player_Repeat
  
'---------------------------------------------------------------------------------------
'Hacer mientras se lea algo en el archivo .ini
 frmPopUp.fileBmps.Pattern = strPathern
 Do While strAlbums <> "\"
   i = i + 1
   strAlbums = Read_INI("albums", "Album_" & i, "\", , True)
   If strAlbums <> "\" Then
     If Dir(strAlbums & "\") <> "" Then  '// Si existe el directorio
       frmPopUp.fileBmps.Path = strAlbums
       If frmPopUp.fileBmps.ListCount > 0 Then '// Si hay mp3's
         CopyMp3Totales = CopyMp3Totales + frmPopUp.fileBmps.ListCount
         intMp3 = intMp3 + 1
         If intMp3 = 1 Then
           MusicMp3.picAlbum(intMp3).ToolTipText = strAlbums
         Else
           Load MusicMp3.picAlbum(intMp3)
           MusicMp3.picAlbum(intMp3).ToolTipText = strAlbums
         End If
       End If
     End If
   End If
 Loop
 
' TotalAlbumS = intMp3 + 1
 CopyTotalAlbums = intMp3 + 1
If intMp3 > 0 Then MusicMp3.Process_Albums False
 
 strRes = Read_INI("Configuration", "AlbumPlaying", 1, , True)
 If CInt(strRes) > 0 And CInt(strRes) <= (intMp3 - 2) Then
   MusicMp3.Play_Album CInt(strRes)
 ElseIf intMp3 > 0 Then
        MusicMp3.Play_Album 1
     End If
  
  strRes = Read_INI("Configuration", "TrackNumber", 0, , True)
 If CInt(strRes) >= 0 Then
   MusicMp3.ListaRep.Selected(CInt(strRes)) = True
   MusicMp3.ListaRep.ListIndex = CInt(strRes)
 End If
    

 strRes = Read_INI("Configuration", "RandomizeAlbum", "false", , True)
  If LCase(strRes) = "true" And intMp3 > 0 Then
    frmPopUp.Menu_Aleatorio_Album
  Else
     strRes = Read_INI("Configuration", "RandomizeCollection", "false", , True)
       If LCase(strRes) = "true" And intMp3 > 1 Then
         frmPopUp.Menu_Aleatorio_Coleccion
       End If
  End If
End If

End Sub

Sub Save_Settings_INI()
 Dim Fnum As Integer, i As Integer
 Dim ArchivoINI As String
 Dim intClave As Integer

 On Error Resume Next
 
  '// delete systray icons
  
   QuitarIcono MusicMp3.Text1.hwnd
   
  If PlayerTrayIcon.Previous = True Then QuitarIcono MusicMp3.txtSTIcon(0).hwnd
     
  If PlayerTrayIcon.Play = True Then QuitarIcono MusicMp3.txtSTIcon(1).hwnd
     
  If PlayerTrayIcon.Pause = True Then QuitarIcono MusicMp3.txtSTIcon(2).hwnd
     
  If PlayerTrayIcon.Stop = True Then QuitarIcono MusicMp3.txtSTIcon(3).hwnd
     
  If PlayerTrayIcon.Next = True Then QuitarIcono MusicMp3.txtSTIcon(4).hwnd
 
   
 On Error GoTo Bitch
ArchivoINI = Path_Exe(PathExe) & App.EXEName & ".ini"

If Dir(ArchivoINI) <> "" Then '// si existe el archivo borrarlo
 SetAttr ArchivoINI, vbNormal
 Kill ArchivoINI
End If
    Fnum = FreeFile  '// numeroaleatorio para asignar al archivo
    Open ArchivoINI For Output As Fnum
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, "|  cONFIGURATION fILE fOR mUSIC mP3 pLAYER X    |"
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, ""
    Print #Fnum, "[Configuration]" '// Seccion principal
    
      If OpcionesMusic.Splash = True Then '// mostrar splash screen
        Print #Fnum, "SplashScreen=true"
      Else
        Print #Fnum, "SplashScreen=false"
      End If
      
      If OpcionesMusic.Instancias = True Then '// Permitir Multiples instancias
        Print #Fnum, "MulInstances=true"
      Else
        Print #Fnum, "MulInstances=false"
      End If

    Print #Fnum, "Skin=" & strSkinSeleccionado  '// Skin seleccionado
    
    If bolMiniMascara = True Then
      Print #Fnum, "MX=" & frmMini.left
    Else
      Print #Fnum, "MX=" & MusicMp3.left
    End If
    
    If bolMiniMascara = True Then
      Print #Fnum, "MY=" & frmMini.Top
    Else
      Print #Fnum, "MY=" & MusicMp3.Top
    End If
    
    Print #Fnum, "Volume=" & CInt(MusicMp3.imgNormal(16).Top)
     
     If bolMiniMascara = True Then '// Si esta la minimascara
        Print #Fnum, "Mini=true"
     Else
        Print #Fnum, "Mini=false"
     End If

     If OpcionesMusic.NoAlteraR = True Then
       intClave = 0
     ElseIf OpcionesMusic.Mosaico = True Then
           intClave = 1
         ElseIf OpcionesMusic.Centrar = True Then
               intClave = 2
             Else
               intClave = 3
             End If
  
    Print #Fnum, "Wallpaper=" & intClave '// Estilo del Wallpaper
     
      If OpcionesMusic.Proporcional = True Then '// check proporcional
        Print #Fnum, "Proportional=true"
      Else
        Print #Fnum, "Proportional=false"
      End If
      
      If OpcionesMusic.Directorio = True Then '// check habilitar menu en explorer
        Print #Fnum, "Directory=true"
      Else
        Print #Fnum, "Directory=false"
      End If
      
      If Trim(OpcionesMusic.Language) = "" Then '// lenguaje
        Print #Fnum, "Language=Spanish"
      Else
        Print #Fnum, "Language=" & OpcionesMusic.Language
      End If
      
      If Trim(OpcionesMusic.MP3FILE) = True Then '// mp3 files
        Print #Fnum, "PlayMP3=True"
      Else
        Print #Fnum, "PlayMP3=False"
      End If
      
      If Trim(OpcionesMusic.WMAFILE) = True Then '// wma files
        Print #Fnum, "PlayWMA=True"
      Else
        Print #Fnum, "PlayWMA=False"
      End If
      
      If Trim(OpcionesMusic.WAVFILE) = True Then '// wav files
        Print #Fnum, "PlayWAV=True"
      Else
        Print #Fnum, "PlayWAV=False"
      End If

      
    Print #Fnum, "Alpha=" & OpcionesMusic.Alpha  '// cantidad de trasparencia
      
      If OpcionesMusic.SiempreTop = True Then
        Print #Fnum, "AlwaysTop=true"
      Else
        Print #Fnum, "AlwaysTop=false"
      End If
      
      If PlayerTrayIcon.Previous = True Then   '// systray icon Previous
        Print #Fnum, "SysTrayPrevious=true"
      Else
        Print #Fnum, "SysTrayPrevious=false"
      End If
      
      If PlayerTrayIcon.Play = True Then   '// systray icon Play
        Print #Fnum, "SysTrayPlay=true"
      Else
        Print #Fnum, "SysTrayPlay=false"
      End If
      
      If PlayerTrayIcon.Pause = True Then   '// systray icon Pause
        Print #Fnum, "SysTrayPause=true"
      Else
        Print #Fnum, "SysTrayPause=false"
      End If
      
      If PlayerTrayIcon.Stop = True Then   '// systray icon stop
        Print #Fnum, "SysTrayStop=true"
      Else
        Print #Fnum, "SysTrayStop=false"
      End If
      
      If PlayerTrayIcon.Next = True Then    '// systray icon next
        Print #Fnum, "SysTrayNext=true"
      Else
        Print #Fnum, "SysTrayNext=false"
      End If
      
      If frmPopUp.mnuIntro.Checked = True Then  '// Seleccionado intro
        Print #Fnum, "Intro=True"
      Else
        Print #Fnum, "Intro=False"
      End If
      
      If frmPopUp.mnuSilencio.Checked = True Then '// seleccionado mute
        Print #Fnum, "Mute=True"
      Else
        Print #Fnum, "Mute=False"
      End If
     
      If frmPopUp.mnuRepetir.Checked = True Then '// seleccionado repetir
        Print #Fnum, "Repeat=True"
      Else
        Print #Fnum, "Repeat=False"
      End If
    
      If frmPopUp.mnuAleatorioTodaColec.Checked = True Then 'Seleccionado aleatorio en toda la coleccion
        Print #Fnum, "RandomizeCollection=True"
      Else
        Print #Fnum, "RandomizeCollection=False"
      End If
    
      If frmPopUp.mnuAleatorioActAlbum.Checked = True Then '// Seleccionado aleatorio actual album
        Print #Fnum, "RandomizeAlbum=True"
      Else
        Print #Fnum, "RandomizeAlbum=False"
      End If
    
    Print #Fnum, "AlbumPlaying=" & intActiveAlbum  '// Album Reproduciendo
    Print #Fnum, "TrackNumber=" & MusicMp3.ListaRep.ListIndex '// Track Playing
    Print #Fnum, ""
      
    
    '// Seccion para almecenar los albums actuales reproduciendo
    Print #Fnum, "[Albums]"
     For i = 1 To TotalAlbumS
       Print #Fnum, "Album_" & i & "=" & MusicMp3.picAlbum(i).ToolTipText
     Next i
    Close Fnum
    
Exit Sub
Bitch:

End Sub

Public Function Path_Exe(Opcion As Sel_Option) As String
  On Error Resume Next
  Dim strRuta As String
   strRuta = App.Path
   If Right(strRuta, 1) <> "\" Then strRuta = strRuta & "\"
   If Opcion = 0 Then Path_Exe = strRuta
   If Opcion = 1 Then Path_Exe = strRuta & "MMp3Player\Skins\" & strSkinSeleccionado & "\"
End Function

Public Sub Always_on_Top()
 Const flag As Long = SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOSIZE
  If OpcionesMusic.SiempreTop = True Then
    If bolMiniMascara = True Then
      SetWindowPos frmMini.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
      SetWindowPos MusicMp3.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    End If
  Else
    If bolMiniMascara = True Then
      SetWindowPos frmMini.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    Else
      SetWindowPos MusicMp3.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    End If
  End If
End Sub

Public Sub Change_Mask(MiniMask As Boolean)
 On Error Resume Next
 Dim FormLeft As Long, formTop As Long
If MiniMask = True Then
  bolMiniMascara = True
  MusicMp3.Visible = False
  '//------------------------//
  FormLeft = MusicMp3.left + (MusicMp3.imgNormal(12).left * Screen.TwipsPerPixelX)
  FormLeft = FormLeft - (frmMini.picNormal(5).left * Screen.TwipsPerPixelX) + (frmMini.picNormal(5).ScaleWidth * Screen.TwipsPerPixelX)
  frmMini.left = FormLeft
  
  formTop = MusicMp3.Top + (MusicMp3.imgNormal(12).Top * Screen.TwipsPerPixelY)
  formTop = formTop - (frmMini.picNormal(5).Top * Screen.TwipsPerPixelY)
  
  frmMini.Top = formTop
  frmMini.Visible = True
  Always_on_Top
  Make_Transparent frmMini.hwnd, OpcionesMusic.Alpha
  If MusicMp3.bolToyBuscando = True Then
    MusicMp3.Timer_Texto.Enabled = False
    MusicMp3.Scroll_Text arryLanguage(68), frmMini.picScroll
  Else
    MusicMp3.Timer_Texto.Enabled = False
    MusicMp3.Scroll_Text ScrollText, frmMini.picScroll
  End If
Else
   frmMini.Visible = False
   
   FormLeft = frmMini.left + (frmMini.picNormal(5).left * Screen.TwipsPerPixelX)
   FormLeft = FormLeft - (MusicMp3.imgNormal(12).left * Screen.TwipsPerPixelX) - (MusicMp3.imgNormal(12).ScaleWidth * Screen.TwipsPerPixelX)
   MusicMp3.left = FormLeft
   
   formTop = frmMini.Top + (frmMini.picNormal(5).Top * Screen.TwipsPerPixelY)
   formTop = formTop - (MusicMp3.imgNormal(12).Top * Screen.TwipsPerPixelY)

   MusicMp3.Top = formTop
   MusicMp3.Visible = True
   bolMiniMascara = False
   Always_on_Top
   Make_Transparent MusicMp3.hwnd, OpcionesMusic.Alpha
   MusicMp3.Timer_Texto.Enabled = False
   MusicMp3.Scroll_Text ScrollText, MusicMp3.picScroll
End If

End Sub

'+----------------------------------------------------------------------------------------+
'|             TRASPARENCIA                                                               |
'+----------------------------------------------------------------------------------------+

Sub Make_Transparent(LHwnD As Long, Porcentaje As Integer)
 On Error GoTo Hell
  '// procedimento para hacer transparente en porcentaje los formularios
  '// parametros
  '// [LHwnD] -> Manejador para a kual aplikar el efekto
  '// [Porcentaje] -> pus que va ser el ...che porcentaje
  
  '// only work with win 2000 and later
  
  Dim OSV As OSVersionInfo
    
  '/* Get OS compatability flag
  OSV.OSVSize = Len(OSV)
  If GetVersionEx(OSV) <> 1 Then Exit Sub
       
  If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
  If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
    Call SetWindowLong(LHwnD, GWL_EXSTYLE, GetWindowLong(LHwnD, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(LHwnD, 0, (Porcentaje * 255) / 100, LWA_ALPHA)
  End If
Exit Sub
Hell:
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  PROCEDIMIENTO PARA ARRASTRAR EL FORMULARIO SOLO DEKLARARLO EN MOUSE DOWN             |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub FormDrag(TheForm As Form)
  ReleaseCapture
  Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

