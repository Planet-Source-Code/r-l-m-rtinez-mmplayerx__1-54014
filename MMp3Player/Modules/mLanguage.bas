Attribute VB_Name = "mLanguage"
Option Explicit

Public arryLanguage() As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  IDIOMA                                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Sub Load_Language(strLang As String)
 On Error Resume Next
 Dim Linenr
 Dim InputData
 Dim i As Integer
 Dim strRuta As String, strTemp As String
 ReDim arryLanguage(83)
 arryLanguage(1) = "    New Find"
 arryLanguage(2) = "    Covert Front"
    arryLanguage(3) = "    Change ListRep/Cover Front"
    arryLanguage(4) = "    Put Cover Front as Wallpaper"
    arryLanguage(5) = "    Maximize Cover front"
 arryLanguage(6) = "    Browsers "
    arryLanguage(7) = "    Explore Files"
    arryLanguage(8) = "    Albums Browser"
    arryLanguage(9) = "    Edit Track(s) Tag"
   arryLanguage(10) = "    Karaoke Function"
 arryLanguage(11) = "    Players Controls"
 arryLanguage(12) = "  Volume"
 arryLanguage(13) = "+   Increase Volume"
 arryLanguage(14) = "-   Decrease Volume"
 arryLanguage(15) = "Z   Previous Track"
 arryLanguage(16) = "X   Play"
 arryLanguage(17) = "C   Pause"
 arryLanguage(18) = "V   Stop"
 arryLanguage(19) = "B   Next Track"
 arryLanguage(20) = "<   Previous Album/Folder"
 arryLanguage(21) = ">   Next Album/Folder"
 arryLanguage(22) = "I   Intro 10 seg."
 arryLanguage(23) = "R   Repeat Track"
 arryLanguage(24) = "S   Mute"
 arryLanguage(25) = "  Randomize"
 arryLanguage(26) = "Q   Current Album/Folder"
 arryLanguage(27) = "W   Whole Albums"
 arryLanguage(28) = "A   Skip Backward 5 sec."
 arryLanguage(29) = "D   Skip Forward 5 sec."
 arryLanguage(30) = "    Options "
 arryLanguage(31) = "    Skins "
 arryLanguage(32) = " << Skins Browser >>"
 arryLanguage(33) = "    Alpha Mode"
 arryLanguage(34) = " Custom"
 arryLanguage(35) = "    About..."
 arryLanguage(36) = "    Minimize"
 arryLanguage(37) = "    Change Mask"
 arryLanguage(38) = "    Exit"
 '// form options
 arryLanguage(39) = " Wallpaper"
 arryLanguage(40) = " Skins"
 arryLanguage(41) = " Alpha"
 arryLanguage(42) = " Application"
 arryLanguage(43) = " Player"
 arryLanguage(44) = " Enable right click menu in drives and directories"
 arryLanguage(45) = " Options Wallpaper"
 arryLanguage(46) = " No Alter."
 arryLanguage(47) = " Strech."
 arryLanguage(48) = " Center."
 arryLanguage(49) = " Tile."
 arryLanguage(50) = " Proportional."
 arryLanguage(51) = " Alpha (Only win 2000 or later.)"
 arryLanguage(52) = " Alpha: "
 arryLanguage(53) = " Language"
 arryLanguage(54) = " Application"
 arryLanguage(55) = " Always on Top."
 arryLanguage(56) = " Show Splash Screen."
 arryLanguage(57) = " Multiple Instances."
 arryLanguage(58) = " Play Files"
 arryLanguage(59) = " .mp3 Files."
 arryLanguage(60) = " .wma Files."
 arryLanguage(61) = " .wav Files."
 arryLanguage(62) = " Show in system tray"
 arryLanguage(63) = " Previous Track Icon."
 arryLanguage(64) = " Play Icon."
 arryLanguage(65) = " Pause Icon."
 arryLanguage(66) = " Stop Icon."
 arryLanguage(67) = " Next Track Icon."
 arryLanguage(68) = " [ Searching... ]"
 arryLanguage(69) = " [ No Mp3's files found ]"
 arryLanguage(70) = " Ok"
 arryLanguage(71) = " Apply"
 arryLanguage(72) = " Cancel"
 arryLanguage(73) = " error reading file"
 arryLanguage(74) = " Current Cover Front"
 arryLanguage(75) = " Searching files in:"
 arryLanguage(76) = " Select a directory for search."
 '// form lyrics
 arryLanguage(77) = " [ No lyrics found ]"
 '// form tags
 arryLanguage(78) = " Select All"
 arryLanguage(79) = " Tags"
 arryLanguage(80) = " Lyrics"
 arryLanguage(81) = " Add"
 arryLanguage(82) = " Undo"
 arryLanguage(83) = " Multiple track are selected. Select checkboxes to apply changes to ALL selected tracks"
   
  strRuta = Path_Exe(PathExe) & "MMp3Player\Language\" & strLang & ".lng"
   If Dir(strRuta) <> "" Then
    Open strRuta For Input As #2

     Linenr = -1
     Do While Not EOF(2)
       Line Input #2, InputData
        i = i + 1
        If i > 84 Then Exit Do
        If Trim(InputData) <> "" Or Len(Trim(InputData)) > 3 Then
          Linenr = Linenr + 1
          strTemp = left(arryLanguage(Linenr), 1)
          strTemp = Trim(strTemp) & "  " & InputData
          arryLanguage(Linenr) = strTemp
          If Linenr > 38 Then arryLanguage(Linenr) = Trim(strTemp)
        End If
     Loop
    Close #2
   End If
 With frmPopUp
   .mnuNuevaBusqueda.Caption = arryLanguage(1)
   .mnuCFront.Caption = arryLanguage(2)
   .mnuCambiarListaCaratula.Caption = arryLanguage(3)
   MusicMp3.imgNormal(10).ToolTipText = Trim(arryLanguage(3))
   .mnuWallpapper.Caption = arryLanguage(4)
   .mnuMCaratula.Caption = arryLanguage(5)
   .mnuBrowsers.Caption = arryLanguage(6)
   .mnuExplorar.Caption = arryLanguage(7)
   .mnuExpAlbum.Caption = arryLanguage(8)
   .mnuTagEditor.Caption = arryLanguage(9)
   .mnuLyrics.Caption = arryLanguage(10)
   .mnuControles.Caption = arryLanguage(11)
   .mnuVolumen.Caption = arryLanguage(12)
   .mnuSubirVolumen.Caption = arryLanguage(13)
   .mnuBajarVolumen.Caption = " " & arryLanguage(14)
   .mnuTrackAnterior.Caption = arryLanguage(15)
   MusicMp3.imgNormal(0).ToolTipText = Trim(Right(arryLanguage(15), Len(arryLanguage(15)) - 1))
   frmMini.picNormal(0).ToolTipText = Trim(Right(arryLanguage(15), Len(arryLanguage(15)) - 1))
   .mnuReproducir.Caption = arryLanguage(16)
   MusicMp3.imgNormal(1).ToolTipText = Trim(Right(arryLanguage(16), Len(arryLanguage(16)) - 1))
   frmMini.picNormal(1).ToolTipText = Trim(Right(arryLanguage(16), Len(arryLanguage(16)) - 1))
   .mnuPausa.Caption = arryLanguage(17)
   MusicMp3.imgNormal(2).ToolTipText = Trim(Right(arryLanguage(17), Len(arryLanguage(17)) - 1))
   frmMini.picNormal(2).ToolTipText = Trim(Right(arryLanguage(17), Len(arryLanguage(17)) - 1))
   .mnuDetener.Caption = arryLanguage(18)
   MusicMp3.imgNormal(3).ToolTipText = Trim(Right(arryLanguage(18), Len(arryLanguage(18)) - 1))
   frmMini.picNormal(3).ToolTipText = Trim(Right(arryLanguage(18), Len(arryLanguage(18)) - 1))
   .mnuSigTrack.Caption = arryLanguage(19)
   MusicMp3.imgNormal(4).ToolTipText = Trim(Right(arryLanguage(19), Len(arryLanguage(19)) - 1))
   frmMini.picNormal(4).ToolTipText = Trim(Right(arryLanguage(19), Len(arryLanguage(19)) - 1))
   .mnuAnteriorAlbum.Caption = arryLanguage(20)
   MusicMp3.imgNormal(9).ToolTipText = Trim(Right(arryLanguage(20), Len(arryLanguage(20)) - 1))
   .mnuSigAlbum.Caption = arryLanguage(21)
   MusicMp3.imgNormal(11).ToolTipText = Trim(Right(arryLanguage(21), Len(arryLanguage(21)) - 1))
   .mnuIntro.Caption = arryLanguage(22)
   MusicMp3.imgNormal(5).ToolTipText = Trim(Right(arryLanguage(22), Len(arryLanguage(22)) - 1))
   .mnuSilencio.Caption = arryLanguage(24)
   MusicMp3.imgNormal(6).ToolTipText = Trim(Right(arryLanguage(24), Len(arryLanguage(24)) - 1))
   .mnuRepetir.Caption = arryLanguage(23)
   MusicMp3.imgNormal(7).ToolTipText = Trim(Right(arryLanguage(23), Len(arryLanguage(23)) - 1))
   .mnuOrdenAleatorio.Caption = arryLanguage(25)
   MusicMp3.imgNormal(8).ToolTipText = Trim(arryLanguage(25))
   .mnuAleatorioActAlbum.Caption = arryLanguage(26)
   .mnuAleatorioTodaColec.Caption = arryLanguage(27)
   .mnuAtras5Seg.Caption = arryLanguage(28)
   .mnuAdelante5Seg.Caption = arryLanguage(29)
   .mnuOpciones.Caption = arryLanguage(30)
   .mnuSkins.Caption = arryLanguage(31)
   .mnuExpSkins.Caption = arryLanguage(32)
   .mnuWOpacity.Caption = arryLanguage(33)
   .mnuAlphaPer.Caption = Trim(arryLanguage(34))
   .mnuAcercaDe.Caption = arryLanguage(35)
   .mnuMinimizar.Caption = arryLanguage(36)
   .mnuCambiarMascaras.Caption = arryLanguage(37)
   .mnuSalir.Caption = arryLanguage(38)
   
   MusicMp3.imgNormal(12).ToolTipText = Trim(arryLanguage(36))
   MusicMp3.imgNormal(13).ToolTipText = Trim(arryLanguage(37))
   frmMini.picNormal(5).ToolTipText = Trim(arryLanguage(36))
   MusicMp3.imgNormal(14).ToolTipText = Trim(arryLanguage(38))
   frmMini.picNormal(6).ToolTipText = Trim(arryLanguage(38))
   
   If bolOpcionesShow = True Then Load_Language_Options
   
   If bolDirectoriosShow = True Then
      frmDirectorios.Caption = arryLanguage(8) & " [ " & TotalAlbumS & " Albums ]"
      frmDirectorios.mnuExpArchivos.Caption = "  " & Trim(arryLanguage(7))
      frmDirectorios.mnuEditTags.Caption = "  " & Trim(arryLanguage(9))
      frmDirectorios.mnuPlay.Caption = "  " & Trim(Right(arryLanguage(16), Len(arryLanguage(16)) - 1))
   End If

   If bolCaratulaShow = True Then frmCaratula.Caption = arryLanguage(73)
   If bolAcercaShow = True Then frmAcerca.Caption = arryLanguage(35)
   If bolTagsShow = True Then Load_Language_Tags
        '//change language at systray icons
     If PlayerTrayIcon.Previous = True Then CambiarIcono MusicMp3.txtSTIcon(0).hwnd, MusicMp3.ImageList.ListImages(1).ExtractIcon.Handle, MusicMp3.imgNormal(0).ToolTipText & " - MMp3Player"
     
     If PlayerTrayIcon.Play = True Then CambiarIcono MusicMp3.txtSTIcon(1).hwnd, MusicMp3.ImageList.ListImages(2).ExtractIcon.Handle, MusicMp3.imgNormal(1).ToolTipText & " - MMp3Player"
     
     If PlayerTrayIcon.Pause = True Then CambiarIcono MusicMp3.txtSTIcon(2).hwnd, MusicMp3.ImageList.ListImages(3).ExtractIcon.Handle, MusicMp3.imgNormal(2).ToolTipText & " - MMp3Player"
     
     If PlayerTrayIcon.Stop = True Then CambiarIcono MusicMp3.txtSTIcon(3).hwnd, MusicMp3.ImageList.ListImages(4).ExtractIcon.Handle, MusicMp3.imgNormal(3).ToolTipText & " - MMp3Player"
     
     If PlayerTrayIcon.Next = True Then CambiarIcono MusicMp3.txtSTIcon(4).hwnd, MusicMp3.ImageList.ListImages(5).ExtractIcon.Handle, MusicMp3.imgNormal(4).ToolTipText & " - MMp3Player"
  
 End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Language_Options()
 With frmOpciones
   .Caption = Trim(arryLanguage(30))
   .TabStrip1.Tabs(1).Caption = arryLanguage(39)
   .TabStrip1.Tabs(2).Caption = arryLanguage(40)
   .TabStrip1.Tabs(3).Caption = arryLanguage(41)
   .TabStrip1.Tabs(4).Caption = arryLanguage(42)
   .TabStrip1.Tabs(5).Caption = arryLanguage(43)
   .chkDir.Caption = arryLanguage(44)
   '//walpaper
   .Frame1.Caption = arryLanguage(45)
   .optWallpaper(0).Caption = arryLanguage(46)
   .optWallpaper(3).Caption = arryLanguage(47)
   .optWallpaper(2).Caption = arryLanguage(48)
   .optWallpaper(1).Caption = arryLanguage(49)
   .chkProporcional.Caption = arryLanguage(50)
   '//alpha
   .Frame3.Caption = arryLanguage(51)
   .Label1(2).Caption = arryLanguage(52)
   '//language
   .Frame2.Caption = arryLanguage(53)
   '// application
   .Frame5.Caption = arryLanguage(54)
   .chkSiemTop.Caption = arryLanguage(55)
   .chkSplash.Caption = arryLanguage(56)
   .chkinstancias.Caption = arryLanguage(57)
   '// format files
   .Frame6.Caption = arryLanguage(58)
   .chkMP3.Caption = arryLanguage(59)
   .chkWMA.Caption = arryLanguage(60)
   .chkWAV.Caption = arryLanguage(61)
   '// system tray icon
   
   .Frame7.Caption = arryLanguage(62)
   .chkPIcon(0).Caption = arryLanguage(63)
   .chkPIcon(1).Caption = arryLanguage(64)
   .chkPIcon(2).Caption = arryLanguage(65)
   .chkPIcon(3).Caption = arryLanguage(66)
   .chkPIcon(4).Caption = arryLanguage(67)
   
   '//buttons
   .cmdOk.Caption = arryLanguage(70)
   .cmdApply.Caption = arryLanguage(71)
   .cmdCancel.Caption = arryLanguage(72)
  End With
End Sub

Sub Load_Language_Tags()
 With frmTags
    .cmdOk.Caption = arryLanguage(70)
    .cmdCancel.Caption = arryLanguage(72)
    .cmdApply.Caption = arryLanguage(71)
    .cmdSelAll.Caption = arryLanguage(78)
    .TabStrip.Tabs(1).Caption = arryLanguage(79)
    .TabStrip.Tabs(2).Caption = arryLanguage(80)
    .cmdAdd.Caption = arryLanguage(81)
    .cmdUndo.Caption = arryLanguage(82)
 End With
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

