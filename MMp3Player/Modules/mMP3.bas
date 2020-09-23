Attribute VB_Name = "mMP3"


Option Explicit
Public Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public L As Integer
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000

Public Declare Function SetFilePointer Lib "Kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetEndOfFile Lib "Kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1
Public Const ZERO = 0


Public Type VBRinfo
  VBRrate As String
  VBRlength As String
End Type

Public Type ID3v1Tag
  id As String * 3
  Title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type


Public Type MPEGInfo
  BITRATE As String
  CHANNELS As String
  COPYRIGHT As String
  CRC As String
  EMPHASIS As String
  FREQ As String
  LAYER As String
  LENGTH As String
  MPEG As String
  ORIGINAL As String
  SIZE As String
  TRACKNUM As String
  Title As String
  Artist As String
  Album As String
  Year As String
  Comment As String
  Genre As Variant
  LYRICS As String
End Type

Private strPathFile As String
Public HasLyrics3Tag As Boolean
Public HasID3v1Tag As Boolean
Private LSZ As String
Private copyMPEGinfo As MPEGInfo
Private s As String
Private posLyrics As Long
'// funcion para cargar las caracteristikas del archivo

Private Mp3Length As Long

Public Function Load_MPEGInfo(FileName As String) As MPEGInfo
 
 strPathFile = FileName

 copyMPEGinfo.Album = ""
 copyMPEGinfo.Artist = ""
 copyMPEGinfo.BITRATE = ""
 copyMPEGinfo.CHANNELS = ""
 copyMPEGinfo.Comment = ""
 copyMPEGinfo.COPYRIGHT = ""
 copyMPEGinfo.CRC = ""
 copyMPEGinfo.EMPHASIS = ""
 copyMPEGinfo.FREQ = ""
 copyMPEGinfo.Genre = ""
 copyMPEGinfo.LAYER = ""
 copyMPEGinfo.LENGTH = ""
 copyMPEGinfo.LYRICS = ""
 copyMPEGinfo.MPEG = ""
 copyMPEGinfo.ORIGINAL = ""
 copyMPEGinfo.SIZE = ""
 copyMPEGinfo.Title = ""
 copyMPEGinfo.TRACKNUM = ""
 copyMPEGinfo.Year = ""


'// obtener tags
 Read_Tags
 
 Load_MPEGInfo.Album = Trim(copyMPEGinfo.Album)
 Load_MPEGInfo.Artist = Trim(copyMPEGinfo.Artist)
 Load_MPEGInfo.BITRATE = Trim(copyMPEGinfo.BITRATE)
 Load_MPEGInfo.CHANNELS = Trim(copyMPEGinfo.CHANNELS)
 Load_MPEGInfo.Comment = Trim(copyMPEGinfo.Comment)
 Load_MPEGInfo.COPYRIGHT = Trim(copyMPEGinfo.COPYRIGHT)
 Load_MPEGInfo.CRC = Trim(copyMPEGinfo.CRC)
 Load_MPEGInfo.EMPHASIS = Trim(copyMPEGinfo.EMPHASIS)
 Load_MPEGInfo.FREQ = Trim(copyMPEGinfo.FREQ)
 Load_MPEGInfo.Genre = Trim(copyMPEGinfo.Genre)
 Load_MPEGInfo.LAYER = Trim(copyMPEGinfo.LAYER)
 Load_MPEGInfo.LENGTH = Trim(copyMPEGinfo.LENGTH)
 Load_MPEGInfo.LYRICS = Trim(copyMPEGinfo.LYRICS)
 Load_MPEGInfo.MPEG = Trim(copyMPEGinfo.MPEG)
 Load_MPEGInfo.ORIGINAL = Trim(copyMPEGinfo.ORIGINAL)
 Load_MPEGInfo.SIZE = Trim(copyMPEGinfo.SIZE)
 Load_MPEGInfo.Title = Trim(copyMPEGinfo.Title)
 Load_MPEGInfo.TRACKNUM = Trim(copyMPEGinfo.TRACKNUM)
 Load_MPEGInfo.Year = Trim(copyMPEGinfo.Year)

End Function

Private Function Read_Tags()

On Error GoTo errorhandler

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' use the filename to get ID3 info
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngFilesize As Long
    Dim fn As Integer
    Dim lngHeaderPosition As Long
    Dim Tag1 As ID3v1Tag
    Dim Tag2 As String
    Dim strID3v2 As String * 3
    Dim LyricEndID As String * 6
    Dim sFileExt As String
    Dim i As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    fn = FreeFile
    
    Open strPathFile For Binary As #fn                      'Open the file so we can read it
    lngFilesize = LOF(fn)                                   'Size of the file, in bytes
  sFileExt = UCase$(Right$(strPathFile, 3))

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check for an ID3v1 tag
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
    'ID3v1 tag
        
        Get #fn, lngFilesize - 127, Tag1.id
        
        If Tag1.id = "TAG" Then 'If "TAG" is present, then we have a valid ID3v1 tag and will extract all available ID3v1 info from the file
            Get #fn, , Tag1.Title   'Always limited to 30 characters
            Get #fn, , Tag1.Artist  'Always limited to 30 characters
            Get #fn, , Tag1.Album   'Always limited to 30 characters
            Get #fn, , Tag1.Year    'Always limited to 4 characters
            Get #fn, , Tag1.Comment 'Always limited to 30 characters
            Get #fn, , Tag1.Genre   'Always limited to 1 byte (?)
            
            HasID3v1Tag = True
            'Populate the form with the ID3v1 info
                copyMPEGinfo.Title = Replace(Trim(Tag1.Title), Chr(0), "")
                copyMPEGinfo.Artist = Replace(Trim(Tag1.Artist), Chr(0), "")
                copyMPEGinfo.Album = Replace(Trim(Tag1.Album), Chr(0), "")
                copyMPEGinfo.Year = Replace(Trim(Tag1.Year), Chr(0), "")
                copyMPEGinfo.Comment = Trim(Replace(Trim(Tag1.Comment), Chr(0), ""))
                copyMPEGinfo.Genre = Tag1.Genre + 1
        Else

          HasID3v1Tag = False
          
        End If
        
       'lyrics3 tag
   If HasID3v1Tag = True Then
     Get #fn, lngFilesize - 136, LyricEndID 'look for a lyrics 3 tag
   Else
     Get #fn, lngFilesize - 8, LyricEndID 'look for a lyrics 3 tag
   End If

           'lyrics3 tag?
   If LyricEndID = "LYRICS" Then 'got one, go get it
      HasLyrics3Tag = GetLyrics3Tag()
   Else
      HasLyrics3Tag = False
   End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Close
    getMPEGInfo
    
    
    Exit Function
        
errorhandler:
    err.Clear
    Close
    Resume Next
End Function


'// sakar los valores del mpeginfo

Private Sub getMPEGInfo()
  Dim Buf As String * 4096
  Dim infoStr As String * 3
  Dim lpVBRinfo As VBRinfo
  Dim tmpByte As Byte
  Dim tmpNum As Byte
  Dim i As Integer
  Dim designator As Byte
  Dim baseFreq As Single
  Dim vbrBytes As Long
  Dim HH As Long, MM As Long, SS As Long
  Dim tmp As String
  Dim R As Double
  
  Open strPathFile For Binary As #1
    Get #1, 1, Buf
  Close #1
  
  For i = 1 To 4092
    If Asc(Mid(Buf, i, 1)) = &HFF Then
      tmpByte = Asc(Mid(Buf, i + 1, 1))
      If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
        Exit For
      End If
    End If
  Next i
  If i = 4093 Then
    Exit Sub
  Else
    infoStr = Mid(Buf, i + 1, 3)
    'Getting info from 2nd byte(MPEG,Layer type and CRC)
    tmpByte = Asc(Mid(infoStr, 1, 1))
    
    'Getting CRC info
    If ((tmpByte Mod 16) Mod 2) = 0 Then
      copyMPEGinfo.CRC = "Yes"
    Else
      copyMPEGinfo.CRC = "No"
    End If
    
    'Getting MPEG type info
    If Between(tmpByte, &HF2, &HF7) Then
      copyMPEGinfo.MPEG = "MPEG 2.0"
      designator = 1
    Else
      copyMPEGinfo.MPEG = "MPEG 1.0"
      designator = 2
    End If
    
    'Getting layer info
    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
      copyMPEGinfo.LAYER = "layer 3"
    Else
      If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
        copyMPEGinfo.LAYER = "layer 2"
      Else
        copyMPEGinfo.LAYER = "layer 1"
      End If
    End If
    
    'Getting info from 3rd byte(Frequency, Bit-rate)
    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    'Getting frequency info
    If Between(tmpByte Mod 16, &H0, &H3) Then
      baseFreq = 22.05
    Else
      If Between(tmpByte Mod 16, &H4, &H7) Then
        baseFreq = 24
      Else
        baseFreq = 16
      End If
    End If
    copyMPEGinfo.FREQ = baseFreq * designator & " Hz"
    
    'Getting Bit-rate
    tmpNum = tmpByte \ 16 Mod 16
    If designator = 1 Then
      If tmpNum < &H8 Then
        copyMPEGinfo.BITRATE = tmpNum * 8
      Else
        copyMPEGinfo.BITRATE = 64 + (tmpNum - 8) * 16
      End If
    Else
      If tmpNum <= &H5 Then
        copyMPEGinfo.BITRATE = (tmpNum + 3) * 8
      Else
        If tmpNum <= &H9 Then
          copyMPEGinfo.BITRATE = 64 + (tmpNum - 5) * 16
        Else
          If tmpNum <= &HD Then
            copyMPEGinfo.BITRATE = 128 + (tmpNum - 9) * 32
          Else
            copyMPEGinfo.BITRATE = 320
          End If
        End If
      End If
    End If
    Mp3Length = FileLen(strPathFile) \ (Val(copyMPEGinfo.BITRATE) / 8) \ 1000
    If Mid(Buf, i + 36, 4) = "Xing" Then
      vbrBytes = Asc(Mid(Buf, i + 45, 1)) * &H10000
      vbrBytes = vbrBytes + (Asc(Mid(Buf, i + 46, 1)) * &H100&)
      vbrBytes = vbrBytes + Asc(Mid(Buf, i + 47, 1))
      GetVBRrate strPathFile, vbrBytes, lpVBRinfo
      copyMPEGinfo.BITRATE = lpVBRinfo.VBRrate
      
       '/* time
       HH = CLng(lpVBRinfo.VBRlength) \ 3600  '/* hours
       MM = CLng(lpVBRinfo.VBRlength) \ 60 Mod 60 '/* Minutes
       SS = CLng(lpVBRinfo.VBRlength) Mod 60  '/* Seconds
 
        If HH > 0 Then tmp = Format$(HH, "00:")
       copyMPEGinfo.LENGTH = Trim(tmp & Format$(MM, "00:") & Format$(SS, "00"))

    Else
      copyMPEGinfo.BITRATE = copyMPEGinfo.BITRATE & " Kbps"
      '/* time
       HH = Mp3Length \ 3600  '/* hours
       MM = Mp3Length \ 60 Mod 60 '/* Minutes
       SS = Mp3Length Mod 60  '/* Seconds
 
        If HH > 0 Then tmp = Format$(HH, "00:")
       copyMPEGinfo.LENGTH = Trim(tmp & Format$(MM, "00:") & Format$(SS, "00"))
      
    End If
    
    'Getting info from 4th byte(Original, Emphasis, Copyright, Channels)
    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16
    
    
    'Getting Copyright bit
    If tmpNum \ 8 = 1 Then
      copyMPEGinfo.COPYRIGHT = " Yes"
      tmpNum = tmpNum - 8
    Else
      copyMPEGinfo.COPYRIGHT = " No"
    End If
    
    'Getting Original bit
    If (tmpNum \ 4) Mod 2 Then
      copyMPEGinfo.ORIGINAL = " Yes"
      tmpNum = tmpNum - 4
    Else
      copyMPEGinfo.ORIGINAL = " No"
    End If
    
    'Getting Emphasis bit
    Select Case tmpNum
      Case 0
        copyMPEGinfo.EMPHASIS = " None"
      Case 1
        copyMPEGinfo.EMPHASIS = " 50/15 microsec"
      Case 2
        copyMPEGinfo.EMPHASIS = " invalid"
      Case 3
        copyMPEGinfo.EMPHASIS = " CITT j. 17"
    End Select
    
    'Getting channel info
    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
      Case 0
        copyMPEGinfo.CHANNELS = " Stereo"
      Case 1
        copyMPEGinfo.CHANNELS = " Joint Stereo"
      Case 2
        copyMPEGinfo.CHANNELS = " 2 Channel"
      Case 3
        copyMPEGinfo.CHANNELS = " Mono"
    End Select
  End If
  

   R = FileLen(strPathFile) / 1024
   copyMPEGinfo.SIZE = CLng(R / 1024 * 100) / 100 & " MB"
   'copyMPEGinfo.SIZE = CLng(R / 1024 / 1024 * 100) / 100 & "  GB"
   'copyMPEGinfo.SIZE = CLng(R) & "  KB"

End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBRinfo)
  Dim i As Long
  Dim ok As Boolean

  i = 0
  byteRead = byteRead - &H39
  Do
    If byteRead > 0 Then
      i = i + 1
      byteRead = byteRead - 38 - Deljivo(i)
    Else
      ok = True
    End If
  Loop Until ok
  lpVBRinfo.VBRlength = Trim(Str(i))
  lpVBRinfo.VBRrate = Trim(Str(Int(8 * FileLen(lpMP3File) / (1000 * i)))) & " Kbit (VBR)"
End Sub

Private Function Deljivo(ByVal Num As Long) As Byte
  If Num Mod 3 = 0 Then
    Deljivo = 1
  Else
    Deljivo = 0
  End If
End Function

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean
  If accNum >= accDown And accNum <= accUp Then
    Between = True
  Else
    Between = False
  End If
End Function


Public Function WriteTag(strFileName As String, cID3v1Tag As ID3v1Tag, strLyrics) As Boolean
   
   Dim WholeTag As String
   Dim TagSize As String * 6
   Dim Position As Long
   Dim MoveMP3Tag As Boolean
   Dim fn As Integer
   Dim UseOldInfo As Boolean
   Dim strAuthor As String
   Dim NewLyr As Boolean
   WholeTag = ""
   
   On Error GoTo Hell
      
   '// you can add more tags(fields) at file for exemple i add LYR -> Lyrics
   '// (but you can add more ex: AUT -> author)
  If strLyrics <> "" Then NewLyr = True
   
   '//Author field
'   strAuthor = "Raul Martinez"
'   If strAuthor <> "" Then NewLyr = True
   
   'build the tag
   If NewLyr = True Then
      WholeTag = "LYRICSBEGIN"
      If strLyrics <> "" Then
         WholeTag = WholeTag & "LYR" & Format(Len(strLyrics), "00000") & strLyrics
      End If
      
'      If strAuthor <> "" Then
'         WholeTag = WholeTag & "AUT" & Format(Len(strAuthor), "00000") & strAuthor
'      End If
      
      TagSize = Format(Len(WholeTag), "000000")
      
      'append the end identifier
      WholeTag = WholeTag & TagSize & "LYRICS200"

   End If
   
   'prepare for writing
   fn = FreeFile
   Open strFileName For Binary As #fn
   
   If HasID3v1Tag = True Then
      'set to just before the current id3 tag.
      Position = LOF(fn) - 127
   Else
      Position = LOF(fn) + 1
   End If
   
   'if there is a Lyrics3 tag, then go back to the beginning of the old one
   If HasLyrics3Tag Then Position = posLyrics
   

   ' write the lyrics3 tag if there is one...
   If WholeTag <> "" Then
      Put #fn, Position, WholeTag
      Position = Seek(fn)
   End If
   
   'write the id3tag
   
   Put #fn, Position, cID3v1Tag
   
   'set the last byte of the file
   Position = Seek(fn) - 1
   Close
   'make sure this is the end of the file which is needed if this tag is smaller than the old tag.
   SetFileLength strFileName, Position
   Exit Function
Hell:
End Function

Public Function GetLyrics3Tag() As Boolean
 On Error GoTo Hell
   Dim Position As Long
   Dim FieldData As String   '// save Text of field
   Dim FieldID As String * 3 '// save Field ex: LYR, ALB, AUT... etc...
   Dim LengthField As String
   Dim fn As Integer
   Dim SIZE As Long '// size of File
   Dim TagType As String * 9 '//save end tag lyrics
   Dim Byte11Buffer As String * 11 '// save LYRICSBEGIN
   Dim Byte5Buffer As String * 5   '// Length of Field
   Dim lEndLyr As Long
   
   '//reset size of tag
   LengthField = "000000"
   
   '//open the file
   fn = FreeFile
   Open strPathFile For Binary As #fn
   
   'get filesize
   SIZE = LOF(fn)
   
   'get the tag END
   If HasID3v1Tag = True Then
     'get the tag END [ (127 -> ID3v1Tag) + (9  -> LYRICS200) ] = 136
     lEndLyr = SIZE - 136
     Get #fn, lEndLyr, TagType
   Else
     lEndLyr = SIZE - 8
     Get #fn, lEndLyr, TagType
   End If
   
   'if tag is valid then
   If TagType = "LYRICS200" Then
      
      'get the size of the tag  ( 136 - 6 ) = 142
      Get #fn, lEndLyr - 6, LengthField
      
      'set the position to the first byte of Lyrics
      Position = lEndLyr - 6 - Val(LengthField)
      
      'get the beginning of tag
      Get #fn, Position, Byte11Buffer
      'save beginning lyrics tag
       posLyrics = Position
       
      If Byte11Buffer <> "LYRICSBEGIN" Then
         'invalid Lyrics3 version 2 tag! we don't support version 1...
         Close
         Exit Function
      End If
      
      'first field ( Position + LYRICSBEGIN )
      Position = Position + 11
      
      'keep getting fields until we get to the end of the tag
      Do Until Position >= lEndLyr - 9
      
         'the field id -> LYR-ALB-AUT. etc...
         Get #fn, Position, FieldID
         
         'the size of the field
         Get #fn, Position + 3, Byte5Buffer
         LengthField = Val(Byte5Buffer)
                  
         'make room for the data
         FieldData = Space(Val(LengthField))
         'get the data
         Get #fn, Position + 8, FieldData
         'and fill the approprate field
         Select Case FieldID
            Case "LYR" '// Lyrics
               copyMPEGinfo.LYRICS = Trim$(FieldData)
'            Case "AUT" '// Author
'               MsgBox "Copyright : " & Trim$(FieldData)
         End Select
         
         'now set the postion to the beginning of the next field
         Position = Position + 8 + Val(LengthField)
      Loop
      'set the flag
      GetLyrics3Tag = True
   End If
   Exit Function
Hell:
   'close all open files
   Close
End Function

Private Sub SetFileLength(strFileName As String, ByVal NewLength As Long)

   'Will cut the length of a file to the length specified.
   
   Dim hFile As Long
   Dim L As Long
   Dim lpSecurity As SECURITY_ATTRIBUTES
   'if file is smaller than or equal to requsted length, exit.
   If FileLen(strFileName) <= NewLength Then Exit Sub
   'open the file
   hFile = CreateFile(strFileName, GENERIC_WRITE, ZERO, lpSecurity, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
   'if file not open exit
   If hFile = -1 Then Exit Sub
   'seek to position
   L = SetFilePointer(hFile, NewLength, ZERO, ZERO)
   'and mark here as end of file
   SetEndOfFile hFile
   'close the file
   L = CloseHandle(hFile)
End Sub



