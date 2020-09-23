Attribute VB_Name = "mRegistry"
Option Explicit

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  explorador pequeño para directorios usado en nueva busqueda                           |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Type BrowseInfo   '// estructura para la api
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
   
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
 
  Private m_lngRetVal As Long
  
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|      OBTENER EL DIRECTORIO DEL WINDOWS
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Public Declare Function GetWindowsDirectoryA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| ACTUALIZAR AL ...CHE WINDOWS LO DEL FONDO DEL ESCRITORIO                              |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
  
  Const SPI_SETDESKWALLPAPER = 20
  Const SPIF_SENDWININICHANGE = &H2
  Const SPIF_UPDATEINIFILE = &H1


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| OBTENER EL AREA DEL ESCRITORIO SIN LA TASK BAR                                        |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function SystemGetWorkArea Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
  
  Public Const API_NULL_HANDLE = 0
  Public Const SPI_GETWORKAREA = 48
  
  Type RECT
    left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| Constants required for values in the keys
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Private Const REG_NONE As Long = 0                  ' No value type
  Private Const REG_SZ As Long = 1                    ' nul terminated string
  Private Const REG_EXPAND_SZ As Long = 2             ' nul terminated string w/enviornment var
  Private Const REG_BINARY As Long = 3                ' Free form binary
  Private Const REG_DWORD As Long = 4                 ' 32-bit number
  Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4   ' 32-bit number (same as REG_DWORD)
  Private Const REG_DWORD_BIG_ENDIAN As Long = 5      ' 32-bit number
  Private Const REG_LINK As Long = 6                  ' Symbolic Link (unicode)
  Private Const REG_MULTI_SZ As Long = 7              ' Multiple Unicode strings
  Private Const REG_RESOURCE_LIST As Long = 8         ' Resource list in the resource map
  Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9 ' Resource list in the hardware description
  Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
' Registry Specific Access Rights
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_ALL_ACCESS As Long = &H3F

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| Constants required for key locations in the registry
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Public Const HKEY_CLASSES_ROOT As Long = &H80000000
  Public Const HKEY_CURRENT_USER As Long = &H80000001
  Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Public Const HKEY_USERS As Long = &H80000003
  Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
  Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
  Public Const HKEY_DYN_DATA As Long = &H80000006

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
' Constants required for return values (Error code checking)
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Private Const ERROR_SUCCESS As Long = 0
  Private Const ERROR_ACCESS_DENIED As Long = 5
  Private Const ERROR_NO_MORE_ITEMS As Long = 259

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
' Open/Create constants
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Private Const REG_OPTION_NON_VOLATILE As Long = 0
  Private Const REG_OPTION_VOLATILE As Long = &H1

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
' Declarations required to access the Windows registry
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
  
  Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
  
  Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
  
  Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
             lpType As Long, lpData As Any, lpcbData As Long) As Long
  
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
             ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Dim PoniendoWallpaper As Boolean

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Function regDelete_Sub_Key(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String, _
                                  ByVal strRegSubKey As String)
    
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the sub key.  If it does not exist, then ignore it.
      m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
  
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean
    
  Dim lngKeyHandle As Long
  lngKeyHandle = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regDoes_Key_Exist = False
  Else
      regDoes_Key_Exist = True
  End If
  m_lngRetVal = RegCloseKey(lngKeyHandle)
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Function regQuery_A_Key(ByVal lngRootKey As Long, _
                               ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String) As Variant
    
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String
  lngKeyHandle = 0
  lngBufferSize = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Query the registry and determine the data type.
' --------------------------------------------------------------
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                         lngDataType, ByVal 0&, lngBufferSize)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Make the API call to query the registry based on the type
' of data.
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data (most common)
              ' Preload the receiving buffer area
              strBuffer = Space(lngBufferSize)
      
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                     ByVal strBuffer, lngBufferSize)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Strip out the string data
                  intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                  If intPosition > 0 Then
                      ' if we found one, then save everything up to that point
                      regQuery_A_Key = left(strBuffer, intPosition - 1)
                  Else
                      ' did not find one.  Save everything.
                      regQuery_A_Key = strBuffer
                  End If
              End If
              
         Case REG_DWORD:    ' Numeric data (Integer)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                     lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Save the captured data
                  regQuery_A_Key = lngBuffer
              End If
         
         Case Else:    ' unknown
              regQuery_A_Key = ""
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
    
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngKeyValue As Long
  Dim strKeyValue As String
  
' --------------------------------------------------------------
' Determine the type of data to be updated
' --------------------------------------------------------------
  If PoniendoWallpaper = True Then
     lngDataType = REG_SZ
  Else
    If IsNumeric(varRegData) Then
      lngDataType = REG_DWORD
    Else
      lngDataType = REG_SZ
    End If
  End If
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    
' --------------------------------------------------------------
' Update the sub key based on the data type
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data
              strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          ByVal strKeyValue, Len(strKeyValue))
                                   
         Case REG_DWORD:    ' numeric data
              lngKeyValue = CLng(varRegData)
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
                                   
  End Select
  
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)

  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Create the key.  If it already exist, ignore it.
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)

' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Function regDelete_A_Key(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean
    
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Preset to a failed delete
' --------------------------------------------------------------
  regDelete_A_Key = False
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the key
      m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
      
      ' If the value returned is equal zero then we have succeeded
      If m_lngRetVal = 0 Then regDelete_A_Key = True
      
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| CAMBIAR EL WALLPAPPER DEL ESCRITORIO                                                  |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub PoneRWallPapeR(Tipo As String)
On Error GoTo Hell
Dim NuevoPaper As String, strEstilo As String
 '// buskar la ruta en donde se guardara la imagen .bmp para el wallpaper
    NuevoPaper = DirectoriOWindowS & "MusicMp3.bmp"
    PoniendoWallpaper = True
    '// ponerla con los parametros deseados
    If Tipo = "Centro" Or Tipo = "Clear" Then
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"
    ElseIf Tipo = "Mosaico" Then
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "1"
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0"
    ElseIf Tipo = "Expandido" Then
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", "0"
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2"
    End If
    
    '// Actualizar al windows para que se muestre el wallpaper
    If Tipo = "Clear" Then
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, " ", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    Else
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, NuevoPaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    End If
   PoniendoWallpaper = False
Exit Sub
Hell:
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| CAMBIAR EL WALLPAPPER DEL ESCRITORIO  AL ORIGINAL                                     |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub PoneRWallPapeROriginaL()
 On Error GoTo Hell
  '//  Actualizar el wallpaper al originar al salir
    PoniendoWallpaper = True
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallpaper", OriginalTileWallpaper
       regCreate_Key_Value HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", OriginalWallpaperStyle
       SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, OriginalRutaWallpaper, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
    PoniendoWallpaper = False
 Exit Sub
Hell:
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'OBTENER EL DIRECTORIO DEL WINDOWS
Public Function DirectoriOWindowS() As String
 On Error GoTo Hell
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0) ' cuanto espacio tiene la ruta
   s = Space(i) 'poner los espacios
   'llamar ala api y almacenar la ruta en s
   Call GetWindowsDirectoryA(s, i)
    s = left$(s, i - 1) 'quitar el ultimo caracter
     If Len(s) > 0 Then
       If Right$(s, 1) <> "\" Then
         s = s + "\"
       End If
     Else
       s = "C:\WINDOWS\"
     End If
   DirectoriOWindowS = s
Exit Function
Hell:
End Function


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  FUNCION PARA EXPLORADOR DE DIRECTORIOS                                                |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Function Explorador_Para_Directorios(ByVal lngHwnd As Long, ByVal strMensaje As String) As String
    On Error GoTo Hell
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    '// Ajustar las propiedades de la api con la estructura apropiada
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strMensaje, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '// mostrar el Explorador...
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        '// Krear string con 260 spacion nullos para almacenar la direccion / path
        strPath = String(MAX_PATH, 0)

        '// Almacenar la direccion seleccionada en strPath
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        '// Liberar memoria
        Call CoTaskMemFree(lngIDList)

        '// Buskar el primer carakter nullo en la cadena
        '// Para obtener la direccion
        intNull = InStr(strPath, vbNullChar)
        '// Asegurarnos que este bien
        If intNull > 0 Then
            '// Poner ahora si el ...che path
            strPath = left(strPath, intNull - 1)
        End If
    End If

    '// regresar la bendita direccion Chet :P
    Explorador_Para_Directorios = strPath
 Exit Function
 
Hell:
    '// Retornar null en error
    Explorador_Para_Directorios = Empty

End Function

