Attribute VB_Name = "XPTheme"
Option Explicit

Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Function XPStyle(Optional AutoRestart As Boolean = True, Optional CreateNew As Boolean) As Boolean
 '// cargar ini controles necesario para que funcione el xp stylo
 InitCommonControls

On Error Resume Next
 Dim XML             As String
 Dim ManifestCheck   As String
 Dim strManifest     As String
 Dim FreeFileNo      As Integer

If AutoRestart = True Then CreateNew = False

'// poner el xmp en una string

XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & _
  "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & _
  "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & _
  "    processorArchitecture=""x86"" " & vbCrLf & _
  "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & _
  "    name=""EXENAME""/>" & vbCrLf & _
  "    <description>EXEDESCRIBTION</description>" & vbCrLf & _
  "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & _
  "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & _
  "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & _
  "         version=""6.0.0.0""" & vbCrLf & _
  "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & _
  "         language=""*""" & vbCrLf & _
  "         processorArchitecture=""x86""/>" & vbCrLf & _
  "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & _
  "</assembly>" & vbCrLf & "")

'// poner el nombre el archivo manifest
strManifest = Path_Exe(PathExe) & App.EXEName & ".exe.manifest"

'// Chekar los atributos
ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)

'// Si no se encuentra hacerlo... el archivo :D
If ManifestCheck = "" Or CreateNew = True Then
  '// Replazar la cadena "EXENAME" con el nombre del archivo exe
  XML = Replace(XML, "EXENAME", App.EXEName & ".exe")
  '// Replazar la version
  XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
  '// Replazar la ExeDescripcion con la valida
  XML = Replace(XML, "EXEDESCRIBTION", App.FileDescription)
  
  FreeFileNo = FreeFile
  
  '// si si se encontro borrarlo
  If ManifestCheck <> "" Then
    SetAttr strManifest, vbNormal
    Kill strManifest
  End If
  
  '// abrir el archivo
  Open strManifest For Binary As #(FreeFileNo)
     '// uses 'put' to set the file content.. note that 'put' (binary mode)
     '// is much faster than 'print'(output mode)
     Put #(FreeFileNo), , XML
  Close #(FreeFileNo)  '// cerrar archivo
  
  '// Poner nuevos atributos
  SetAttr strManifest, vbHidden + vbSystem
  
  If ManifestCheck = "" Then
    XPStyle = False
  Else
    XPStyle = True
  End If
  
  '// si no estaba y restar = true pus recargarlo
  If AutoRestart = True Then
    '// Recargar el programa de nuevo con el xp stilo
    Shell Path_Exe(PathExe) & App.EXEName & ".exe", vbNormalFocus
    End
  End If
  
Else  '// si ya existe el archivo manifest
  XPStyle = True
End If

End Function
