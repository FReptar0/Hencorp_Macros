Attribute VB_Name = "IniFunc"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function sGetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
Dim sTemp As String * 256
Dim nLength As Integer

sTemp = Space$(256)
nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
sGetINI = Left$(sTemp, nLength)
End Function

Public Sub writeINI(sINIFile As String, sSection As String, sKey As String, sValue As String)
Dim n As Integer
Dim sTemp As String

sTemp = sValue
'reemplazar los caracteres CR/LF con espacios
For n = 1 To Len(sValue)
   If Mid$(svale, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then
      Mid$(sValue, n) = " "
   End If
Next n
      
n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)

End Sub

