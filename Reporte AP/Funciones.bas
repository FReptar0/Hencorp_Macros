Attribute VB_Name = "Funciones"
Public cnDB1 As Connection
Public cnDB As New ADODB.Connection
Public BoolExito As Boolean


Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
  
  
Private Declare Function GetWindowLong Lib "user32" _
                                Alias "GetWindowLongA" _
                                (ByVal hwnd As Long, _
                                ByVal nIndex As Long) _
                                As Long
                                  
Private Declare Function SetWindowLong Lib "user32" _
                                Alias "SetWindowLongA" _
                                (ByVal hwnd As Long, _
                                ByVal nIndex As Long, _
                                ByVal dwNewLong As Long) _
                                As Long
                                  
Private Declare Function SetWindowPos Lib "user32" _
                                (ByVal hwnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal cx As Long, _
                                ByVal cy As Long, _
                                ByVal wFlags As Long) _
                                As Long
  
Public Const WS_MAXIMIZEBOX = &H10000
  
Private Const GWL_STYLE = (-16)
  
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
  
Const FLAG As Long = SWP_NOACTIVATE Or SWP_FRAMECHANGED _
                     Or SWP_NOSIZE Or SWP_NOMOVE
Public usuario As String
Public Empresa As String
Public EmpresaNombre As String
Public Sub Maximizar(ByVal Hwnd_Ventana As Long, _
                     ByVal Flags As Long, _
                     Optional ByVal Accion As Boolean = True)
      
    Dim ret        As Long
      
    ' Obtiene el estilo actual
    ret = GetWindowLong(Hwnd_Ventana, GWL_STYLE)
      
    ' asigna los flags al estilo actual dependiendo de la acción
    If Accion Then
        ret = ret Or Flags
    Else
        ret = ret And Not Flags
    End If
      
    ' aplica el nuevo estilo
    SetWindowLong Hwnd_Ventana, GWL_STYLE, ret
      
    ' si la ventana está visible ...
    If IsWindowVisible(Hwnd_Ventana) Then
          
        ' Es necesario ya que si no solo se verá el camio si se repinta la ventana
        SetWindowPos Hwnd_Ventana, 0, 0, 0, 0, 0, FLAG
    End If
End Sub
Public Function ChecaNullInt(cadena As Variant) As String
ChecaNullInt = IIf(IsNull(cadena), 0, cadena)
End Function

Public Function ChecaNullStr(cadena As Variant) As String
ChecaNullStr = IIf(IsNull(cadena), "", cadena)
End Function
Public Function Formato(cadena As Currency) As Currency
Formato = Format(Round(cadena, 2), "$###,##0.00")
End Function

Public Function conecta()

Set cnDB = New Connection

Dim sINIFile As String
Dim sPassword As String
Dim sServer As String
Dim sUser As String
n = False
sINIFile = App.Path & "\config.ini"

'leer el nombre del archivo ini

dbinformacion = sGetINI(sINIFile, "settings", "dbinformacion", "?")
If dbinformacion = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. L"
   Exit Function
End If
sServer = sGetINI(sINIFile, "settings", "Server", "?")
If sServer = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. B"
   Exit Function
End If
sUser = sGetINI(sINIFile, "settings", "User", "?")
If sUser = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. C"
   Exit Function
End If
sPassword = sGetINI(sINIFile, "settings", "Password", "?")
If sPassword = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. D"
   Exit Function
End If
'sPassword = Desencriptar(sPassword)
'MsgBox "debug Form_Load sServer " & sServer
'MsgBox "debug Form_Load dbinformacion " & dbinformacion
'MsgBox "debug Form_Load sUser " & sUser
'MsgBox "debug Form_Load sPassword " & sPassword
On Error GoTo ErrorSQL
cnDB.Open "Driver={SQL Server};Server=" & sServer & ";Database=" & dbinformacion & ";UID=" & sUser & ";PWD=" & sPassword
Exit Function
ErrorSQL:
    MsgBox "connection error" & Err.Number & " - " & Err.Description, vbCritical, "Error SQL"
End
Exit Function
End Function
Public Function Desconectar()
On Error Resume Next
cnDB.Close
Set cnDB = Nothing
cnDB2.Close
Set cnDB2 = Nothing
cnDB3.Close
Set cnDB3 = Nothing
End Function
Public Function Encriptar(cadena) As String
Dim CRYPT_KEY As String
Dim i As Double
Dim j As Double
Dim ls_enctext As String

CRYPT_KEY = "$#C@%&#G%@J#P%#C%#A%#J%#E%#D%#L%#O%#%#%#%&*"

j = Len(cadena)
For i = 1 To j
    ls_enctext = ls_enctext & Mid(CRYPT_KEY, (i Mod 10) + 1, 1)
    ls_enctext = ls_enctext & CStr(Chr(255 - Asc(Mid(cadena, i, 1))))
Next i

Encriptar = ls_enctext

End Function
Public Function Desencriptar(cadena_encriptada) As String
Dim i As Double
Dim j As Double
Dim ls_encchar As String
Dim ls_temp As String
Dim ls_unasstr As String
ls_unasstr = "** Encryption Error **"
Dim lb_ok As Boolean
Dim CRYPT_KEY As String

CRYPT_KEY = "$#C@%&#G%@J#P%#C%#A%#J%#E%#D%#L%#O%#%#%#%&*"
lb_ok = True
j = Len(cadena_encriptada)

If Not j Mod 2 = 1 Then
   ls_temp = ""
   For i = 2 To (j + 1) Step 2
      ls_encchar = Mid(cadena_encriptada, i - 1, 1)
      If Mid(CRYPT_KEY, i / 2 Mod 10 + 1, 1) <> ls_encchar Then
        lb_ok = False
        Exit Function
      End If
      ls_encchar = Mid(cadena_encriptada, i, 1)
      ls_temp = ls_temp & CStr(Chr(255 - Asc(ls_encchar)))
   Next
End If

If lb_ok Then ls_unasstr = ls_temp

Desencriptar = ls_unasstr

End Function

