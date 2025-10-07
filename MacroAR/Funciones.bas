Attribute VB_Name = "Funciones"

Public cnDB As New ADODB.Connection
Public cnDB2 As New ADODB.Connection
Public filePath As String
Public v1 As Variant
Public v2 As Variant
Public Error As Boolean
Public AsentarLote As Boolean

Public sDatabase As String
Public RutaCuerpo As String
Public RutaResponse As String
Public cartID As String
Public clientID As String
Public UserSage As String
Public PassSage As String
Public Rbitacora As String
Public PorAsentar As String
Public bitacora As String

Public mSession As AccpacSession
Public mSessMgr As AccpacSessionMgr ' this is useful if you need to use the AccpacMeter
Public usuario As String
Public Gempresa As String
Public empresanombre As String
Public RutaAPI As String
Public Exempt As String
Public dbinformacion As String
Public fecha As String
Public hora As String
Public Manual As String

Public Function conecta()
Set cnDB = New Connection
Dim sPassword As String
Dim sServer As String
Dim sUser As String
Dim sINIFile As String
n = False

 sINIFile = App.Path & "\Config.ini"
 sServer = sGetINI(sINIFile, "settings", "server", "?")
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
   UserSage = sGetINI(sINIFile, "settings", "userSage", "?")
If UserSage = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. I"
   Exit Function
End If
 PassSage = sGetINI(sINIFile, "settings", "PassSage", "?")
If PassSage = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. J"
   Exit Function
End If
 Rbitacora = sGetINI(sINIFile, "settings", "LOG", "?")
If Rbitacora = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. K"
   Exit Function
End If
 dbinformacion = sGetINI(sINIFile, "settings", "dbinformacion", "?")
If dbinformacion = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. L"
   Exit Function
End If
 PorAsentar = sGetINI(sINIFile, "settings", "AsientaAR", "?")
If PorAsentar = "?" Then
   MsgBox "The INI file doesn't exist, please call your System Administrator. m"
   Exit Function
End If
On Error GoTo ErrorSQL
cnDB.Open "Driver={SQL Server};Server=" & sServer & ";Database=" & dbinformacion & ";UID=" & sUser & ";PWD=" & sPassword
Exit Function
ErrorSQL:
MsgBox "connection error" & Err.Number & " - " & Err.Description, vbCritical, "Error SQL"
bitacora = Rbitacora & "logMacroAP.txt"
Set v1 = CreateObject("Scripting.FileSystemObject")
Set v2 = v1.OpenTextFile(bitacora, 8, True)

v2.WriteLine "************************************************************************************************"
v2.WriteLine "Error Taxes" & "  " & Date & "  " & Time & "  " & usuario & "  " & "  -   " & "  ErrorSQL "
v2.WriteLine "No. Error: " & Err.Number & ".  Description: " & Err.Description
v2.WriteLine "************************************************************************************************"
v2.Close
End

Exit Function
End Function
Public Function conectaSage(sDatabase As String)
Set cnDB2 = New Connection
Dim sPassword As String
Dim sServer As String
Dim sUser As String
Dim sINIFile As String
n = False

 sINIFile = App.Path & "\Config.ini"
 sServer = sGetINI(sINIFile, "settings", "server", "?")
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
On Error GoTo ErrorSQL
cnDB2.Open "Driver={SQL Server};Server=" & sServer & ";Database=" & sDatabase & ";UID=" & sUser & ";PWD=" & sPassword
Exit Function
ErrorSQL:
'MsgBox "connection error" & Err.Number & " - " & Err.Description, vbCritical, "Error SQL"
bitacora = Rbitacora & "logMacroAP.txt"
Set v1 = CreateObject("Scripting.FileSystemObject")
Set v2 = v1.OpenTextFile(bitacora, 8, True)

v2.WriteLine "************************************************************************************************"
v2.WriteLine "Error Taxes" & "  " & Date & "  " & Time & "  " & usuario & "  " & "  -   " & "  ErrorSQL "
v2.WriteLine "No. Error: " & Err.Number & ".  Description: " & Err.Description
v2.WriteLine "************************************************************************************************"
v2.Close
End

Exit Function
End Function

Public Sub Pause(ByVal nSecond As Single)
   
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer

      dummy = DoEvents()
      ' if we cross midnight, back up one day
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub

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
Public Function existebanco(banco As String) As Boolean
Dim ssql As String
Dim rsValidaBanco As Recordset
Set rsValidaBanco = New Recordset
ssql = "select BANK from BKACCT where INACTIVE=0 and BANK='" & banco & "'"
rsValidaBanco.Open ssql, cnDB2, adOpenForwardOnly, adLockReadOnly
If rsValidaBanco.EOF = False And rsValidaBanco.BOF = False Then
    existebanco = True
Else
    existebanco = False
End If
rsValidaBanco.Close
Set rsValidaBanco = Nothing
End Function
