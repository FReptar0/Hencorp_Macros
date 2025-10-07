VERSION 5.00
Begin VB.Form frmGL2100 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Macro genera lotes GL v1"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6240
   Icon            =   "frmGL2100.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   795
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenera 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Genera lotes GL"
      Height          =   795
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmGL2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xPointer As Integer
Private Sub cmdGenera_Click()
cmdGenera.Enabled = False
Screen.MousePointer = vbHourglass
AsientoGL
cmdsalir_Click
End Sub

Private Sub cmdsalir_Click()
Unload Me
Unload frmMenu
End Sub

Private Sub Form_Load()

On Error GoTo ErrorLoad
Dim lSignonID As Long
lSignonID = 0 ' MUST be initialized to 0 since you don't have a signon ID yet
Set mSessMgr = New AccpacSessionMgr
With mSessMgr
 .AppID = "GL"
 .AppVersion = "70A"
 .ProgramName = "GL2100"
 .ServerName = "" ' empty string if running on local computer
 .CreateSession "", lSignonID, mSession ' first argument is the object handle (if you don't have one, pass "")
End With ' mSessMgr
Call mSession.GetSignonInfo(usuario, Gempresa, empresanombre)
sDatabase = Gempresa
'MsgBox "debug Form_Load se va a conectar a sql"
Call conecta
'MsgBox "debug Form_Load se coencto a sql"
'Call conectaSage(Gempresa)
Me.Caption = Gempresa & " - " & Me.Caption & " *"
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Rbitacora) Then
   fso.CreateFolder Rbitacora
End If
'MsgBox "debug Form_Load"
Exit Sub
ErrorLoad:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ErrorLoad"
    bitacora = Rbitacora & "logMacroGL.txt"
    Set v1 = CreateObject("Scripting.FileSystemObject")
    Set v2 = v1.OpenTextFile(bitacora, 8, True)
    
    v2.WriteLine "************************************************************************************************"
    v2.WriteLine "Error Macro GL" & "  " & Date & "  " & Time & "  " & usuario & "  " & "  -   " & "  ErrorLoad "
    v2.WriteLine "No. Error: " & Err.Number & ".  Description: " & Err.Description
    v2.WriteLine "************************************************************************************************"
    v2.Close
    End
Exit Sub
End Sub
