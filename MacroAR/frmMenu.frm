VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Macro AR"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAutomatico 
      Caption         =   "Automatico"
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Manual"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAutomatico_Click()
Call conecta
Dim ssql1 As String
Dim rsEmpresa As Recordset
Set rsEmpresa = New ADODB.Recordset
ssql1 = "select distinct COMPANYID from AR_RA "
rsEmpresa.Open ssql1, cnDB, adOpenStatic, adLockReadOnly
If rsEmpresa.EOF = False And rsEmpresa.BOF = False Then
    rsEmpresa.MoveFirst
    Do While Not rsEmpresa.EOF
        Gempresa = Trim(rsEmpresa!CompanyID)
        AsientoAR
        rsEmpresa.MoveNext
    Loop
End If
Set rsEmpresa = Nothing
Unload Me
End Sub

Private Sub cmdManual_Click()
frmMenu.Hide
frmAR3100.Show
End Sub

Private Sub Form_Load()
Manual = Command
'MsgBox Manual
If Trim(Manual) = "" Then ' si no tiene argumento se abre manual
    cmdManual_Click
Else
    cmdAutomatico_Click
End If
End Sub
