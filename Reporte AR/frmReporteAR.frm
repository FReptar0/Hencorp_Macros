VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReporteAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Macro AR"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmReporteAR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Rago Fecha (DATERMIT)"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118554625
         CurrentDate     =   41775
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118554625
         CurrentDate     =   41775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton OptError 
         Caption         =   "Con error"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton OptSinError 
         Caption         =   "Sin Error"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formato"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton optExcel 
         Caption         =   "Formato Excel"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkMostrar 
      Caption         =   "Mostar reporte al terminar la impresión"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdExporta 
      Caption         =   "Generar Reporte"
      Default         =   -1  'True
      Height          =   615
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Exporta información de master actual a archivo Excel"
      Top             =   3360
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdExporta 
      Left            =   1920
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReporteAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mSession As AccpacSession
Private mSessMgr As AccpacSessionMgr ' this is useful if you need to use the AccpacMeter

Private Sub ExportaExcel()
On Error GoTo Error
Dim ssql As String
Dim ssql2 As String
Dim ssql3 As String

Dim rsDocumentos As Recordset
Set rsDocumentos = New Recordset

Dim appExcel As Excel.Application
Dim wrkExcel As Excel.Workbook
Dim shtExcel As Excel.Worksheet

Dim intCol As Long
Dim intRow As Long
Dim strNombre As String

Dim NoPag As Long
Dim RutaRespaldo As String
NoPag = 1

FechaA = IIf(Month(DTPicker1.Value) < 10, "0" & Month(DTPicker1.Value), Month(DTPicker1.Value)) & "/" & IIf(Day(DTPicker1.Value) < 10, "0" & Day(DTPicker1.Value), Day(DTPicker1.Value)) & "/" & Year(DTPicker1.Value)
Fechab = IIf(Month(DTPicker2.Value) < 10, "0" & Month(DTPicker2.Value), Month(DTPicker2.Value)) & "/" & IIf(Day(DTPicker2.Value) < 10, "0" & Day(DTPicker2.Value), Day(DTPicker2.Value)) & "/" & Year(DTPicker2.Value)

If OptSinError = True Then
    ssql = "select a.LOTE as lotecreado,a.ASIENTO as asientocreado,a.CODEPAYM as CodigoPago, * from AR_RA a "
    ssql = ssql & "left outer join AR_MR b on a.CNTBTCH=b.CNTBTCH and a.CNTITEM=b.CNTITEM and a.COMPANYID=b.COMPANYID "
    ssql = ssql & " where a.COMPANYID='" & Empresa & "' and  a.ESTADO='Completo' and (a.DATERMIT   between '" & FechaA & "' AND '" & Fechab & "' )order by a.LOTE,a.ASIENTO"
Else
    ssql = "select a.LOTE as lotecreado,a.ASIENTO as asientocreado,a.CODEPAYM as CodigoPago,  * from AR_RA a "
    ssql = ssql & "left outer join AR_MR b on a.CNTBTCH=b.CNTBTCH and a.CNTITEM=b.CNTITEM and a.COMPANYID=b.COMPANYID "

    If OptError = True Then
        ssql = ssql & " where a.COMPANYID='" & Empresa & "' and  a.ESTADO='Error' and (a.DATERMIT   between '" & FechaA & "' AND '" & Fechab & "' )order by a.LOTE,a.ASIENTO"
    ElseIf OptTodos Then
        ssql = ssql & " WHERE (a.DATERMIT  between '" & FechaA & "' AND '" & Fechab & "') and a.COMPANYID='" & Empresa & "' order by a.CNTBTCH,a.CNTITEM"
    End If
End If
    'MsgBox Fechab
Debug.Print ssql
rsDocumentos.Open ssql, cnDB, adOpenForwardOnly, adLockReadOnly
If rsDocumentos.EOF = False And rsDocumentos.BOF = False Then
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        
        cdExporta.Filter = "Archivo Excel (.xlsx)|*.xlsx"
        cdExporta.ShowSave
        If Trim(cdExporta.FileName) <> "" Then
            RutaRespaldo = (cdExporta.FileName)
        End If
        
        On Error Resume Next
        Set appExcel = GetObject(, "Excel.Application")
        If appExcel Is Nothing Then
            Set appExcel = CreateObject("Excel.Application")
            If appExcel Is Nothing Then
                MsgBox "Cannot Open Microsoft Excel For Export", vbCritical
                Exit Sub
            End If
        End If
    
        intNewSheets = appExcel.SheetsInNewWorkbook
        appExcel.SheetsInNewWorkbook = 1
        Set wrkExcel = appExcel.Workbooks.Add
        appExcel.SheetsInNewWorkbook = intNewSheets
        Set shtExcel = wrkExcel.Sheets(1)
       
        With shtExcel
    
    ' EncabezadoMSD Macro
    '**************************************************************************************************************************************************************
    '**************************************************************************************************************************************************************
    '**************************************************************************************************************************************************************
 
        .Columns("A:A").ColumnWidth = 10
        .Columns("B:B").ColumnWidth = 10
        .Columns("C:C").ColumnWidth = 10
        .Columns("D:D").ColumnWidth = 40
        .Columns("E:E").ColumnWidth = 10
        .Columns("F:F").ColumnWidth = 10
        .Columns("G:G").ColumnWidth = 10
        .Columns("H:H").ColumnWidth = 10
        .Columns("i:i").ColumnWidth = 10
        .Columns("j:j").ColumnWidth = 10
        .Columns("k:k").ColumnWidth = 40
        .Columns("l:l").ColumnWidth = 40
        .Columns("m:m").ColumnWidth = 8
        .Columns("n:n").ColumnWidth = 10
        .Columns("o:o").ColumnWidth = 10
        .Columns("p:p").ColumnWidth = 30
        .Columns("q:q").ColumnWidth = 10
        .Columns("r:r").ColumnWidth = 10
        .Columns("s:s").ColumnWidth = 40
        .Columns("t:t").ColumnWidth = 10
        .Columns("u:u").ColumnWidth = 40
        .Columns("v:v").ColumnWidth = 40
        .Columns("w:w").ColumnWidth = 10

        .Range("C1").FormulaR1C1 = Empresa & "-Reporte Macro AR"
        .Range("F1").FormulaR1C1 = "De Fecha (DATERMIT): " & (FechaA)
        .Range("i1").FormulaR1C1 = "A Fecha (DATERMIT):  " & (Fechab)
  
        .Range("A3").FormulaR1C1 = "Lote Creado"
        .Range("B3").FormulaR1C1 = "Asiento Creado"
        .Range("C3").FormulaR1C1 = "Estado"
        .Range("D3").FormulaR1C1 = "Resultado"
        .Range("E3").FormulaR1C1 = "Fecha"
        .Range("F3").FormulaR1C1 = "Hora"
        .Range("G3").FormulaR1C1 = "COMPANYID"
        .Range("H3").FormulaR1C1 = "CNTBTCH"
        .Range("i3").FormulaR1C1 = "CNTITEM"
        .Range("J3").FormulaR1C1 = "CNTLINE"
        .Range("K3").FormulaR1C1 = "TEXTRMIT"
        .Range("L3").FormulaR1C1 = "TEXTPAYOR"
        .Range("M3").FormulaR1C1 = "IDBANK"
        .Range("N3").FormulaR1C1 = "CODECURN"
        .Range("O3").FormulaR1C1 = "CODEPAYM"
        .Range("P3").FormulaR1C1 = "DATEDEP"
        .Range("Q3").FormulaR1C1 = "IDRMIT"
        .Range("R3").FormulaR1C1 = "DATERMIT"
        .Range("S3").FormulaR1C1 = "BATCHDESC"
        .Range("T3").FormulaR1C1 = "DATEBATCH"
        .Range("U3").FormulaR1C1 = "DATEPOST"
        .Range("V3").FormulaR1C1 = "TXTRMITREF"
        .Range("W3").FormulaR1C1 = "IDACCT"
        .Range("X3").FormulaR1C1 = "GLREF"
        .Range("Y3").FormulaR1C1 = "GLDESC"
        .Range("Z3").FormulaR1C1 = "AMTDISTTC"




        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 8

        '**************************************************************************************************************************************************************
            j = 4 'FILA INICIAL
            rsDocumentos.MoveFirst
            Do While rsDocumentos.EOF = False
                For i = 1 To 26
                    Select Case i
                        Case 1
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!lotecreado)
                        Case 2
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!asientocreado)
                        Case 3
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!estado)
                        Case 4
                            .Cells(j, i).Value = "'" & IIf(IsNull(Trim(rsDocumentos!resultado)), "sin procesar", Trim(rsDocumentos!resultado))
                        Case 5
                            .Cells(j, i).Value = "'" & FormatoFecha(Trim(rsDocumentos!fecha))
                        Case 6
                            .Cells(j, i).Value = "'" & Left(Trim(rsDocumentos!hora), 8)
                        Case 7
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CompanyID)
                        Case 8
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CNTBTCH)
                        Case 9
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CNTITEM)
                        Case 10
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CNTLINE)
                        Case 11
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!TEXTRMIT)
                        Case 12
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!TEXTPAYOR)
                        Case 13
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!IDBANK)
                        Case 14
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CODECURN)
                        Case 15
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!CodigoPago)
                        Case 16
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!DATEDEP)
                        Case 17
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!IDRMIT)
                        Case 18
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!DATERMIT)
                        Case 19
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!BATCHDESC)
                        Case 20
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!DATEBATCH)
                        Case 21
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!DATEPOST)
                        Case 22
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!TXTRMITREF)
                        Case 23
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!IDACCT)
                        Case 24
                            .Cells(j, i).Value = "'" & Trim(rsDocumentos!GLREF)
                        Case 25
                            .Cells(j, i).Value = Trim(rsDocumentos!GLDESC)
                        Case 26
                            .Cells(j, i).Value = Trim(rsDocumentos!AMTDISTTC)
                        Case Else
                            'CASE ELSE
                    End Select
                Next
                rsDocumentos.MoveNext
                j = j + 1
                X = j
            Loop
'**************************************************************************************************************************************************************
   
        appExcel.ActiveWorkbook.SaveAs FileName:= _
            RutaRespaldo, FileFormat _
            :=xlOpenXMLWorkbook, CreateBackup:=False
            
        If chkMostrar.Value = 1 Then
            appExcel.DisplayFullScreen = True
            appExcel.Visible = True
            appExcel.WindowState = xlNormal
        Else
            appExcel.ActiveWorkbook.Close savechanges:=False
        End If

   End With
   Set shtExcel = Nothing
   Set wrkExcel = Nothing
   Set appExcel = Nothing
   MsgBox "Se ha creado con éxito  el archivo " & RutaRespaldo, vbInformation + vbOKOnly, "Éxito"
Else
   MsgBox "No existen datos para generar el reporte'.", vbCritical + vbOKOnly, "Sin Información"
End If
rsDocumentos.Close
Set rsDocumentos = Nothing
Screen.MousePointer = vbNormal
Me.Enabled = True
Exit Sub
Error:
If Err.Number <> 32755 Then
    MsgBox Err.Description & " " & Err.Number
    Select Case Err.Number
        Case cdlCancel
    End Select
End If
End Sub

Private Sub cmdExporta_Click()
On Error GoTo Error
Screen.MousePointer = vbHourglass
'cdExporta.Filter = "Archivo Excel (.xlsx)|*.xlsx"
'cdExporta.ShowSave

'If Trim(cdExporta.FileName) <> "" Then
    Call ExportaExcel '(cdExporta.FileName)
'End If
Exit Sub
Error:
If Err.Number <> 32755 Then
    MsgBox Err.Description & " " & Err.Number
    Select Case Err.Number
        Case cdlCancel
    End Select
End If
End Sub
Private Sub cmdSalir_Click()
Desconectar
Unload Me
End Sub
Private Sub Form_Load()
Dim lSignonID As Long
Dim rptName As String


lSignonID = 0 ' MUST be initialized to 0 since you don't have a signon ID yet
Set mSessMgr = New AccpacSessionMgr
With mSessMgr
    .AppID = "AS"
    .AppVersion = "70A"
    .ProgramName = "as1000"
    .ServerName = "" ' empty string if running on local computer
    .CreateSession "", lSignonID, mSession ' first argument is the object handle (if you don't have one, pass "")
End With ' mSessMgr

Call mSession.GetSignonInfo(usuario, Empresa, EmpresaNombre)
Caption = Empresa & "-" & Me.Caption
'MsgBox "debug Form_Load usuario " & usuario
'MsgBox "debug Form_Load Empresa " & Empresa
'MsgBox "debug Form_Load EmpresaNombre " & EmpresaNombre
'MsgBox "debug Form_Load se va a conectar a sql"
conecta
'MsgBox "debug Form_Load paso  sql"
DTPicker1.Value = Date
DTPicker2.Value = Date
OptTodos.Value = True
End Sub

Private Function FormatoFecha(intFecha As String) As String
Dim dia As Double
Dim mes As Double
Dim anio As Double
Dim strDia As String
Dim strMes As String
Dim strAnio As String

strAnio = Left(intFecha, 4)
strMes = Mid(intFecha, 5, 2)
strDia = Right(intFecha, 2)

'Fecha = CDate(dia & "/" & mes & "/" & anio)
FormatoFecha = strMes & "/" & strDia & "/" & strAnio
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

