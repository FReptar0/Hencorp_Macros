Attribute VB_Name = "MacroAR"
Public Sub AsientoAR()
Error = False
Dim loteCreado As String
Dim AsientoCreado As String
Dim lote As String
Dim Asiento As String
Dim ConError As Boolean
Dim mensajeerrorDet As String
Dim mensajeerrorHea As String
Dim UserID As String
Dim Password As String
Dim OrgID As String
Dim SessionDate As Date
Dim dbCmp As AccpacDBLink
Dim objSession As AccpacCOMAPI.AccpacSession
Dim objCompany As AccpacCOMAPI.AccpacDBLink
Dim ssql As String
Dim ssql2 As String
Dim DATEBUSDia As String
Dim DATEBUSMes As String
Dim DATEBUSAnio As String
Dim DATERMITDia As String
Dim DATERMITMes As String
Dim DATERMITAnio As String
Dim DATEDEPDia As String
Dim DATEDEPMes As String
Dim DATEDEPAnio As String

Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim loteant As String
Dim B As Integer
'    Set rsJournal_Headers = New Recordset
'    Set rsJournal_Details = New Recordset
Dim rsJournal_Details As Recordset
Dim rsJournal_Headers As Recordset
Set rsJournal_Headers = New ADODB.Recordset
Set rsJournal_Details = New ADODB.Recordset
UserID = UserSage '"test"
Password = PassSage '"test1"
OrgID = Gempresa '"geodat"
SessionDate = Date
On Error GoTo ErrorMacroAR
ConError = False
If usuario = "" Then
    usuario = UserID
End If
Dim objSignOn As AccpacSignonManager.AccpacSignonMgr
Set objSession = AccpacCOMAPI.AccpacSession

objSession.Init "", "AR", "AR3100", "70A"
objSession.Open UserID, Password, OrgID, SessionDate, 0, ""

If objSession.IsOpened Then
   Set objCompany = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
   Set dbCmp = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
End If


On Error GoTo ErrorMacroAR

' TODO: To increase efficiency, comment out any unused DB links.
Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
Set mDBLinkCmpRW = OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)

Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
Set mDBLinkSysRW = OpenDBLink(DBLINK_SYSTEM, DBLINK_FLG_READWRITE)

Dim temp As Boolean
Dim ARRECMAC1batch As AccpacCOMAPI.AccpacView
Dim ARRECMAC1batchFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0041", ARRECMAC1batch
Set ARRECMAC1batchFields = ARRECMAC1batch.Fields

Dim ARRECMAC1header As AccpacCOMAPI.AccpacView
Dim ARRECMAC1headerFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0042", ARRECMAC1header
Set ARRECMAC1headerFields = ARRECMAC1header.Fields

Dim ARRECMAC1detail1 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail1Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0044", ARRECMAC1detail1
Set ARRECMAC1detail1Fields = ARRECMAC1detail1.Fields

Dim ARRECMAC1detail2 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0045", ARRECMAC1detail2
Set ARRECMAC1detail2Fields = ARRECMAC1detail2.Fields

Dim ARRECMAC1detail3 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail3Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0043", ARRECMAC1detail3
Set ARRECMAC1detail3Fields = ARRECMAC1detail3.Fields

Dim ARRECMAC1detail4 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail4Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0061", ARRECMAC1detail4
Set ARRECMAC1detail4Fields = ARRECMAC1detail4.Fields

Dim ARRECMAC1detail5 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail5Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0406", ARRECMAC1detail5
Set ARRECMAC1detail5Fields = ARRECMAC1detail5.Fields

Dim ARRECMAC1detail6 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail6Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0170", ARRECMAC1detail6
Set ARRECMAC1detail6Fields = ARRECMAC1detail6.Fields

Dim ARRECMAC1detail7 As AccpacCOMAPI.AccpacView
Dim ARRECMAC1detail7Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0085", ARRECMAC1detail7
Set ARRECMAC1detail7Fields = ARRECMAC1detail7.Fields

ARRECMAC1batch.Compose Array(ARRECMAC1header)

ARRECMAC1header.Compose Array(ARRECMAC1batch, ARRECMAC1detail3, ARRECMAC1detail1, ARRECMAC1detail5, ARRECMAC1detail6, ARRECMAC1detail7)

ARRECMAC1detail1.Compose Array(ARRECMAC1header, ARRECMAC1detail2, ARRECMAC1detail4)

ARRECMAC1detail2.Compose Array(ARRECMAC1detail1)

ARRECMAC1detail3.Compose Array(ARRECMAC1header)

ARRECMAC1detail4.Compose Array(ARRECMAC1batch, ARRECMAC1header, ARRECMAC1detail3, ARRECMAC1detail1, ARRECMAC1detail2)

ARRECMAC1detail5.Compose Array(ARRECMAC1header)

ARRECMAC1detail6.Compose Array(ARRECMAC1header)

ARRECMAC1detail7.Compose Array(ARRECMAC1header)


Dim ARPAYMPOST2 As AccpacCOMAPI.AccpacView
Dim ARPAYMPOST2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AR0049", ARPAYMPOST2
Set ARPAYMPOST2Fields = ARPAYMPOST2.Fields


ARRECMAC1batch.RecordClear

ARRECMAC1batchFields("CODEPYMTYP").PutWithoutVerification ("CA")      ' Tipo de Lote

ARRECMAC1headerFields("CODEPYMTYP").PutWithoutVerification ("CA")     ' Tipo de Lote
ARRECMAC1detail3Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Tipo de Lote
ARRECMAC1detail1Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Tipo de Lote
ARRECMAC1detail2Fields("CODEPAYM").PutWithoutVerification ("CA")      ' Tipo de Lote
ARRECMAC1detail4Fields("PAYMTYPE").PutWithoutVerification ("CA")      ' Tipo de Lote
ARRECMAC1detail4.Cancel
'AsentarLote = True
ssql = "select * from AR_RA a "
ssql = ssql & " where a.COMPANYID='" & Gempresa & "' and   ((a.ESTADO<>'Completo' and a.ESTADO<>'Error') or a.ESTADO is null )  order by a.CNTBTCH,a.CNTITEM,a.COMPANYID"
rsJournal_Headers.Open ssql, cnDB, adOpenStatic, adLockReadOnly
If rsJournal_Headers.EOF = False And rsJournal_Headers.BOF = False Then
    rsJournal_Headers.MoveFirst
    Do While rsJournal_Headers.EOF = False
        If (loteant <> Trim(rsJournal_Headers!CNTBTCH) Or loteant = "") Then
            If (loteant <> "" And AsentarLote = True And PorAsentar = "SI") Then
                ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("3")       ' Comando Procesar
                ARRECMAC1batch.Process
                
                ARRECMAC1batchFields("BATCHSTAT").PutWithoutVerification ("7")        ' Estado de Lote
                
                ARRECMAC1batch.Update
                ARRECMAC1batch.Update
                
                ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("1")       ' Comando Procesar
                
                ARRECMAC1batch.Process
                
                ARPAYMPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)          ' Asentar Lote Desde
                ARPAYMPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)          ' Asentar Lote Hasta
                
                ARPAYMPOST2.Process
                ARRECMAC1batch.Read
                ARRECMAC1detail4.Cancel
                ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("-9999999")   ' N° de Asiento
                ARRECMAC1header.Browse "", 1
                ARRECMAC1header.Fetch
                ARRECMAC1detail4.Cancel
            End If
            AsentarLote = True
            Call conectaSage(Gempresa)
            fecha = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
            hora = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
            If existebanco(Trim(rsJournal_Headers!IDBANK)) = False Then
                Call AgregaError("No existe Banco " & Trim(rsJournal_Headers!IDBANK), "1", "0", "0", Trim(rsJournal_Headers!CNTBTCH), Trim(rsJournal_Headers!CNTITEM), fecha, hora)
                If rsJournal_Headers.State = 1 Then
                    rsJournal_Headers.Close
                End If
            End
            End If
            ARRECMAC1batch.Browse "((CODEPYMTYP = ""CA"") AND ((BATCHSTAT = 1) OR (BATCHSTAT = 7)))", 1
            
            ARRECMAC1batchFields("CODEPYMTYP").PutWithoutVerification ("CA")      ' Tipo de Lote
            ARRECMAC1batchFields("CNTBTCH").PutWithoutVerification ("0")          ' N° de Lote
            
            ARRECMAC1batch.RecordCreate 1
            
            ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("2")       ' Comando Procesar
            
            ARRECMAC1batch.Process
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1header.RecordCreate 2
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1batchFields("BATCHDESC").PutWithoutVerification (Trim(IIf(IsNull(rsJournal_Headers!BATCHDESC), "", rsJournal_Headers!BATCHDESC)))   ' Descripción
            ARRECMAC1batch.Update
            
            Dia = Mid(Trim(rsJournal_Headers!DATEBATCH), 4, 2)
            Mes = Left(Trim(rsJournal_Headers!DATEBATCH), 2)
            Anio = Right(Trim(rsJournal_Headers!DATEBATCH), 4)
            ARRECMAC1batchFields("DATEBTCH").value = DateSerial(Anio, Mes, Dia)     ' Fecha de Lote

            ARRECMAC1batch.Update
            ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("0")         ' N° de Asiento
            ARRECMAC1header.RecordCreate 2
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1batchFields("IDBANK").value = Trim(rsJournal_Headers!IDBANK) '"FCBANK"                       ' Cód. de Banco
            ARRECMAC1batch.Update
            ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("0")         ' N° de Asiento
            ARRECMAC1header.RecordCreate 2
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1batchFields("CODECURN").value = Trim(rsJournal_Headers!CODECURN) '"USD"                        ' Moneda de Banco Predet.
            ARRECMAC1batch.Update
            ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("0")         ' N° de Asiento
            ARRECMAC1header.RecordCreate 2
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("0")       ' Comando Procesar
            ARRECMAC1batch.Process
            ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("0")         ' N° de Asiento
            ARRECMAC1header.RecordCreate 2
            ARRECMAC1detail4.Cancel
            temp = ARRECMAC1header.Exists
            
            DATEDEPDia = Mid(Trim(rsJournal_Headers!DATEDEP), 4, 2)
            DATEDEPMes = Left(Trim(rsJournal_Headers!DATEDEP), 2)
            DATEDEPAnio = Right(Trim(rsJournal_Headers!DATEDEP), 4)
            ARRECMAC1batchFields("DEPDATE").PutWithoutVerification (DateSerial(DATEDEPAnio, DATEDEPMes, DATEDEPDia))  ' Fch. Depósito
            ARRECMAC1batch.Update
        End If

        ARRECMAC1headerFields("TEXTRMIT").value = Trim(rsJournal_Headers!TEXTRMIT) '"asiento descripcion"       ' Descripción del Asiento
        
        ARRECMAC1headerFields("RMITTYPE").value = "5"                         ' Tipo Trans. de Recibo
        
        ARRECMAC1detail1.Cancel
        ARRECMAC1headerFields("PROCESSCMD").PutWithoutVerification ("0")      ' Cód. de Comando Procesar
        ARRECMAC1header.Process
        
        DATERMITDia = Mid(Trim(rsJournal_Headers!DATERMIT), 4, 2)
        DATERMITMes = Left(Trim(rsJournal_Headers!DATERMIT), 2)
        DATERMITAnio = Right(Trim(rsJournal_Headers!DATERMIT), 4)
        
        ARRECMAC1headerFields("DATERMIT").value = DateSerial(DATERMITAnio, DATERMITMes, DATERMITDia)    ' Payment Date/Adjustment Date
                
        
        DATEBUSDia = Mid(Trim(rsJournal_Headers!DATEPOST), 4, 2)
        DATEBUSMes = Left(Trim(rsJournal_Headers!DATEPOST), 2)
        DATEBUSAnio = Right(Trim(rsJournal_Headers!DATEPOST), 4)
        ARRECMAC1headerFields("DATEBUS").value = DateSerial(DATEBUSAnio, DATEBUSMes, DATEBUSDia)      ' Fch. Asentam.
        
        ARRECMAC1headerFields("TEXTPAYOR").value = Trim(rsJournal_Headers!TEXTPAYOR) '"pagador"                  ' Pagador
        ARRECMAC1headerFields("TXTRMITREF").value = Trim(rsJournal_Headers!TXTRMITREF) '"referencia"              ' Referencia de Asiento
        
        ARRECMAC1headerFields("CODEPAYM").value = Trim(rsJournal_Headers!CODEPAYM) '"CASH"                      ' Cód. de Pago
        ARRECMAC1headerFields("IDRMIT").value = Trim(rsJournal_Headers!IDRMIT) '"987654"                      ' N° de Cheque/Recibo
        ARRECMAC1headerFields("AMTRMIT").value = rsJournal_Headers!AMTRMIT '"48.440"                     ' Mto. de Recibo Bancario
        
        temp = ARRECMAC1detail3.Exists
        ARRECMAC1detail3.RecordClear
        temp = ARRECMAC1detail3.Exists
        ARRECMAC1detail3.RecordCreate 0

        loteCreado = ARRECMAC1headerFields("CNTBTCH").value
        lote = Trim(rsJournal_Headers!CNTBTCH)
        Asiento = Trim(rsJournal_Headers!CNTITEM)
        B = 1
        ssql2 = "select * from AR_RA a "
        ssql2 = ssql2 & " left outer join AR_MR b on a.CNTBTCH=b.CNTBTCH and a.CNTITEM=b.CNTITEM and a.COMPANYID=b.COMPANYID"
        ssql2 = ssql2 & " where a.COMPANYID='" & Gempresa & "' and  b.CNTBTCH='" & rsJournal_Headers!CNTBTCH & "' and  b.CNTITEM='" & rsJournal_Headers!CNTITEM & "'  order by b.CNTBTCH,b.CNTITEM,b.CNTLINE,b.COMPANYID"
        rsJournal_Details.Open ssql2, cnDB, adOpenStatic, adLockReadOnly
        If rsJournal_Details.EOF = False And rsJournal_Details.BOF = False Then
            rsJournal_Details.MoveFirst
            Do While rsJournal_Details.EOF = False

                ARRECMAC1detail3Fields("GLDESC").value = Trim(rsJournal_Details!GLDESC) '"descripcion detalle"        ' Descripción del LM
                
                ARRECMAC1detail3Fields("IDACCT").value = Trim(rsJournal_Details!IDACCT) '"1000"                       ' N° de Cuenta
                ARRECMAC1detail3Fields("AMTDISTTC").value = rsJournal_Details!AMTDISTTC '"48.440"                  ' Monto de Dist.
                ARRECMAC1detail3Fields("GLREF").value = Trim(rsJournal_Details!GLREF) '"referencia detalle"          ' Referencia del LM
                
                ARRECMAC1detail3.Insert
                
                ARRECMAC1detail3Fields("CNTLINE").PutWithoutVerification ("-" & B)     ' N° de Línea
                
                ARRECMAC1detail3.Read
                temp = ARRECMAC1detail3.Exists
                ARRECMAC1detail3.RecordCreate 0
                B = B + 1

            rsJournal_Details.MoveNext
            Loop
        Else
            ARRECMAC1headerFields("GLREF").value = Trim(IIf(IsNull(rsJournal_Details!GLDESC), "", rsJournal_Details!GLREF)) 'Trim(rsJournal_Details!GLREF)  ' Referencia
            rsJournal_Details.Close
        End If
        If rsJournal_Details.State = 1 Then
            rsJournal_Details.Close
        End If
        loteant = Trim(rsJournal_Headers!CNTBTCH)
        
        ARRECMAC1detail3.Read
        ARRECMAC1header.Insert
        ARRECMAC1detail4.Cancel
        temp = ARRECMAC1header.Exists
        ARRECMAC1batch.Read
        AsientoCreado = ARRECMAC1headerFields("CNTITEM").value
        ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("0")         ' N° de Asiento
        ARRECMAC1header.RecordCreate 2
        ARRECMAC1detail4.Cancel
        temp = ARRECMAC1header.Exists
        
        If ConError = True Then
            AsentarLote = False
            Call AgregaError(mensajeerrorDet, "1", loteCreado, AsientoCreado, lote, Asiento, fecha, hora)
        Else
            Call AgregaCompleto(loteCreado, AsientoCreado, lote, Asiento, fecha, hora)
        End If
        ConError = False
        mensajeerrorDet = ""
        rsJournal_Headers.MoveNext
    Loop
    If (loteant <> "" And AsentarLote = True And PorAsentar = "SI") Then
        ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("3")       ' Comando Procesar
        ARRECMAC1batch.Process
        
        ARRECMAC1batchFields("BATCHSTAT").PutWithoutVerification ("7")        ' Estado de Lote
        
        ARRECMAC1batch.Update
        ARRECMAC1batch.Update
        
        ARRECMAC1batchFields("PROCESSCMD").PutWithoutVerification ("1")       ' Comando Procesar
        
        ARRECMAC1batch.Process
        
        ARPAYMPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)          ' Asentar Lote Desde
        ARPAYMPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)          ' Asentar Lote Hasta
        
        ARPAYMPOST2.Process
        ARRECMAC1batch.Read
        ARRECMAC1detail4.Cancel
        ARRECMAC1headerFields("CNTITEM").PutWithoutVerification ("-9999999")   ' N° de Asiento
        ARRECMAC1header.Browse "", 1
        ARRECMAC1header.Fetch
        ARRECMAC1detail4.Cancel
    End If
rsJournal_Headers.Close
Set rsJournal_Headers = Nothing
Set rsJournal_Details = Nothing
End If

Exit Sub
    
ErrorMacroAR:
    Error = True
    Dim lCount As Long
    Dim lIndex As Long
    Dim errorTexto As String
    AsentarLote = False
    ConError = True
    bitacora = Rbitacora & "logMacroAR.txt"
    Set v1 = CreateObject("Scripting.FileSystemObject")
    Set v2 = v1.OpenTextFile(bitacora, 8, True)
    
    v2.WriteLine "************************************************************************************************"
    v2.WriteLine "error MacroAP" & "  " & Date & "  " & Time & "  " & usuario & "  " & Order
    If Errors Is Nothing Then
        v2.WriteLine "No. Error: " & Err.Number & ".  Descripcion: " & Err.Description
        'MsgBox Err.Description
        errorTexto = Err.Number & "  " & Err.Description
        If InStr(mensajeerrorDet, errorTexto) = 0 Then
            mensajeerrorDet = mensajeerrorDet & Err.Number & "  " & Err.Description
        End If
    Else
        lCount = Errors.Count
    
        If lCount = 0 Then
            v2.WriteLine "No. Error: " & Err.Number & ".  Descripcion: " & Err.Description
            errorTexto = Err.Number & "  " & Err.Description
            If InStr(mensajeerrorDet, errorTexto) = 0 Then
                mensajeerrorDet = mensajeerrorDet & Err.Number & "  " & Err.Description
            End If
            'MsgBox Err.Description
        Else
            For lIndex = 0 To lCount - 1
                v2.WriteLine "No. Error: " & Err.Number & ".  Descripcion: " & Errors.Item(lIndex)  'Err.Description
                'MsgBox Errors.Item(lIndex)
                errorTexto = Err.Number & "  " & Errors.Item(lIndex)
                If InStr(mensajeerrorDet, errorTexto) = 0 Then
                    mensajeerrorDet = mensajeerrorDet & Err.Number & "  " & Errors.Item(lIndex)
                End If
            Next
            Errors.Clear
        End If
    End If
    v2.WriteLine "************************************************************************************************"
    v2.Close
    Resume Next

End Sub

Public Function RoundUpAlways(ByVal value As Double, Optional ByVal decimals As Integer = 2) As Double
    Dim factor As Double
    factor = 10 ^ decimals
    
    ' Multiplica, suma 0.5 para redondear hacia arriba siempre, y luego divide
    RoundUpAlways = Int(value * factor + 0.5) / factor
End Function

Public Sub AgregaError(Error As String, tipo As String, loteCreado As String, AsientoCreado As String, lote As String, Asiento As String, fecha As String, hora As String)
Dim rsErrorHeader As Recordset
Dim rsErrorDetalle As Recordset
Set rsErrorHeader = New ADODB.Recordset
Set rsErrorDetalle = New ADODB.Recordset
Dim ssql3 As String
Error = Replace(Error, "'", "")                 ' Quita comillas simples
Error = Replace(Error, vbCrLf, " ")             ' Quita saltos de línea combinados
Error = Replace(Error, vbCr, " ")               ' Quita retorno de carro
Error = Replace(Error, vbLf, " ")               ' Quita salto de línea
Error = Trim(Error)
Error = Left(Error, 255)
'If tipo = "1" Then ' detalle
    ssql3 = "update AR_MR set COMENTARIO = '" & Error & "' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTITEM='" & Asiento & "'"
    cnDB.Execute ssql3
'Else ' header
    ssql3 = " update AR_RA set RESULTADO = '" & Error & "', ESTADO='Error' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTITEM='" & Asiento & "'"
    cnDB.Execute ssql3

'End If
'rsErrorHeader.Close
'rsErrorDetalle.Close
Set rsErrorHeader = Nothing
Set rsErrorDetalle = Nothing
End Sub

Public Sub AgregaCompleto(loteCreado As String, AsientoCreado As String, lote As String, Asiento As String, fecha As String, hora As String)
Dim rsErrorHeader As Recordset
Set rsErrorHeader = New ADODB.Recordset
Dim ssql3 As String
'MsgBox "debug AgregaCompleto"
ssql3 = " update AR_RA set RESULTADO = '', ESTADO='Completo' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTITEM='" & Asiento & "'"
cnDB.Execute ssql3

ssql3 = "update AR_MR set COMENTARIO = 'Ok' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'"
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTITEM='" & Asiento & "'"
cnDB.Execute ssql3
'rsErrorHeader.Close
Set rsErrorHeader = Nothing
End Sub
