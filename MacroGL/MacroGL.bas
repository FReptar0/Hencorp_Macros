Attribute VB_Name = "MacroGL"
Public Sub AsientoGL()
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
Dim DOCDATEDia As String
Dim DOCDATEMes As String
Dim DOCDATEAnio As String
Dim DATEENTRYDia As String
Dim DATEENTRYMes As String
Dim DATEENTRYAnio As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim loteant As String
'    Set rsJournal_Headers = New Recordset
'    Set rsJournal_Details = New Recordset
Dim rsJournal_Details As Recordset
Dim rsJournal_Headers As Recordset
Set rsJournal_Headers = New ADODB.Recordset
Set rsJournal_Details = New ADODB.Recordset
'MsgBox "debug AsientoGL pasa a conectarse con el usuario"
UserID = UserSage '"test"
Password = PassSage '"test1"
OrgID = Gempresa '"geodat"
SessionDate = Date
'MsgBox "debug AsientoGL se coencto con el usuario"
On Error GoTo ErrorMacroGL
ConError = False
Dim objSignOn As AccpacSignonManager.AccpacSignonMgr
Set objSession = AccpacCOMAPI.AccpacSession
If usuario = "" Then
    usuario = UserID
End If
objSession.Init "", "GL", "GL2100", "70A"
objSession.Open UserID, Password, OrgID, SessionDate, 0, ""

If objSession.IsOpened Then
   Set objCompany = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
   Set dbCmp = objSession.OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)
End If
' TODO: To increase efficiency, comment out any unused DB links.
Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
Set mDBLinkCmpRW = OpenDBLink(DBLINK_COMPANY, DBLINK_FLG_READWRITE)

Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
Set mDBLinkSysRW = OpenDBLink(DBLINK_SYSTEM, DBLINK_FLG_READWRITE)

Dim temp As Boolean
Dim GLBATCH1batch As AccpacCOMAPI.AccpacView
Dim GLBATCH1batchFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "GL0008", GLBATCH1batch
Set GLBATCH1batchFields = GLBATCH1batch.Fields

Dim GLBATCH1header As AccpacCOMAPI.AccpacView
Dim GLBATCH1headerFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "GL0006", GLBATCH1header
Set GLBATCH1headerFields = GLBATCH1header.Fields

Dim GLBATCH1detail1 As AccpacCOMAPI.AccpacView
Dim GLBATCH1detail1Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "GL0010", GLBATCH1detail1
Set GLBATCH1detail1Fields = GLBATCH1detail1.Fields

Dim GLBATCH1detail2 As AccpacCOMAPI.AccpacView
Dim GLBATCH1detail2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "GL0402", GLBATCH1detail2
Set GLBATCH1detail2Fields = GLBATCH1detail2.Fields

GLBATCH1batch.Compose Array(GLBATCH1header)

GLBATCH1header.Compose Array(GLBATCH1batch, GLBATCH1detail1)

GLBATCH1detail1.Compose Array(GLBATCH1header, GLBATCH1detail2)

GLBATCH1detail2.Compose Array(GLBATCH1detail1)


Dim GLPOST2 As AccpacCOMAPI.AccpacView
Dim GLPOST2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "GL0030", GLPOST2
Set GLPOST2Fields = GLPOST2.Fields

ssql = "select * from GL_JH a "
ssql = ssql & " where a.COMPANYID='" & Gempresa & "' and   ((a.ESTADO<>'Completo' and a.ESTADO<>'Error') or a.ESTADO is null )  order by a.BATCHID,a.BTCHENTRY,a.COMPANYID"
rsJournal_Headers.Open ssql, cnDB, adOpenStatic, adLockReadOnly
If rsJournal_Headers.EOF = False And rsJournal_Headers.BOF = False Then
    rsJournal_Headers.MoveFirst
    Do While rsJournal_Headers.EOF = False
        If (loteant <> Trim(rsJournal_Headers!BATCHID) Or loteant = "") Then
            If (loteant <> "" And AsentarLote = True And PorAsentar = "SI") Then
                GLBATCH1batchFields("PROCESSCMD").PutWithoutVerification ("2")       ' Bloquear Opc. Lote
                
                GLBATCH1batch.Process
                
                GLBATCH1batchFields("RDYTOPOST").PutWithoutVerification ("1")        ' Listo para Asentar
                
                GLBATCH1batch.Update
                GLBATCH1batch.Update
                
                GLBATCH1batchFields("PROCESSCMD").PutWithoutVerification ("0")       ' Bloquear Opc. Lote

                GLBATCH1batch.Update
                GLBATCH1batch.Process
                
                GLPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)         ' Desde Número de Lote
                GLPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)         ' Hasta Número de Lote
                
                GLPOST2.Process
                GLBATCH1batch.Read
                GLBATCH1batch.Read
                GLBATCH1detail1Fields("TRANSNBR").PutWithoutVerification ("0000000020")   ' Número de Transacción

            
            End If
            GLBATCH1batch.Browse "((BATCHSTAT = ""1"" OR BATCHSTAT = ""6"" OR BATCHSTAT = ""9""))", 1
            GLBATCH1batch.RecordCreate 1
            GLBATCH1batch.Read

            GLBATCH1batchFields("PROCESSCMD").PutWithoutVerification ("1")        ' Bloquear Opc. Lote

            GLBATCH1batch.Process
            GLBATCH1headerFields("BTCHENTRY").PutWithoutVerification ("")         ' N° de Asiento
            GLBATCH1header.Browse "", 1
            GLBATCH1header.Fetch

            GLBATCH1headerFields("BTCHENTRY").value = "00000"                     ' N° de Asiento

            GLBATCH1header.RecordCreate 2
            temp = GLBATCH1header.Exists
            GLBATCH1batchFields("BTCHDESC").PutWithoutVerification (Trim(rsJournal_Headers!BATCHDESC))    ' Descripción
            GLBATCH1batch.Update
            
            DOCDATEDia = Mid(Trim(rsJournal_Headers!DOCDATE), 4, 2)
            DOCDATEMes = Left(Trim(rsJournal_Headers!DOCDATE), 2)
            DOCDATEAnio = Right(Trim(rsJournal_Headers!DOCDATE), 4)
            GLBATCH1headerFields("DOCDATE").value = DateSerial(DOCDATEAnio, DOCDATEMes, DOCDATEDia)      ' Fecha del Doc.
        End If
        AsentarLote = True
        DOCDATEDia = Mid(Trim(rsJournal_Headers!DOCDATE), 4, 2)
        DOCDATEMes = Left(Trim(rsJournal_Headers!DOCDATE), 2)
        DOCDATEAnio = Right(Trim(rsJournal_Headers!DOCDATE), 4)
        GLBATCH1headerFields("DOCDATE").value = DateSerial(DOCDATEAnio, DOCDATEMes, DOCDATEDia)      ' Fecha del Doc.
        fecha = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
        hora = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
        loteCreado = GLBATCH1headerFields("BATCHID").value
'        DATEENTRYMes = Mid(Trim(rsJournal_Headers!DATEENTRY), 5, 2)
'        DATEENTRYDia = Right(Trim(rsJournal_Headers!DATEENTRY), 2)
'        DATEENTRYAnio = Left(Trim(rsJournal_Headers!DATEENTRY), 4)
        
        DATEENTRYDia = Mid(Trim(rsJournal_Headers!DATEENTRY), 4, 2)
        DATEENTRYMes = Left(Trim(rsJournal_Headers!DATEENTRY), 2)
        DATEENTRYAnio = Right(Trim(rsJournal_Headers!DATEENTRY), 4)
        
        temp = GLBATCH1detail1.Exists
        GLBATCH1detail1.RecordClear
        temp = GLBATCH1detail1.Exists
        GLBATCH1detail1.RecordCreate 0

        'GLBATCH1headerFields("DATEENTRY").value = DateSerial(2025, 5, 19)     ' Fch. Asentam.
        GLBATCH1headerFields("DATEENTRY").value = DateSerial(DATEENTRYAnio, DATEENTRYMes, DATEENTRYDia)       ' Fch. Asentam.
        GLBATCH1headerFields("FSCSPERD").value = Trim(rsJournal_Headers!FSCSPERD)                         ' Período Fiscal
        GLBATCH1headerFields("JRNLDESC").PutWithoutVerification (Trim(rsJournal_Headers!JRNLDESC))      ' Descripción
    
        GLBATCH1detail1.Read
        temp = GLBATCH1detail1.Exists
        GLBATCH1detail1.RecordCreate 0
                
        lote = Trim(rsJournal_Headers!BATCHID)
        Asiento = Trim(rsJournal_Headers!BTCHENTRY)
        ssql2 = "select * from GL_JH a "
        ssql2 = ssql2 & " left outer join GL_JD b on a.BATCHID=b.BATCHNBR and a.BTCHENTRY=b.JOURNALID and a.COMPANYID=b.COMPANYID"
        ssql2 = ssql2 & " where a.COMPANYID='" & Gempresa & "' and  b.BATCHNBR='" & rsJournal_Headers!BATCHID & "' and  b.JOURNALID='" & rsJournal_Headers!BTCHENTRY & "'  order by b.BATCHNBR,b.JOURNALID,b.TRANSNBR,b.COMPANYID"
        rsJournal_Details.Open ssql2, cnDB, adOpenStatic, adLockReadOnly
        If rsJournal_Details.EOF = False And rsJournal_Details.BOF = False Then
            rsJournal_Details.MoveFirst
            Do While rsJournal_Details.EOF = False
                GLBATCH1detail1Fields("TRANSREF").value = Trim(IIf(IsNull(rsJournal_Details!TRANSREF), "", rsJournal_Details!TRANSREF)) 'Trim(rsJournal_Details!TRANSREF)  ' Referencia
                GLBATCH1detail1Fields("TRANSDESC").value = Trim(rsJournal_Details!TRANSDESC) ' Descripción
                GLBATCH1detail1Fields("ACCTID").value = Trim(rsJournal_Details!ACCTID)                       ' Número de Cuenta
                GLBATCH1detail1Fields("PROCESSCMD").PutWithoutVerification ("0")      ' Cambios de proceso
                GLBATCH1detail1.Process
                
                GLBATCH1detail1Fields("SCURNAMT").value = Trim(rsJournal_Details!TRANSAMT)                    ' Monto de Moneda de Origen
'                Mes = Mid(Trim(rsJournal_Details!TRANSDATE), 5, 2)
'                Dia = Right(Trim(rsJournal_Details!TRANSDATE), 2)
'                Anio = Left(Trim(rsJournal_Details!TRANSDATE), 4)
                Dia = Mid(Trim(rsJournal_Headers!DATEENTRY), 4, 2)
                Mes = Left(Trim(rsJournal_Headers!DATEENTRY), 2)
                Anio = Right(Trim(rsJournal_Headers!DATEENTRY), 4)
                'GLBATCH1detail1Fields("TRANSDATE").value = DateSerial(2025, 5, 11)    ' Fecha del Libro Diario
                GLBATCH1detail1Fields("TRANSDATE").value = DateSerial(Anio, Mes, Dia)    ' Fecha del Libro Diario
                GLBATCH1detail1.Insert
                
                GLBATCH1detail1Fields("TRANSNBR").PutWithoutVerification (Trim(rsJournal_Details!TRANSNBR))   ' Número de Transacción

                GLBATCH1detail1.Read
                temp = GLBATCH1detail1.Exists
                GLBATCH1detail1.RecordCreate 0
                
                rsJournal_Details.MoveNext
            Loop
        Else
            GLBATCH1detail1Fields("TRANSREF").value = Trim(IIf(IsNull(rsJournal_Details!TRANSREF), "", rsJournal_Details!TRANSREF)) 'Trim(rsJournal_Details!TRANSREF)  ' Referencia
            rsJournal_Details.Close
        End If
        If rsJournal_Details.State = 1 Then
            rsJournal_Details.Close
        End If
        loteant = Trim(rsJournal_Headers!BATCHID)
        GLBATCH1header.Insert
        GLBATCH1header.Read
        temp = GLBATCH1header.Exists
        GLBATCH1batch.Read
        AsientoCreado = GLBATCH1headerFields("BTCHENTRY").value
        If ConError = True Then
            AsentarLote = False
            Call AgregaError(mensajeerrorDet, "1", loteCreado, AsientoCreado, lote, Asiento, fecha, hora)
        Else
            Call AgregaCompleto(loteCreado, AsientoCreado, lote, Asiento, fecha, hora)
        End If
        ConError = False
        mensajeerrorDet = ""
        GLBATCH1headerFields("BTCHENTRY").value = "00000"                     ' N° de Asiento
        GLBATCH1header.RecordCreate 2
        temp = GLBATCH1header.Exists
        rsJournal_Headers.MoveNext
    Loop
    If (loteant <> "" And AsentarLote = True And PorAsentar = "SI") Then
        GLBATCH1batchFields("PROCESSCMD").PutWithoutVerification ("2")       ' Bloquear Opc. Lote
                
        GLBATCH1batch.Process
        
        GLBATCH1batchFields("RDYTOPOST").PutWithoutVerification ("1")        ' Listo para Asentar
        
        GLBATCH1batch.Update
        GLBATCH1batch.Update
        
        GLBATCH1batchFields("PROCESSCMD").PutWithoutVerification ("0")       ' Bloquear Opc. Lote
    
        GLBATCH1batch.Update
        GLBATCH1batch.Process
        
        GLPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)         ' Desde Número de Lote
        GLPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)         ' Hasta Número de Lote
        
        GLPOST2.Process
        GLBATCH1batch.Read
        GLBATCH1batch.Read
        GLBATCH1detail1Fields("TRANSNBR").PutWithoutVerification ("0000000020")   ' Número de Transacción
 
    End If

rsJournal_Headers.Close
'rsJournal_Details.Close
Set rsJournal_Headers = Nothing
Set rsJournal_Details = Nothing
End If
Exit Sub
    
ErrorMacroGL:
    Error = True
    Dim lCount As Long
    Dim lIndex As Long
    Dim errorTexto As String
    ConError = True
    AsentarLote = False
    bitacora = Rbitacora & "logMacroGL.txt"
    Set v1 = CreateObject("Scripting.FileSystemObject")
    Set v2 = v1.OpenTextFile(bitacora, 8, True)
    
    v2.WriteLine "************************************************************************************************"
    v2.WriteLine "error MacroGL" & "  " & Date & "  " & Time & "  " & usuario & "  " & Order
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
'MsgBox "debug AgregaError"
Error = Replace(Error, "'", "")
Error = Left(Error, 255)
'If tipo = "1" Then ' detalle
    ssql3 = "update GL_JD set COMENTARIO = '" & Error & "' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  BATCHNBR='" & lote & "' and JOURNALID='" & Asiento & "'"
    cnDB.Execute ssql3
'Else ' header
    ssql3 = " update GL_JH set RESULTADO = '" & Error & "', ESTADO='Error' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  BATCHID='" & lote & "' and BTCHENTRY='" & Asiento & "'"
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
ssql3 = " update GL_JH set RESULTADO = '', ESTADO='Completo' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  BATCHID='" & lote & "' and BTCHENTRY='" & Asiento & "'"
cnDB.Execute ssql3

ssql3 = "update GL_JD set COMENTARIO = 'Ok' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'"
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  BATCHNBR='" & lote & "' and JOURNALID='" & Asiento & "'"
cnDB.Execute ssql3
'rsErrorHeader.Close
Set rsErrorHeader = Nothing
End Sub


