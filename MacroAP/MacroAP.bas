Attribute VB_Name = "MacroAP"
Public Sub AsientoAP()
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
'MsgBox "debug AsientoAP pasa a conectarse con el usuario"
UserID = UserSage '"test"
Password = PassSage '"test1"
OrgID = Gempresa '"geodat"
SessionDate = Date
'MsgBox "debug AsientoAP se coencto con el usuario"
On Error GoTo ErrorMacroAP
ConError = False
If usuario = "" Then
    usuario = UserID
End If
Dim objSignOn As AccpacSignonManager.AccpacSignonMgr
Set objSession = AccpacCOMAPI.AccpacSession

objSession.Init "", "AP", "AP3000", "70A"
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
Dim APPAYMENT1batch As AccpacCOMAPI.AccpacView
Dim APPAYMENT1batchFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0030", APPAYMENT1batch
Set APPAYMENT1batchFields = APPAYMENT1batch.Fields

Dim APPAYMENT1header As AccpacCOMAPI.AccpacView
Dim APPAYMENT1headerFields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0031", APPAYMENT1header
Set APPAYMENT1headerFields = APPAYMENT1header.Fields

Dim APPAYMENT1detail1 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail1Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0033", APPAYMENT1detail1
Set APPAYMENT1detail1Fields = APPAYMENT1detail1.Fields

Dim APPAYMENT1detail2 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0034", APPAYMENT1detail2
Set APPAYMENT1detail2Fields = APPAYMENT1detail2.Fields

Dim APPAYMENT1detail3 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail3Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0032", APPAYMENT1detail3
Set APPAYMENT1detail3Fields = APPAYMENT1detail3.Fields

Dim APPAYMENT1detail4 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail4Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0048", APPAYMENT1detail4
Set APPAYMENT1detail4Fields = APPAYMENT1detail4.Fields

Dim APPAYMENT1detail5 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail5Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0170", APPAYMENT1detail5
Set APPAYMENT1detail5Fields = APPAYMENT1detail5.Fields

Dim APPAYMENT1detail6 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail6Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0406", APPAYMENT1detail6
Set APPAYMENT1detail6Fields = APPAYMENT1detail6.Fields

Dim APPAYMENT1detail7 As AccpacCOMAPI.AccpacView
Dim APPAYMENT1detail7Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0069", APPAYMENT1detail7
Set APPAYMENT1detail7Fields = APPAYMENT1detail7.Fields

APPAYMENT1batch.Compose Array(APPAYMENT1header)

APPAYMENT1header.Compose Array(APPAYMENT1batch, APPAYMENT1detail3, APPAYMENT1detail1, APPAYMENT1detail6, APPAYMENT1detail5, APPAYMENT1detail7)

APPAYMENT1detail1.Compose Array(APPAYMENT1header, APPAYMENT1detail2, APPAYMENT1detail4)

APPAYMENT1detail2.Compose Array(APPAYMENT1detail1)

APPAYMENT1detail3.Compose Array(APPAYMENT1header)

APPAYMENT1detail4.Compose Array(APPAYMENT1batch, APPAYMENT1header, APPAYMENT1detail3, APPAYMENT1detail1, APPAYMENT1detail2)

APPAYMENT1detail5.Compose Array(APPAYMENT1header)

APPAYMENT1detail6.Compose Array(APPAYMENT1header)

APPAYMENT1detail7.Compose Array(APPAYMENT1header)


Dim APPAYMPOST2 As AccpacCOMAPI.AccpacView
Dim APPAYMPOST2Fields As AccpacCOMAPI.AccpacViewFields
mDBLinkCmpRW.OpenView "AP0040", APPAYMPOST2
Set APPAYMPOST2Fields = APPAYMPOST2.Fields


APPAYMENT1batch.RecordClear

APPAYMENT1batchFields("PAYMTYPE").PutWithoutVerification ("PY")       ' Batch Selector

APPAYMENT1headerFields("BTCHTYPE").PutWithoutVerification ("PY")      ' Batch Type
APPAYMENT1detail3Fields("BATCHTYPE").PutWithoutVerification ("PY")    ' Batch Type
APPAYMENT1detail1Fields("BATCHTYPE").PutWithoutVerification ("PY")    ' Batch Type
APPAYMENT1detail2Fields("BATCHTYPE").PutWithoutVerification ("PY")    ' Batch Type
temp = APPAYMENT1header.Exists
temp = APPAYMENT1header.Exists
temp = APPAYMENT1header.Exists

ssql = "select * from AP_PA a "
ssql = ssql & " where a.COMPANYID='" & Gempresa & "' and   ((a.ESTADO<>'Completo' and a.ESTADO<>'Error') or a.ESTADO is null )  order by a.CNTBTCH,a.CNTENTR,a.COMPANYID"
rsJournal_Headers.Open ssql, cnDB, adOpenStatic, adLockReadOnly
If rsJournal_Headers.EOF = False And rsJournal_Headers.BOF = False Then
    rsJournal_Headers.MoveFirst
    Do While rsJournal_Headers.EOF = False
        If (loteant <> Trim(rsJournal_Headers!CNTBTCH) Or loteant = "") Then
            If (loteant <> "" And AsentarLote = True And PorAsentar = "SI") Then
                APPAYMENT1batch.Read
                
                APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("1")      ' Código de Comando Procesar
                
                APPAYMENT1batch.Process
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1header.Read
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("2")      ' Código de Comando Procesar
                APPAYMENT1batch.Process
                
                APPAYMENT1batchFields("BATCHSTAT").PutWithoutVerification ("7")       ' Estado de Lotes
                
                APPAYMENT1batch.Update
                APPAYMENT1batch.Update
                
                APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("0")      ' Código de Comando Procesar

                APPAYMPOST2Fields("TYPEBTCH").PutWithoutVerification ("PY")           ' Tipo de Lote
                APPAYMPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)          ' Asentar Lote Desde
                APPAYMPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)          ' Asentar Lote Hasta
                
                APPAYMPOST2.Process
                APPAYMENT1batch.Read
            End If
            Call conectaSage(Gempresa)
            fecha = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
            hora = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
            If existebanco(Trim(rsJournal_Headers!IDBANK)) = False Then
                Call AgregaError("No existe Banco " & Trim(rsJournal_Headers!IDBANK), "1", "0", "0", Trim(rsJournal_Headers!CNTBTCH), Trim(rsJournal_Headers!CNTENTR), fecha, hora)
                If rsJournal_Headers.State = 1 Then
                    rsJournal_Headers.Close
                End If
            End
            End If
            AsentarLote = True
            APPAYMENT1batch.Browse "((PAYMTYPE = ""PY"") AND ((BATCHSTAT = 1) OR (BATCHSTAT = 7) OR (BATCHSTAT = 8)))", 1
            
            APPAYMENT1batchFields("PAYMTYPE").PutWithoutVerification ("PY")       ' Batch Selector
            APPAYMENT1batchFields("CNTBTCH").PutWithoutVerification ("0")         ' Batch Number
            
            APPAYMENT1batch.RecordCreate 1
            
            APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("1")      ' Process Command Code
            
            APPAYMENT1batch.Process
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            APPAYMENT1header.RecordCreate 2
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            APPAYMENT1batchFields("BATCHDESC").PutWithoutVerification (Trim(rsJournal_Headers!BATCHDESC))   ' Description
            APPAYMENT1batch.Update
            
            
            Dia = Mid(Trim(rsJournal_Headers!DATEBATCH), 4, 2)
            Mes = Left(Trim(rsJournal_Headers!DATEBATCH), 2)
            Anio = Right(Trim(rsJournal_Headers!DATEBATCH), 4)
            APPAYMENT1batchFields("DATEBTCH").value = DateSerial(Anio, Mes, Dia)     ' Batch Date
            
            APPAYMENT1batch.Update
            APPAYMENT1header.RecordCreate 2
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            APPAYMENT1batchFields("IDBANK").value = Trim(rsJournal_Headers!IDBANK)                      ' Bank Code
            APPAYMENT1batch.Update
            APPAYMENT1header.RecordCreate 2
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
            temp = APPAYMENT1header.Exists
        End If

        
        APPAYMENT1headerFields("TEXTRMIT").value = Trim(rsJournal_Headers!TEXTRMIT) '"descripcion asiento"      ' Entry Description
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("RMITTYPE").value = "4"                         ' Payment Trans. Type Trim(rsJournal_Headers!RMITTYPE)
        
        APPAYMENT1detail1.Cancel
        APPAYMENT1headerFields("PROCESSCMD").PutWithoutVerification ("0")     ' Process Command Code
        APPAYMENT1header.Process
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        DATERMITDia = Mid(Trim(rsJournal_Headers!DATERMIT), 4, 2)
        DATERMITMes = Left(Trim(rsJournal_Headers!DATERMIT), 2)
        DATERMITAnio = Right(Trim(rsJournal_Headers!DATERMIT), 4)
        
        APPAYMENT1headerFields("DATERMIT").value = DateSerial(DATERMITAnio, DATERMITMes, DATERMITDia)    ' Payment Date/Adjustment Date
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        DATEBUSDia = Mid(Trim(rsJournal_Headers!DATEPOST), 4, 2)
        DATEBUSMes = Left(Trim(rsJournal_Headers!DATEPOST), 2)
        DATEBUSAnio = Right(Trim(rsJournal_Headers!DATEPOST), 4)

        APPAYMENT1headerFields("DATEBUS").value = DateSerial(DATEBUSAnio, DATEBUSMes, DATEBUSDia)     ' Posting Date
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("NAMERMIT").value = Trim(rsJournal_Headers!NAMERMIT) '"remit to"                 ' Vendor / Payee Name
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("PAYMCODE").value = "CHECK"                 ' Payment Code
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("SWPRNTRMIT").value = "0"                      ' Check Print Required
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("IDRMIT").value = Trim(rsJournal_Headers!IDRMIT) '"000000000000012365"         ' Check Number
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        
        APPAYMENT1headerFields("TXTRMITREF").value = Trim(rsJournal_Headers!REFERECE) '"referencia header"      ' Entry Reference
        
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1detail3.Exists
        APPAYMENT1detail3.RecordClear
        temp = APPAYMENT1detail3.Exists
        APPAYMENT1detail3.RecordCreate 0
   
        
        loteCreado = APPAYMENT1headerFields("CNTBTCH").value
        lote = Trim(rsJournal_Headers!CNTBTCH)
        Asiento = Trim(rsJournal_Headers!CNTENTR)
        B = 1
        ssql2 = "select * from AP_PA a "
        ssql2 = ssql2 & " left outer join AP_MP b on a.CNTBTCH=b.CNTBTCH and a.CNTENTR=b.CNTRMIT and a.COMPANYID=b.COMPANYID"
        ssql2 = ssql2 & " where a.COMPANYID='" & Gempresa & "' and  b.CNTBTCH='" & rsJournal_Headers!CNTBTCH & "' and  b.CNTRMIT='" & rsJournal_Headers!CNTENTR & "'  order by b.CNTBTCH,b.CNTRMIT,b.CNTLINE,b.COMPANYID"
        rsJournal_Details.Open ssql2, cnDB, adOpenStatic, adLockReadOnly
        If rsJournal_Details.EOF = False And rsJournal_Details.BOF = False Then
            rsJournal_Details.MoveFirst
            Do While rsJournal_Details.EOF = False
                APPAYMENT1detail3Fields("GLDESC").value = Trim(rsJournal_Details!GLDESC) '"descripcion detalle 1"     ' G/L Description
                
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1detail3Fields("IDACCT").value = Trim(rsJournal_Details!IDACCT)                     ' Account Number
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1detail3Fields("AMTDISTTC").value = Trim(rsJournal_Details!AMTDISTTC)               ' Dist. Amount
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1detail3Fields("GLREF").value = Trim(rsJournal_Details!GLREF)          ' G/L Reference
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1detail3.Insert
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                temp = APPAYMENT1header.Exists
                APPAYMENT1detail3Fields("CNTLINE").PutWithoutVerification ("-" & B)    ' Line Number
                
                APPAYMENT1detail3Fields("CNTLINE").PutWithoutVerification ("-" & B)     ' Line Number
                
                APPAYMENT1detail3.Read
                temp = APPAYMENT1detail3.Exists
                APPAYMENT1detail3.RecordCreate 0
                B = B + 1

            rsJournal_Details.MoveNext
            Loop
        Else
            APPAYMENT1detail3Fields("GLREF").value = Trim(IIf(IsNull(rsJournal_Details!GLREF), "", rsJournal_Details!GLREF)) 'Trim(rsJournal_Details!GLREF)  ' Referencia
            rsJournal_Details.Close
        End If
        If rsJournal_Details.State = 1 Then
            rsJournal_Details.Close
        End If
        loteant = Trim(rsJournal_Headers!CNTBTCH)

        
        'APPAYMENT1detail3Fields("CNTLINE").PutWithoutVerification ("-2")      ' Line Number
        
        APPAYMENT1detail3.Read
        APPAYMENT1header.Insert
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        APPAYMENT1batch.Read
        AsientoCreado = APPAYMENT1headerFields("CNTENTR").value
        APPAYMENT1headerFields("CNTENTR").PutWithoutVerification ("0")        ' Entry Number
        APPAYMENT1header.RecordCreate 2
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
                    
        
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
        APPAYMENT1batch.Read
                
        APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("1")      ' Código de Comando Procesar
        
        APPAYMENT1batch.Process
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        APPAYMENT1header.Read
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        temp = APPAYMENT1header.Exists
        APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("2")      ' Código de Comando Procesar
        APPAYMENT1batch.Process
        
        APPAYMENT1batchFields("BATCHSTAT").PutWithoutVerification ("7")       ' Estado de Lotes
        
        APPAYMENT1batch.Update
        APPAYMENT1batch.Update
        
        APPAYMENT1batchFields("PROCESSCMD").PutWithoutVerification ("0")      ' Código de Comando Procesar

        APPAYMPOST2Fields("TYPEBTCH").PutWithoutVerification ("PY")           ' Tipo de Lote
        APPAYMPOST2Fields("BATCHIDFR").PutWithoutVerification (loteCreado)          ' Asentar Lote Desde
        APPAYMPOST2Fields("BATCHIDTO").PutWithoutVerification (loteCreado)          ' Asentar Lote Hasta
        
        APPAYMPOST2.Process
        APPAYMENT1batch.Read
    End If
rsJournal_Headers.Close
Set rsJournal_Headers = Nothing
Set rsJournal_Details = Nothing
End If
Exit Sub
    
ErrorMacroAP:
    Error = True
    Dim lCount As Long
    Dim lIndex As Long
    Dim errorTexto As String
    ConError = True
    AsentarLote = False
    bitacora = Rbitacora & "logMacroAP.txt"
    Set v1 = CreateObject("Scripting.FileSystemObject")
    Set v2 = v1.OpenTextFile(bitacora, 8, True)
    
    v2.WriteLine "************************************************************************************************"
    v2.WriteLine "error MacroAP" & "  " & Date & "  " & Time & "  " & usuario & "  " & Order
    If Errors Is Nothing Then
        v2.WriteLine "No. Error: " & Err.Number & ".  Descripcion: " & Err.Description
        MsgBox Err.Description
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
            MsgBox Err.Description
        Else
            For lIndex = 0 To lCount - 1
                v2.WriteLine "No. Error: " & Err.Number & ".  Descripcion: " & Errors.Item(lIndex)  'Err.Description
                MsgBox Errors.Item(lIndex)
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
    ssql3 = "update AP_MP set COMENTARIO = '" & Error & "' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTRMIT='" & Asiento & "'"
    cnDB.Execute ssql3
'Else ' header
    ssql3 = " update AP_PA set RESULTADO = '" & Error & "', ESTADO='Error' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
    ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTENTR='" & Asiento & "'"
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
ssql3 = " update AP_PA set RESULTADO = '', ESTADO='Completo' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "' ,FECHA='" & fecha & "' ,HORA='" & hora & "', USUARIO='" & usuario & "'   "
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTENTR='" & Asiento & "'"
cnDB.Execute ssql3

ssql3 = "update AP_MP set COMENTARIO = 'Ok' , LOTE='" & loteCreado & "',ASIENTO='" & AsientoCreado & "'"
ssql3 = ssql3 & " where COMPANYID='" & Gempresa & "' and  CNTBTCH='" & lote & "' and CNTRMIT='" & Asiento & "'"
cnDB.Execute ssql3
'rsErrorHeader.Close
Set rsErrorHeader = Nothing
End Sub
