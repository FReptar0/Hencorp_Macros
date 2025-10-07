-- DROP SCHEMA dbo;

CREATE SCHEMA dbo;
-- hencore_accpac.dbo.AP_PA definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.AP_PA;

CREATE TABLE hencore_accpac.dbo.AP_PA (
	BTCHTYPE nchar(2) COLLATE Latin1_General_BIN NULL,
	CNTBTCH decimal(9,0) NOT NULL,
	CNTENTR decimal(7,0) NOT NULL,
	IDRMIT nchar(12) COLLATE Latin1_General_BIN NULL,
	IDVEND nchar(12) COLLATE Latin1_General_BIN NULL,
	DATERMIT nchar(12) COLLATE Latin1_General_BIN NULL,
	TEXTRMIT char(60) COLLATE Latin1_General_BIN NULL,
	NAMERMIT nchar(60) COLLATE Latin1_General_BIN NULL,
	AMTRMIT numeric(19,3) NULL,
	AMTRMITTC numeric(19,3) NULL,
	RMITTYPE smallint NULL,
	DOCTYPE smallint NULL,
	CNTLSTLINE decimal(5,0) NULL,
	FISCYR nchar(4) COLLATE Latin1_General_BIN NULL,
	FISCPER nchar(2) COLLATE Latin1_General_BIN NULL,
	AMTRMITHC numeric(19,3) NULL,
	SWPRINTED smallint NULL,
	CHECKLANG nchar(3) COLLATE Latin1_General_BIN NULL,
	OPERBANK smallint NULL,
	OPERVEND smallint NULL,
	DATEACTVPP nchar(12) COLLATE Latin1_General_BIN NULL,
	SRCEAPPL nchar(2) COLLATE Latin1_General_BIN NULL,
	IDBANK nchar(8) COLLATE Latin1_General_BIN NULL,
	PAYMTYPE smallint NULL,
	AMTNETTC numeric(19,3) NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	ESTADO nchar(15) COLLATE Latin1_General_BIN NULL,
	RESULTADO nchar(255) COLLATE Latin1_General_BIN NULL,
	FECHA decimal(19,0) NULL,
	HORA time NULL,
	LOTE decimal(9,0) NULL,
	ASIENTO decimal(7,0) NULL,
	USUARIO nchar(15) COLLATE Latin1_General_BIN NULL,
	BATCHDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	DATEBATCH nchar(12) COLLATE Latin1_General_BIN NULL,
	DATEPOST nchar(12) COLLATE Latin1_General_BIN NULL,
	REFERECE nchar(60) COLLATE Latin1_General_BIN NULL,
	CONSTRAINT AP_PA_PK PRIMARY KEY (CNTBTCH,CNTENTR)
);


-- hencore_accpac.dbo.AR_RA definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.AR_RA;

CREATE TABLE hencore_accpac.dbo.AR_RA (
	CODEPYMTYP nchar(2) COLLATE Latin1_General_BIN NULL,
	CNTBTCH decimal(9,0) NOT NULL,
	CNTITEM decimal(7,0) NOT NULL,
	IDRMIT nchar(24) COLLATE Latin1_General_BIN NULL,
	IDCUST nchar(12) COLLATE Latin1_General_BIN NULL,
	DATERMIT nchar(12) COLLATE Latin1_General_BIN NULL,
	TEXTRMIT nchar(60) COLLATE Latin1_General_BIN NULL,
	TXTRMITREF nchar(60) COLLATE Latin1_General_BIN NULL,
	AMTRMIT numeric(19,3) NULL,
	AMTRMITTC numeric(19,3) NULL,
	CNTPAYMETR decimal(5,0) NULL,
	AMTPAYMTC numeric(19,3) NULL,
	CODEPAYM nchar(12) COLLATE Latin1_General_BIN NULL,
	CODECURN nchar(3) COLLATE Latin1_General_BIN NULL,
	RMITTYPE smallint NULL,
	DOCTYPE smallint NULL,
	IDINVCMTCH nchar(22) COLLATE Latin1_General_BIN NULL,
	CNTLSTLINE decimal(5,0) NULL,
	FISCYR nchar(4) COLLATE Latin1_General_BIN NULL,
	FISCPER nchar(2) COLLATE Latin1_General_BIN NULL,
	TEXTPAYOR nchar(60) COLLATE Latin1_General_BIN NULL,
	DATERATETC nchar(12) COLLATE Latin1_General_BIN NULL,
	AMTRMITHC numeric(19,3) NULL,
	DOCNBR nchar(22) COLLATE Latin1_General_BIN NULL,
	AMTADJHC numeric(19,3) NULL,
	OPERBANK smallint NULL,
	OPERCUST smallint NULL,
	SRCEAPPL nchar(2) COLLATE Latin1_General_BIN NULL,
	IDBANK nchar(8) COLLATE Latin1_General_BIN NULL,
	CODECURNBC nchar(3) COLLATE Latin1_General_BIN NULL,
	AMTNETTC numeric(19,3) NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	ESTADO nchar(15) COLLATE Latin1_General_BIN NULL,
	RESULTADO nchar(255) COLLATE Latin1_General_BIN NULL,
	FECHA decimal(9,0) NULL,
	HORA time NULL,
	LOTE decimal(9,0) NULL,
	ASIENTO decimal(7,0) NULL,
	USUARIO nchar(15) COLLATE Latin1_General_BIN NULL,
	BATCHDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	CONSTRAINT AR_RA_PK PRIMARY KEY (CNTBTCH,CNTITEM)
);


-- hencore_accpac.dbo.GL_JH definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.GL_JH;

CREATE TABLE hencore_accpac.dbo.GL_JH (
	BATCHID nchar(6) COLLATE Latin1_General_BIN NOT NULL,
	BTCHENTRY nchar(5) COLLATE Latin1_General_BIN NOT NULL,
	SRCELEDGER nchar(2) COLLATE Latin1_General_BIN NULL,
	SRCETYPE nchar(2) COLLATE Latin1_General_BIN NULL,
	FSCSYR numeric(18,0) NULL,
	FSCSPERD numeric(18,0) NULL,
	JRNLDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	DATEENTRY nchar(12) COLLATE Latin1_General_BIN NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	ESTADO nchar(15) COLLATE Latin1_General_BIN NULL,
	RESULTADO nchar(255) COLLATE Latin1_General_BIN NULL,
	FECHA decimal(9,0) NULL,
	HORA time NULL,
	LOTE decimal(9,0) NULL,
	ASIENTO decimal(7,0) NULL,
	USUARIO nchar(15) COLLATE Latin1_General_BIN NULL,
	DOCDATE nchar(12) COLLATE Latin1_General_BIN NULL,
	BATCHDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	CONSTRAINT PK__GL_JH__E353881204A71C70 PRIMARY KEY (BATCHID,BTCHENTRY)
);


-- hencore_accpac.dbo.AP_MP definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.AP_MP;

CREATE TABLE hencore_accpac.dbo.AP_MP (
	BATCHTYPE nchar(2) COLLATE Latin1_General_BIN NULL,
	CNTBTCH decimal(9,0) NOT NULL,
	CNTRMIT decimal(7,0) NOT NULL,
	CNTLINE decimal(5,0) NOT NULL,
	IDACCT nchar(45) COLLATE Latin1_General_BIN NULL,
	GLREF nchar(60) COLLATE Latin1_General_BIN NULL,
	GLDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	AMTDISTTC numeric(19,3) NULL,
	AMTNETTC numeric(19,3) NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	COMENTARIO nchar(255) COLLATE Latin1_General_BIN NULL,
	LOTE numeric(9,0) NULL,
	ASIENTO numeric(7,0) NULL,
	CONSTRAINT AP_MP_PK PRIMARY KEY (CNTBTCH,CNTRMIT,CNTLINE),
	CONSTRAINT AP_MP_AP_PA_FK FOREIGN KEY (CNTBTCH,CNTRMIT) REFERENCES hencore_accpac.dbo.AP_PA(CNTBTCH,CNTENTR)
);


-- hencore_accpac.dbo.AR_MR definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.AR_MR;

CREATE TABLE hencore_accpac.dbo.AR_MR (
	CODEPAYM nchar(2) COLLATE Latin1_General_BIN NULL,
	CNTBTCH decimal(9,0) NOT NULL,
	CNTITEM decimal(7,0) NOT NULL,
	CNTLINE decimal(5,0) NOT NULL,
	AMTDIST numeric(19,3) NULL,
	IDDISTCODE nchar(6) COLLATE Latin1_General_BIN NULL,
	IDACCT nchar(45) COLLATE Latin1_General_BIN NULL,
	GLREF nchar(60) COLLATE Latin1_General_BIN NULL,
	GLDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	AMTDISTTC numeric(19,3) NULL,
	AMTNETTC numeric(19,3) NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	COMENTARIO nchar(255) COLLATE Latin1_General_BIN NULL,
	LOTE decimal(9,0) NULL,
	ASIENTO decimal(7,0) NULL,
	CONSTRAINT AR_MR_PK PRIMARY KEY (CNTBTCH,CNTITEM,CNTLINE),
	CONSTRAINT AR_MR_AR_RA_FK FOREIGN KEY (CNTBTCH,CNTITEM) REFERENCES hencore_accpac.dbo.AR_RA(CNTBTCH,CNTITEM)
);


-- hencore_accpac.dbo.GL_JD definition

-- Drop table

-- DROP TABLE hencore_accpac.dbo.GL_JD;

CREATE TABLE hencore_accpac.dbo.GL_JD (
	BATCHNBR nchar(6) COLLATE Latin1_General_BIN NOT NULL,
	JOURNALID nchar(5) COLLATE Latin1_General_BIN NOT NULL,
	TRANSNBR nchar(10) COLLATE Latin1_General_BIN NOT NULL,
	ACCTID numeric(18,0) NULL,
	COMPANYID nchar(15) COLLATE Latin1_General_BIN NULL,
	TRANSAMT decimal(19,3) NULL,
	TRANSDESC nchar(60) COLLATE Latin1_General_BIN NULL,
	TRANSREF nchar(60) COLLATE Latin1_General_BIN NULL,
	TRANSDATE nchar(15) COLLATE Latin1_General_BIN NULL,
	COMENTARIO nvarchar(255) COLLATE Latin1_General_BIN NULL,
	LOTE decimal(9,0) NULL,
	ASIENTO decimal(7,0) NULL,
	CONSTRAINT PK__GL_JD__DE77667F7D645F6D PRIMARY KEY (BATCHNBR,JOURNALID,TRANSNBR),
	CONSTRAINT FK_GL_DL_GL_JH FOREIGN KEY (BATCHNBR,JOURNALID) REFERENCES hencore_accpac.dbo.GL_JH(BATCHID,BTCHENTRY)
);

INSERT INTO hencore_accpac.dbo.AP_MP (BATCHTYPE,CNTBTCH,CNTRMIT,CNTLINE,IDACCT,GLREF,GLDESC,AMTDISTTC,AMTNETTC,COMPANYID,COMENTARIO,LOTE,ASIENTO) VALUES
	 (N'PY',1,4,20,N'140980                                       ',N'Mariela Peña Pinto - Hencorp Capital - CREDITO DIRECTO      ',N'APLP-ID# 100132 Principal payment CREDITO DIRECTO           ',1000.000,1000.000,N'HBCMX          ',NULL,NULL,NULL),
	 (N'PY',1,5,20,N'140980                                       ',N'Mariela Peña Pinto - Hencorp Capital - CREDITO DIRECTO      ',N'APLP-ID# 91993 Principal payment CREDITO DIRECTO            ',2000.000,2000.000,N'HBCMX          ',NULL,NULL,NULL),
	 (N'PY',1,6,20,N'230380                                       ',N'Banco de Desarrollo Multisectorial de El Salvador (BANDESAL)',N'APLender-ID# 1364 Interest payment                          ',6041.100,6041.100,N'HBCMX          ',NULL,NULL,NULL);

INSERT INTO hencore_accpac.dbo.AP_PA (BTCHTYPE,CNTBTCH,CNTENTR,IDRMIT,IDVEND,DATERMIT,TEXTRMIT,NAMERMIT,AMTRMIT,AMTRMITTC,RMITTYPE,DOCTYPE,CNTLSTLINE,FISCYR,FISCPER,AMTRMITHC,SWPRINTED,CHECKLANG,OPERBANK,OPERVEND,DATEACTVPP,SRCEAPPL,IDBANK,PAYMTYPE,AMTNETTC,COMPANYID,ESTADO,RESULTADO,FECHA,HORA,LOTE,ASIENTO,USUARIO,BATCHDESC,DATEBATCH,DATEPOST,REFERECE) VALUES
	 (N'PY',1,4,N'99250514    ',NULL,N'5/14/2025   ',N'APLP-ID# 100132 Principal payment                           ',N'Mariela Peña Pinto                                          ',1000.000,1000.000,1,1,1,N'2025',N'N5',1000.000,1,N'ENG',1,1,N'5/14/2025   ',N'AP',N'AGRHBCA ',2,1000.000,N'HBCMX          ',NULL,NULL,NULL,NULL,NULL,NULL,NULL,N'LTAP-CAP-Agricula-14-05-2025                                ',N'5/14/2025   ',N'5/14/2025   ',N'APLP-ID# 100132 Principal payment                           '),
	 (N'PY',1,5,N'99250514    ',NULL,N'5/14/2025   ',N'APLP-ID# 91993 Principal payment                            ',N'Mariela Peña Pinto                                          ',2000.000,2000.000,4,1,1,N'2025',N'N5',2000.000,1,N'ENG',1,1,N'5/14/2025   ',N'AP',N'AGRHBCA ',2,2000.000,N'HBCMX          ',NULL,NULL,NULL,NULL,NULL,NULL,NULL,N'LTAP-CAP-Agricula-14-05-2025                                ',N'5/14/2025   ',N'5/14/2025   ',N'APLP-ID# 91993 Principal payment                            '),
	 (N'PY',1,6,N'99250514    ',NULL,N'5/14/2025   ',N'APLender-ID# 1364 Interest payment                          ',N'Banco de Desarrollo Multisectorial de El Salvador (BANDESAL)',6041.100,6041.100,4,1,1,N'2025',N'N5',6041.100,1,N'ENG',1,1,N'5/14/2025   ',N'AP',N'AGRHBCA ',2,6041.100,N'HBCMX          ',NULL,NULL,NULL,NULL,NULL,NULL,NULL,N'LTAP-CAP-Agricula-14-05-2025                                ',N'5/14/2025   ',N'5/14/2025   ',N'APLender-ID# 1364 Interest payment                          ');

INSERT INTO hencore_accpac.dbo.AR_MR (CODEPAYM,CNTBTCH,CNTITEM,CNTLINE,AMTDIST,IDDISTCODE,IDACCT,GLREF,GLDESC,AMTDISTTC,AMTNETTC,COMPANYID,COMENTARIO,LOTE,ASIENTO) VALUES
	 (N'CA',1,1,20,215000.000,NULL,N'140580                                       ',N'Hencorp Capital - CREDITO DIRECTO                           ',N'AR-ID#31996 - Comercial Exportadora SA de CV - Pago Principa',215000.000,215000.000,N'HBCMX          ',NULL,NULL,NULL),
	 (N'CA',1,1,40,268.750,NULL,N'140680                                       ',N'Hencorp Capital - CREDITO DIRECTO                           ',N'AR-ID#31996 - Comercial Exportadora SA de CV - Pago de Inter',268.750,268.750,N'HBCMX          ',NULL,NULL,NULL),
	 (N'CA',1,2,20,2600.000,NULL,N'280180                                       ',N'Hencorp Capital - CREDITO DIRECTO                           ',N'AR-ID#32376 - SINGUIL de R.L. - Pago de Comision            ',2600.000,2600.000,N'HBCMX          ',NULL,NULL,NULL);

INSERT INTO hencore_accpac.dbo.AR_RA (CODEPYMTYP,CNTBTCH,CNTITEM,IDRMIT,IDCUST,DATERMIT,TEXTRMIT,TXTRMITREF,AMTRMIT,AMTRMITTC,CNTPAYMETR,AMTPAYMTC,CODEPAYM,CODECURN,RMITTYPE,DOCTYPE,IDINVCMTCH,CNTLSTLINE,FISCYR,FISCPER,TEXTPAYOR,DATERATETC,AMTRMITHC,DOCNBR,AMTADJHC,OPERBANK,OPERCUST,SRCEAPPL,IDBANK,CODECURNBC,AMTNETTC,COMPANYID,ESTADO,RESULTADO,FECHA,HORA,LOTE,ASIENTO,USUARIO,BATCHDESC) VALUES
	 (N'CA',1,1,N'91250514                ',NULL,N'05/14/2025  ',N'AR-ID#31996 - Comercial Exportadora SA de CV - Pago de Inter',N'Hencorp Capital CREDITO DIRECTO                             ',215268.750,215268.750,20,215268.750,N'WT          ',N'USD',5,1,NULL,1,N'25  ',N'5 ',N'Comercial Exportadora SA de CV                              ',N'5/14/2025   ',215268.750,NULL,NULL,1,1,N'AR',N'CUSHBCA ',N'USD',215268.750,N'HBCMX          ',NULL,NULL,NULL,NULL,NULL,NULL,NULL,N'HBC-14/05/2025                                              '),
	 (N'CA',1,2,N'91250514                ',NULL,N'05/14/2025  ',N'AR-ID#32376 - SINGUIL de R.L. - Pago de Comision            ',N'Hencorp Capital CREDITO DIRECTO                             ',2600.000,2600.000,20,2600.000,N'WT          ',N'USD',5,1,NULL,1,N'25  ',N'5 ',N'SINGUIL de R.L.                                             ',N'5/14/2025   ',2600.000,NULL,NULL,1,1,N'AR',N'CUSHBCA ',N'USD',2600.000,N'HBCMX          ',NULL,NULL,NULL,NULL,NULL,NULL,NULL,N'HBC-14/05/2025                                              ');

