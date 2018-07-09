<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<%
'-----------------------------------------------------------------------------
Function getCalidad(pPto, pIDCamion, pDTContable)
	Dim strSQL, rsCalidad, rtrn
	
	rtrn = ""	
	strSQL = "select dsrubro, VLBONREBAJA , VLMERMA " &_
			 "from (Select convert(date, getdate(), 126) dtcontable, A.* from RUBROSVISTEOCAMIONES A union Select * from HRUBROSVISTEOCAMIONES B) HRVC " & _
			 "   inner join dbo.RUBROS RU on HRVC.cdrubro=ru.cdrubro " & _
			 "   where idcamion='" & pIDCamion & "' and dtcontable='" & pDTContable & "'" & _
			 "      and hrvc.sqcalada = (select max(sqcalada) from (Select convert(date, getdate(), 126) dtcontable, A.* from RUBROSVISTEOCAMIONES A union Select * from HRUBROSVISTEOCAMIONES B) T where dtcontable=hrvc.DTCONTABLE and IDCAMION=hrvc.IDCAMION)"	
	Call executeQueryDb(pPto, rsCalidad, "OPEN", strSQL)
	While Not rsCalidad.EOF
		rtrn = rtrn & Trim(rsCalidad("dsrubro")) & " " & Trim(rsCalidad("VLBONREBAJA")) & " (" & Trim(rsCalidad("VLMERMA")) & "%) - "
		rsCalidad.MoveNext
	Wend
	If Len(rtrn) > 0 Then rtrn = Left(rtrn, Len(rtrn) - 2)
	getCalidad = rtrn
End Function
'-----------------------------------------------------------------------------
Function drawTable(rs)
	Dim datos, pesoBruto, pesoTara
	
	while not rs.eof								        		       
	
		datos = rs("CIRCUITO") & ":center|"
		datos = datos & GF_FN2DTE(rs("DTINGRESO"))  & ":center|"
		datos = datos & rs("PRODUCTODS") & "|"
		datos = datos & rs("CDCHAPACAMION") & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("REMITENTECUIT"))) & " " & Trim(rs("REMITENTEDS")) & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("CYO1CUIT")))  & " " & Trim(rs("CYO1DS")) & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("CYO2CUIT")))  & " " & Trim(rs("CYO2DS")) & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("CorredorCUIT"))) & " " & Trim(rs("CorredorDs")) & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("ClienteCUIT"))) & " " & Trim(rs("ClienteDS")) & "|"
		datos = datos & rs("PROCEDENCIADS") & "|"
		datos = datos & rs("ESTADODS") & "|"
		datos = datos & rs("ACEPTACIONDS") & "|"
		if ((CInt(rs("CDACEPTACION")) = ACEPTACION_CONFORME) or (CInt(rs("CDACEPTACION")) = ACEPTACION_AUT_ENTREGADOR) or (CInt(rs("CDACEPTACION")) = ACEPTACION_REBAJA_CONVENIDA) or (CInt(rs("CDACEPTACION")) = ACEPTACION_ANALISIS)) then
			datos = datos & rs("GRADO") & ":center|"
		else
			datos = datos & "|"
		end if
		datos = datos & getCalidad(pto, rs("IDCamion"),  rs("DTContable")) & "|"
		datos = datos & rs("CARTAPORTE") & "|"
		datos = datos & GF_STR2CUIT(Trim(rs("ENTREGADORCUIT"))) & " " & Trim(rs("ENTREGADORDS")) & "|"
		datos = datos & rs("vlNetoOrigen") & "|"
		pesoBruto = 0
		if (rs("VLBRUTOBZA") <> "") then pesoBruto = CDbl(rs("VLBRUTOBZA"))
		pesoTara = 0
		if (rs("VLTARABZA") <> "") then pesoTara = CDbl(rs("VLTARABZA"))
		datos = datos & pesoBruto & "|"
		datos = datos & pesoTara & "|"
		datos = datos & pesoBruto - pesoTara & "|"
		datos = datos & rs("CTG") & "|"
		datos = datos & Left(rs("COSECHA"), 4) & "/" & Right(rs("COSECHA"), 4) & "|"
		datos = datos & rs("NUCUPO")
		Call XLS_AddDataRow(datos, "|")		
		rs.MoveNext()
	wend
End Function
'-----------------------------------------------------------------------------
Function generateReportFile(pto, entregador, fechaDesde, fechaHasta)

	Dim filename, rsC, sp_ret, rsD, rsDH, rsCH

	logMig.info("Inicia generacion de reporte.")
	
	Set sp_ret = executeProcedureDb(pto, rsD, "CAMIONESDESCARGA_X_ENTREGADOR", entregador & "||" & fechaDesde & "||" & fechaHasta)
	logMig.info("CamionesDescarga leida OK.")
	Set sp_ret = executeProcedureDb(pto, rsC, "CAMIONESCARGA_X_ENTREGADOR", entregador & "||" & fechaDesde & "||" & fechaHasta)	
	logMig.info("CamionesCarga leida OK.")
	Set sp_ret = executeProcedureDb(pto, rsDH, "HCAMIONESDESCARGA_X_ENTREGADOR", entregador & "||" & fechaDesde & "||" & fechaHasta)
	logMig.info("HCamionesCarga leida OK.")
	Set sp_ret = executeProcedureDb(pto, rsCH, "HCAMIONESCARGA_X_ENTREGADOR", entregador & "||" & fechaDesde & "||" & fechaHasta)	
	logMig.info("HCamionesCarga leida OK.")
	
	generateReportFile = ""	
	if ((not rsC.eof) or (not rsD.eof) or (not rsCH.eof) or (not rsDH.eof)) then 				
		filename = "Camiones_" & pto & "_" & entregador & "_" & fechaDesde		
		xls_mode = XLS_FILE_MODE
		generateReportFile = XLS_Open(filename)
		Call XLS_AddDataTitle("CIRCUITO,INGRESO,PRODUCTO,PATENTE,TITULAR,INTERMEDIARIO,RTE COMERCIAL,CORREDOR,DESTINATARIO,PROCEDENCIA,ESTADO,ACEPTACION,GRADO,CALIDAD/RUBROS,CARTA PORTE,ENTREGADOR,NETO ORIGEN,BRUTO BALANZA,TARA BALANZA,NETO BALANZA,CTG,COSECHA,NU CUPO", "")
		Call drawTable(rsD)
		Call drawTable(rsDH)
		Call drawTable(rsC)
		Call drawTable(rsCH)
		Call XLS_closeXLS()
	end if
	
	logMig.info("Reporte Generado")
End Function
'**************************************************************
'***				COMIENZO DE LA PAGINA					***
'**************************************************************
Dim entregador, pto, fechaDesde, fechaHasta, myReportFile, logMig

pto = GF_PARAMETROS7("pto", "", 6)
entregador = GF_PARAMETROS7("ent", "", 6)
fechaDesde = GF_PARAMETROS7("fd", "", 6)
fechaHasta = GF_PARAMETROS7("fh", "", 6)

Call GP_CONFIGURARMOMENTOS()

if (fechaDesde = "") then fechaDesde = Left(session("MmtoDato"), 8)
if (fechaHasta = "") then fechaHasta = fechaDesde

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "ENTRAGADORES_DESCARGAS_" & left(session("MmtoDato"),8)

logMig.info("********** INICIA PROCESO  (Tarea: " & TASK_POS_DESC_X_ENT & ") **********")
logMig.info("Puerto: " & pto)
logMig.info("Desde : " & GF_FN2DTE(fechaDesde))
logMig.info("Hasta : " & GF_FN2DTE(fechaHasta))

myReportFile = generateReportFile(pto, entregador, fechaDesde, fechaHasta)

if (myReportFile <> "") and (xls_mode = XLS_FILE_MODE)	then
	logMig.info("Enviando mails....")
	Call SendMail(TASK_POS_DESC_X_ENT, entregador, "ADM Agro S.R.L. - Info Entregador " & pto & " - " & GF_FN2DTE(fechaDesde), "Se adjunta el archvio con las descargas.", myReportFile)    
	logMig.info("Mail enviados OK!")
end if

logMig.info("***************     FIN     ***************")
%>
