<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosFacturacionCalidad.asp"-->
<%

'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'****************************************************

Dim pto, myHoy, logMig, flagBach, myHasta, strSQL, rs, myDt, rs2, strSQLH, myTipCbt
Dim rsCtrl, msg
'On Error Resume Next

Call GP_CONFIGURARMOMENTOS

if (session("usuario") = "") then session("usuario") = "SYNC"
pto = GF_PARAMETROS7("pto", "", 6)
myHoy = Left(session("MmtoDato"), 8)

Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "FACT-ACOND-SYNC-" & pto & "-" & myHoy	

logMig.info("####################################################")
logMig.info("		-PUERTO       :  "& pto)
logMig.info("		-MOMENTO      :  "& GF_FN2DTE(Left(session("MmtoDato"),8)))
logMig.info("		-USUARIO      :  "& session("Usuario"))	
logMig.info("####################################################")		

	
	'Se carga la tabla de conversiones para el cliente
    Call cargarTablaConversion(CUIT_ADM, pto)      
	
	'myDt = GF_FN2DTCONTABLE(GF_DTEADD(myHoy, -3, "M"))
	myDt = "2017-09-13"
	strSQL= "Select Year(DTCONTABLE)*10000+MONTH(DTCONTABLE)*100+DAY(DTCONTABLE) FECHA,* from FACTURACIONSERVICIOS FS" &_
			" where (ESTADO = " & FACTURA_CALIDAD_PROFORMA_PTO &_
			"		or (ESTADO = " & FACTURA_CALIDAD_PROFORMA_BSAS & " and SUCCBT=0)" &_
			"		or (ESTADO = " & FACTURA_CALIDAD_PRE_CANCELADA & "))" &_
			"		and DTCONTABLE >= '" & myDt & "'"
	Call executeQueryDb(pto, rs, "OPEN", strSQL)
	while (not rs.eof)				
		'Averiguo si la carta de porte - Servicio ya fue migrada a Bs As, si no exsiste la inserto.
		strSQL = "Select * from FAC002A where IDREGISTRO=" & rs("IDREGISTRO") & " and codcia='" & rs("codcia") & "'"		
		Call executeQueryDb(DBSITE_SQL_MAGIC, rs2, "OPEN", strSQL)
		if (rs2.eof) then
			'La tiene solo el puerto, si no esta cancelado lo doy de alta en Bs As.
			if (CInt(rs("ESTADO")) <> FACTURA_CALIDAD_PRE_CANCELADA) then
				strSQL = "Insert into FAC002A([IDREGISTRO], [tipoTransporte], [dtContable], [nudocumento], [IDTransporte], [codconce], [descripcion], [cdProducto], [cuitCliente], [cdRubro], [vlRubro], [ptoCalidad], [kilos], [merma], [codmone], [precio], [importe], [codcia], [tipcbt], [letra], [succbt], [nrocbt], [fecalta], [usualta], [estado]) " &_
					 " values(" & rs("IDREGISTRO") & ", " & rs("tipoTransporte") & ", '" & GF_FN2DTCONTABLE(rs("fecha")) & "', '" & rs("nudocumento") & "', '"& rs("IDTransporte") &"', " & rs("codconce") & ", '" & rs("descripcion") & "', " & convertirDatoPuerto(CONV_KEY_PRODUCTO, rs("cdProducto"), msg) & ", '" & rs("cuitCliente") & "', " & rs("cdRubro") & ", " & rs("vlRubro") & ", " & rs("ptoCalidad") & "," & rs("kilos") & ", " & rs("merma") & ", " & rs("codmone") & ", " & rs("precio") & ", " & rs("importe") & ", '" & rs("codcia") & "', " & rs("tipcbt") & ", '" & rs("letra") & "', " & rs("succbt") & ", " & rs("nrocbt") & ", '" & GF_FN2DTCONTABLE(rs("fecha")) & "', '" & rs("usualta") & "', " & FACTURA_CALIDAD_PROFORMA_BSAS &  ")"	
			    logMig.info(strSQL)
				Call executeQueryDb(DBSITE_SQL_MAGIC, rsX, "EXEC", strSQL)
				strSQL="Update FACTURACIONSERVICIOS set ESTADO=" & FACTURA_CALIDAD_PROFORMA_BSAS & " where IDREGISTRO=" & rs("IDREGISTRO")
				logMig.info(strSQL)
				Call executeQueryDb(pto, rsX, "EXEC", strSQL)
			else
				'Se pre-cancelo en puerto, esto solo pasa si ya vino a bs as, es raro que no la encuentre.
				'Este caso no deberia darse nunca, pero por las dudas se implementa.
				strSQL="Update FACTURACIONSERVICIOS set ESTADO=" & FACTURA_CALIDAD_CANCELADA & " where IDREGISTRO=" & rs("IDREGISTRO")
				logMig.info(strSQL)
				Call executeQueryDb(pto, rsX, "EXEC", strSQL)
			end if
		else
			'El registro lo tiene Bs As!
			'Si esta facturado, actualizo el valor de la factura en el puerto.
			if (CInt(rs2("SUCCBT")) > 0) then								
				'Si ya se facturo pero en el puerto se cancelo, entonces se va a actualizar la factura ya emitida y se dejara preparada la nota de crédito pra anular dicha factura.				
				if (CInt(rs("ESTADO")) = FACTURA_CALIDAD_PRE_CANCELADA) then					
					myTipCbt = TIPO_CBTE_EMITIDO_NCR
					if (CInt(rs("tipcbt")) = TIPO_CBTE_EMITIDO_NCR) then myTipCbt = TIPO_CBTE_EMITIDO_FAC	
					strSQL="Insert into FACTURACIONSERVICIOS([tipoTransporte], [dtContable], [nudocumento], [IDTransporte], [codconce], [descripcion], [cdProducto], [cuitCliente], [cdRubro], [vlRubro], [ptoCalidad], [kilos], [merma], [codmone], [precio], [importe], [codcia], [tipcbt], [letra], [succbt], [nrocbt], [tipcbtrel], [letrarel], [succbtrel], [nrocbtrel], [fecalta], [usualta], [estado]) " &_
							" Select [tipoTransporte], [dtContable], [nudocumento], [IDTransporte], [codconce], [descripcion], [cdProducto], [cuitCliente], [cdRubro], [vlRubro], [ptoCalidad], [kilos], [merma], [codmone], [precio], [importe], [codcia], " & myTipCbt & ", '', 0, 0, " & rs2("TIPCBT") & ", '" & rs2("letra") & "', " & rs2("SUCCBT") & ", " & rs2("NROCBT") & ", '" & GF_FN2DTCONTABLE(myHoy) & "', '" & session("usuario") & "', " & FACTURA_CALIDAD_PROFORMA_PTO & " from FACTURACIONSERVICIOS where IDREGISTRO=" & rs("IDREGISTRO") 
					logMig.info(strSQL)
					Call executeQueryDb(pto, rsX, "EXEC", strSQL)
				end if		
				strSQL="Update FACTURACIONSERVICIOS set letra='" & rs2("letra") & "', SUCCBT=" & rs2("SUCCBT") & " , NROCBT=" & rs2("NROCBT") & ", ESTADO=" & FACTURA_CALIDAD_FACTURADA & " where IDREGISTRO=" & rs("idregistro")
				logMig.info(strSQL)
				Call executeQueryDb(pto, rsX, "EXEC", strSQL)				
			else
				'No se facturó - Si esta cancelado lo elimino de Bs As.
				if (CInt(rs("ESTADO")) = FACTURA_CALIDAD_PRE_CANCELADA) then
					strSQL="Delete from FAC002A  where IDREGISTRO=" & rs("IDREGISTRO") & " and codcia='" & rs("codcia") & "'"
					logMig.info(strSQL)
					Call executeQueryDb(DBSITE_SQL_MAGIC, rsX, "EXEC", strSQL)
					strSQL="Update FACTURACIONSERVICIOS set ESTADO=" & FACTURA_CALIDAD_CANCELADA & " where IDREGISTRO=" & rs("IDREGISTRO")
					logMig.info(strSQL)
					Call executeQueryDb(pto, rsX, "EXEC", strSQL)
				end if
			end if
		end if
		rs.MoveNext()
	wend		

logMig.info("####################################################")
logMig.info("--- 			  FIN DEL PROCESO 				  ---")
logMig.info("####################################################")	

%>