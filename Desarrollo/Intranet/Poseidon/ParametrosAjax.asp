<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<%
Function eliminarParametro(ppto, pcodParam)
	dim strSQL, rs
	strSQL = "DELETE  FROM parametros WHERE cdparametro = '"& pcodParam &"'"
	GF_BD_Puertos ppto, rs, "EXEC",strSql 		
End function
Function eliminarParametroExtra(ppto, pcodParam)
	dim strSQL, rs
	strSQL = "DELETE  FROM TBLPARAMETROSEXTRA WHERE cdparametro = '"& pcodParam &"'"
	GF_BD_Puertos ppto, rs, "EXEC",strSql 		
End function
'-----------------------------------------------------------------------------------
Function updateParametro(pcdParam,pnomParam,pvalParam,ppto)
	dim strSQL, rs
	strSQL = "Update PARAMETROS set DSPARAMETRO ='" & pnomParam &"',VLPARAMETRO ='"& pvalParam &"' where CDPARAMETRO='"& pcdParam &"'" 
	GF_BD_Puertos ppto, rs, "EXEC",strSQL 		
End function
'--------------------------------------------------------------------------------------------------
Function guardarParametro(pcdParam,pnomParam,pvalParam,ppto)
	dim strSQL, rs
	strSQL = "INSERT INTO PARAMETROS(CDPARAMETRO,DSPARAMETRO,VLPARAMETRO) "
	strSQL = strSQL & " VALUES ('"& pcdParam &"','"& pnomParam &"','"& pvalParam &"')"	
	GF_BD_Puertos ppto, rs, "EXEC",strSQL 		
End Function
'--------------------------------------------------------------------------------------------------
Function guardarParametroExtra(pcdParam,peditable,ppuesto,ppto)
	dim strSQL, rs
	strSQL = "INSERT INTO TBLPARAMETROSEXTRA(CDPARAMETRO,EDITABLE,PUESTO) "
	strSQL = strSQL & " VALUES ('"& pcdParam &"','"& peditable &"',"& ppuesto &")"	
	GF_BD_Puertos ppto, rs, "EXEC",strSQL 		
End Function
'--------------------------------------------------------------------------------------------------
Function updateParametroExtra(pcdParam,peditable,ppuesto,ppto)
	dim strSQL, rs
	strSQL = "Update TBLPARAMETROSEXTRA set EDITABLE ='" & peditable &"',PUESTO ="& ppuesto &" where CDPARAMETRO='"& pcdParam &"'"
	GF_BD_Puertos ppto, rs, "EXEC",strSQL 		
End function
'--------------------------------------------------------------------------------------------------
Dim accion,nomParam,cdParam,valParam,ppto,esedit,idpuesto,rsPar,nomParam_old,valParam_old,esedit_old,idpuesto_old
Dim v_puerto,myHoy 


cdParam = Trim(GF_PARAMETROS7("cdParam", "", 6))
nomParam = Trim(GF_PARAMETROS7("nomParam", "",6))
valParam = Trim(GF_PARAMETROS7("valParam", "",6))
accion = GF_PARAMETROS7("accion",0,6)
ppto = GF_PARAMETROS7("pto","",6)
esedit = GF_PARAMETROS7("editable","",6)
idpuesto = GF_PARAMETROS7("puesto",0,6)
nomParam_old = GF_PARAMETROS7("nomParam_old", "",6)
valParam_old = GF_PARAMETROS7("valParam_old", "",6)
esedit_old = GF_PARAMETROS7("editable_old","",6)
idpuesto_old = GF_PARAMETROS7("puesto_old",0,6)


Set logParam = new classLog
if(esedit)then 
	esedit = PARAMETRO_EDITABLE
else
	esedit = PARAMETRO_NO_EDITABLE
end if
select case UCASE(ppto)
	case TERMINAL_TRANSITO
		v_puerto = NOMBRE_RUTA_PARAMETRO_TRA
	case TERMINAL_ARROYO		
		v_puerto = NOMBRE_RUTA_PARAMETRO_ARR
	case TERMINAL_PIEDRABUENA
		v_puerto = NOMBRE_RUTA_PARAMETRO_LPB
end select

myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))
call startLog(HND_FILE, MSG_INF_LOG)
logParam.fileName = v_puerto & myHoy


'**************************************************************************
if(accion = ACCION_COMPROBAR_PARAMETRO)then 
set rsPar = leerParametros(ppto,cdParam,"",0,"",false)
	if rsPar.eof then	
	v_parametroNoExistente = PARAMETRO_NO_EXISTENTE	
	call guardarParametro(cdParam,nomParam,valParam,ppto)	
	'Si todo esta bien para guardar registro en el log que se dio de alta tal parametro	
	logParam.info(" ALTA - " & cdParam & " - Parametro '" & cdParam & "' creado!")
	logParam.info(" ALTA - " & cdParam & " - Descripcion: '"& nomParam&"'")
	logParam.info(" ALTA - " & cdParam & " - Valor: '"& valParam &"'")	
		
	response.write v_parametroNoExistente
	end if	
response.end
end if	
'--------------------------------------------------------------
if(accion = ACCION_MODIFICAR_PARAMETRO)then 
	Call updateParametro(cdParam,nomParam,valParam,ppto)
	if(nomParam_old <> nomParam)then 		
		logParam.info(" MODIFICACION - " & cdParam & " - Se modifico la descripcion de '"&nomParam_old&"' a '"&nomParam&"'")
		if(valParam_old <> valParam)then
			logParam.info(" MODIFICACION - " & cdParam & " - Se modifico el valor del parametro de '"&valParam_old&"' a '"&valParam&"'")
		end if	
	else
		if(valParam_old <> valParam)then
			logParam.info(" MODIFICACION - " & cdParam & " - Se modifico el valor del parametro de '"&valParam_old&"' a '"&valParam&"'")
		end if	
	end if		
	response.end
end if
'--------------------------------------------------------------
if(accion = ACCION_MODIFICAR_PARAMETRO_EXTRA)then 
	if(tieneParametrosExtra(cdParam, ppto))then	
	'comprueba si tiene parametro extra cargado, caso contrario lo crea nuevo
		Call updateParametroExtra(cdParam,esedit,idpuesto,ppto)	
		if(esedit <> esedit_old)then 			
			logParam.info(" MODIFICACION - " & cdParam & " - Se modifico si es editable de '"&esedit_old&"' a '"&esedit&"'")
			if(idpuesto_old <> idpuesto)then 
				logParam.info(" MODIFICACION - " & cdParam & " - Se modifico el Puesto de '"& obtenerNombrePuesto(idpuesto_old,ppto) &"' a '"& obtenerNombrePuesto(idpuesto,ppto) &"'")
			end if			
		else
			if(idpuesto_old <> idpuesto)then 				
				logParam.info(" MODIFICACION - " & cdParam & " - Se modifico el Puesto de '"& obtenerNombrePuesto(idpuesto_old,ppto) &"' a '"& obtenerNombrePuesto(idpuesto,ppto) &"'")
			end if
		end if				
	else
		'si el parametro que se quiere modificar antes nunca cargo el EXTRA, y ahora lo 
		' quiere poner entonces se tendra que dar de alta en la Tabla Extra
		call guardarParametroExtra(cdParam,esedit,idpuesto,ppto)
		if(esedit <> esedit_old)then 
			logParam.info(" MODIFICACION - " & cdParam & " - Se cargo el campo Editable con valor '"&esedit&"'")
			if(idpuesto_old <> idpuesto)then 
				logParam.info(" MODIFICACION - " & cdParam & " - Se cargo el Puesto con valor '"& obtenerNombrePuesto(idpuesto,ppto) &"'")
			end if			
		else
			logParam.info(" MODIFICACION - " & cdParam & " - Se cargo el Puesto con valor '"& obtenerNombrePuesto(idpuesto,ppto) &"'")
		end if			
	end if
response.end
end if
'--------------------------------------------------------------
if(accion = ACCION_AGREGAR_PARAMETRO_EXTRA)then	
	'para un nuevo parametro que cargo todo(cabecera y extra)
	call guardarParametroExtra(cdParam,esedit,idpuesto,ppto)
	logParam.info(" ALTA - " & cdParam & " - Editable: '"& esedit &"'")
	logParam.info(" ALTA - " & cdParam & " - Puesto: '"& obtenerNombrePuesto(idpuesto,ppto) &"'")
	response.end	
end if
'--------------------------------------------------------------
if(accion = ACCION_ELIMINAR_PARAMETRO)then
	call eliminarParametro(ppto, cdParam)
	if(tieneParametrosExtra(cdParam, ppto))then
		call eliminarParametroExtra(ppto, cdParam)
	end if
	logParam.info(" BAJA -  " & cdParam & " - Se dió de baja el parametro!")	
	response.end	
end if

'**************************************************************************

%>