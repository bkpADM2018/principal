<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->

<%

	dim strSql, oConn, krRelacionJefe, dsRelacionJefe, g_rsResponsablesSectores
	dim mailResponsable, asuntoMail, descripcionMail
	dim kcSector, kmSector, dsSector
	
	Call GP_CONFIGURARMOMENTOS()
	
	Call GF_MGKS ("SR", "ESJEFEDE", krRelacionJefe, dsRelacionJefe)
	strSQL= "select sro2km, sro2kc, sro2ds, sro2kr, Sectores.Sector as Sector from RelacionesConsulta, (select Sector from Profesionales group by Sector) as Sectores where RelacionesConsulta.sro3kr = Sectores.Sector and RelacionesConsulta.sro1kr = " & krRelacionJefe
	'response.write strSQL
	call GF_BD_CONTROL (g_rsResponsablesSectores,oConn,"OPEN",strSQL)
	if g_rsResponsablesSectores.eof then
		response.write "Recordset de Sectores es vacio"
	else
		while not g_rsResponsablesSectores.eof 			
			if esMomentoDeEnviarPlanilla(g_rsResponsablesSectores("Sector")) then 
				'tomar mail responsable sector			
				mailResponsable = GF_DT1 ("READ", "SGEMAIL", "", "", g_rsResponsablesSectores("sro2km"), g_rsResponsablesSectores("sro2kc"))
				'armar cuerpo de mensaje
				'response.write mailResponsable & "<br>"
				descripcionMail = getCuerpoMensaje(g_rsResponsablesSectores("Sector"), g_rsResponsablesSectores("sro2kr"))
				'armar asunto de mail
				asuntoMail = "Confirmacion de Permisos de Sector "
				Call GF_MGKR (CInt(g_rsResponsablesSectores("Sector")), "", "", dsSector)
				asuntoMail = asuntoMail & dsSector
				'enviar mail
				mailResponsable = "scalisij@toepfer.com"
				Call GP_ENVIAR_MAIL (asuntoMail, descripcionMail,"Soporte.Ar@toepfer.com",mailResponsable)
				'registrar confirmacion pendiente...
				Call registrarConfirmacionPendiente(g_rsResponsablesSectores("Sector"),g_rsResponsablesSectores("sro2kr"))
			end if
			g_rsResponsablesSectores.movenext
		wend
	end if
'------------------------------------------------------------------------------------------
function getCuerpoMensaje(p_sector, p_responsable)
	dim cuerpoMensaje
	cuerpoMensaje = "El soporte técnico de la empresa requiere que se confirmen los permisos asignados al personal de su sector sobre los recursos y aplicaciones informáticas, por favor ingrese a "
	cuerpoMensaje = cuerpoMensaje & "http://bai-vm-intra-1/ActisaIntra/AUPAuditoria.asp?pSector=" & p_sector & "&p_responsable=" & p_responsable &"&p_accion=CONFIRMA "
	cuerpoMensaje = cuerpoMensaje & " para confirmar esta información." & chr(13) & chr(13)
	cuerpoMensaje = cuerpoMensaje & "Gracias"  & chr(13) & chr(13) & "Soporte tecnico." & chr(13) & "Alfred C. Toepfer Intl. Arg. S.R.L."
	getCuerpoMensaje = cuerpoMensaje
end function
'------------------------------------------------------------------------------------------
function registrarConfirmacionPendiente (p_sector, p_responsable)
	dim strSql, oConn, rs
	strSql = "Insert Into ConfirmacionesPermisos values(" & p_sector & ", " & p_responsable & ", 1,null, '" & session("MmtoSistema") & "')"
	'Response.Write "<hr>" & strSql
	call GF_BD_CONTROL (rs,oConn,"EXEC",strSql)
end function
'------------------------------------------------------------------------------------------
function esMomentoDeEnviarPlanilla(p_sector)
	'controlar, si ya han pasado 6 meses desde que el responsable confirmo la planilla
	dim sql, oConn, rs
	sql = "select ConfPendientes, MmtoEnvio from ConfirmacionesPermisos where KrSector = " & p_sector & " order by MmtoEnvio desc"
	'response.write sql
	call GF_BD_CONTROL (rs,oConn,"OPEN",sql)
	if not rs.eof then
		'response.write "<br>" & rs("ConfPendientes") & "/" & rs("MmtoEnvio") & "/" & session("MmtoSistema")
		'response.end
		if (rs("ConfPendientes") = 0) or (GF_DTEDIFF(rs("MmtoEnvio"),session("MmtoSistema"),"M") >= 4) then 
			esMomentoDeEnviarPlanilla = true
		else
			esMomentoDeEnviarPlanilla = false
		end if
	else
		esMomentoDeEnviarPlanilla = true
	end if	
end function
%>
</body>
</html>
