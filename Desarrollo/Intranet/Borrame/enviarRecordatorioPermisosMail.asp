<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->

<%
'Leer todos los registros de sectores.
'Si hay alguno, que no se ha confirmado a tiempo, mandar de vuelta mail o bien al mismo usuario o bien al usuario alternativo, si la cantidad de ConfPendientes es mayor al permitido

	dim strSql, oConn, g_rsConfirmaciones
	dim mailResponsable, asuntoMail, descripcionMail
	dim kmResponsable, kcResponsable, dsResponsable
	dim krRelacionJefe, strSQLResponsableSector, g_rsResponsableSector
	dim dsSector
	
	GP_CONFIGURARMOMENTOS

	Call GF_MGKS ("SR", "ESJEFEDE", krRelacionJefe, dsRelacionJefe)	
	strSQL= "select * from ConfirmacionesPermisos where mmtoenvio=(Select max(mmtoEnvio) from ConfirmacionesPermisos) order by MmtoEnvio desc"
	'Response.Write strSQL
	call GF_BD_CONTROL (g_rsConfirmaciones,oConn,"OPEN",strSQL)
	if g_rsConfirmaciones.eof then
		response.write "Recordset de Confirmaciones es vacio"
	else
		while not g_rsConfirmaciones.eof
		'response.write g_rsConfirmaciones ("KrSector") &  " " & g_rsConfirmaciones ("ConfPendientes") & "<br>"
					
			if g_rsConfirmaciones("ConfPendientes") > 0 and g_rsConfirmaciones("ConfPendientes")<=3 then
					'response.write " entro en > 0 y menor < 3"
					Call enviarRecordatorioConfirmacion()				
					'actualizar campo ConfPendientes (sumar 1)	
					Call actualizarConfirmacionesPendientes(g_rsConfirmaciones("KrSector"))
			else 			
				if g_rsConfirmaciones("ConfPendientes") > 3 then
					'enviar mail de recordatorio de confirmacion de permisos al responsable de sector
					Call enviarRecordatorioConfirmacion()
					'enviar mail de notificacion de falta de confirmacion a administrador
					Call enviarNotificacionFaltaConfirmacion()
				end if
			end if
			g_rsConfirmaciones.movenext
		wend
	end if
'-----------------------------------------------------------------------------------------
function getCuerpoMensaje(p_sector, p_responsable)
	dim cuerpoMensaje
	cuerpoMensaje = "El soporte técnico de la empresa requiere que se confirmen los permisos asignados al personal de su sector sobre los recursos y aplicaciones informáticas, por favor ingrese a "
	cuerpoMensaje = cuerpoMensaje & "http://bai-des-1/ActisaIntra/AUPAuditoria.asp?pSector=" & p_sector & "&p_responsable=" & p_responsable &"&p_accion=CONFIRMA "
	cuerpoMensaje = cuerpoMensaje & " para confirmar esta información." & chr(13) & chr(13)
	cuerpoMensaje = cuerpoMensaje & "Gracias"  & chr(13) & chr(13) & "Soporte tecnico." & chr(13) & "Alfred C. Toepfer Intl. Arg. S.R.L."
	getCuerpoMensaje = cuerpoMensaje
end function
'-----------------------------------------------------------------------------------------
function actualizarConfirmacionesPendientes(p_sector)
	dim sql, Conn, rs
	'Response.Write "<HR>Select * from ConfirmacionesPermisos where KrSector=" & p_sector & " and MmtoEnvio = (Select max(MmtoEnvio) from ConfirmacionesPermisos where KrSector=" & p_sector & ")<br>"
	sql = "UPDATE ConfirmacionesPermisos SET ConfPendientes = ConfPendientes + 1 where KrSector=" & p_sector & " and MmtoEnvio = (Select max(MmtoEnvio) from ConfirmacionesPermisos where KrSector=" & p_sector & ")" 
	'response.write "<br> Va a ejecutar esto " & sql
	call GF_BD_CONTROL (rs,oConn,"EXEC",sql)
end function
'-----------------------------------------------------------------------------------------
function enviarRecordatorioConfirmacion()
	'traer el responsable del sector actual
	strSQLResponsableSector= "select sro2km, sro2kc, sro2ds, sro2kr, Sectores.Sector as Sector from RelacionesConsulta, (select Sector from Profesionales group by Sector) as Sectores where RelacionesConsulta.sro3kr = Sectores.Sector and RelacionesConsulta.sro1kr = " & krRelacionJefe & " and Sectores.Sector=" & g_rsConfirmaciones("KrSector")
	call GF_BD_CONTROL (g_rsResponsableSector,oConn,"OPEN",strSQLResponsableSector)
	if g_rsResponsableSector.eof then
		response.write " Sector con Kr " & g_rsConfirmaciones("KrSector") & " no posee un responsable"
	else
		'enviar mail de recordatorio de confirmacion de permisos al responsable del sector
		'conseguir mail del responsable
		mailResponsable = GF_DT1 ("READ", "SGEMAIL", "", "", g_rsResponsableSector("sro2km"), g_rsResponsableSector("sro2kc"))
		'armar cuerpo del mensaje
		descripcionMail = getCuerpoMensaje(g_rsConfirmaciones("KrSector"), g_rsResponsableSector("sro2kr"))
		'armar asunto de mail
		asuntoMail = "Recordatorio de la confirmacion de Permisos de Sector "
		Call GF_MGKR (CInt(g_rsConfirmaciones("KrSector")), "", "", dsSector)
		asuntoMail = asuntoMail & dsSector
		'enviar mail
		'mailResponsable = "bacariniE@toepfer.com"
		Call GP_ENVIAR_MAIL (asuntoMail, descripcionMail,"Soporte.Ar@toepfer.com",mailResponsable)
		response.write "<br>Mail enviado a: " & mailResponsable & "<br>"
	end if	
end function
'-----------------------------------------------------------------------------------------
function enviarNotificacionFaltaConfirmacion()
	mailAdministrador = GF_DT1 ("READ", "SGEMAIL", "", "", "SP", "AUPSEC")
	descripcionMail = "No se han confirmado los permisos de sector " & dsSector & ", responsable: " & g_rsResponsableSector("sro2ds")
	asuntoMail = "Notificacion de falta de confirmacion de permisos del Sector " & dsSector
	Call GP_ENVIAR_MAIL (asuntoMail, descripcionMail,"Soporte.Ar@toepfer.com",mailAdministrador)
end function
'-----------------------------------------------------------------------------------------
%>
</body>
</html>
