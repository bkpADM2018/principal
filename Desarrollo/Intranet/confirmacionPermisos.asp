<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->

<html>
<head>
<%
	dim strSql, oConn, rs
	sector=GF_Parametros7("p_sector","",6)
	responsable = GF_Parametros7("p_responsable","",6)	
	strSQL= "update ConfirmacionesPermisos set ConfPendientes=ConfPendientes-1, KrUltimoUsuario=" & responsable & ", MmtoConf='" & session("MmtoSistema") & "' where KrSector=" & sector & " and MmtoEnvio = (Select max(MmtoEnvio) from ConfirmacionesPermisos where KrSector=" & sector & ")" 
	'Response.Write strSQL
	call GF_BD_CONTROL (rs,oConn,"EXEC",strSQL)		
%>
<script type="text/javascript">
	window.close();
</script>
</head>
<body>
</body>
</html>
