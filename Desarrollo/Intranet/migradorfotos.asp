<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosXML.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<% call ProcedimientoControl("MIGFOTOS") 
'**********************************************************************************************
sub levantarParametros(p_accion, p_prefijo, p_sucursal, p_anio, p_inicio)	
	p_accion = GF_Parametros7("accion","",6)
	p_prefijo = GF_Parametros7("prefijo","",6)
	p_sucursal = GF_Parametros7("sucursal","",6)
	p_anio = GF_Parametros7("anio", "", 6)
	p_inicio = GF_Parametros7("inicio","",6)
end sub
'**********************************************************************************************
sub reemplazarNombres(p_prefijo, p_sucursal, p_anio, p_inicio)
	dim fs, arch, FileCollection, cont, ext
	dim strOrigen, strDestino, file, id, num

	set fs = server.createObject("Scripting.FileSystemObject")
	strOrigen = server.mapPath("fotos fiesta sin pasar")
	strDestino = server.MapPath("fotos fiesta pasadas")	
	set FileCollection = fs.getFolder(strOrigen).files
	cont = cint(p_inicio)
	for each file in FileCollection
		ext = right(file.name, 3)
		id = p_prefijo & p_sucursal & p_anio
		'response.write strDestino & "\Fiesta" & id & "-" & num & "." & ext & "<br>"
		file.move(strDestino & "\" & id & "-" & cont & "." & ext) 
		cont = cont + 1
	next
end sub
'**********************************************************************************************
dim accion, anio, sucursal, id, inicio
dim conn, rs, strSQL

call levantarParametros(accion, prefijo, sucursal, anio, inicio)
if accion="migrar" then
	call reemplazarNombres(prefijo, sucursal,anio, inicio)
end if
%>
<html>
<head>
	<title></title>
	<link rel="stylesheet" href="CSS/ActisaIntra-1.css">
	<script language="javascript">
		function migrarFotos()
		{
			if (form1.anio.value != '')
			{
				document.form1.accion.value = 'migrar';
				document.form1.submit();
			}
			else	
				alert('<%=GF_Traducir("Falta completar el año")%>');
		}		
	</script>
</head>

<body>
<%call GF_TITULO("TablaMG.gif", GF_Traducir("Migrador de Fotos"))%>
<form name="form1" action="migradorfotos.asp">
	<table align="center">
		<tr>
			<td align="right"><b><%=GF_Traducir("Prefijo")%>:</b></td>
			<td><input type="text" name="prefijo" value="Fiesta"></td>
		</tr>
		<tr>
			<td align="right"><b><%=GF_Traducir("Sucursal")%>:</b></td>
			<td>
				<select id="sucursal" name="sucursal">
				<%strSQL = "select mg_kc, mg_ds from mg where mg_km='MS' order by mg_ds asc"
				call GF_BD_Control(rs, conn, "OPEN", strSQL)
				while not rs.eof 
					id = GF_DT1("READ","shortds", "","","ms",rs("mg_kc"))
					if (id = "?") then id = ""%>
					<option value="<%=id%>"><%=rs("mg_ds")%></option>
					<%rs.movenext
				wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="right"><b><%=GF_Traducir("Año")%>:</b></td>
			<td>
				20<input type="text" maxlength="2" size="2" name="anio">
			</td>
		</tr>
		<tr>
			<td align="right"><b><%=GF_Traducir("Comenzar numeración desde")%>:</b></td>
			<td><input type="text" name="inicio" size="3" maxlength="3" value="1"></td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<input type="button" value="Migrar" onClick="javascript:migrarFotos();">
			</td>
		</tr>
	</table>
	<input type="hidden" name="accion" value="">
</form>
</body>
</html>