<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'-----------------------------------------------------------------------------------------------
Function getCandidatos() 
	Dim strSQL, rs, myWhere,conn
	'Ajusto Paginacion
	strSQL="Select idProfesional, Nombre, CDUSUARIO  from WFPROFESIONAL where EGRESOVALIDO = 'F'  "
	if not destinatarios.eof then
		strSQL = strSQL & " and CDUSUARIO not in( '" & destinatarios("user") & "'"
		destinatarios.movenext
		while not destinatarios.eof
			strSQL = strSQL & ", '" & destinatarios("user") & "'"
			destinatarios.movenext
		wend
		destinatarios.movefirst
		strSQL = strSQL & " )"
	end if	
	strSQL = strSQL & " order by Nombre "
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getCandidatos = rs
End Function
'-----------------------------------------------------------------------------------------------
Function getListaUsuarios(usuarios) 
	Dim lista
	lista = ""
	'Ajusto Paginacion
	if usuarios.recordcount > 0 then usuarios.movefirst
	if not usuarios.eof then		
		while not usuarios.eof
			lista = lista & usuarios("user") & "|"
			usuarios.movenext
		wend
		usuarios.movefirst		
	end if	
	getListaUsuarios = lista
End Function
'-----------------------------------------------------------------------------------------------
Function getDestinatarios(pIdAlmacen)
	Dim strSQL,rs
	strSQL = "select * from TBLMAILSALERTASALMACENES where idalmacen = " & pIdAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	set getDestinatarios = rs
End Function
'-----------------------------------------------------------------------------------------------
Dim candidatos,idAlmacen,rsAlmacenes,seleccionados,accion,rs,strSQL,idAlmacenTemp,destinatarios

idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)

seleccionados = GF_PARAMETROS7("seleccionados","",6)
accion = GF_PARAMETROS7("accion","",6)

if (accion = ACCION_GRABAR) then
	auxSeleccionados = split(seleccionados,"|")
	idAlmacenTemp = idAlmacen * -1
		
	strSQL= "Update TBLMAILSALERTASALMACENES set IDALMACEN=" & idAlmacenTemp & " where IDALMACEN=" & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)

	for i = 0 to ubound(auxSeleccionados)
		if (cstr(auxSeleccionados(i)) <> "") then
			strSQL = "insert into TBLMAILSALERTASALMACENES (idalmacen,[user],email) values("&idalmacen&",'"&auxSeleccionados(i)&"','"&getUserMail(auxSeleccionados(i))&"')"
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	next

	'Borro todos los responsables de almacen
	strSQL= "Delete from TBLMAILSALERTASALMACENES where IDALMACEN=" & idAlmacenTemp
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end if

Set rsAlmacenes = obtenerListaAlmacenes(idAlmacen)

set destinatarios = getDestinatarios(idAlmacen)
set candidatos = getCandidatos()

%>

<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css">

		<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>

		<script>
			function agregar(){
				
				$("#candidatos option:selected").each(function(){
						$("#destinatarios").append($(this));		
						$("#seleccionados").val($("#seleccionados").val() + $(this).val()+"|");
					});
				
				$("#destinatarios option:selected").attr("selected", "");

			}
			function quitar(){
				
				$("#destinatarios option:selected").each(function(){
					auxStr = $("#seleccionados").val().replace($(this).val()+"|","");
					$("#seleccionados").val(auxStr);
				});

				$("#candidatos").append($("#destinatarios option:selected"));
				$("#candidatos option:selected").attr("selected", "");
			}
		</script>
	</head>
	<body>
		<table width="100%">
			<tr>
				<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/almacenes/campana-32x32.png" width="48" height="48"> <% =GF_TRADUCIR("Alertas de Almacen") %></td>
			</tr>
		</table>
		<br />

		<table width="90%" align="center">
			<tr>
				<td>
					<table width="100%" align="center">
						<tr>
							<td class="reg_header" width="100px" align="right">Codigo</td>
							<td><strong><%=rsAlmacenes("CDALMACEN")%></strong></td>
						</tr>
						<tr>
							<td class="reg_header" align="right">Descripcion</td>
							<td><%=rsAlmacenes("DSALMACEN")%></td>
						</tr>
						<tr>
							<td class="reg_header" align="right">Division</td>
							<td><%= getDivisionDS(getDivisionAlmacen(rsAlmacenes("IDALMACEN")))%></td>
						</tr>

					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table align="center">
						<tr>
							<td>
								<strong>Candidatos</strong>
							</td>
							<td>&nbsp;</td>
							<td><strong>Destinatarios</strong></td>
						</tr>
						<tr>
							<td width="210px">
								<select size="20"  multiple="multiple" id="candidatos" name="candidatos" style="width:200pt;">
									<%
									if not candidatos.eof then
										while (not candidatos.eof)	%>
											<option value="<% =candidatos("CDUSUARIO") %>"><% =candidatos("Nombre") %></option>
									<%		candidatos.MoveNext()
										wend%>
									<%end if%>
								</select>
							</td>
							<td width="20px">
								<table>
									<tr height="50%">
										<td valign="middle" style="vertical-align:middle">
											<img src="images/A_NEXT.gif" style="cursor:pointer;" onclick="agregar();">
										</td>
									</tr>
									<tr height="50%">
										<td style="vertical-align:middle">
											<img src="images/A_PREV.gif" style="cursor:pointer;" onclick="javascript:quitar();">
										</td>
									</tr>
								</table>
							</td>
							<td width="210px">
								<select size="20"  multiple="multiple" id="destinatarios" name="destinatarios" style="width:200pt;">
									<%
									if not destinatarios.eof then
										while (not destinatarios.eof)	%>
											<option value="<% =destinatarios("user") %>"><% =getUserDescription(destinatarios("user")) %></option>
									<%		
											destinatarios.MoveNext()
											
										wend%>
									<%end if%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="3" align="right">
								<form type="post" action="almacenPropalertasAlmacen.asp">
									<input type="submit" value="Aceptar">
									<input type="hidden" name="accion" id="accion" value="<%=ACCION_GRABAR%>">
									<input type="hidden" name="idAlmacen" id="idAlmacen" value="<%=idAlmacen%>">
									<input type="hidden" name="seleccionados" id="seleccionados" value="<%=getListaUsuarios(destinatarios)%>">
								</form>		
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>
