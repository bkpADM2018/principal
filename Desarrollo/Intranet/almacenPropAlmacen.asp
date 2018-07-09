<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'Call controlAccesoCM("CMADMREM")
'-----------------------------------------------------------------------------------------------
Function getUsuarios(idAlmacen) 
	Dim strSQL, rs, myWhere,conn
	strSQL = "select * from tblalmacenesusuario "
	strSql = strSQL & " where idAlmacen = " & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getUsuarios = rs
End Function
'-----------------------------------------------------------------------------------------------
Function getCandidatos() 
	Dim strSQL, rs, myWhere,conn
	'Ajusto Paginacion
	strSQL="Select idProfesional, Nombre, CDUSUARIO  from WFPROFESIONAL where EGRESOVALIDO = 'F'  "
	if not usuarios.eof then
		strSQL = strSQL & " and CDUSUARIO not in( '" & usuarios("CDUSUARIO") & "'"
		usuarios.movenext
		while not usuarios.eof
			strSQL = strSQL & ", '" & usuarios("CDUSUARIO") & "'"
			usuarios.movenext
		wend
		usuarios.movefirst
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
			lista = lista & usuarios("CDUSUARIO") & "-" & usuarios("NIVEL") & ";"
			usuarios.movenext
		wend
		usuarios.movefirst		
	end if	
	getListaUsuarios = lista
End Function
'-----------------------------------------------------------------------------------------------
Sub grabarResponsablesAlmacen(idAlmacen, listaCandidatosAsignar) 
	Dim strSQL, rs, idx, mySplit
	dim arrListaCandidatosAsignar
	idAlmacenTemp = idAlmacen*-1
	arrListaCandidatosAsignar = split(listaCandidatosAsignar, ";", -1)
	'Borro todos los responsables de almacen (Se borran al finalizar por si se produce algun error.
	strSQL= "Update TBLALMACENESUSUARIO set IDALMACEN=" & idAlmacenTemp & " where IDALMACEN=" & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	'Grabo los responsables
	for i = 0 to ubound(arrListaCandidatosAsignar)
		if arrListaCandidatosAsignar(i) <> "" then
			mySplit = split(arrListaCandidatosAsignar(i),"-")
			strSQL = "Insert into tblalmacenesusuario (idalmacen, cdusuario, nivel) values(" & idAlmacen & ", '" & mySplit(0) & "', '" & mySplit(1) & "')"			
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		end if
	next
	'Borro todos los responsables de almacen
	strSQL= "Delete from TBLALMACENESUSUARIO where IDALMACEN=" & idAlmacenTemp
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
End sub
'-----------------------------------------------------------------------------------------------
Sub grabarDatosAlmacen(byref idAlmacen, cdAlmacen, dsAlmacen, idDivision) 
	Dim strSQL, rs, oConn
		'Borro todos los responsables de almacen
	if idAlmacen = 0 then
	'es nuevo, insertar
		strSQL = "insert into tblalmacenes (cdalmacen, dsalmacen, idDivision, estado, momento, cdusuario) values ( "
		strSQL = strSQL & "'" & cdAlmacen & "', '" & dsAlmacen & "', " & idDivision & ", 1, " & session("MmtoSistema") & ", '" & session("Usuario") & "')"		
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		strSQL = "select max(idAlmacen) as idalmacenNuevo from tblalmacenes "
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		idAlmacen = rs("idAlmacenNuevo")
	else
	'actualizar
		strSQL = "update tblalmacenes set dsalmacen = '" & dsAlmacen & "', iddivision = " & idDivision
		strSQL = strSQL & " , momento = " & session("MmtoSistema") & " , cdusuario = '" & session("Usuario") & "' "
		strSQL = strSQL & " where idAlmacen = " & idAlmacen
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if		
End sub
'**********************************************************
'***	COMIENZO DE PAGINA
'**********************************************************
Dim idAlmacen
Dim rsAlmacenes
Dim candidatos, usuarios
dim dsUsuario
dim listaCandidatosAsignar
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
cdAlmacen = GF_PARAMETROS7("cdAlmacen","",6)
dsAlmacen = GF_PARAMETROS7("dsAlmacen","",6)
idDivision = GF_PARAMETROS7("idDivision",0,6)
listaCandidatosAsignar = GF_PARAMETROS7("seleccionados","",6)
'Response.Write listaCandidatosAsignar
'Response.End 
accion = GF_PARAMETROS7("accion","",6)
'si hay responsables de almacen seleccionados  para grabar, se procede a insertarlos en BD
if accion = ACCION_GRABAR then 
	grabarDatosAlmacen idAlmacen, cdAlmacen, dsAlmacen, idDivision
	grabarResponsablesAlmacen idAlmacen, listaCandidatosAsignar
end if
Set usuarios = getUsuarios(idAlmacen)
Set candidatos = getCandidatos()
GP_ConfigurarMomentos
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Sistema de Compras</title>

<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/diagram.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">

var refPopUpAlmacen;

	function canSubmit(acc) {	
	<%if idAlmacen = 0 then%>
		var cdAlmacen = document.getElementById("cdAlmacen").value;
		if (cdAlmacen == '') return alert('Debe ingresar el codigo de Almacen');
	<%end if%>
		var dsAlmacen = document.getElementById("dsAlmacen").value;
		if (dsAlmacen == '') return alert('Debe ingresar la descripcion de Almacen');
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();		
	}
	
	function almacenOnLoad() {	
		refPopUpAlmacen = getObjPopUp('popupAlmacenes');		
	<% if (accion <> "") then %>
		refPopUpAlmacen.hide();		
	<% end if %>
		pngfix();
	}
	function agregar(){
		//agrega 
		var optUsuario;
		var objSelectCandidatos = document.getElementById("candidatos");
		var longitud = (objSelectCandidatos == null)?0:objSelectCandidatos.length;
		var seleccionados = document.getElementById("seleccionados").value;
		var auxText;
		var i;
		for (i=0;i<longitud;i++) {
			optUsuario = objSelectCandidatos.options[i];
			if (optUsuario.selected) {
				agregarAUsuarios(optUsuario.value, optUsuario.text);
				for (i=0;i<document.frmSel.level.length;i++){
					if (document.frmSel.level[i].checked){
						auxText = document.frmSel.level[i].value;
						break;
					}	
				}				
				seleccionados = seleccionados + optUsuario.value + "-" + auxText + ";"				
				objSelectCandidatos.removeChild(optUsuario);
				i--;
				longitud--;
			}
		}
		document.getElementById("seleccionados").value = seleccionados;
	}
	function agregarAUsuarios(cdUsuario, dsUsuario){
		var objSelectUsuarios = document.getElementById("usuarios");
		var optNueva = objSelectUsuarios.appendChild(document.createElement('option'));
		optNueva.value = cdUsuario;
		var auxText, auxText2;
	    var i;
		for (i=0;i<document.frmSel.level.length;i++){
			if (document.frmSel.level[i].checked){
				auxText = document.frmSel.level[i].value;
				break;
			}	
		}
		if (auxText=="<% =ALMACEN_ADMIN %>"){
			auxText2 = "Admin";
		}
		else if(auxText=="<% =ALMACEN_AUDITOR %>"){
			auxText2 = "Auditor";
		}
		else if(auxText=="<% =ALMACEN_SOLICITANTE %>"){
			auxText2 = "Solicitante";
		}
		else if(auxText=="<% =ALMACEN_SOLICITANTE_CONTROL %>"){
			auxText2 = "Solicitante + Ctrl Stock";
		}
		else{
			auxText2 = "Usuario";
		}
		optNueva.text = dsUsuario + " - " + auxText2;
	}
	function quitar(){
		var optUsuario;
		var objSelectUsuarios = document.getElementById("usuarios");
		var longitud = (objSelectUsuarios == null)?0:objSelectUsuarios.length;
		var seleccionados = document.getElementById("seleccionados").value;
		var auxText = new String();
		var auxText2, mySplit;
		for (i=0;i<longitud;i++) {
			optUsuario = objSelectUsuarios.options[i];
			if (optUsuario.selected) {
				auxText = optUsuario.text; 
				mySplit = auxText.split("-");
				agregarACandidatos(optUsuario.value, mySplit[0]);
				var str = mySplit[1];
				str = str.replace(/^\s*|\s*$/g,"");
				if (str=="Admin"){
					auxText2 = "<% =ALMACEN_ADMIN %>";
				}
				else if(str=="Auditor"){
					auxText2 = "<% =ALMACEN_AUDITOR %>";
				}
				else if(str=="Solicitante"){
					auxText2 = "<% =ALMACEN_SOLICITANTE %>";
				} 
				else if(str=="Solicitante + Ctrl Stock"){
					auxText2 = "<% =ALMACEN_SOLICITANTE_CONTROL %>";
				}
				else{
					auxText2 = "<% =ALMACEN_USUARIO %>";
				}
				seleccionados = seleccionados.replace(optUsuario.value + "-" + auxText2 + ";","");				
				objSelectUsuarios.removeChild(optUsuario);
				i--;
				longitud--;
			}
		}
		document.getElementById("seleccionados").value = seleccionados;
	}
	function agregarACandidatos(cdCandidato, dsCandidato){
		var objSelectCandidatos = document.getElementById("candidatos");
		var optNueva = objSelectCandidatos.appendChild(document.createElement('option'));
		optNueva.value = cdCandidato;
		optNueva.text = dsCandidato;
	}
</script>
</head>
<body onLoad="almacenOnLoad()">
<table width="100%">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/almacenes/warehouses-48x48.png"> <% =GF_TRADUCIR("Propiedades de Almacen") %></td>
	</tr>
</table>
<form id="frmSel" name="frmSel" action="almacenPropAlmacen.asp?idAlmacen=<%=idAlmacen%>" method="POST">	

	<table width="50%" align="center">
		<tr><td colspan="3">
			<table width="100%" align="left">
			<tr>
				<td class="reg_Header" width="15%" align="right"><% =GF_TRADUCIR("Codigo") %></td>
				<td><b>
					<%	
					if idAlmacen <> 0 then
						Set rsAlmacenes = obtenerListaAlmacenes(idAlmacen)
						response.write rsAlmacenes("CDALMACEN")
					else	%>
						<input type="text" id="cdAlmacen"  size="10" maxlength="10" name="cdAlmacen" value="">					
					<%
					end if
					%>
				</b></td>
			</tr>
			<tr>
				<td class="reg_Header" align="right" ><% =GF_TRADUCIR("Descripcion") %></td>
				<td><b>
				<%if idAlmacen <> 0 then%>
					<input type="text" id="dsAlmacen" name="dsAlmacen" maxlength="50" value="<%=rsAlmacenes("DSALMACEN")%>">
				<%else%>
					<input type="text" id="dsAlmacen" maxlength="50" name="dsAlmacen" value="">
				<%end if%>
				</b></td>
			</tr>
			<tr>
				<td class="reg_Header" align="right"><% =GF_TRADUCIR("División") %></td>
				<td>
				<%strSQL="Select * from TBLDIVISIONES"
				Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
				%>
				<select id="idDivision" name="idDivision">
				<%		while (not rsDivision.eof) %>										
				<option value="<% =rsDivision("IDDIVISION") %>"
					<%if idAlmacen <> 0 then%>
						<% if (CInt(idDivision) = CInt(rsDivision("IDDIVISION")) or (CInt(rsAlmacenes("iddivision")) = CInt(rsDivision("iddivision")))) then response.write "selected='true'" %>
					<%end if%>
					><% =rsDivision("DSDIVISION") %></option>
				<%			rsDivision.MoveNext()
				wend	%>								
				</select>			
				</td>
			</tr>
			<tr>
				<td class="reg_Header" align="right"><% =GF_TRADUCIR("Nivel") %></td>
				<td colspan="2">
					<input style="cursor:pointer;" type="radio" value="<% =ALMACEN_USUARIO %>" id=level name=level checked><%=GF_Traducir("Usuario")%>
					<input style="cursor:pointer;" type="radio" value="<% =ALMACEN_AUDITOR %>" id=level name=level><%=GF_Traducir("Auditor")%>
					<input style="cursor:pointer;" type="radio" value="<% =ALMACEN_ADMIN %>" id=level name=level><%=GF_Traducir("Admin")%>
					<input style="cursor:pointer;" type="radio" value="<% =ALMACEN_SOLICITANTE_CONTROL %>" id=Radio1 name=level><%=GF_Traducir("Solicitante + Ctrl Stock")%>
					<input style="cursor:pointer;" type="radio" value="<% =ALMACEN_SOLICITANTE %>" id=level name=level><%=GF_Traducir("Solicitante")%>
				</td>
			</tr>
			</table>
		</td></tr>


		<tr>
			<td><b><% =GF_TRADUCIR("Candidatos") %></b></td>
			<td>&nbsp;</td>
			<td><b><% =GF_TRADUCIR("Responsables Almacen") %></b></td>
		</tr>		
		<tr>
			<td width="45%" align="left" rowspan="2">
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
			<td width="10%">
				<table>
					<tr height="50%">
						<td valign="middle" style="vertical-align:middle">
							<img src="images/A_NEXT.gif" style="cursor:pointer;" onclick="javascript:agregar();">
						</td>
					</tr>
					<tr height="50%">
						<td style="vertical-align:middle">
							<img src="images/A_PREV.gif" style="cursor:pointer;" onclick="javascript:quitar();">
						</td>
					</tr>
				</table>
			</td>				
			<td width="45%" align="right" rowspan="2">
				<select size="20"  multiple="multiple" id="usuarios" name="usuarios" style="width:230pt;">
					<%
					if not usuarios.eof then
						while (not usuarios.eof)
						dsUsuario = getUserDescription(usuarios("CDUSUARIO"))
						%>
							<option value="<% =usuarios("CDUSUARIO") %>">
								<%
								select case (ucase(usuarios("NIVEL")))
									case ALMACEN_ADMIN
										auxText = "Admin"
									case ALMACEN_AUDITOR	
										auxText = "Auditor"
									case ALMACEN_SOLICITANTE_CONTROL	
										auxText = "Solicitante + Ctrl Stock"
									case ALMACEN_SOLICITANTE	
										auxText = "Solicitante"										
									case else
										auxText = "Usuario"
								end select
								Response.write dsUsuario & " - " & auxText 
								%>
							</option>
					<%		usuarios.MoveNext()
						wend%>
					<%end if%>
				</select>
			</td>
		</tr>
		<tr><td>&nbsp;</td><tr>
		<tr>
			<td></td>
			<td align="right" colspan="3	">
				<table>	
					<tr><td align="right">
						<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" onClick="javascript:canSubmit('<%=ACCION_GRABAR%>');">
						<input type="button" value="<% =GF_TRADUCIR("Cancelar") %>" onClick="refPopUpAlmacen.hide()">
					</td></tr>
				</table>
			</td>		
		</tr>			
	</table>
	<input type="hidden" id="seleccionados" name="seleccionados" value="<%=getListaUsuarios(usuarios)%>" size=50>
	<input type="hidden" id="accion" name="accion" value="">
</form>
</body>
</html>