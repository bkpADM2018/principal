<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)
'---------------------------------------------------------------------------------------------------------------
'Funciones para Factores de conversion

Function eliminarFactoresViejos(idOrigen, idDestino)
	Dim strSQL, rs, conn
	
	strSQL = "DELETE FROM TBLUNIDADESCONVERSION WHERE IDUNIDADORIG = " & idOrigen & " and IDUNIDADDEST = " & idDestino
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	
End Function

Function grabarFactoresConversion(idUnidadOrigen, conversionString) 
	
	Dim filas, campos, i, origen, destino, factor, temp
	
	filas = Split(conversionString, "|")
	for i = 0 To UBound(filas)-1		
		campos = Split(filas(i),";")		
		'Se elimina el factor anterior.
		Call eliminarFactoresViejos(idUnidadOrigen, campos(0))
		Call eliminarFactoresViejos(campos(0), idUnidadOrigen)
		'Se crea el nuevo
		factor = campos(1)
		
		origen = idUnidadOrigen
		destino = campos(0)
		if (factor < 1) then
			temp = origen
			origen = destino 
			destino = temp
			factor = 1/factor
		end if
		strSQL = "INSERT INTO TBLUNIDADESCONVERSION(IDUNIDADORIG, IDUNIDADDEST, FACTOR, CDUSUARIO, MOMENTO)"
		strSQL = strSQL & " VALUES (" & origen & ", " & destino & ", " & factor & ", '" & session("Usuario") & "', " & session("MmtoSistema") & ")"
		'response.write strSQL
		'response.end
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	next	
	
End Function
'---------------------------------------------------------------------------------------------------------------
'Funciones para la uinidad

Function controlar(idUnidad, cdUnidad, dsUnidad, idTipoUnidad, abreviatura)
	Dim strSQL, rs, conn
	
	controlar = RESPUESTA_OK
	if (idUnidad = 0) then
		'Valido el codigo de unidad
		if (cdUnidad = "") then
			controlar = CODIGO_VACIO
		else		
			strSQL="Select * from TBLUNIDADES where CDUNIDAD='" & cdUnidad & "'"
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
			if (not rs.eof) then controlar = CODIGO_EXISTE
		end if
	end if
	if (controlar = RESPUESTA_OK) then					
		'Valido el tipo de unidad
		if (idTipoUnidad = 0) then controlar = TIPOUNIDAD_NO_EXISTE		
	end if

End Function

Function accionGrabar(idUnidad, cdUnidad, dsUnidad, idTipoUnidad, abreviatura)
	Dim strSQL, rs, conn
	
	if (idUnidad = 0) then
		'Es una unidad nueva
		strSQL="Insert into TBLUNIDADES(CDUNIDAD, DSUNIDAD, IDTIPOUNIDAD, ABREVIATURA, ESTADO, REFERENCIAS, CDUSUARIO, MOMENTO)"
		strSQL = strSQL & " values('" & cdUnidad & "', '" & dsUnidad & "', " & idTipoUnidad & ", '" & abreviatura & "', " & ESTADO_ACTIVO & ", 0, '" & session("Usuario") & "', " & session("MmtoSistema") & ")"
	else
		'Es una modificacion
		strSQL="Update TBLUNIDADES Set CDUNIDAD='" & cdUnidad & "', DSUNIDAD='" & dsUnidad & "', IDTIPOUNIDAD=" & idTipoUnidad & ", ABREVIATURA='" & abreviatura & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema")		
		strSQL = strSQL & " where IDUNIDAD=" & idUnidad
	end if
	'response.write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	accionGrabar = true
	
End Function

Function accionConsulta(idUnidad, ByRef cdUnidad, ByRef dsUnidad, ByRef idTipoUnidad, ByRef abreviatura)
	
	Dim strSQL, rs, conn
	
	strSQL="Select * from TBLUNIDADES where IDUNIDAD=" & idUnidad
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then		
		cdUnidad = Trim(rs("CDUNIDAD"))
		dsUnidad = Trim(rs("DSUNIDAD"))
		idTipoUnidad = rs("IDTIPOUNIDAD")
		abreviatura = Trim(rs("ABREVIATURA"))		
	end if
	
	
End Function
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim idUnidad, accion, errMsg, conversion, idxFactores, rsFactores
Dim cdUnidad, dsUnidad, idTipoUnidad, abreviatura, strSQL, conn, rsTipoUnidad

idUnidad = GF_PARAMETROS7("idUnidad",0,6)
cdUnidad = UCase(GF_PARAMETROS7("codigo","",6))
dsUnidad = UCase(GF_PARAMETROS7("descripcion","",6))
idTipoUnidad = GF_PARAMETROS7("idTipoUnidad",0,6)
abreviatura = GF_PARAMETROS7("abreviatura","",6)
accion = GF_PARAMETROS7("accion","",6)
conversiones = GF_PARAMETROS7("conversiones","",6)

Call GP_ConfigurarMomentos
if (accion = ACCION_GRABAR) then
	errMsg = controlar(idUnidad, cdUnidad, dsUnidad, idTipoUnidad, abreviatura)
	if (errMsg = RESPUESTA_OK) then
		Call accionGrabar(idUnidad, cdUnidad, dsUnidad, idTipoUnidad, abreviatura)
		'Se graban los factores de conversion	
		if (conversiones <> "") then Call grabarFactoresConversion(idUnidad, conversiones)
		accion = ACCION_CERRAR
	else
		setError(errMsg)
	end if
else
	Call accionConsulta(idUnidad, cdUnidad, dsUnidad, idTipoUnidad, abreviatura)
end if
if (accion = "") then accion = ACCION_GRABAR

%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/dhtmlxgrid.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/dhtmlxcommon.js"></script> 
<script type="text/javascript" src="scripts/dhtmlxgrid.js"></script> 
<script type="text/javascript" src="scripts/dhtmlxgridcell.js"></script> 
<script type="text/javascript">
var refPopUpUnidad;	
var	dg;
var conversiones = new Array();
var rowErr = new Array();

function codigoOnBlur(ref) {
	if (ref.value == "") {
		document.getElementById("aceptar").disabled = true;
	} else {
		document.getElementById("aceptar").disabled = false;
	}			
}

function tipoOnChange(sel) {
	dg.clearAll();	
	return true;		
}
	
//Marca la existencia de error en la fila indicada.
function addError(rowId) {
	rowErr[rowId] = true;
}

//Anula la existencia de error en la fila indicada.
function removeError(rowId) {
	rowErr[rowId] = false;
}
//Indica si aun hay alguna fila donde existan errores.
function existError() {
	var ret = false;
	for (var i in rowErr) if (rowErr[i]) ret = true; 
	return ret;
}

function factorOnChange(stage, rowId, colIdx, valNew, valOld) {		
	if (stage == 2) {
		//Se procesa el cambio realizado.						
		if (valNew != valOld) {			
			var td = document.getElementById("errMsg")
			//Se controla que el valor ingresado sea numerico. 	
			if (controlNumero(valNew)) {
				//El valor es correcto.
				removeError(rowId);
				conversiones[rowId] = editarNumero(valNew, 4);
				//Si no hay errores se habilita la opcion de guardar los datos.			
				if (!existError()) {
					td.innerHTML = "";
					td.className="";
					document.getElementById("aceptar").disabled = false;
				}
				dg.cells(rowId, colIdx).setValue = conversiones[rowId];	
				dg.setRowColor(rowId,"");
				dg.setRowTextStyle(rowId,"");										
			} else {
				//Se notifica el error y se bloquea el boton de aceptar.
				addError(rowId);
				dg.setRowColor(rowId,"red");
				dg.setRowTextStyle(rowId,"color:white; font-weight:bold;");					
				td.innerHTML = "El valor ingresado no es valido.";
				td.className="TDERROR";
				document.getElementById("aceptar").disabled = true;
			}				
		}
	} else if (stage == 0) {
		dg.cells(rowId, colIdx).setValue = "";
	}
	return true;
}
	
function loadFactores() {		
	dg = new dhtmlXGridObject('factores'); 
	dg.imgURL = "images/compras/"; 
	//Formato
	dg.setSkin("light");
	dg.setHeader("Unidad, Factor"); 
	dg.setInitWidths("320,60"); 
	dg.setColAlign("left,right"); 
	dg.setColTypes("ro,ed"); 
	dg.setColSorting("str,int"); 
	//eventos
	dg.setOnEditCellHandler(factorOnChange);
	//Inicio el grid
	dg.init(); 				
	<%
	strSQL="Select * From VWFACTORESCONVERSION F inner join TBLUNIDADES U ON IDUNIDADDEST=IDUNIDAD where IDUNIDADORIG=" & idUnidad & " and ESTADO=" & ESTADO_ACTIVO
	Call executeQueryDB(DBSITE_SQL_INTRA, rsFactores, "OPEN", strSQL)
		while (not rsFactores.eof)
			idxFactores = 0
	%>			
			dg.addRow(<% =rsFactores("IDUNIDADDEST") %>, "<% =rsFactores("DSUNIDADDEST") %>," + editarNumero('<% =rsFactores("FACTOR") %>',3), <% =idxFactores %>);	
	<% 		rsFactores.MoveNext()
		wend %>
}

function fsubmit() {	
	var cnv = document.getElementById("conversiones");
	var str = "";		
	for (var k in conversiones) if (conversiones[k] != undefined) str = str + k + ";" + conversiones[k] + "|";   
	cnv.value = str;		
	document.forms["frmSel"].submit();
}

function unidadOnLoad() {
	refPopUpUnidad = getObjPopUp('popupUnidad');
	<%  'Solo se cierra cuando es una modificacion, dado que en un alta, se cargan las unidades del mismo
		'tipo para que complete los factores.
		if (accion = ACCION_CERRAR) then %>
		refPopUpUnidad.hide();
	<%  end if %>
	//Cargo la lista de factores de conversion.
	loadFactores();
	var elem = document.getElementById("codigo");
	if (elem.type != "hidden")
		elem.focus();
	else
		document.getElementById("descripcion").focus();
}
function nuevoTipoUnidad(){
	var refPopupNewTypeOfUnit = new winPopUp('popupNewTypeOfUnit', 'comprasPropTipoUnidad.asp', 380, 180, "Nuevo Tipo", '');
}
</script>
</head>
<body onLoad="unidadOnLoad()">
<form name="frmSel" method="post" action="comprasPropUnidad.asp">
<table  width="100">
	<tr>
		<td class="title_sec_section" colspan="2"><img align="absMiddle" src="images/compras/units-32x32.png"><% =GF_TRADUCIR("Propiedades de Unidad") %></td>
	</tr>	
	<tr>
		<td colspan="2" id="errMsg"><% call showErrors() %></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<table>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Unidad") %></td>
					<td colspan="3"><% =idUnidad %></td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Nombre") %></td>
					<td> <% if (idUnidad = 0) then %>
						<input type="text" id="codigo" name="codigo" maxlength="10" size="10" value="<% =cdUnidad %>" onblur="codigoOnBlur(this)" onkeypress="return controlSalto(this, event)">
						<% else %> 
							<b><% =cdUnidad %></b> 
							<input type="hidden" id="codigo" name="codigo" value="<% =cdUnidad %>">
						<% end if %>
					</td>
					<td class="reg_header"><% =GF_TRADUCIR("Abreviatura") %></td>
					<td align="right"> <input type="text" id="abreviatura" name="abreviatura" maxlength="5" size="5" value="<% =abreviatura %>" onkeypress="return controlSalto(this, event)"></td>
				</tr>				
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Descripcion") %></td>
					<td colspan="3"><input type="text" id="descripcion" name="descripcion" maxlength="50" size="45" value="<% =dsUnidad %>" onkeypress="return controlSalto(this, event)"></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Tipo Unidad") %></td>
					<td colspan="3">
						
						<Select id="idTipoUnidad" name="idTipoUnidad" onChange="tipoOnChange(this)">
							<%
							strSQL="Select * from TBLTIPOSUNIDAD"
							Call executeQueryDB(DBSITE_SQL_INTRA, rsTipoUnidad, "OPEN", strSQL)
							while (not rsTipoUnidad.eof)
							%>
								<option value="<% =rsTipoUnidad("IDTIPOUNIDAD") %>" <% if (idTipoUnidad = rsTipoUnidad("IDTIPOUNIDAD")) then response.write "selected=true" %>><% =rsTipoUnidad("DSTIPOUNIDAD") %>							
							<% 		rsTipoUnidad.MoveNext()
								wend %>
						</select>
						<!--<a href="comprasPropTipoUnidad.asp?idUnidad=<% =idUnidad %>"><img align="absMiddle" src="images/compras/add-16x16.png" title="<% =GF_TRADUCIR("Agregar un tipo de Unidad") %>"></a>-->
						<a onclick="javascript:nuevoTipoUnidad()"><img style="cursor:pointer;" align="absMiddle" src="images/compras/add-16x16.png" title="<% =GF_TRADUCIR("Agregar un tipo de Unidad") %>"></a>
						
					</td>
				</tr>								
			</table>
		</td>
	</tr>	
	<tr>
		<td></td>
		<td class="TDNOHAY"><% =GF_TRADUCIR("Factores de Conversion entre Unidades") %></td>
	</tr>			
	<tr>
		<td></td>
		<td>
			<div id="factores" style="width:390;height:120"></div>
		</td>
	</tr>
	<tr>
		<td></td>
		<td align="right">
			<table>	
				<tr><td>
					<%  if (not isAuditor(SIN_DIVISION)) then %>
					<input type="button" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idUnidad = 0) then response.write "disabled=true" %> onClick="fsubmit()">
					<%	end if %>
				</td></tr>
			</table>
		</td>		
	</tr>
</table>
<input type="hidden" name="accion" value="<% =accion %>">
<input type="hidden" name="idUnidad" value="<% =idUnidad %>">
<input type="hidden" name="conversiones" id="conversiones" value="">
</form>
</body>
</html>