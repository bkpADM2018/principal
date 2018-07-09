<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHTML.asp"-->
<%
Const REGISTROS_LISTA = 5
Call comprasControlAccesoCM(RES_OBR)
'--------------------------------------------------------------------------------------------
Function getDetalleLista(IdLista)
	Dim strSQL,rs,listUsuer
	listUsuer = ""
	strSQL = " SELECT A.CDUSER, A.EMAIL  FROM TBLMAILLSTSDETALLE A "	
	strSQL = strSQL & " WHERE A.IDLISTA = " & IdLista
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		while not rs.eof
			listUsuer = listUsuer & rs("CDUSER") & "|" & getUserDescription(rs("CDUSER")) &";"
			rs.MoveNext()	
		wend		
		listUsuer = left(listUsuer,Len(listUsuer)-1)
	end if
	response.write listUsuer
	response.end
End function
'--------------------------------------------------------------------------------------------
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
Dim accion, errMsg, idObra, cdObra, dsObra, IdDivision 
Dim rsDivision, conn,strSQL,dsDivision
Dim cont ,IdLista,DsLista,cdUsuario,dsUsuario,cantUser,cdUser
Dim flagLoad,rsDet,cdSolicitante,dsSolicitante,CdLista

flagLoad = false

IdLista = GF_PARAMETROS7("IdLista", 0, 6)
DsLista = GF_PARAMETROS7("DsLista", "", 6)
IdDivision = GF_PARAMETROS7("IdDivision", "", 6)
accion = GF_PARAMETROS7("accion","",6)
cantUsuario = GF_PARAMETROS7("cantUsuario", 0, 6)
cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)
CdLista = GF_PARAMETROS7("CdLista","",6)

if(accion = ACCION_PROCESAR)then 
	Call getDetalleLista(IdLista)
else	
	if(IdLista > 0)then 
		flagLoad = true		
		dsSolicitante = getUserDescription(cdSolicitante)
		dsDivision = getDivisionDS(IdDivision)
	end if
end if
	
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
<link rel="stylesheet" href="css/jquery.fileupload-ui.css"	 type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">


<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/JQueryUpload2.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery.fileupload.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery.fileupload-ui.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

<script type="text/javascript">
var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
var calendar;
var lastUsuario = 0;
var msUsuario =new Array() ;
var CODIGO_USUARIO = 'cdUsuario';
var msSolicitante;
var ch = new channel();

function submitInfo() {
	document.getElementById("frmSel").submit();
}

function seleccionarUsuario(plastUsuario,ms) {	
	var desc = ms.getSelectedItem();
	if (desc.indexOf('-') != -1) {
		var arr = desc.split('-');
		document.getElementById(CODIGO_USUARIO + (plastUsuario) ).value = arr[0];		
		ms.setValue(arr[1]);
	} else {		
		if (desc == "") document.getElementById(CODIGO_USUARIO + (plastUsuario ) ).value = "";
	}		
}
	
function guardarListaCorreo() {	
	document.getElementById("frmSel").action= "comprasListaCorreoGrabar.asp";
	document.getElementById("frmSel").target= "ifrmLista";
	document.getElementById("frmSel").submit();	
}
function CerrarPopUp(){
	var refPopUpObra;	
	refPopUpObra = getObjPopUp('popupNuevaLista');	
	refPopUpObra.hide();	
}

function bodyOnLoad() {
	<% if(flagLoad)then %>		
		ch.bind("comprasPopUpListaCorreo.asp?idLista=<%=IdLista%>&accion=<%=ACCION_PROCESAR%>","comprasLoadSubLista_callBack(<%=IdLista%>)");
		ch.send();
	<% else %>
		agregarLineaUsuario();
	<% end if %>
}


/*		CARGO EN EL ARRAY LAS SUBLISTA QUE TIENE LA LISTA 			*/
function comprasLoadSubLista_callBack(pIdLista){	
	var i = 0;
	var listSubLista;
	listSubLista = ch.response();
	if(listSubLista.length > 0){
		var arr = listSubLista.split(";");
		for (i in arr) {
			var vals = arr[i].split("|");
			agregarLineaUsuario();
			document.getElementById("IdEstado"+i).value = <%=ESTADO_ANULACION%>;			
			document.getElementById("divUsuario_" + i ).innerHTML = vals[1];				
			document.getElementById("cdUsuario" + i ).value = vals[0];		
		} 
	}
}


function agregarLineaUsuario() {
		var tblsubListas = document.getElementById("subListas");		
		var rowsubListas = tblsubListas.insertRow(lastUsuario);			
		rowsubListas.className="reg_header_navdos";
		var colCodigo = rowsubListas.insertCell(0);		
		var colEstado = rowsubListas.insertCell(1);
		
		/* 				CREO EL DIV USUARIO QUE USARA EL MAGIC						*/
		var divCodigo = document.createElement('div');
		divCodigo.id =  "divUsuario_" + lastUsuario;
		//divCodigo.setAttribute("colspan","2");
		divCodigo.align = 'center';
		colCodigo.appendChild(divCodigo);
		/*				CREO EL INPUT USUARIO QUE GUARDARA EL VALOR 				*/
		var hiddenCdUsuario = document.createElement('input');
		hiddenCdUsuario.id =  'cdUsuario' + lastUsuario;
		hiddenCdUsuario.name =  'cdUsuario' + lastUsuario;
		hiddenCdUsuario.type = "hidden";
		colCodigo.appendChild(hiddenCdUsuario);
		colCodigo.setAttribute('width', '85%');
				
		<% if(flagLoad)then %>
			/*				CREO LA IMAGEN PARA BORRAR 							*/
			var img = document.createElement("img");
			img.id= "imgDel" + lastUsuario;
			img.src="images/compras/remove-16x16.png";
			img.className="cursorStyle";
			img.setAttribute('onclick', "deleteSubLista(" + lastUsuario + ");");
			img.setAttribute('title', "Eliminar");
			img.setAttribute('style', "cursor:pointer");
			colEstado.appendChild(img);
			
		<% end if %>
		/*	CREO EL INPUT QUE GUARDARA EL ESTADO DE LA FILA (ALTA, BAJA)	*/
		var hiddenEstado   = document.createElement('input');
		hiddenEstado.id    = 'IdEstado' + lastUsuario;		
		hiddenEstado.name    = 'IdEstado' + lastUsuario;		
		hiddenEstado.value = <%=ESTADO_ACTIVO%>;//1
		hiddenEstado.type  = "hidden";
		
		colEstado.appendChild(hiddenEstado);
		colEstado.align = 'center';
		colEstado.setAttribute('width', '15%');
		
		rowsubListas.id = 'rowsubListas' + lastUsuario;
		tblsubListas.appendChild(rowsubListas);
		/*		LE AGREGO AL NUEVO CONTROL LA PROPIEDAD DEL MAGIC			*/		
		msUsuario[lastUsuario] = new MagicSearch("", "divUsuario_" + lastUsuario, 30, 2, "comprasStreamElementos.asp?tipo=personas");
		msUsuario[lastUsuario].setToken(";");		
		msUsuario[lastUsuario].onBlur = 'seleccionarUsuario(' + lastUsuario + ')';		
				
		lastUsuario++;
		document.getElementById("cantUsuario").value = lastUsuario;				
	}	

	/*	LOS ESTADOS DE CADA FILA SON IDENTIFICADOS CON 
	 *		ESTADO ACTIVO (1)   : SON LOS NUEVOS QUE SE AGREGAN  
	 *		ESTADO BAJA (2)	 	: SON LOS QUE SE ELIMINAN 
	 *		ESTADO ANULACION (3) : SON LOS QUE VIENEN CARGADOS POR DEFECTO
	 */
	
	function deleteSubLista(i){
		document.getElementById('IdEstado'+i).value = <%=ESTADO_BAJA%>;//2
		document.getElementById('rowsubListas'+i).style.display = 'none';
	}

	function resultadoCarga_callback(rtrn){				
		switch(rtrn) {
			case "<%=DESCRIPCION_VACIA%>":
				document.getElementById("msjResultado").innerHTML = "<%=GF_TRADUCIR("Se debe ingresar la descripcion")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break			
			case "<%=RESPONSABLE_NO_EXISTE%>":
				document.getElementById("msjResultado").innerHTML ="<%=GF_TRADUCIR("El responsable no existe o no esta habilitado para operar.")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break			
			case "<%=FALTA_RESPONSABLE%>":
				document.getElementById("msjResultado").innerHTML ="<%=GF_TRADUCIR("Falta especificar un responsable.")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break	
			case "<%=CODIGO_EXISTE%>":
				document.getElementById("msjResultado").innerHTML = "<%=GF_TRADUCIR("El codigo no esta habilitado")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break	
			case "<%=CODIGO_VACIO%>":
				document.getElementById("msjResultado").innerHTML = "<%=GF_TRADUCIR("El codigo no existe")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break	
			case "<%=DIVISION_NO_EXISTE%>":
				document.getElementById("msjResultado").innerHTML = "<%=GF_TRADUCIR("Debe seleccionar la division")%>";
				document.getElementById("msjResultado").className = "TDBAJAS";
			break
			default:
				document.getElementById("msjResultado").innerHTML ="<%=GF_TRADUCIR("Se guardo correctamente.")%>";
				document.getElementById("msjResultado").className = "TDSUCCESS";
			break
		}		
	}	
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<form name="frmSel" id="frmSel" method="post" action="comprasPopUpListaCorreo.asp">
<table class="reg_header" id="Listas" align="center" width="95%" border="0" >				
	<tr>
		<td colspan="2"><div id="msjResultado" ></td>
	</tr>
	<tr>
		<td class="reg_header_nav" colspan="2" align="center"><% =GF_TRADUCIR("Datos de la Lista") %></td>				
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Codigo") %></td>	
		<td>		
			<input type="text" id="CdLista" name="CdLista" maxlength="10" size="30" value="<% =CdLista %>"></td>
		</td>	
	</tr>
	<tr>
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Descripcion") %></td>	
		<td>		
			<input type="text" id="DsLista" name="DsLista" maxlength="50" size="50" value="<% =DsLista %>"></td>
		</td>	
	</tr>
	<tr>	
		<td class="reg_header_navdos"><% =GF_TRADUCIR("Division") %></td>
		<% if(flagLoad)then %>
			<td ><% =dsDivision %></td>
		<% else 
			strSQL = "Select IDDIVISION, DSDIVISION from TBLDIVISIONES "
			Call executeQueryDB(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", strSQL)
		%>
            <td>
				<select style="z-index:-1;" name="idDivision">
			        <option SELECTED value="<% =SIN_DIVISION %>">- <% =GF_TRADUCIR("Seleccione") %> -
					<%		while (not rsDivisiones.eof)								
								if (CLng(rsDivisiones("IDDIVISION")) = CLng(idDivision)) then selected = "selected"
					%>
				  	<option value="<% =rsDivisiones("IDDIVISION") %>" <% =selected %>><% =rsDivisiones("DSDIVISION") %>
					<%			rsDivisiones.MoveNext()
							wend 	%>
				</select>
            </td>				
		<% end if %>
	</tr>	
	<tr>
		<td colspan="2">	
			<table width="100%" >
				<tr>					
					<td width="85%" class="reg_header_nav" align="center"><% =GF_TRADUCIR("Usuario") %></td>
					<td width="15%" class="reg_header_nav" align="center">.</td>
				</tr>
				<tr>
					<td colspan="2">
						<table width="100%" id="subListas"></table>
					</td>	
				</tr>
				<tr>
					<td></td>
					<td align="center">
						<img src="images/compras/add-16x16.png" title="Agregar" onClick="agregarLineaUsuario();" style="cursor:pointer">
					</td>
				</tr>
				<tr>
					<td align="right" colspan="2">						
						<input type="button" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" onClick="javascript:guardarListaCorreo()">
						<input type="button" id="cancelar" name="cancelar" value="<% =GF_TRADUCIR("Cancelar") %>" onClick="javascript:CerrarPopUp()">
					</td>
				</tr>
			</table>
		</td>	
	</tr>	
</table>

<input type="hidden" name="accion" value="<% =ACCION_SUBMITIR %>">
<input type="hidden" id="IdLista" name="IdLista" value="<% =IdLista %>">
<input type="hidden" id="idDivision" name="idDivision" value="<% =IdDivision %>">
<input type="hidden" id="cantUsuario" name="cantUsuario"  value="0">
</form>
<iframe name="ifrmLista" id="ifrmLista" width="1px" height="1px" style="visibility:hidden"></iframe>		
</body>
</html>