<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%

Const CTST_MIN_ARTICULOS     = 20
'-----------------------------------------------------------------------------------------
'Genera un numero aleatorio entre 0 y 1, a eso lo multiplico por el la cantidad maxima de 
'registro y obtengo X casos posibles de articulos
Function elegirIndiceArticulo(pcant, pmax, pmin)	    
	Dim  resultado,	strSeleccionados, vecSeleccionados(), i, cant
		
    cant = pcant
    if (cant > (pmax-pmin)) then cant = pmax-pmin

	strSeleccionados = ";"   'Se utilizará para mejorar el método de selección aleatoria.
	Redim vecSeleccionados(cant)
	i = -1
	while (i < cant)
	    'response.Write "inicia: " & i & "|" & cant & "<br>"
	    'Se elige un numero al azar.
	    Randomize()
	    resultado =  cInt(((pmax-pmin) * Rnd) + pmin) 
	    'response.Write "Resultado: " & resultado & "<br>"
	    'Si este numero ya no fue elegido se toma como válido. Sino, se intenta con el próximo en orden ascendente.
	    while (InStr(1,strSeleccionados, ";" & resultado & ";") <> 0)	 
	        resultado = resultado + 1
	        if (resultado > pmax) then resultado = pmin
	        'response.Write "No encontrado, siguiente: " & resultado & "<br>"	         
	    wend
	    'Se encontro un nuevo numero!
	    strSeleccionados = strSeleccionados & resultado & ";"	    	    
	    'response.Write "Guardado! " & strSeleccionados & "<br>"
	    i= i + 1
	    vecSeleccionados(i) = resultado	    
	wend	
	'if (x = 1000) then response.Write "Mal!<br>"
	elegirIndiceArticulo = vecSeleccionados
End Function
'------------------------------------------------------------------------------------------
Function getArticulosRegistrados(precioMinimo, precioMaximo, pchkStock, idAlmacen)
	Dim strSQL ,rs,oConn, tCambio
	
	tCambio = CDbl(getTipoCambio(MONEDA_DOLAR, ""))
	strSQL = " SELECT IDARTICULO FROM TBLREPORTESTOCKWF where CDUSUARIO='" & session("Usuario") & "' and IDALMACEN= " & idAlmacen 
	if (precioMinimo > 0) then strSQL = strSQL & " and VLUPESOS >= " & precioMinimo * tCambio * 100
	if (precioMaximo > 0) then strSQL = strSQL & " and VLUPESOS <= " & precioMaximo * tCambio * 100	
	if (pchkStock = CTST_ARTICULO_CON_STOCK) then strSQL = strSQL & " and (EXISTENCIA > 0 or SOBRANTE > 0)"	
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getArticulosRegistrados = rs	
End function
'---------------------------------------------------------------------------------
Function cargarArticulos(cantArticulos, precioMinimo, precioMaximo, pchkStock, idAlmacen)
dim item,rtrn,min, max,rsArticulos,v_art, vec

Set rsArticulos = getArticulosRegistrados(precioMinimo,precioMaximo, pchkStock, idAlmacen)
v_art = ""
if (not rsArticulos.eof) then
    min = 0
    max = rsArticulos.recordcount-1

    vec = elegirIndiceArticulo(CantidadArticulos, max, min)		
    
    For each item in vec	
		    rsArticulos.moveFirst()				
		    rsArticulos.move(item)
		    v_art = v_art & rsArticulos("IDARTICULO") & ";"			
    next
    'Saco el ultimo punto y coma.
    if (Len(v_art) > 1) then v_art = Left(v_art, Len(v_art)-1)      
end if
cargarArticulos = v_art
end function
'---------------------------------------------------------------------------------------
Dim rsAlmacenes,IdAlmacen,pSeleccion,CantidadArticulos,accion,idControl_new,controlOK,rtrn,index,vArticulos()
Dim i,CtSt_cdResponsable,CtSt_dsResponsable, precioMinimo, chkStock, precioMaximo, pChkStock

pSeleccion = GF_PARAMETROS7("tipoReporte","",6)
IdAlmacen = GF_PARAMETROS7("rAlmacen",0,6)
CantidadArticulos = GF_PARAMETROS7("cantArticulostxt",0,6)
precioMinimo = GF_PARAMETROS7("precioMinimo",0,6)
precioMaximo = GF_PARAMETROS7("precioMaximo",0,6)
pChkStock = GF_PARAMETROS7("chkStock","",6)
accion = GF_PARAMETROS7("accion","",6)
CtSt_cdResponsable = GF_PARAMETROS7("CtSt_cdResponsable","",6)
CtSt_dsResponsable = getUserDescription(CtSt_cdResponsable)

controlOK = false
if (pSeleccion = "") then pSeleccion = CTST_SELECCION_MANUAL
if(accion = ACCION_CONTROLAR) then
	'Controles
	if(CantidadArticulos < CTST_MIN_ARTICULOS)then	setError(CANTIDAD_ATICULOS_CTST)
	if(IdAlmacen = 0)then setError(ALMACEN_NO_EXISTE)	
	if(CtSt_cdResponsable = "")then setError(FALTA_RESPONSABLE)
	'Si no hubo error, se graban los articulos
	if (not hayError())then	controlOK = true		
end if
if(accion = ACCION_PROCESAR)then
    if(pSeleccion = CTST_SELECCION_AUTOMATICA)then
		response.write cargarArticulos(CantidadArticulos, precioMinimo,precioMaximo,  pChkStock, IdAlmacen)
	end if
	response.End
end if

%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<style type="text/css">.pp{font: normal normal 18px Times;text-align: center;style="color:blue;"}</style>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
var ch= new channel();
var pmsj;
var refPopUpCtSt;
var ITEM_ID = "item";
var lastArticulos = 0;
function BodyOnLoad(){
	refPopUpCtSt = getObjPopUp('popupNuevoCtSt');	
	<%	if (controlOK) then 
			 if (pSeleccion = CTST_SELECCION_MANUAL) then %>				
				//window.open("almacenCtStSelecciones.asp?IdAlmacen=<%=IdAlmacen%>&tipoReporte=<%=CTST_SELECCION_MANUAL%>", "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700", false);				
				var respuesta = window.showModalDialog("almacenCtStSelecciones.asp?IdAlmacen=<%=IdAlmacen%>&tipoReporte=<%=CTST_SELECCION_MANUAL%>&cdResponsable=<%=CtSt_cdResponsable%>","_blank","dialogHeight:500px;dialogLeft=400px;dialogWidth:700px;center=yes;status:no;scroll:no")
				if(respuesta == true) cerrarPopUp();	
			<% else %>			
				loadingAutomatica();
				//Se arma el workfile con os artículos y sus precios.
				ch.bind("almacenReporteStockAjax.asp?categoria=-1&incluir=true&almacen=<% =IdAlmacen %>&fechaBusqueda=<% = GF_FN2DTE(Left(session("MmtoDato"), 8)) %>", "armadoWFCallback()");
				ch.send();
			<% end if
		end if %>
	}
	
function crearLinea(pLinea){
	var myForm = document.getElementById('frmPop');
	var iCodigo = document.createElement('input');
	iCodigo.type = 'hidden';
	iCodigo.id = ITEM_ID + pLinea;
	iCodigo.name = ITEM_ID + pLinea;
	myForm.appendChild(iCodigo);			
}

function armadoWFCallback() {
    ch.bind("almacenPropCtStNuevo.asp?accion=<% =ACCION_PROCESAR%>&tipoReporte=<% =pSeleccion %>&rAlmacen=<% =IdAlmacen %>&chkStock=<%=pChkStock%>&precioMinimo=<% =precioMinimo %>&precioMaximo=<% =precioMaximo %>&cantArticulostxt=<% =CantidadArticulos %>&CtSt_cdResponsable=<% =CtSt_cdResponsable %>", "cargaAutoCallback()");    
	ch.send();
}

function cargaAutoCallback() {
    var resp = ch.response();    
    if (resp != "") {
        var vArticulos = resp.split(";");
        for (var i=0; i < vArticulos.length; i++) {
            //		crea la linea HTML dinamico					
		    crearLinea(lastArticulos);
    		//		le asigna el valor del vector cargado 		
		    $("#item"+lastArticulos).val (vArticulos[i]);
		    lastArticulos ++;
	    }		
	    document.getElementById("chkStock").value = '<%=pChkStock%>'
	    document.getElementById("frmPop").action= "almacenCtStGrabar.asp";
	    document.getElementById("frmPop").target= "ifrmAut";	    
	    document.getElementById("frmPop").submit();	    
	} else {
	    refPopUpCtSt.resize('400', '250');
	    document.getElementById('ControlGeneration').style.visibility = "hidden";
	    document.getElementById('ControlGeneration').style.position = "absolute";   
	    document.getElementById('RegularBody').style.visibility = "visible";	    
	    document.getElementById('RegularBody').style.position = "relative";
	    document.getElementById("frmPop").action= "almacenPropCtStNuevo.asp";
	    document.getElementById("frmPop").target= "";
	    document.getElementById("errDisplay").className="reg_Header_Error";    	    	    
	    document.getElementById("errDisplay").innerHTML="<% =POCOS_ARTICULOS & " - " & errMessage(POCOS_ARTICULOS) %>";
	}
}

function resultadoCarga_callback(pMsj,idControl){	
	document.getElementById('ControlGeneration').style.visibility = "hidden";
	document.getElementById('ControlGeneration').style.position = "absolute";		
	if (pMsj == "<% =RESPUESTA_OK %>") {	 	   	    
	    document.getElementById('SaveResponse').style.visibility = "visible";
	    document.getElementById('SaveResponse').style.position = "relative";
	    document.getElementById('newControl').innerHTML = idControl;
	  } else {	    
	    refPopUpCtSt.resize('400', '250');
	    document.getElementById('RegularBody').style.visibility = "visible";	    
	    document.getElementById('RegularBody').style.position = "relative";
	    document.getElementById("frmPop").action= "almacenPropCtStNuevo.asp";
	    document.getElementById("frmPop").target= "";
	    document.getElementById("errDisplay").className="reg_Header_Error";    	    	    
	    document.getElementById("errDisplay").innerHTML=pMsj;
	  }
}	

function submitirPagina(){	
   	document.getElementById('frmPop').submit();
}
		
function FinalizarCarga_CallBack(){
	cerrarPopUp();
}
	
function loadingAutomatica(){
	refPopUpCtSt.resize('250', '150');	
	document.getElementById('ControlGeneration').style.visibility = "visible";
	document.getElementById('ControlGeneration').style.position = "relative";
	document.getElementById('RegularBody').style.visibility = "hidden";	
	document.getElementById('RegularBody').style.position = "absolute";	
}

function ocultarPopUp(){
	refPopUpCtSt.hide();
}


function cerrarPopUp(){
	parent.window.submitInfo();
}

function mostrarCantArticulos(){
    refPopUpCtSt.resize('400','250');
	document.getElementById("myCantArticulos").style.display = "block";
}

function ocultarCantArticulos(){    
	document.getElementById("myCantArticulos").style.display = "none";	
	refPopUpCtSt.resize('400','200');
}

	$(function() {
			
				$( "#CtSt_Responsable" ).autocomplete({
				minLength: 2,
				source: "comprasStreamElementos.asp?tipo=JQPersonas",
				focus: function( event, ui ) {
					$( "#CtSt_Responsable" ).val(ui.item.nombre);
					return false;
				},
				select: function( event, ui ) {
					$( "#CtSt_Responsable"    ).val (ui.item.nombre);
					$( "#CtSt_cdResponsable"  ).val (ui.item.cdusuario );					
					return false;
				},
				change: function( event, ui ) {
					if (!ui.item)
					{						
						$( "#CtSt_cdResponsable"  ).val ("");						
					}
				}				
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
					.appendTo( ul );
			};
		});

</script>
</head>
<body onLoad="BodyOnLoad()">
<form id="frmPop" name="frmPop" action="almacenPropCtStNuevo.asp" method="POST">	
<div id="RegularBody" style="visibility: visible; position: relative">
<table width="100%" align="center">
	<tr>		
		<td colspan="2"><div id="errDisplay"><% call showErrors() %></div></td>
	</tr>		
	<tr>		
		<td align="right" ><% =GF_TRADUCIR("Responsable") %>:</td>
		<td align="left" >
			<input id="CtSt_Responsable" name="CtSt_Responsable"  style="width:185px" value="<%=CtSt_dsResponsable%>">						
		</td>		
	</tr>	
	<tr >
		<td  align="right"><% =GF_TRADUCIR("Almacen") %>:</td>
		<td>
		<%Set rsAlmacenes = obtenerListaAlmacenesUsuario()%>
        
			<select id="rAlmacen" name="rAlmacen">
				<option value="0">- <% =GF_TRADUCIR("Seleccione") %> -</option>
				<%
				while (not rsAlmacenes.eof)	%>
				<option value="<% =rsAlmacenes("IDALMACEN") %>" <% if (rsAlmacenes("IDALMACEN") = idAlmacen) then response.write "selected='true'" %>><% =GF_TRADUCIR(rsAlmacenes("CDALMACEN")) %> - <% =GF_TRADUCIR(rsAlmacenes("DSALMACEN")) %></option>
				<%
				rsAlmacenes.MoveNext()
				wend 	
				%>		
			</select>		
			<br></br>	
        </td>	
	</tr>	
	<tr>		
		<td align="right" ><% =GF_TRADUCIR("Seleccion") %>:</td>
	</tr>	
	<tr>
		<td colspan="2" align="center">
			<input onclick="mostrarCantArticulos()" type="radio" id="tipoReporte" name="tipoReporte" value="<%= CTST_SELECCION_AUTOMATICA%>"	<% if (pSeleccion = CTST_SELECCION_AUTOMATICA) then response.write "checked='checked'" %>/><label for="tipoReporte">Automatica</label> &nbsp&nbsp&nbsp
			<input onclick="ocultarCantArticulos()" type="radio" id="tipoReporte" name="tipoReporte" value="<%= CTST_SELECCION_MANUAL%>"		<% if (pSeleccion = CTST_SELECCION_MANUAL) then response.write "checked='checked'" %>/><label for="tipoReporte">Manual</label>
			<br></br>	
		</td>
	</tr>	
	<tr>	
		<td colspan="2" align="center">
			<table id="myCantArticulos" align="center" style="<% if (pSeleccion = CTST_SELECCION_MANUAL) then response.write "display:none" %>" width="100%">
				<tr>
					<td align="right">
						<% =GF_TRADUCIR("Cantidad") %>:
						<input size="4" type="text" id="cantArticulostxt" name="cantArticulostxt" value="<%=CantidadArticulos%>">
					</td>
					<td align="right" >
						<% =GF_TRADUCIR("Solo art. c/stock") %>:
						<input type="checkbox" id="chkStock" name="chkStock" value="<%=CTST_ARTICULO_CON_STOCK%>" checked>
					</td>										
				</tr>
				<tr>
					<td align="right" >
						<% =GF_TRADUCIR("Precio Minimo (u$s)") %>:
						<input size="4" type="text" id="precioMinimo" name="precioMinimo" value="<% =precioMinimo %>" onkeypress="return controlIngreso(this, event, 'I')">
					</td>														
					<td align="right" >
						<% =GF_TRADUCIR("Precio Maximo (u$s)") %>:
						<input size="4" type="text" id="precioMaximo" name="precioMaximo" value="<% =precioMaximo %>" onkeypress="return controlIngreso(this, event, 'I')">
					</td>										
				</tr>
			</table>	
		</td>
	</tr>		
	
	<tr > 						
		<td  colspan="2" align="right">
			<INPUT type="button" value="Aceptar" onclick="submitirPagina()">										
			<INPUT id="Cancelar" type="button" value="Cancelar" onclick="cerrarPopUp()">										
		</td>		
	</tr>
</table>
</div>
<div id="ControlGeneration" style="visibility: hidden; position: absolute">
    <table style='color:#0B6121;' width='100%' align='center' id='Table1' border='0' class='reg_header round_border_all'>
	    <tr>
	        <td align='center'><strong>Generando listado</strong></td>
	    </tr>
	    <tr>
	        <td align='center'><strong>de articulos</strong></td>
	    </tr>
	    <tr>
	        <td align='center'><img src='images/loading_bar_green.gif'></td>
	    </tr>
	    <tr>
	        <td align='center'><strong>Aguarde unos </strong></td>
	    </tr>
	    <tr>
	        <td align='center'><strong>instantes...</strong></td>
	    </tr>
	</table>
</div>
<div id="SaveResponse" style="visibility: hidden; position: absolute">
    <table width='100%' align='center' id='tbLoading' border=0 class='.pp'>
	    <tr>
	        <td align='center'><h3>El Id de Control asignado es </h3></td>
	    </tr>
	    <tr>
	        <td align='center' ><h3><span style='color:blue;'><div id="newControl"></div></span></h3></td>
	    </tr>
</table>
</div>
<input type="hidden" id="accion" name="accion" value="<% =ACCION_CONTROLAR %>">
<input type="hidden" name="CtSt_cdResponsable" id="CtSt_cdResponsable" value="<%=CtSt_cdResponsable%>">
<input type="hidden" id="IdAlmacen" name="IdAlmacen" value="<%=IdAlmacen %>">
<input type="hidden" id="cantArticulos" name="cantArticulos"  value="<%=CantidadArticulos%>">
</form>
<iframe name="ifrmAut" id="ifrmAut" width="1px" height="1px" style="visibility:hidden"></iframe>
</body>
</html>