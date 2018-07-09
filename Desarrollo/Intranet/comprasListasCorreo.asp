<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%
'-------------------------------------------------------------------------------
Function cargarListasCorreo(pDivision,pCdLista,DsLista)
	dim rs, cn, strSQL,myWhere
	call buscarFiltrosListaCorreo(myWhere,pDivision,pCdLista,DsLista)
	strSQL = " SELECT A.IDLISTA,A.CDLISTA, A.DSLISTA, A.IDDIVISION, B.DSDIVISION ,A.CDUSUARIO FROM TBLMAILLSTCABECERA A "
	strSQL = strSQL & " INNER JOIN TBLDIVISIONES B ON B.IDDIVISION = A.IDDIVISION "
	strSQL = strSQL & myWhere &" ORDER BY A.IDLISTA "
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set cargarListasCorreo = rs
End function
'-------------------------------------------------------------------------------
Function buscarFiltrosListaCorreo(ByRef myWhere,pDivision,pCdLista,pDsLista)
	if (pCdLista <> "") then Call mkWhere(myWhere, "A.CDLISTA", Trim(UCase(pCdLista)), "LIKE", 3)
	if (pDivision > 0) then	Call mkWhere(myWhere, "A.IDDIVISION", pDivision, "=", 1)
	if (pDsLista <> "") then Call mkWhere(myWhere, "A.DSLISTA", pDsLista, "LIKE", 3)
	buscarFiltrosListaCorreo = myWhere
End function
'-------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
    if (not isEmpty(p_strValue)) then
         if (isEmpty(p_strParam)) then
           p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
           p_strParam = p_strParam & p_strKey & "=" & p_strValue
      end if
End Function
'-------------------------------------------------------------------------------
Function EliminarListaCorreo(pIdLista)
	dim rs, cn, strSQL,myWhere,strSQL1
	strSQL = " DELETE FROM TBLMAILLSTCABECERA WHERE IDLISTA = "&pIdLista
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)
	strSQL1 = " DELETE FROM TBLMAILLSTSDETALLE WHERE IDLISTA = "&pIdLista
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL1)
End Function
'-------------------------------------------------------------------------------
'******************************************************************************
'								INICIO PAGINA
'******************************************************************************
Dim IdLista,DsLista, hayBusqueda,dsSolicitante,mostrar,paginaActual,lineasTotales
Dim reg,busquedaActiva,accion,idDivision,CdLista

IdLista = GF_PARAMETROS7("IdLista", 0, 6)
call addParam("IdLista", IdLista, params)
CdLista = GF_PARAMETROS7("CdLista", "", 6)
call addParam("CdLista", CdLista, params)
DsLista = GF_PARAMETROS7("DsLista", "", 6)
call addParam("DsLista", DsLista, params)
accion = GF_PARAMETROS7("accion", "", 6)
idDivision = GF_PARAMETROS7("idDivision", 0, 6)
call addParam("idDivision", idDivision, params)

hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", busquedaActiva, params)
if busquedaActiva = 1 then hayBusqueda = true

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

if(accion = ACCION_BORRAR)then  call EliminarListaCorreo(IdLista)	

Set rsList = cargarListasCorreo(idDivision,CdLista,DsLista)
Call setupPaginacion(rsList, paginaActual, mostrar)
lineasTotales = rsList.recordcount


%>
<html>
<head>
<title><%=GF_TRADUCIR("Lista Correos")%></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<style type="text/css">
.labelStyle {
	font-weight: bold;
	text-align: center;
}
.numberStyle {
	font-weight: bold;
	font-size: 14px;
}
.divOculto {
	display: none;
}
</style>
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/script_fechas.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">	
	var ch = new channel();
	
	function bodyOnLoad() {
		tb = new Toolbar('toolbar', 6,'images/almacenes/');
		tb.addButton("Previous-16x16.png", "Volver", "Volver()");
		tb.addButton("refresh-16x16.png", "Recargar", "submitInfo()");		
		var swt = tb.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.addButton("add-16x16.png", "Nueva Lista", "loadPopUpNew()");				
		tb.draw();
		<%	if (hayBusqueda) then %>
				tb.changeState(swt);
		<%	end if 
		 	if (not rsList.eof) then %>
				var pgn = new Paginacion("paginacion");				
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "comprasListasCorreo.asp<% =params %>");					
		<%	end if 	%>		
		autocompleteDsLista();	
		pngfix();	
	}
	
	function buscarOn() {
		document.getElementById("busqueda").className = "";	
		document.getElementById("busquedaActiva").value = "1";	
	}
	
	function buscarOff() {
		document.getElementById("busqueda").className = "divOculto";		
		document.getElementById("busquedaActiva").value = "0";		
	}
	
	function EliminarLista(pIdLista){
		if (confirm("Esta seguro que desea eliminar la lista seleccionada?")){
			ch.bind("comprasListasCorreo.asp?accion=<%=ACCION_BORRAR%>&IdLista="+ pIdLista ,"loadPopUpNew_callback()");
			ch.send();
		}	
	}
	function loadPopUpNew_callback(){
		submitInfo();
	}
	
	function Volver(){
		location.href = "comprasAdministracion.asp";
	}
	function submitInfo() {
		document.getElementById("frmSel").submit();
	}

	function loadPopUpNew(id) {
		var puw = new winPopUp('popupNuevaLista','comprasPopUpListaCorreo.asp?','375','250','Nueva Lista Correo', "loadPopUpNew_callback()");		
	}
	
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	
	function abrirListaCorreo(pIdLista, pDsLista, pIdDivision,pCdLista){
		var puw = new winPopUp('popupNuevaLista','comprasPopUpListaCorreo.asp?IdLista='+pIdLista+'&DsLista='+pDsLista+'&IdDivision='+pIdDivision+'&CdLista='+pCdLista+'&accion=<%=ACCION_CONTROLAR%>','375','250','Editar Lista Correo', 'loadPopUpNew_callback()');
	}
	
	

	function autocompleteDsLista()
	{	
		$(function() {
		$( "#DsLista" ).autocomplete({
		minLength: 2,
		source: function(request,response){
			$.ajax({
				url: "comprasStreamElementos.asp",
				dataType: "json",
			data: {				
				term : request.term,
				Tipo : "JQListaCorreo",
				DsLista : document.getElementById("DsLista").value
				 },
		    success: function(data) {				
				response(data);
				}
			});	
		},		
		focus: function( event, ui ) {
				$( "#DsLista").val(ui.item.dslista);
				return false;
			},
		select: function( event, ui ) {
				$( "#DsLista").val (ui.item.dslista);				
				return false;
			}		
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
			.data( "item.autocomplete", item )
			.append( "<a><font style='font-size:10;'>" + item.dslista + "</font></a>" )
			.appendTo( ul );
		};
	});
		
		
		
		
		
		
		
		
		
		
	}
	
	
</script>
</head>
<body onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Lista de Correos") %>
<div id="toolbar"></div>
<br>		
<form id="frmSel" name="frmSel" action="comprasListasCorreo.asp" method="POST">
<div id="busqueda" class="divOculto">
	<br><br>	
	<table id="tblBusqueda" width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
           <td width="8"><img src="images/marco_r2_c3.gif"></td>
           <td align="right"></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="images/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="95%" align="center" border="0">
                            <tr>								
								<td width="15%" align="right"><% = GF_TRADUCIR("Codigo") %>:</td>
								<td width="20%">
									<input type="text"  id="CdLista" name="CdLista" value="<%=CdLista%>">
								</td>
								<%
								strSQL = "Select IDDIVISION, DSDIVISION from TBLDIVISIONES "
                                Call executeQueryDb(DBSITE_SQL_INTRA, rsDivisiones, "OPEN", strSQL)
								%>                                
                                <td align="right"><% =GF_TRADUCIR("Division") %>:</td>
                                <td>                                
									<select style="z-index:-1;" name="idDivision">
									        <option SELECTED value="<% =SIN_DIVISION %>">- <% =GF_TRADUCIR("Seleccione") %> -
									<%		while (not rsDivisiones.eof)		
												selected = ""										
												if (CLng(rsDivisiones("IDDIVISION")) = CLng(idDivision)) then selected = "selected"
									%>
												<option value="<% =rsDivisiones("IDDIVISION") %>" <% =selected %>><% =rsDivisiones("DSDIVISION") %>                                        
									<%			rsDivisiones.MoveNext()
											wend 	
									%>
									</select>
                                </td>					
                            </tr>
                            <tr>
								<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
								<td >
									<input type="text"  id="DsLista" name="DsLista" value="<%=DsLista%>">									
								</td>
                            </tr>
							<tr>																							
								<td colspan="4" align="center"><br></br><input type="submit" value="Buscar" id="submit1" name="submit1" onclick="submitInfo();"></td>
							</tr>		
								
                     </table>
	           </td>
	           <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="images/marco_r3_c1.gif"></td>
	           <td width="100%" align="center" colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="images/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	</div> 	
	<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">
	<input type="hidden" name="accion" id="accion" value="<%=ACCION_SUBMIT%>">		
	
<br>
	
	<br>
<table class="reg_Header" align="center" width="65%" border="0">	
	<tr><td colspan="5"><div id="paginacion"></div></td></tr>				
	<tr>
		<td width="15%" class="reg_header_nav"  align="center"><%=GF_Traducir("Codigo")%></td>
		<td width="45%" class="reg_header_nav"  align="center"><%=GF_Traducir("Descripcion")%></td>		
		<td width="30%" class="reg_header_nav"  align="center"><%=GF_Traducir("Division")%></td>	
		<td width="5%" class="reg_header_nav"   align="center"><%=GF_Traducir(".")%></td>		
		<td width="5%" class="reg_header_nav"   align="center"><%=GF_Traducir(".")%></td>		
	</tr>
	<% 
	if rsList.eof then %>
		<tr>
			<td align="center" colspan="8">
				<%=GF_Traducir("No se encontraron resultados")%>
			</td>
		</tr>
	<% end if
		while ((not rsList.eof) and (CInt(reg) < CInt(mostrar)))
			reg = reg + 1 %>
			<tr class="reg_Header_navdos" onMouseOver="javascript:lightOn(this)" onMouseOut="javascript:lightOff(this)">			
				<td  align="center" ><%=rsList("CDLISTA")%></td>
				<td  align="left" ><%=rsList("DSLISTA")%></td>
				<td  align="left" ><%=rsList("DSDIVISION")%></td>
				<td  align="center" onclick="abrirListaCorreo(<%=rsList("IDLISTA")%>,'<%=rsList("DSLISTA")%>','<%=rsList("IDDIVISION")%>','<%=rsList("CDLISTA")%>')"><img title="Editar" src="images/compras/edit-16x16.png" style="cursor:pointer"></td>
				<td  align="center" onclick="EliminarLista(<%=rsList("IDLISTA")%>)"><img title="Eliminar" src="images/compras/cancel-16x16.png" style="cursor:pointer"></td>
			</tr>
		<%rsList.movenext
		wend %>
</table>
<input type="hidden" id="IdLista" name="IdLista" value="<% =IdLista %>">
</form>

</body>
</html>