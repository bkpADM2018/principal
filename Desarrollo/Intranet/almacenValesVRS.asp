<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->

<%
Const ID_ARTICULO_ORIGEN = "item0"
Const ID_ARTICULO_DESTINO = "item1"
Const SALDO_ORIGEN = "saldo0"
Const SALDO_DESTINO = "saldo1"
Const STOCK_ACTUAL_ORIGEN = "amount0"
Const STOCK_ACTUAL_DESTINO = "amount1"

'-------------------------------------------------------------------------------------------------
Function getComboAlmacenes(idAlmacen)

	Set rsAlmacenes = obtenerListaAlmacenesUsuario()
%>					
	<select id="idalmacen" name="idalmacen" onChange="cambioAlmacen(this)">
		<option value = 0>- Seleccione -</option>
		<%
		while not rsAlmacenes.EoF
			%>
			<option value="<%=rsAlmacenes("idalmacen")%>" <% if (cdbl(rsAlmacenes("idalmacen")) = idalmacen) then response.write "selected"%>>
				<%=rsAlmacenes("dsalmacen")%>
			</option>
			<%
			rsAlmacenes.MoveNext()
		wend
		%>
	</select>
<%
End Function
'-------------------------------------------------------------------------------------------------
Function grabarValeReclasificacion(pIdAlmacen, pIdArticuloOrigen, pIdArticuloDestino, pCantidad)
    Dim nroVale
        
	'1ro obtengo el stock actual del artículo origen leido.
	strSql = "select * from TBLARTICULOSDATOS WHERE IDARTICULO = " & pIdArticuloOrigen & " and IDALMACEN=" & pIdAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
    if (not rs.eof) then        
        VS_existencia = CDbl(rs("EXISTENCIA"))
        VS_sobrante =  CDbl(rs("SOBRANTE"))
        '2do Averiguo cuanto corresponde a sobrante y cuanto a Existencia. Priorizo la reclasificacion de sobrantes ya que no tienen costo contable.
        if (pCantidad <= VS_sobrante) then
            VS_existencia = 0
            VS_sobrante =  pCantidad
        else
            'VS_sobrante se toma completo y la existencia es lo que queda de cantidad sin que lo tome el sobrante.
            VS_existencia = pCantidad-VS_sobrante
        end if
        'Grabo el vale.        
        Call grabarHeaderVale(nroVale,0)
        'Grabo detalle destino.        
        VS_idArticulo = pIdArticuloDestino
        VS_saldo = VS_existencia + VS_sobrante
        Call grabarValeDetalle(nroVale, 0)
        Call actualizarStock()
        'Grabo detalle origen.
        VS_idArticulo = pIdArticuloOrigen
        VS_saldo = VS_saldo * -1
        VS_existencia = VS_existencia * -1
        VS_sobrante =  VS_sobrante * -1
        Call grabarValeDetalle(nroVale, 0)
        Call actualizarStock()        
    end if
    grabarValeReclasificacion = nroVale
End Function
'*****************************************************************
'*************        COMIENZO DE PAGINA            **************
'*****************************************************************
Dim nroVale,idalmacen,esNuevo,rsAlmacenes

nroVale = GF_PARAMETROS7("nroVale",0,6)
idArticuloOrigen = GF_PARAMETROS7(ID_ARTICULO_ORIGEN,"",6)
idArticuloDestino = GF_PARAMETROS7(ID_ARTICULO_DESTINO,"",6)
dsArticuloOrigen = GF_PARAMETROS7("articuloOrigen","",6)
dsArticulODestino = GF_PARAMETROS7("articuloDestino","",6)
stockOrigen = GF_PARAMETROS7(STOCK_ACTUAL_ORIGEN,"",6)
stockDestino = GF_PARAMETROS7(STOCK_ACTUAL_DESTINO,"",6)
unidadDestino = GF_PARAMETROS7("unidadDestino","",6)
unidadOrigen = GF_PARAMETROS7("unidadOrigen","",6)
categoriaOrigen = GF_PARAMETROS7("categoriaOrigen","",6)
categoriaDestino = GF_PARAMETROS7("categoriaDestino","",6)
idalmacen = GF_PARAMETROS7("idalmacen",0,6)
accion = GF_PARAMETROS7("accion","",6)
mover = GF_PARAMETROS7("Mover",2,6)
stockFinOrigen = GF_PARAMETROS7(SALDO_ORIGEN,2,6)
stockFinDestino = GF_PARAMETROS7(SALDO_DESTINO,2,6)

Call GP_ConfigurarMomentos()

call initHeaderVale(nroVale)
VS_cdSolicitante = session("Usuario")
VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
Call initArticulosVale(nroVale)

if not (isFormSubmit()) then
	VS_idAlmacen = idalmacen 
else
	VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))		
	'Controlar el Vale
	controlOK = controlarVale(nroVale)
	if (mover > 0) then
		if ((accion = ACCION_GRABAR) and (controlOK)) then			
			nroVale = grabarValeReclasificacion(idalmacen, idArticuloOrigen, idArticuloDestino, mover)			
			Call grabarRegistroFirmas(nroVale)
			Call grabarPreciosVigentesPorArticulo(nroVale)
			Call ActualizarPrecios(nroVale, CODIGO_VS_RECLASIFICACION_STOCK)
			Response.Redirect "almacenAjustes.asp"
		end if
	else
		Call setError(CANTIDAD_NO_NEGATIVA)
	end if 
end if
%>
<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
		<link href="css/Toolbar.css" rel="stylesheet" type="text/css" />

		<script type="text/javascript" src="scripts/channel.js"></script>
		<script type="text/javascript" src="scripts/Toolbar.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
		
		<style>
			.ui-autocomplete-loading { background: white url('images/loading_small_green.gif') right center no-repeat; }

			.ui-autocomplete-category {
				font-weight: bold;
				padding: .2em .4em;
				margin: auto;
				text-align:center;
				line-height: 1.5;
			}
			
		</style>
		<script>
			var ENABLED = "activar";
			var DISABLED = "deshabilitar";
			
			ch = new channel();
			
			var flagLetSave = false;
			
			function crearAutoCompletes() {
				lastCategory = "";
				$( "#articuloOrigen" ).autocomplete({
					minLength: 2,
					//El source se setea al seleccionar un almacen
					//source: "comprasStreamElementos.asp?tipo=JQArticulos",
					focus: function( event, ui ) {
						$( "#articuloOrigen" ).val(ui.item.dsarticulo);
						return false;
					},
					select: function( event, ui ) {
						$( "#articuloOrigen"    ).val (ui.item.dsarticulo);
						$( "#<%=ID_ARTICULO_ORIGEN%>"    ).val (ui.item.idarticulo);
						$( "#unidadOrigen"	  ).val (ui.item.abreviatura);
						$( "#categoriaOrigen"	  ).val(ui.item.dscategoria);
						
						$("#<%=STOCK_ACTUAL_ORIGEN%>").val(ui.item.stock);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item)
						{
							lastCategory = "";
							$( "#articuloOrigen" ).val("");
							$( "#<%=ID_ARTICULO_ORIGEN%>"  ).val ("");
							$( "#unidadOrigen"   ).val("");
							$( "#categoriaOrigen"  ).val("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					if (item.stock == null) {
						item.stock = 0;
					}
					
					li_Item = $( "<li></li>" )
								.data( "item.autocomplete", item )
								.append( "<a><font style='font-size:10;'>" + item.idarticulo + " - " + item.dsarticulo + " - "+item.stock+ " ["+item.abreviatura+"]</font></a>" )
								.appendTo( ul );
								
					if (lastCategory != item.idcategoria) {
						lastCategory = item.idcategoria;
						return $(ul)
							.append( "<li class='ui-autocomplete-category'>" + item.dscategoria + "</li>" ).append(
								li_Item
							);
					} else {
						return li_Item;
					}
					
				};
				
				
				$( "#articuloDestino" ).autocomplete({
					minLength: 2,
					//El source se setea al seleccionar un almacen
					//source: "comprasStreamElementos.asp?tipo=JQArticulos",
					focus: function( event, ui ) {
						$( "#articuloDestino" ).val(ui.item.dsarticulo);
						return false;
					},
					select: function( event, ui ) {
						$( "#articuloDestino"    ).val (ui.item.dsarticulo);
						$( "#<%=ID_ARTICULO_DESTINO%>"    ).val (ui.item.idarticulo);
						$( "#unidadDestino"	  ).val (ui.item.abreviatura);
						$( "#categoriaDestino"	  ).val (ui.item.dscategoria);

						$("#<%=STOCK_ACTUAL_DESTINO%>").val(ui.item.stock);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item)
						{
							lastCategory = "";
							$( "#articuloDestino" ).val("");
							$( "#<%=ID_ARTICULO_DESTINO%>"  ).val ("");
							$( "#unidadDestino"   ).val("");
							$( "#categoriaDestino"  ).val("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					li_Item = $( "<li></li>" )
								.data( "item.autocomplete", item )
								.append( "<a>" + item.idarticulo + " - <font style='font-size:10;'>" + item.dsarticulo + " - "+item.stock +" ["+item.abreviatura+"]</font></a>" )
								.appendTo( ul );
								
					if (lastCategory != item.idcategoria) {
						lastCategory = item.idcategoria;
						return $(ul)
							.append( "<li class='ui-autocomplete-category'>" + item.dscategoria + "</li>" ).append(
								li_Item
							);
					} else {
						return li_Item;
					}
				};
			}
			
			function bodyOnLoad()
			{
				deshabilitarElementos();
				crearAutoCompletes();
				
				<% if ((accion = ACCION_GRABAR) or (accion = ACCION_CONTROLAR)) then %>
					habilitarElementos()
				<% end if %>
				
				var tb = new Toolbar('toolbar', 5, "images/almacenes/");	
				tb.addButton("Home-16x16.png", "Home", "location.href = 'almacenIndex.asp'");		
				tb.addButtonSAVE("Guardar", "submitir('<% =ACCION_GRABAR %>')");
				tb.addButtonCONFIRM("Controlar",  "submitir('<% =ACCION_CONTROLAR %>')");
				tb.addButton("Setting_folder-16x16.png", "Ajustes", "irAjustes()");										
				tb.draw();		
			}
			function irAjustes() {
				location.href = "almacenAjustes.asp";
			}
			
			function submitir(acc)
			{
				document.getElementById("accion").value = acc;
				$("#myForm").submit();
			}
			
			function cambioAlmacen(me)
			{
				limpiar();
				if ($(me).val() == 0){
					deshabilitarElementos();
				}else{
					habilitarElementos();
					var almacen = $(me).val();
					//No utilizar la siguiente forma porque al realizar el cambio realiza el pedido de datos
					//$( "#articuloOrigen" ).autocomplete("source" , "comprasStreamElementos.asp?tipo=JQArticulos" );
					$( "#articuloOrigen" ).autocomplete("option", "source" , "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen="+almacen );
					$( "#articuloDestino" ).autocomplete("option", "source" , "comprasStreamElementos.asp?tipo=JQArticulos&idAlmacen="+almacen );
				}
			}
			
			function limpiar()
			{
				$("#articuloOrigen").val("");
				$("#<%=ID_ARTICULO_ORIGEN%>").val("");
				$("#<%=STOCK_ACTUAL_ORIGEN%>").val("");
				$("#unidadOrigen").val("");
				$("#categoriaOrigen").val("");
				$("#articuloDestino").val("");
				$("#<%=ID_ARTICULO_DESTINO%>").val("");
				$("#<%=STOCK_ACTUAL_DESTINO%>").val("");
				$("#unidadDestino").val("");
				$("#categoriaDestino").val("");
				
			}
			
			function habilitarElementos(){
				if ($("#idalmacen").val() != 0){
					toggleStatus($("#articuloOrigen"),ENABLED);
					toggleStatus($("#articuloDestino"),ENABLED);
					toggleStatus($("#mover"),ENABLED);
				}
			}

			function deshabilitarElementos(){
				toggleStatus($("#articuloOrigen"),DISABLED);
				toggleStatus($("#articuloDestino"),DISABLED);
				toggleStatus($("#mover"),DISABLED);
			}
			
			function calcularFinal()
			{
				var stockO = $("#<%=STOCK_ACTUAL_ORIGEN%>").val();
				var stockD = $("#<%=STOCK_ACTUAL_DESTINO%>").val();
				var stockM = $("#mover").val();
				
				if (stockO != "" && stockD != "" && stockM != "")
				{
					$("#<%=SALDO_ORIGEN%>").val(Number(stockO) - Number(stockM));
					$("#<%=SALDO_DESTINO%>").val(Number(stockD) + Number(stockM));
				}								
			}
			
			function toggleStatus(elem,action) { 
				if (action == DISABLED) {
					$(elem).attr('disabled', true); 
				} else { 
					$(elem).removeAttr('disabled'); 
				}
			} 
			
		</script>
		<style>
			.deshabilitado{
				background-color: #DCDCDC;
				border: 1px solid #DCDCDC;
				color:#006400;
				font-weight:bold;
			}
		</style>
	</head>
	<body onLoad="bodyOnLoad()">
		<% call GF_TITULO2("kogge64.gif","Vale Reclasificacion de Stock") %>	
		<div id="toolbar"></div>
		<br />
		<%call showErrors()%>
		<br />
		<form action="almacenValesVRS.asp" method="get" id="myForm">
			<table align="center"  class="reg_header">
				<tr>
					<th width="150px" colspan="2" class="reg_header_nav ui-corner-tl">Almacen </th>
					<th align="left"><% call getComboAlmacenes(idalmacen)%></th>
			  </tr>
				<tr>
					<td colspan="3">
						<div id="articulos">
							<table cellpadding="2" cellspacing="2" class="reg_header" align="center">
							  

							  <tr>
								<th colspan="4" class="reg_header_nav ui-corner-top">Articulos</th>
							  </tr>
							  <tr>
								<th colspan="2" class="reg_header_nav">Origen</th>
								<th colspan="2" class="reg_header_nav">Destino</th>
							  </tr>
							  <tr>
								<td colspan="2" align="center"><input type="text" id="articuloOrigen" name="articuloOrigen" value="<%=dsArticuloOrigen%>">
								  <input type="hidden" id="<%=ID_ARTICULO_ORIGEN%>" name="<%=ID_ARTICULO_ORIGEN%>" value="<%=idArticuloOrigen%>"></td>
								<td colspan="2" align="center"><input type="text" id="articuloDestino" name="articuloDestino" value="<%=dsArticuloDestino%>">
								  <input type="hidden" id="<%=ID_ARTICULO_DESTINO%>" name="<%=ID_ARTICULO_DESTINO%>" value="<%=idArticuloDestino%>"></td>
							  </tr>
							  <tr>
								<td width="70" class="reg_header_navdos">Stock inicial</td>
								<td width="136" align="right" class="reg_header_navdos">
								<input type="text" size="5" readonly style="text-align:right" class="deshabilitado" id="<%=STOCK_ACTUAL_ORIGEN%>" name="<%=STOCK_ACTUAL_ORIGEN%>" value="<%=stockOrigen%>">&nbsp;							</td>
								<td width="134" class="reg_header_navdos">
									&nbsp;
								<input type="text" size="5" class="deshabilitado" readonly id="<%=STOCK_ACTUAL_DESTINO%>" name="<%=STOCK_ACTUAL_DESTINO%>" value="<%=stockDestino%>">					        </td>
								<td width="70" align="right" class="reg_header_navdos">Stock inicial</td>
							  </tr>
							  <tr>
							    <td class="reg_header_navdos">Mover</td>
						        <td colspan="2" align="center" class="reg_header_navdos"><input name="mover" style="text-align:center" type="text" id="mover" size="6" onBlur="calcularFinal()" value="<%=mover%>"></td>
						        <td align="right" class="reg_header_navdos">Mover</td>
							  </tr>
							  <tr>
							    <td class="reg_header_navdos">Stock final</td>
							    <td align="right" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" style="text-align:right" name="<%=SALDO_ORIGEN%>" id="<%=SALDO_ORIGEN%>" value="<%=stockFinOrigen%>">								</td>                              
							    <td align="left" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" style="text-align:left" name="<%=SALDO_DESTINO%>" id="<%=SALDO_DESTINO%>" value="<%=stockFinDestino%>">								</td>                        
							    <td align="right" class="reg_header_navdos">Stock final</td>
						      </tr>
							  <tr>
								<td class="reg_header_navdos">Unidad</td>
								<td align="right" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" style="text-align:right" name="unidadOrigen" id="unidadOrigen" value="<%=unidadOrigen%>">
								<td align="left" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" name="unidadDestino" id="unidadDestino" value="<%=unidadDestino%>">
								<td align="right" class="reg_header_navdos">Unidad</td>
							  </tr>
							  <tr>
								<td class="reg_header_navdos">Categoria</td>
								<td align="right" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" style="text-align:right" name="categoriaOrigen" id="categoriaOrigen" value="<%=categoriaOrigen%>">								</td>
								<td align="left" class="reg_header_navdos">
									<input type="text" readonly class="deshabilitado" name="categoriaDestino" id="categoriaDestino" value="<%=categoriaDestino%>">								</td>
								<td align="right" class="reg_header_navdos">Categoria</td>
							  </tr>
							</table>
					  </div>			  </td>
				</tr>
			</table>
			<input type="hidden" id="accion" name="accion" value="">
			<input type="hidden" id="cdVale" name="cdVale" value="<%=CODIGO_VS_RECLASIFICACION_STOCK%>">
		</form>
</body>
</html>