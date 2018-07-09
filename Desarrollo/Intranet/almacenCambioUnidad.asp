<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosFechas.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp"-->
<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp"-->
<!-- #include file="Includes/procedimientosVales.asp"-->

<%
Call initAccessInfo(RES_ADM_AL)
'--------------------------------------------------------------------------------------------------
Function crearStockNuevoArticulo(pIdArticuloViejo,pIdArticuloNuevo,pFactor)
	
	strSQL =          "INSERT "
	strSQL = strSQL & "INTO   tblarticulosdatos "
	strSQL = strSQL & "       ( "
	strSQL = strSQL & "              IDARTICULO  , "
	strSQL = strSQL & "              EXISTENCIA  , "
	strSQL = strSQL & "              SOBRANTE    , "
	strSQL = strSQL & "              CDUSUARIO   , "
	strSQL = strSQL & "              MOMENTO     , "
	strSQL = strSQL & "              IDALMACEN   , "
	strSQL = strSQL & "              STOCKMINIMO , "
	strSQL = strSQL & "              STOCKMAXIMO , "
	strSQL = strSQL & "              COMPRAMINIMA, "
	strSQL = strSQL & "              COMPRAMAXIMA, "
	strSQL = strSQL & "              CDINTERNO  "
	strSQL = strSQL & "       ) "
	strSQL = strSQL & "       (SELECT '"&pIdArticuloNuevo&"' IDARTICULO , "
	strSQL = strSQL & "              EXISTENCIA*"&pFactor&"  , "
	strSQL = strSQL & "              SOBRANTE*"&pFactor&"    , "
	strSQL = strSQL & "              CDUSUARIO   , "
	strSQL = strSQL & "              MOMENTO     , "
	strSQL = strSQL & "              IDALMACEN   , "
	strSQL = strSQL & "              STOCKMINIMO*"&pFactor&" , "
	strSQL = strSQL & "              STOCKMAXIMO*"&pFactor&" , "
	strSQL = strSQL & "              COMPRAMINIMA, "
	strSQL = strSQL & "              COMPRAMAXIMA, "
	strSQL = strSQL & "              CDINTERNO  "
	strSQL = strSQL & "       FROM    tblarticulosdatos "
	strSQL = strSQL & "       WHERE   idarticulo = " & pIdArticuloViejo
	strSQL = strSQL & "       )"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
End Function 
'--------------------------------------------------------------------------------------------------
Function cambiarConversion(pUnidadOrigen,pUnidadDestino,pFactor)
	strSQL = "select * from tblunidadesconversion where idunidadorig =" & pUnidadOrigen 
	strSQL = strSQL & " and idunidaddest = " & pUnidadDestino
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.EoF) then
		'Ya existe el factor para las unidades
		if (cdbl(rs("factor")) <> cdbl(pFactor)) then
			'actualizo el factor con el nuevo 
			strSQL = "Update tblunidadesconversion set factor = " &pFactor
			strSQL = strSQL & " where idunidadorig = " & pUnidadOrigen
			strSQL = strSQL & " and idunidaddest = " & pUnidadDestino
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	else
		'No existe el factor para las unidades, lo creo
		strSQL = "insert into tblunidadesconversion (idunidadorig,idunidaddest,factor,cdusuario,momento)"
		strSQL = strSQL & " values("&pUnidadOrigen&","&pUnidadDestino&","&pFactor&",'"&session("Usuario")&"',"&session("MmtoSistema")&")"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
Function crearValeAjuste(pIdArt,pIdNewArt,pFc)
		strSQL =	" SELECT EXISTENCIA, SOBRANTE, IDALMACEN, IDARTICULO " & _
					" FROM TBLARTICULOSDATOS " & _ 
					" WHERE IDARTICULO IN(" & pIdArt & ") " & _
					" AND (EXISTENCIA<>0 OR SOBRANTE<>0) ORDER BY IDALMACEN"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
		while not rs.EoF
				'cdvale 		= CODIGO_VS_AJUSTE_STOCK
				cdvale 		= CODIGO_VS_RECLASIFICACION_STOCK
				fecha 		= left(session("MmtoSistema"),8)
				idAlmacen 	= rs("IDALMACEN")
				idobra 		= 0
				cdusuario 	= session("Usuario")
				momento 	= session("MmtoSistema")
				estado 		= ESTADO_ACTIVO
				idsector 	= 0
				nrovale 	= getNumeracionVale(rs("IDALMACEN"))
				cdsolic 	= session("Usuario")

				'Creo la cabecera
				strSQL = "Insert into tblvalescabecera"
				strSQL = strSQL & " (CDVALE,FECHA,IDALMACEN,IDOBRA,CDUSUARIO,MOMENTO,PARTIDAPENDIENTE,IDBUDGETAREA,IDBUDGETDETALLE,ESTADO,IDSECTOR,NRVALE,CDSOLICITANTE)"
				strSQL = strSQL & " values('"&cdvale&"',"&fecha&","&idAlmacen&","&idobra&",'"&cdusuario&"',"&momento&",0,0,0,"&estado&","&idsector&",'"&nrovale&"','"&cdsolic&"')"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
				
				strSQL = "select max(idvale) idvale from tblvalescabecera where cdvale = '" & cdvale & "'"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				
				idAjuste = rs("idvale")
				
				'creo el comentario
				abrevD = GF_PARAMETROS7("abrevD","",6)
				abrevO = GF_PARAMETROS7("abrevO","",6)
				
				strSQL = "insert into tblvalescomentarios (idvale,comentario) values("&idAjuste&",'Cambio de unidad del articulo. Factor de conversion: 1 "&abrevO&" = "&pFc&" "&abrevD&"')"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs0, "EXEC", strSQL)
				
				'creo las firmas
				strSQL = "insert into tblvalesfirmas (IDVALE,SECUENCIA,CDUSUARIO) values("&idAjuste&","&VS_FIRMA_RESPONSABLE&",'"&cdusuario&"') "
				Call executeQueryDB(DBSITE_SQL_INTRA, rs0, "EXEC", strSQL)
				strSQL = "insert into tblvalesfirmas (IDVALE,SECUENCIA,CDUSUARIO) values("&idAjuste&","&VS_FIRMA_GERENTE&",'"&VS_NO_USER&"') "
				Call executeQueryDB(DBSITE_SQL_INTRA, rs0, "EXEC", strSQL)
				strSQL = "insert into tblvalesfirmas (IDVALE,SECUENCIA,CDUSUARIO) values("&idAjuste&","&VS_FIRMA_COORD_AUDIT&",'"&VS_NO_USER&"') "
				Call executeQueryDB(DBSITE_SQL_INTRA, rs0, "EXEC", strSQL)
				
				'creo el detalle
				strSQL = "select * from tblarticulosprecios where idarticulo = " & pIdArt & _
						 " and mmtoprecio = (select max(mmtoprecio) from tblarticulosprecios where idarticulo = "&pIdArt&" and iddivision = " & getDivisionAlmacen(idAlmacen) & ")" & _
						 " and iddivision = " & getDivisionAlmacen(idAlmacen)
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
				vlupesos = 0
				vludolares = 0
				if not rs2.eof then
					vlupesos = cdbl(rs2("VLUPESOS"))
					vludolares = cdbl(rs2("VLUDOLARES")) 
				end if
				
				strSQL = "select * from tblarticulosdatos where idarticulo = " & pIdArt & _
						 " and momento = (select max(momento) from tblarticulosdatos where idarticulo = "&pIdArt&" and idalmacen = " & idAlmacen & ")" & _
						 " and idalmacen = " & idAlmacen
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)						 
				
				stockOri = cdbl(rs2("EXISTENCIA"))
				stockSobranteOri = cdbl(rs2("SOBRANTE"))
			
				stockNuevo = stockOri * pfc
				stockSobranteNuevo = stockSobranteOri * pfc
				
				valorpesos = vlupesos * ( stockOri / stockNuevo )
				valordolares = vludolares * ( stockOri / stockNuevo )
				
				strSQL = "insert into tblvalesdetalle (IDVALE,IDARTICULO,CANTIDAD,EXISTENCIA,SOBRANTE,VLUPESOS,VLUDOLARES)"
				strSQL = strSQL & " values("&idAjuste&","&pIdArt&","&(stockOri+stockSobranteOri)*-1&","&stockOri*-1&","&stockSobranteOri*-1&","&vlupesos&","&vludolares&")"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "EXEC", strSQL)	
				
				strSQL = "insert into tblvalesdetalle (IDVALE,IDARTICULO,CANTIDAD,EXISTENCIA,SOBRANTE,VLUPESOS,VLUDOLARES)"
				strSQL = strSQL & " values("&idAjuste&","&pIdNewArt&","&stockSobranteNuevo+stockNuevo&","&stockNuevo&","&stockSobranteNuevo&","&valorpesos&","&valordolares&")"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "EXEC", strSQL)	
				
				'Creo el precio del nuevo articulo
				if (valordolares <> 0) then
					tipoCambio = ROUND(valorpesos/valordolares,3)
				else
					tipoCambio = "0"
				end if
				
				strSQL = "insert into tblarticulosprecios (MMTOPRECIO,IDDIVISION,IDARTICULO,VLUPESOS,VLUDOLARES,TIPOCAMBIO)"
				strSQL = strSQL & " values("&session("MmtoSistema")&","&getDivisionAlmacen(idAlmacen)&","&pIdNewArt&","&valorpesos&","&valordolares&","&tipoCambio&")"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "EXEC", strSQL)	
				
				'Borrando el stock del articulo original
				strSQL = "update tblarticulosdatos set existencia = 0, sobrante = 0 where idarticulo = " & pIdArt & " and idalmacen = " & idAlmacen
				Call executeQueryDB(DBSITE_SQL_INTRA, rs2, "EXEC", strSQL)	
				
			rs.MoveNext
		wend
End Function
'--------------------------------------------------------------------------------------------------
Function crearArticulo(pIdArticulo)
		'se da de baja el articulo viejo
		strSQL = "update tblarticulos set estado = " & ESTADO_BAJA & " where idarticulo =" & pIdArticulo
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
		
		'creo el nuevo registro haciendo una copia del viejo
		strSQL =          "INSERT "
		strSQL = strSQL & "INTO   tblarticulos "
		strSQL = strSQL & "       ( "
		strSQL = strSQL & "              IDCATEGORIA           , "
		strSQL = strSQL & "              DSARTICULO            , "
		strSQL = strSQL & "              IDUNIDAD              , "
		strSQL = strSQL & "              ESTADO                , "
		strSQL = strSQL & "              BIENUSO               , "
		strSQL = strSQL & "              CDCUENTA              , "
		strSQL = strSQL & "              CDUSUARIO             , "
		strSQL = strSQL & "              MOMENTO               , "
		strSQL = strSQL & "              CDCUENTAGASTOS        , "
		strSQL = strSQL & "              CDCUENTASAF           , "
		strSQL = strSQL & "              CCOSTOS               , "
		strSQL = strSQL & "              MMTOULTIMACOMPRA      , "
		strSQL = strSQL & "              VLUPESOSULTIMACOMPRA  , "
		strSQL = strSQL & "              VLUDOLARESULTIMACOMPRA, "
		strSQL = strSQL & "              IDPIC "
		strSQL = strSQL & "       ) "
		strSQL = strSQL & "       (SELECT IDCATEGORIA           , "
		strSQL = strSQL & "               DSARTICULO            , "
		strSQL = strSQL & "               IDUNIDAD              , "
		strSQL = strSQL & "               ESTADO                , "
		strSQL = strSQL & "               BIENUSO               , "
		strSQL = strSQL & "               CDCUENTA              , "
		strSQL = strSQL & "               CDUSUARIO             , "
		strSQL = strSQL & "               MOMENTO               , "
		strSQL = strSQL & "               CDCUENTAGASTOS        , "
		strSQL = strSQL & "               CDCUENTASAF           , "
		strSQL = strSQL & "               CCOSTOS               , "
		strSQL = strSQL & "               MMTOULTIMACOMPRA      , "
		strSQL = strSQL & "               VLUPESOSULTIMACOMPRA  , "
		strSQL = strSQL & "               VLUDOLARESULTIMACOMPRA, "
		strSQL = strSQL & "               IDPIC "
		strSQL = strSQL & "       FROM    tblarticulos "
		strSQL = strSQL & "       WHERE   idarticulo = " & pIdArticulo
		strSQL = strSQL & "       )"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
		
		'obtengo el id del registro nuevo
		strSQL = "select max(idarticulo) AS lastId from tblarticulos"
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		idNewArticulo = 0
		if not rs.eof then idNewArticulo = rs("lastId")
			
		
		'actualizo los valores del nuevo registro
		GP_ConfigurarMomentos
		
		strSQL = "update tblarticulos set "
		strSQL = strSQL & " estado 				   = "&ESTADO_ACTIVO&","
		strSQL = strSQL & " idunidad			   = "&idunidades&","
		strSQL = strSQL & " momento				   = "&session("MmtoSistema")&","
		strSQL = strSQL & " VLUPESOSULTIMACOMPRA   = 0,"
		strSQL = strSQL & " VLUDOLARESULTIMACOMPRA = 0"  
		strSQL = strSQL & " where idarticulo = " & idNewArticulo
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
		
		crearArticulo = idNewArticulo
		
End Function

'**************************************************************************************************
'**************************************************************************************************
'**************************************************************************************************
'*                                    INICIO DE PAGINA                                            *
'**************************************************************************************************
'**************************************************************************************************
'**************************************************************************************************

	Dim accion, idOrigen, idDestino, strSQL, rs, conn,idArticulo,dsArticulo,IdNewArticulo,flagMostrarId
	dim myArticulo,myArticuloId ,myCategoria,myIdUnidadesO,myAbrevO,myUnidadesO,myIUnidadesO,myIdUnidades,myAbrev,myFc
	dim myPopUp
	accion = GF_PARAMETROS7("accion","",6)
	idOrigen = GF_PARAMETROS7("idOrigen",0,6)
	idDestino = GF_PARAMETROS7("idDestino",0,6)
	myPopUp = GF_PARAMETROS7("pPopUp",0,6)
	
	if (accion = ACCION_CALCULAR) then
		strSQL = "select * from tblunidadesconversion where IDUNIDADORIG = " & idOrigen & " and IDUNIDADDEST = " & idDestino
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
		
		if (not rs.EoF) then response.write rs("FACTOR")
			
		response.end
	end if
	flagMostrarId = false
	
	idArticulo 	= GF_PARAMETROS7("idArticulo",0,6)
	idunidades 	= GF_PARAMETROS7("idunidades",0,6)
	idunidadesO = GF_PARAMETROS7("idunidadesO",0,6)
	fc 			= GF_PARAMETROS7("fc",10,6)
	if (idArticulo = 0) then setError("0023")
	if (idunidades = 0) then setError("0110")
	if (fc = 0) then setError("0111")
	if (idunidades = idunidadesO) then setError("0112")


			myArticulo 		= GF_PARAMETROS7("articulo","",6)
			myArticuloId 	= GF_PARAMETROS7("idarticulo",0,6)
			myCategoria 	= GF_PARAMETROS7("icategoria","",6)
			myIdUnidadesO 	= GF_PARAMETROS7("idunidadesO",0,6)
			myAbrevO 		= GF_PARAMETROS7("abrevO","",6)
			myUnidadesO 	= GF_PARAMETROS7("iunidadesO","",6)
			myIUnidadesO 	= GF_PARAMETROS7("iunidadesO","",6)
			myIdUnidades 	= GF_PARAMETROS7("idunidades",0,6)
			myAbrev 		= GF_PARAMETROS7("abrevD","",6)
			myFc 			= GF_PARAMETROS7("fc","",6)
	
	if (accion = ACCION_GRABAR) then
		if (not hayError()) then
			IdNewArticulo = crearArticulo(idArticulo)
			call cambiarConversion(idunidadesO,idunidades,fc)
			Call crearStockNuevoArticulo(idArticulo,idNewArticulo,fc)
			Call crearValeAjuste(idArticulo,IdNewArticulo,fc)
			flagMostrarId = true
			
			myArticulo 		= ""
			myArticuloId 	= 0
			myCategoria 	= ""
			myIdUnidadesO 	= 0
			myAbrevO 		= ""
			myUnidadesO 	= ""
			myIUnidadesO 	= ""
			myIdUnidades 	= 0
			myAbrev 		= ""
			myFc 			= ""
		else

		end if
	end if
%>

<html>
<head>
	<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">	
	<script type="text/javascript" src="scripts/controles.js"></script>
    <script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	<script type="text/javascript" src="scripts/jQueryAutocomplete.js"></script>
    <script type="text/javascript" src="scripts/botoneraPopUp.js"></script>
	
	<script type="text/javascript">
		ch = new channel();
		var factorAnterior;
		$(function() {
			$( "#articulo" ).autocomplete({
				minLength: 2,
				source: "comprasStreamElementos.asp?tipo=JQArticulos",
				focus: function( event, ui ) {
					$( "#articulo" ).val(ui.item.dsarticulo);
					return false;
				},
				select: function( event, ui ) {
					$( "#articulo"    ).val (ui.item.dsarticulo);
					$( "#idarticulo"  ).val (ui.item.idarticulo );
					$( "#unidadesO"   ).html(ui.item.dsunidad);
					$( "#iunidadesO"  ).val(ui.item.dsunidad);
					$( "#idunidadesO" ).val (ui.item.idunidad);
					$( "#abrevO"	  ).val (ui.item.abreviatura);
					$( "#unidades"    ).val (ui.item.dsunidad);
					$( "#idunidades"  ).val (ui.item.idunidad);
					$( "#abrevD"	  ).val (ui.item.abreviatura);
					$( "#categoria"	  ).html(ui.item.dscategoria);
					$( "#icategoria"  ).val(ui.item.dscategoria);
					limpiarFC();
					return false;
				},
				change: function( event, ui ) {
					document.getElementById("results").innerHTML = "";
					if (!ui.item)
					{
						$( "#articulo" ).val("");
						$( "#idarticulo"  ).val ("");
						$( "#unidadesO"   ).html("");
						$( "#iunidadesO"  ).html("");
						$( "#idunidadesO" ).val ("");
						$( "#abrevO"	  ).val ("");
						$( "#categoria"	  ).html("");
						$( "#icategoria"  ).val("");
						//document.getElementById("results").innerHTML = "";
					}
					else{
						//realizarConsulta(ui.item.idarticulo);
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.idarticulo + " - <font style='font-size:10;'>" + item.dsarticulo + "</font></a>" )
					.appendTo( ul );
			};
		});
	
	function getFactorConversion()
	{
		ch.bind("almacenCambioUnidad.asp?accion=<%=ACCION_CALCULAR%>&idOrigen="+ $("#idunidadesO").val() + "&idDestino="+ $("#idunidades").val(),"callback_fc()");
		ch.send();
	}
	
	function callback_fc()
	{
		$("#fc").val(ch.response());
		changeFC();
	}
	
	function limpiarFC()
	{
		$("#fc").val("");
		$("#lfc").html("");
	}
	
	function changeFC()
	{
		//if (factorAnterior!= $("#fc").val()){
		if ($("#abrevO").val() != "" && $("#fc").val() != 0 && $("#fc").val() != "" && $("#abrevD").val() != ""){
			$("#lfc").html("Resultado: 1 "+ $("#abrevO").val() +" = "+$("#fc").val()+" "+$("#abrevD").val());	
			realizarConsulta($("#idarticulo").val(), $("#fc").val());
			factorAnterior = $("#fc").val();
			}
		else{
			$("#lfc").html("");
			}
		//	}
	}
	function realizarConsulta(idArticulo,factorConversion){
		if (factorConversion!=undefined){
		document.getElementById("imgLoading").style.position = "relative";
		document.getElementById("imgLoading").style.visibility  = "visible";
		document.getElementById("lblLoading").style.position = "relative";
		document.getElementById("lblLoading").style.visibility  = "visible";
		ch.bind("almacenCambioUnidadAjax.asp?idArticulo=" + idArticulo + "&factorConversion=" + factorConversion, "realizarConsultaCallback()");
		ch.send();			
		}
	}
	function realizarConsultaCallback(){
		document.getElementById("imgLoading").style.position = "absolute";
		document.getElementById("imgLoading").style.visibility  = "hidden";
		document.getElementById("lblLoading").style.position = "absolute";
		document.getElementById("lblLoading").style.visibility  = "hidden";
		document.getElementById("results").innerHTML = ch.response(); 
	}		
	function selectUnidades()
			{
				var comboUnidades = document.getElementById("unidades").value;
				comboUnidades = comboUnidades.split("|");
				$( "#idunidades" ).val(comboUnidades[0]);
				$( "#abrevD"     ).val(comboUnidades[1]);
				
				limpiarFC();
				document.getElementById("results").innerHTML = ""; 
				getFactorConversion();
			}
	
	function submitir(pAccion)
	{
		document.getElementById("accion").value=pAccion;
		document.getElementById("form1").submit();
	}
	
	function bodyOnLoad()
	{
		var tb = new Toolbar('toolbar', 6, "images/almacenes/");
		//tb.addButton("Home-16x16.png", "Home", "irA('almacenIndex.asp')");		
		<%if (flagMostrarId) then%>
			tb.addButton("Previous-16x16.png", "Volver", "irA('almacenCambioUnidad.asp')");
		<%else%>
			tb.addButton("accept-16x16.png", "Controlar", "submitir('<%=ACCION_CONTROLAR%>');");	
			tb.addButton("save-16x16.png", "Guardar", "submitir('<%=ACCION_GRABAR%>')");			
		<%end if%>
		<%if (myPopUp=0) then%>
			tb.addButton("Setting_folder-16x16.png", "Ajustes", "irA('almacenAjustes.asp')");										
		<%end if%>	
		changeFC();
		tb.draw();	
	}
	function irA(pLink) {
		document.location.href = pLink;
	}
	</script>
</head>
<body onload="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Sistema de Almacences - Cambio de Unidad") %>	
	
	<div id="toolbar"></div>
	<br>
	<%if (flagMostrarId) then%>
    	<table height='120px' align="center" class="reg_header" cellpadding=5>
    		<tr valign='middle'>
    			<td><img src='images/almacenes/items-32x32.png'></td>
    			<td><p>El nuevo Id del articulo es: <%=IdNewArticulo%></p></td>
    		</tr>
    	</table>
    <%
		response.end
	end if%>
    
	<div class="ui-state-highlight ui-corner-all" style="text-align:center; width:520px; height:25px; margin:0 auto 0 auto;">
		<table>
			<tr>
				<td><span class="ui-icon ui-icon-alert" style="float: left; margin-left: 10px; margin-right: 10px;"></span></td>
				<td>Este movimiento generará un nuevo articulo, bloqueando el actual para su uso.</td>
			</tr>
		</table>
		
		
	</div>
	<br>
	<table align="center" width="550px">	<tr><td><%=showErrors()%></td></tr>	</table>
	<br>
	<form method="GET" id="form1">
		<table class="ui-widget-content ui-corner-all " align="center" width="60%">
			<tr>
				<th colspan=2 class="ui-widget-header ui-corner-top">
					Cambio Unidad Articulo
				</th>
			</tr>
			<tr>
				<td width="150px" align="right" class="reg_header_navdos">
					Articulo
				</td>
				<td>
					<span class="ui-widget">
						<input id="articulo" name="articulo"  style="width:350px" value="<%=myArticulo%>">
						<input type="hidden" name="idarticulo" id="idarticulo" value="<%=myArticuloId%>">
					</span>
				</td>
			</tr>
			<tr>
				<td width="150px" align="right" class="reg_header_navdos" height="20px">
					Categoria
				</td>
				<td>
					<label id="categoria" name="categoria"><%=myCategoria%></label>
					<input type="hidden" name="icategoria" id="icategoria" value="<%=myCategoria%>">
				</td>
			</tr>
			<tr>
				<td width="150px" align="right" class="reg_header_navdos" height="20px">
					Unidad Origen
					<input type="hidden" name="idunidadesO" id="idunidadesO" value="<%=myIdunidadesO%>">
					<input type="hidden" name="abrevO" id="abrevO" value="<%=myAbrevO%>">
				</td>
				<td>
					<label id="unidadesO" name="unidadesO"><%=myUnidadesO%></label>
					<input type="hidden" name="iunidadesO" id="iunidadesO" value="<%=myIUnidadesO%>">
				</td>
			</tr>
			<tr>
				<td width="150px" align="right" class="reg_header_navdos">
					Unidad Destino
					<input type="hidden" name="idunidades" id="idunidades" value="<%=myIdUnidades%>">
					<input type="hidden" name="abrevD" id="abrevD" value="<%=myAbrev%>">
				</td>
				<td>
                    <% 	strSQL = "select * from tblunidades where estado = "  & ESTADO_ACTIVO & " order by dsunidad"
						Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
					%>
                    <select id="unidades" name="unidades" onChange="selectUnidades()">
                    	<option value="0">- Seleccione -</option>
                        <% while not rs.EoF 
							mySelected = ""
							if myIdUnidades = rs("idunidad") then mySelected = "Selected"%>
								<option value="<%=rs("idunidad")%>|<%=rs("ABREVIATURA")%>" <%=mySelected%>><%=rs("dsunidad")%></option>
						<%
    							rs.MoveNext
							wend
						%>
                    </select>
				</td>
			</tr>
			<tr>
				<td width="150px" align="right" class="reg_header_navdos">
					Factor Conversion
				</td>
				<td>
					<span class="ui-widget">
						<input size="10" maxlength="7" id="fc" name="fc" style="width:200px" onkeypress="return controlIngreso(this,event,'N')" onBlur="changeFC();" value="<%=myFc%>">
					</span>
				</td>
			</tr>
			<tr>
				<td colspan="2" class="reg_header" align="center">
					<label id="lfc" name="lfc"></label>&nbsp;
				</td>
			</tr>
		</table>
		<br>
		<table align="center" width="90%" border="0">
			<tr>
				<td align="center">
					<img style="position:absolute;visibility:hidden;" id="imgLoading" src="images/Loading4.gif">
					<div style="position:absolute;visibility:hidden;" id="lblLoading"><b><br>Aguarde por favor...</b></div>
					
				</td>
			</tr>
		</table>
		<div id="results"></div>
		<!--<div id="results"></div>-->
		<input type="hidden" name="accion" id="accion" value="<%=accion%>">
		<input type="hidden" name="pPopUp" id="pPopUp" value="<%=myPopUp%>">
	</form>
	
</body>
</html>