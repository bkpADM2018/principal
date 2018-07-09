<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<%
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
Dim pIdAlmacen, mySelect, rsAlmacenes, titleAux, flagUno,myRs,puedeHacerPics
titleAux = "Almacen"
pIdAlmacen = GF_Parametros7("idAlmacen", 0, 6)
set rsAlmacenes = obtenerListaAlmacenesUsuario()
if rsAlmacenes.Eof Then response.redirect "comprasAccesoDenegado.asp"
if rsAlmacenes.recordCount = 1 then
	if not rsAlmacenes.eof then 
		pIdAlmacen = rsAlmacenes("IDALMACEN")
		titleAux = rsAlmacenes("DSALMACEN")
		flagUno = true
	end if
end if	

if pIdAlmacen = 0 then
	pIdAlmacen = rsAlmacenes("IDALMACEN")
end if	
strSQL = "select * from tblusuariopermisos where cdusuario = '"&session("usuario")&"' "
strSQL = strSQL & " and iddivision = " & getDivisionAlmacen(pIdAlmacen)
strSQL = strSQL & " and idrecurso = " & RES_CD 
strSQL = strSQL & " and permiso in ("&SEC_U&","&SEC_A&")"
Call executeQueryDB(DBSITE_SQL_INTRA, myRs, "OPEN", strSQL)
puedeHacerPics = false
if (not Myrs.EoF) then puedeHacerPics = true

%>
<html>
<head>
<link rel="icon" type="image/gif" href="images/kogge256.gif">
<link rel="icon" type="images/3DCLOS~1.ICO" href="http://example.com/image.ico">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/iwin.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>
<script type="text/javascript">
	var popUpPic;

	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 8);
		tb.addButton("almacenes/Exit-16x16.png"			, "Salidas"			  , "irA('almacenValesTitulo.asp?cdVale=VMS&TC=2')");
		tb.addButton("almacenes/Loan-16x16.png"			, "Prestamos"		  , "irA('almacenValesTitulo.asp?cdVale=VMP&TC=2')");
		tb.addButton("almacenes/Entry-16x16.png"		, "Entradas"		  , "irA('almacenValesTitulo.asp?cdVale=VME&TC=2')");
		tb.addButton("almacenes/Transfer-16x16.png"		, "Transferencias"	  , "irA('almacenValesTitulo.asp?cdVale=VMT&TC=2')");		
		tb.addButton("almacenes/items-16x16.png"		, "Articulos"		  , "irA('almacenAdministrarArticulosAlmacen.asp')");
		tb.addButton("almacenes/REM_folder-16x16.png"	, "Remitos"			  , "irA('almacenAdministrarREM.asp')");
		tb.addButton("almacenes/Getting_Vale-16x16.png" , "Consulta de Vales" , "irA('almacenAdministrarVales.asp')");		
		tb.addButton("almacenes/Refresh-16x16.png"		, "Refresh"			  , "refreshFrames1();");		
		tb.draw();
		loadNovedades();
	}
	function loadPopUpArtSalida(idArticulo, idAlmacen, typeOfView, cdVale) {
		var myPage, w, h;
		w=640;
		h=230;
		if (typeOfView == 'F'){
			h = 350;
			w = 750;
			myPage = 'almacenVales.asp?TC=1&cdVale=' + cdVale + '&pmReferencia=' + idArticulo + '&idAlmacen=' + idAlmacen;
		}
		else{
			myPage = 'almacenArtOut.asp?idArticulo=' + idArticulo + '&idAlmacen=' + idAlmacen + '&typeOfView=' + typeOfView;
		}
		var puw = new PopUpWindow('popupArt',myPage, w, h,cdVale + ' - Entregas');		
		puw.onHideEnd = "refreshFrames1()";		
	}
	function loadPopUpAJP(pIdPM) {
		var myPage, w, h;
		h = 350;
		w = 750;
		myPage = "almacenValesAJP.asp?TC=1&pmReferencia=" + pIdPM + "&cdVAle=AJP";
		var puw = new PopUpWindow('popupAJP',myPage, w, h, 'Pedido de Materiales - Ajuste');		
		puw.onHideEnd = "refreshFrames1()";						
	}
	function loadPopUpArtRecepcion(idArticulo, idAlmacen, typeOfView, cdVale) {
		var myPage, w, h;
		w=640;
		h=230;
		if (typeOfView == 'F'){
			h = 350;
			w = 750;
			myPage = 'almacenVales.asp?TC=1&cdVale=' + cdVale + '&pmReferencia=' + idArticulo + '&idAlmacen=' + idAlmacen;
		}
		else{
			myPage = 'almacenArtRecepcion.asp?idArticulo=' + idArticulo + '&idAlmacen=' + idAlmacen + '&typeOfView=' + typeOfView;
		}
		var puw = new PopUpWindow('popupArt',myPage, w, h,cdVale + ' - Recepcion');		
		puw.onHideEnd = "refreshFrames2()";		
	}
	
	function loadPopUpArtTransferencia(idArticulo, idAlmacen, typeOfView, cdVale, idAlmacenDest) {
		var myPage, w, h;
		w=640;
		h=230;
		if (typeOfView == 'F'){
			h = 350;
			w = 750;
			myPage = 'almacenVales.asp?TC=1&cdVale=' + cdVale + '&pmReferencia=' + idArticulo + '&idAlmacen=' + idAlmacen + '&idAlmacenDest=' + idAlmacenDest;
		}
		else{
			myPage = 'almacenArtRec.asp?idArticulo=' + idArticulo + '&idAlmacen=' + idAlmacen + '&typeOfView=' + typeOfView;
		}
		var puw = new PopUpWindow('popupArt',myPage, w, h,cdVale + ' - Recepcion');		
		puw.onHideEnd = "refreshFrames2()";		
	}
	function loadPopUpAJT(pIdPM) {
		var myPage, w, h;
		h = 350;
		w = 750;
		myPage = "almacenValesAJT.asp?TC=1&pmReferencia=" + pIdPM + "&cdVAle=AJT";
		var puw = new PopUpWindow('popupAJT',myPage, w, h, 'Vale - Ajuste de Transferencia');		
		puw.onHideEnd = "refreshFrames1()";						
	}
	function loadPopUpArtEntrada(idArticulo, idAlmacen, typeOfView, cdVale) {
		var myPage, w, h;
		w=640;
		h=230;
		if (typeOfView == 'F'){
			h = 350;
			w = 750;
			myPage = 'almacenVales.asp?TC=1&cdVale=' + cdVale + '&pmReferencia=' + idArticulo + '&idAlmacen=' + idAlmacen;
		}
		else{
			myPage = 'almacenArtIn.asp?idArticulo=' + idArticulo + '&idAlmacen=' + idAlmacen + '&typeOfView=' + typeOfView;
		}
		var puw = new PopUpWindow('popupArt',myPage, w, h,cdVale + ' - Devolucion');				
		puw.onHideEnd = "refreshFrames1()";				
	}
	function loadPopUpAJU(pIdVale) {
		var myPage, w, h;
		h = 350;
		w = 750;
		myPage = "almacenValesAJU.asp?TC=1&pmReferencia=" + pIdVale + "&cdVAle=AJU";
		var puw = new PopUpWindow('popupArt',myPage, w, h, 'Vale - Ajuste');		
		puw.onHideEnd = "refreshFrames1()";						
	}	
	function refreshFrames1() {
		window.frames["IF1"].document.frmSel.submit();
		window.frames["IF2"].document.frmSel.submit();
		window.frames["IF3"].document.frmSel.submit();
	}	
	function refreshFrames2() {
		window.frames["IF3"].document.frmSel.submit();
	}	
	function irA(pPage){
		if (pPage.indexOf('?') == -1) { 
			pPage += '?';
		} else {
			pPage += '&';
		}
		<% 
		if flagUno then %>
			pPage += "idAlmacen=<%=pIdAlmacen%>";
		<% else %>	
			var ls = document.getElementById("idAlmacenT");
			pPage += "idAlmacen=" + ls.options[ls.selectedIndex].value;		
		<% end if %>	
		document.location.href = pPage; 		
	}
	
	function submitPage(){
		var ls = document.getElementById("idAlmacenT");
		document.getElementById("idAlmacen").value = ls.options[ls.selectedIndex].value;
		document.frmSel.submit();
	}	
	function submitPagePre(pIFNumber, pImg){
		var myIF = 'IF' + pIFNumber;
		var myControl = 'verSegun' + pIFNumber;
		window.frames[myIF].document.src = "";
		document.getElementById(pImg).src = 'images/almacenes/loading_small_orange.gif';
		if (window.frames[myIF].document.getElementById(myControl).value == 'F'){
			window.frames[myIF].document.getElementById(myControl).value = 'A';
		}else{
			window.frames[myIF].document.getElementById(myControl).value = 'F';
		}	
		window.frames[myIF].document.frmSel.submit();
	}	
	function iFrameOnLoad(pIFNumber, pImg){
		var myIF = 'IF' + pIFNumber;
		var myControl = 'verSegun' + pIFNumber;
		if (window.frames[myIF].document.getElementById(myControl).value != 'F'){
			document.getElementById(pImg).src = 'images/almacenes/PM-16x16.png';
			document.getElementById(pImg).title = 'Ver por Articulos';
		}else{
			document.getElementById(pImg).src = 'images/almacenes/items-16x16.png';
			document.getElementById(pImg).title = 'Ver por Formularios';
		}
	}

	function picStockFaltante(){

		
		var $currentIFrame = $('#IF4');
		var idArticulos = "";
		var myCheck;
		var listaArticulos = "";
		var listaIDArticulos = "";
		var listaCantidades = "";
		var i = 1;
		

		totalArticulos = $currentIFrame.contents().find("#totalitems").val();
		
		$.each($currentIFrame.contents().find("input:checked"),function (key,value) {
				if ($(value).attr("id") != "todos"){
					
					auxValue = String($(value).val()).split(";")

					
					listaArticulos += "ARTID_"+(i-1)+"="+auxValue[0]+"&";

					listaCantidades += "CAN_"+(i-1)+"="+auxValue[1]+"&";

					listaIDArticulos += auxValue[0]+",";
					i++;
				}
		});
		if (i!=1){
			<% if (puedeHacerPics) then	%>
				var puw = new winPopUp('popupAlmacenes','compraspic.asp?isInPopUp=1&tipocambio=<%=getTipoCambio(MONEDA_DOLAR, "")%>&'+ listaArticulos+listaCantidades+"nroLinea="+i,'900','600','PIC', "loadNovedades()");
			<% else 
				strSQL = "select * from TBLMAILSALERTASALMACENES where idalmacen = " & pIdAlmacen
				Call executeQueryDB(DBSITE_SQL_INTRA, rs4, "OPEN", strSQL)
				auxEmail = ""
				while not rs4.EoF
					if (trim(cstr(rs4("email"))) <> "") then auxEmail = auxEmail & rs4("email") & ";"
					rs4.MoveNext
				wend
				%>
				var puw = new winPopUp('popupAlmacenes','almacenAlertasEmail.asp?idalmacen=<%=pIdAlmacen%>&articulos='+listaIDArticulos,'450','250','Envio Email', "");
			<% end if %>
		} else {
			alert("Debe seleccionar al menos 1 Articulo");
		}
	}
	
	function loadNovedades() {
		//Imagen loading
		document.getElementById("IF4").contentWindow.document.body.innerHTML = "<table width='100%'><tr><th><img src='images/loading_bar_green.gif'></th><tr></table>";
		//carga de la pagina por ajax
		$.ajax({
		  url: "almacenAlertas.asp?idalmacen=<%=pIdAlmacen%>",
		  success: function(data){
		    document.getElementById("IF4").contentWindow.document.body.innerHTML = data;
		  }
		});
	}

	function configAlertas()
	{
		popUpPic = new winPopUp('popupAlmacenes','almacenPropAlertasAlmacen.asp?idAlmacen=<%=pIdAlmacen%>','650','520','Alertas Almacen', "");
	}

	function cerrarPopUpPics()
	{
		$("#popupAlmacenes").dialog("close");
	}

</script>
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.divOculto {
	display: none;
}
</style>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
</head>
<body onLoad="bodyOnLoad()">
<%

call GF_TITULO2("kogge64.gif",titleAux & " - Tablero de Control") 
%>
<div id="toolbar"></div>

<form id="frmSel" name="frmSel" method="post">
<table class="reg_Header2" align="center" width="100%"  border="0">
		<%
		if flagUno then
		%>
				<tr><td colspan="2">&nbsp</td></tr>
		<%
		else
		%>
			<tr>
				<td>
					<font class="big2"><%=GF_TRADUCIR("Almacén de trabajo:")%></font>
						<select onchange="submitPage();" name="idAlmacenT" id="idAlmacenT">
							<%

							while not rsAlmacenes.eof
								if rsAlmacenes("IDALMACEN") = pIdAlmacen then
									mySelected = "SELECTED"
								else
									if pIdAlmacen = 0 then
										pIdAlmacen = rsAlmacenes("IDALMACEN")
									end if
									mySelected = ""
								end if
								%>
								<option title="<%=rsAlmacenes("DSALMACEN")%>" VALUE="<%=rsAlmacenes("IDALMACEN")%>" <%=mySelected%>><%=rsAlmacenes("DSALMACEN")%></option>
								<%
								rsAlmacenes.movenext
							wend
							%>
						</select>		
				</td>
			<tr>	
		<%
		end if
		%>	
		<input type="hidden" name="idAlmacen2" id="idAlmacen2" <%=pIdAlmacen2%>>		
	<tr>
		<td width="50%">
			<table id="TBL1" width="100%" cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width="8px;"><img src="images/marco_SD_FILL.gif"></td>
					<td colspan="2" background="images/marco_SM_FILL.gif"></td>
					<td width="8px;"><img src="images/marco_SI_FILL.gif"></td>
				</tr>
				<tr class="reg_header_nav">
					<td><img src="images/marco_r2_c1.gif"></td>
					<td><%=GF_TRADUCIR("Pendientes de entrega")%></td>					
					<td onclick="submitPagePre('1','ART_OUT_IMG');" align="right"><img id="ART_OUT_IMG" src="images/almacenes/items-16x16.png" style="cursor:pointer;" title="Ver por Articulos" /></td>
					<td><img src="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_MD_FILL.gif"></td>
					<td colspan="2" background="images/marco_MM_FILL.gif"></td>
					<td><img src="images/marco_MI_FILL.gif"></td>
				</tr>				
				<tr>
					<td background="images/marco_r2_c1.gif"></td>
					<td colspan="2">	
						<iframe id="IF1" name="IF1" frameborder=0 src="almacenIFartSalida.asp?idAlmacen=<%=pIdAlmacen%>&typeOfView=<%=verSegun1%>" valign="top" align="center" height="150px" width="100%" border=0 onload="iFrameOnLoad('1','ART_OUT_IMG')"></iframe>
					</td>
					<td background="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_r3_c1.gif"></td>
					<td colspan="2" background="images/marco_r3_c2.gif"></td>
					<td><img src="images/marco_r3_c3.gif"></td>
				</tr>				
			</table>			
		</td>
		<td width="50%">
			<table id="TBL1" width="100%" cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width="8px;"><img src="images/marco_SD_FILL.gif"></td>
					<td colspan="2" background="images/marco_SM_FILL.gif"></td>
					<td width="8px;"><img src="images/marco_SI_FILL.gif"></td>
				</tr>
				<tr class="reg_header_nav">
					<td><img src="images/marco_r2_c1.gif"></td>
					<td><%=GF_TRADUCIR("Pendientes de devolucion")%></td>
					<td onclick="submitPagePre('2','ART_IN_IMG');" align="right"><img id="ART_IN_IMG" src="images/almacenes/items-16x16.png" style="cursor:pointer;" title="Ver por Articulos"></td>
					<td><img src="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_MD_FILL.gif"></td>
					<td colspan="2" background="images/marco_MM_FILL.gif"></td>
					<td><img src="images/marco_MI_FILL.gif"></td>
				</tr>						
				<tr>
					<td background="images/marco_r2_c1.gif"></td>
					<td colspan="2">	
						<iframe name="IF2" id="IF2" frameborder=0 src="almacenIFartEntrada.asp?idAlmacen=<%=pIdAlmacen%>&typeOfView=<%=verSegun2%>" valign="top" align="center" height="150px" width="100%" border=0 onload="iFrameOnLoad('2','ART_IN_IMG')"></iframe>
					</td>
					<td background="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_r3_c1.gif"></td>
					<td colspan="2" background="images/marco_r3_c2.gif"></td>
					<td><img src="images/marco_r3_c3.gif"></td>
				</tr>				
			</table>			
		</td>
	</tr>
	<tr>
		<td width="50%">
			<table id="TBL1" width="100%" cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width="8px;"><img src="images/marco_SD_FILL.gif"></td>
					<td colspan="2" background="images/marco_SM_FILL.gif"></td>
					<td width="8px;"><img src="images/marco_SI_FILL.gif"></td>
				</tr>
				<tr class="reg_header_nav">
					<td><img src="images/marco_r2_c1.gif"></td>
					<td><%=GF_TRADUCIR("Transferencias en curso")%></td>
					<td onclick="submitPagePre('3','ART_TR_IMG');" align="right"><img id="ART_TR_IMG" src="images/almacenes/items-16x16.png" style="cursor:pointer;" title="Ver por Articulos"></td>					
					<td><img src="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_MD_FILL.gif"></td>
					<td colspan="2" background="images/marco_MM_FILL.gif"></td>
					<td><img src="images/marco_MI_FILL.gif"></td>
				</tr>					
				<tr>
					<td background="images/marco_r2_c1.gif"></td>
					<td colspan="2">	
						<iframe name="IF3" id="IF3" frameborder=0 src="almacenIFartTransferencia.asp?idAlmacen=<%=pIdAlmacen%>&typeOfView=<%=verSegun3%>" valign="top" align="center" height="150px" width="100%" border=0 onload="iFrameOnLoad('3','ART_TR_IMG')"></iframe>
					</td>
					<td background="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_r3_c1.gif"></td>
					<td colspan="2" background="images/marco_r3_c2.gif"></td>
					<td><img src="images/marco_r3_c3.gif"></td>
				</tr>				
			</table>			
		</td>
		<td width="50%">
			<table id="TBL1" width="100%" cellpadding=0 cellspacing=0 border=0>
				<tr>
					<td width="8px;"><img src="images/marco_SD_FILL.gif"></td>
					<td COLSPAN="2" background="images/marco_SM_FILL.gif"></td>
					<td width="8px;"><img src="images/marco_SI_FILL.gif"></td>
				</tr>
				<tr class="reg_header_nav">
					<td><img src="images/marco_r2_c1.gif"></td>
					<td><%=GF_TRADUCIR("Novedades")%></td>
					<td align="right"><img src="images/almacenes/campana-16x16.png" onclick="configAlertas()" style="cursor:pointer;" title="Configurar Alertas">&nbsp;<img onclick="picStockFaltante()" id="ART_IN_IMG" src="images/compras/PIC-16x16.png" style="cursor:pointer;" title="Hacer Pic"></td>
					<td><img src="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_MD_FILL.gif"></td>
					<td COLSPAN="2" background="images/marco_MM_FILL.gif"></td>
					<td><img src="images/marco_MI_FILL.gif"></td>
				</tr>				
				<tr>
					<td background="images/marco_r2_c1.gif"></td>
					<td COLSPAN="2">
						<iframe name="IF4" id="IF4" frameborder=0 src="" valign="top" align="center" height="150px" width="100%" border=0 ></iframe>						
					</td>
					<td background="images/marco_r2_c3.gif"></td>
				</tr>
				<tr>
					<td><img src="images/marco_r3_c1.gif"></td>
					<td COLSPAN="2" background="images/marco_r3_c2.gif"></td>
					<td><img src="images/marco_r3_c3.gif"></td>
				</tr>				
			</table>	
		</td>
	</tr>
</table>
<input type="hidden" name="idAlmacen" id="idAlmacen" value="<% =pIdAlmacen %>">
</form>
</body>
</html>