<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosBR.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 07/04/2008                         -->
'<!--               Pagina     : BRExplorador.ASP                   -->
'<!--               Descripcion: Explorador de carpeta con imagenes -->
'<!------------------------------------------------------------------->
dim myContrato, myDoc, myUbicacion, myVista
myContrato = GF_Parametros7("pContrato","",6)
myDoc = GF_Parametros7("pDoc","",6)
myUbicacion = 0
myVista = "I"
'------------------------------------------------------------------------------------------------
function existElement(pElement)
dim i
	for i=0 to ubound(MyBoatAccountOrd)
		if pElement=MyBoatAccountOrd(i) then
			existElement = true
			exit function
		end if
	next
end function
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
%>
<html>
<head>
<meta http-equiv="Cache-Control" content="no-cache, mustrevalidate">
<meta http-equiv="Pragma" content="no-cache">
<Link REL="stylesheet" href="CSS/ActiSAintra-1.css" type="text/css">
<link rel="stylesheet" href="css/tabsC.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="CSS/Toolbar.css" type="text/css">
<script language="javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">	var tabberOptions = {manualStartup:true}; </script>
<script language="javascript" src="Scripts/tabber.js"></script>
<title><%=GF_Traducir("Intranet ActiSA - Documentos Relacionados con el Buque")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="scripts/Toolbar.js"></script>
<script language="JavaScript" src="Scripts/channel.js"></script> 
<script language="JavaScript">
	var pedidos = new Array();
	var ch = new channel();
	var rowPrevious,rowPrevious2,rowPrevious22, rowPrevious3,rowPrevious33, myTableFVL;
	var allFechas = new String();
	var antTab = 0;
	function traerDocumento_callback(pKey)
	{		
		
		document.getElementById("documentosDetalle1").innerHTML = ch.response();
		pedidos[pKey] = ch.response();
		
		tabberAutomatic(tabberOptions);
		pngfix();
		document.getElementById('Tab3').tabber.tabShow(antTab);		
	}
	function traerDocumento(pContrato, pDoc, pView) 
	{
		var flag = false;
		pView=document.getElementById("myVista").value;
		var myKey = pContrato + "-" + pView + "-" + antTab;
		var element, myImg;
			//Buscar si ya esta cargado en el array
			for (element in pedidos){
				if (element==myKey){
					document.getElementById("documentosDetalle1").innerHTML = pedidos[element];
					tabberAutomatic(tabberOptions);
					pngfix();
					flag = true;
				}
			}
			//no esta cargado en el array pedirselo al canal
			if (!flag){	
				var link = "BRCGetDocumentos.asp?pContrato=" + pContrato + "&pDoc=" + pDoc + "&pView=" + pView;
				var param = "traerDocumento_callback('" + myKey + "')";
				ch.bind(link, param);
				ch.send();
			}	
	}		
	function delAntTab() {
		antTab = 0;
		document.getElementById("documentosDetalle1").innerHTML = "<table width='100%'><tr><td align=center><img src='images/loading2.gif'></td></tr></table>";
		}
	function openWin(pPath){
		//alert(pPath);
		window.open(pPath,"","top=1, left=1, scrollbars=yes,status=no,resizable=yes,toolbar=no,location=no,menu=no,width=700,height=700");
	}
	function loadDefault(){
		/*var a = document.getElementById("pmtAno").value;
		if (document.getElementById("TD(" + a + ")") == null){
	    	if (a!='') { alert("No existe documentos para el BA " + a + "!!!"); }
			a = document.getElementById("myFirstAno").value;
			document.getElementById("TD(" + a + ")").onclick();
		}
		else
		{
			document.getElementById("TD(" + a + ")").onclick();
		} */
		traerDocumento(document.getElementById("myContrato").value, document.getElementById("myDoc").value, document.getElementById("myVista").value);
		toolBarGrupos.draw();
		/*
		var a = document.getElementById("myFirstAno").value;
		if (document.getElementById("TD(" + a + ")") == null){
	    	if (a!='') { alert("No existe documentos para el BA " + a + "!!!"); }
			a = document.getElementById("pmtAno").value;
			document.getElementById("TD(" + a + ")").onclick();
		}
		else
		{
			document.getElementById("TD(" + a + ")").onclick();
		} 
		*/
	}
	function fcnResaltar(P_objFila)
	{
		P_objFila.style.background= "#ffeecd";
	}
	function fcnNormal(P_objFila)
	{
		P_objFila.style.background= "#fffaf0";
	}	

	function changeView(){
		var myView;
		if (document.getElementById("myVista").value=="L"){ 
			document.getElementById("myVista").value = "I";
			myView = "I";
		}	
		else
		{
			document.getElementById("myVista").value = "L";
			myView = "L";
		}	
		traerDocumento(document.getElementById("myContrato").value, document.getElementById("myDoc").value, myView) 
	}	
	
		var tabberOptions = {
	  'onClick': function(argsObj) {
	    var t = argsObj.tabber; /* Tabber object */
		var id = t.id; /* ID of the main tabber DIV */
		var i = argsObj.index; /* Which tab was clicked (0 is the first tab) */
		var e = argsObj.event; /* Event object */
		antTab = i;
		}
	}
		/* Barra de herramientas de grupos */
	var toolBarGrupos = new Toolbar("toolBarGrupos",10);
	toolBarGrupos.addSwitcher("search3.gif","Lista","javascript:changeView();","javascript:changeView();");
	toolBarGrupos.addButtonCANCEL("Atras", "javascript:history.back();");
	//toolBarGrupos.addButton(TOOL_REFRESH, "", "");
	
	/* Barra de herramientas de usuarios */
	
	
</script>
</head>
<body onload="tabberAutomatic('');loadDefault();">
<form name="frmMain" action="" method="post">
<%=GF_TITULO2("Contracts64.png","Documentos Relacionados con el Contrato") %>
<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
	<tr>
		<td colspan="3">
			<table border="1"  cellspacing="0" cellpadding="0" width="100%" align="center" rules=rows>

				<tr>
					<td  colspan="6" align="left" class="titu_header11" background="images\fondo_barra.gif">&nbsp;<font class="bigger" color="#FFFFFF"><b><%=GF_TRADUCIR("Nro de Contrato: " & myContrato)%></b></font></td>
				</tr>
					<% 
					call GF_BD_Control_BR (rs, cn, "OPEN", "Select C.NroContrato , C2.dsCompania 'Compania', C.cdOperacion 'Operacion', P.dsProducto 'Producto', C3.dsCliente 'Cliente', C.FechaCierre, C.FechaConfirma  from Contratos C inner join Companias C2 on C.cdCompania=C2.cdCompania inner join Productos P on C.cdProducto=P.cdProducto inner join Clientes C3 on C.cdCliente=C3.cdCliente where NroContrato='" & myContrato & "'")
					if not rs.EOF then
						%>
						<tr>
							<td class="reg_Header_navdos"><%=GF_Traducir("Compañia")%></td>
							<td><%=ucase(trim(rs("Compania")))%></td>
							<td class="reg_Header_navdos"><%=GF_Traducir("Operación")%></td>
							<td><%
								if rs("Operacion") = "C" then 
									Response.Write "COMPRA"
								else
									Response.Write "VENTA"
								end if
							%></td>
							<td class="reg_Header_navdos"><%=GF_Traducir("Producto")%></td>
							<td><%=trim(rs("Producto"))%></td>							
						</tr>
						<tr>
							<td class="reg_Header_navdos"><%=GF_Traducir("Cliente")%></td>
							<td><%=trim(rs("Cliente"))%></td>
							<td class="reg_Header_navdos"><%=GF_Traducir("Fecha de Cierre")%></td>
							<td><%=GF_MDA2DMA(rs("FechaCierre"))%></td>
							<td class="reg_Header_navdos"><%=GF_Traducir("Fecha de Conf.")%></td>
							<td><%=GF_MDA2DMA(rs("FechaConfirma"))%></td>							
						</tr>						
						<%
					end if	
					call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
					%>
				<tr><td colspan="6"><div id="toolBarGrupos"></div></td></tr>				
			</table>	
		</td>	
	</tr>
	<% 
	'CargarFechasVL myBuqueCD
	'CargarLineaBA
	%>
	<tr>
		<td colspan="3">
			<div id="documentosDetalle1"></div>
		</td>
	</tr>
</table>	
</form>
<input type="hidden" size="30" id="myContrato"		value="<%=myContrato%>">
<input type="hidden" size="30" id="myDoc"			value="<%=myDoc%>">
<input type="hidden" size="30" id="myVista"			value="<%=myVista%>">
</body>
</html>