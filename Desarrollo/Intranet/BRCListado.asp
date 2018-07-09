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
'<!--               Pagina     : BRListadoBuques.ASP                -->
'<!--               Descripcion: Listado y busqueda de documentos   -->
'<!------------------------------------------------------------------->
dim MySubFolders0, MySubFolders1, MySubFolders2, MySubFolders3, MySubFolders4
dim cont2, cont, i
dim cn, rs, sql, mySelect
dim myCompania, myOperacion, myAno, MyMes, MyDia, MyTipo, MyFile, MyVisible, MyBA, myHAM, MyLista
myCompania = GF_Parametros7("txtCompania","",6)
myOperacion = GF_Parametros7("txtOperacion","",6)
myAno = GF_Parametros7("txtAno","",6)
MyMes = GF_Parametros7("txtMes","",6)
MyDia = GF_Parametros7("txtDia","",6)
MyTipo = GF_Parametros7("txtTipo","",6)
MyFile = GF_Parametros7("txtFile","",6)
MyBA = GF_Parametros7("txtBA","",6)
MyHAM = GF_Parametros7("chkHAM","",6)
MyVisible = GF_Parametros7("pVisible","",6)
if myVisible = "" then myVisible = "position:relative;visibility:visible;"
dim contractsListL, clientsListL
call CargarListas
'-----------------------------------------------------------------------------------------
sub CargarListas
'Cargar lista de contratos
dim sql, rs, cn
	sql = "Select NroContrato from Contratos order by convert(int,substring(NroContrato,len(NroContrato)-2,4)) desc"
	call GF_BD_Control_BR (rs, cn, "OPEN", sql)
	while not rs.eof	
		contractsListL = contractsListL & ", " & trim(rs("NroContrato"))
		rs.movenext
	wend	
	call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
	contractsListL = replace(contractsListL,"'","")
	'MyLista = replace(MyLista,".","")
	if len(contractsListL)>2 then contractsListL = right(contractsListL,len(contractsListL)-2)
	
	sql = "Select * from Clientes order by dsCliente asc"
	call GF_BD_Control_BR (rs, cn, "OPEN", sql)
	while not rs.eof 
		clientsListL = clientsListL & ", " & trim(rs("dsCliente"))
		rs.movenext
	wend	
	clientsListL = replace(clientsListL,"'","")
	clientsListL = replace(clientsListL,"""","")
	if len(clientsListL)>2 then clientsListL = right(clientsListL,len(clientsListL)-2)
	call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
	
end sub
'Response.Write clientsList
'---------------------------------------------------------------------------------------------
%>
<script language="JavaScript" src="Scripts/channel.js"></script>
<script language="JavaScript" src="Scripts/controles.js"></script>  
<script language="JavaScript" src="Scripts/formato.js"></script>  
<script language="JavaScript" src="Scripts/paginar.js"></script>  
<script language="javascript" src="scripts/pngfix.js"></script>
<script language="javascript" src="scripts/magicSearchObj.js"></script>
<script language="JavaScript">
	var ch = new channel();
	var MScontracts, MSclients, MStest;
	function traerDocumento_callback()
	{		
		document.getElementById("resultados").innerHTML = ch.response();
		pngfix();
	}
	function traerDocumento() 
	{		
		var pContrato, pCompania, pOperacion, pAno, pMes, pDia, pProducto, pCliente, pDoc;
		var myContratoAux = new String();
		pContrato = MScontracts.getSelectedItem();
		pAno = document.getElementById("txtAno").value;
		pMes = document.getElementById("txtMes").value;
		pDia = document.getElementById("txtDia").value;
		pCompania = document.getElementById("txtCompania").value;
		pOperacion = document.getElementById("txtOperacion").value;
		//pProducto = document.getElementById("txtProducto").value;
		pProducto = "";
		pCliente = MSclients.getSelectedItem();
		if (pAno!='' || pMes!='' || pDia!=''){
			if (!controlFecha3P(pAno, pMes, pDia)) { 
				alert("La fecha no es correcta!");
				return 0;			
			}
		}
		pDoc = document.getElementById("txtDocumento").value;
		document.getElementById("resultados").innerHTML = "<table width='100%'><tr><td align=center><img src='images/loading2.gif'></td></tr></table>";				
		var link;
		if (pDoc != ""){
			link = "BRCGetResultado2.asp?pContrato=" + pContrato + "&pCompania=" + pCompania + "&pOperacion=" + pOperacion + "&pAno=" + pAno + "&pMes=" + pMes + "&pDia=" + pDia + "&pDoc=" + pDoc + "&pProducto=" + pProducto + "&pCliente=" + pCliente;
		}
		else{	
			link = "BRCGetResultado.asp?pContrato=" + pContrato + "&pCompania=" + pCompania + "&pOperacion=" + pOperacion + "&pAno=" + pAno + "&pMes=" + pMes + "&pDia=" + pDia + "&pDoc=" + pDoc + "&pProducto=" + pProducto + "&pCliente=" + pCliente;
		}
		var param = "traerDocumento_callback()";
		ch.bind(link, param);
		ch.send();
	}
	function HabilitarFiltros(pForce){
		var filtros;
		filtros = document.getElementById("filtros");
		if (pForce){
			filtros.style.visibility = "hidden";
			filtros.style.position = "absolute";
			document.getElementById("pVisible").value = "Hidden";
			return 0;
		}
		if (filtros.style.visibility == "hidden"){
			filtros.style.visibility = "visible";
			filtros.style.position = "relative";
			document.getElementById("pVisible").value = "position:relative;visibility:visible;";
		}
		else{
			filtros.style.visibility = "hidden";
			filtros.style.position = "absolute";
			document.getElementById("pVisible").value = "position:absolute;visibility:hidden;";
		}
	}
	function openWin(pPath){
		window.open(pPath,"Visualizacion","top=1, left=1, scrollbars=yes,status=no,resizable=yes,toolbar=no,location=no,menu=no,width=500,height=500");
	}
	function fcnResaltar(P_objFila)
	{
		P_objFila.style.background= "#ffeecd";
		P_objFila.style.cursor="hand";
	}
	function fcnNormal(P_objFila)
	{
		P_objFila.style.background= "#fffaf0"
	}
	var o = new Paginacion("myPaging");

	
	function myClear(){
		document.getElementById("txtAno").value = "";
		document.getElementById("txtMes").value = "";
		document.getElementById("txtDia").value = "";
		document.getElementById("txtBuque").value = "";
		document.getElementById("txtDocumento").value = "";
		document.getElementById("txtBA").value = ""; 
		document.getElementById("chkHam").checked = false;
		document.getElementById("chkPro").checked = false;
	}
	function myBuscar(){
		if (MScontracts.getSelectedItem() != ""){
			document.getElementById("txtCompania").value = "";
			document.getElementById("txtAno").value = "";
			document.getElementById("txtMes").value = "";
			document.getElementById("txtDia").value = "";
			document.getElementById("txtOperacion").value = "";
			MSclients.setValue("");
			//document.getElementById("clientsList").value = "";
			//document.getElementById("txtProducto").value = "";
		} 
		traerDocumento();
	}
	function cambiarFoco(pObj, pChr, pObjNext){
		var aux = new String();
		aux = pObj.value;
		if (aux.length==pChr){
			document.getElementById(pObjNext).focus();
		}
	}
	function clearFecha(){
			document.getElementById("txtAno").value = "";
			document.getElementById("txtMes").value = "";
			document.getElementById("txtDia").value = "";
			document.getElementById("txtBA").value = ""; 			
			document.getElementById("chkHam").checked = false;			
	}
	function checkear(element){
			if (document.getElementById(element).checked==true){
				document.getElementById(element).checked = false;
			}
			else{
				document.getElementById(element).checked = true;
			}
			
	}
	function myOnLoad(){
		MScontracts = new MagicSearch(document.getElementById('contractsListL').value,'contractsList',8,4);
		MSclients = new MagicSearch(document.getElementById('clientsListL').value,'clientsList',30,2);
		MStest = new MagicSearch('','testList',30,2,'msOtro.asp');
		MScontracts.getSelectedItem();
	}

</script>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="Expires" CONTENT="-1">
<Link REL="stylesheet" href="CSS/ActiSAintra-1.css" type="text/css">
<link rel="stylesheet" href="CSS/MagicSearch.css" type="text/css">
<title>Intranet ActiSA - Listado de Contratos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body id="myBody" onload="javascript:myOnLoad();traerDocumento();">
<% =GF_TITULO2("Contracts64.png","Listado de Contratos") %>
<form name="frmMain" action="" method="post">
<table border=0 width="100%" align="center" cellpadding=0 cellspacing=0>
	<tr>
		<td colspan="2">
			<table width="100%" border=0 id="filtros" style="<%=myVisible%>" class="reg_header" cellpadding=2 cellspacing=0 rules=groups>
				<tr>
					<td colspan="3" align="left">	<b><%=GF_Traducir("Contrato")			%>	</b></td>
					<td align="left">				<b><%=GF_Traducir("Compañia")			%>	</b></td>
					<td align="left">				<b><%=GF_Traducir("Operación")			%>	</b></td>
					<td align="left">				<b><%=GF_Traducir("Fecha Cierre")		%>	</b></td>
					<td align="left">				<b><%=GF_Traducir("Cliente")			%>	</b></td>
					<td align="left">				<b><%=GF_Traducir("Documento")			%>	</b></td>
					<td align="right">
						<!--<input type="button" name="cmdClear" value="&nbsp;Limpiar&nbsp;" onclick="myClear();">-->
					</td>
				</tr>
				<tr>
					<td colspan="3" align="left"><div id="contractsList"></div></td>
					<td align="left">
						<select name="txtCompania" id="txtCompania">
							<option value=""></option>						
								<%
								sql = "Select * from Companias order by dsCompania asc"
								call GF_BD_Control_BR (rs, cn, "OPEN", sql)
								while not rs.eof	
									if trim(rs("cdCompania")) = myCompania then 
										mySelect = "Selected"
									else
										mySelect = ""
									end if	
									%>
									<option value="<%=trim(rs("cdCompania"))%>" <%=MySelect%>><%=trim(rs("dsCompania"))%></option>
									<%
									rs.movenext
								wend	
								call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
								%>
						</select>						
					</td>
					<td align="left">
						<select name="txtOperacion" id="txtOperacion">
							<option value=""></option>
								<%
									mySelectV = ""
									mySelectC = ""	
									if trim(myOperacion) = "V" then
										mySelectV = "Selected"
									elseif trim(myOperacion) = "C" then
										mySelectC = "Selected"
									end if
								%>
							<option value="C" <%=MySelectV%>>Compra</option>
							<option value="V" <%=MySelectC%>>Venta</option>
						</select>
					</td>
					<td align="left">
						<input type="text" name="txtDia" id="txtDia" size="2" maxlength="2" value="<%=MyDia%>" onkeypress="return controlDatos (this, event, 'N');">
						<input type="text" name="txtMes" id="txtMes" size="2" maxlength="2" value="<%=MyMes%>" onkeypress="return controlDatos (this, event, 'N');">
						<input type="text" name="txtAno" id="txtAno" size="5" maxlength="4" value="<%=MyAno%>" onkeypress="return controlDatos (this, event, 'N');">
					</td>
					<td align="left"><div id="clientsList"></div><div id="testList"></div></td>	
					<td align="left">
						<select name="txtDocumento" id="txtDocumento">
							<option value=""></option>						
								<%
								sql = "Select * from TipoDocumento order by dsTipoDocumento asc"
								call GF_BD_Control_BR (rs, cn, "OPEN", sql)
								while not rs.eof	
									if trim(rs("dsTipoDocumento")) = myTipo then 
										mySelect = "Selected"
									else
										mySelect = ""
									end if	
									%>
									<option value="<%=trim(rs("cdTipoDocumento"))%>" <%=MySelect%>><%=trim(rs("dsTipoDocumento"))%></option>
									<%
									rs.movenext
								wend	
								call GF_BD_Control_BR (rs, cn, "CLOSE", sql)
								%>
						</select>						
					</td>
					<td align="right">
						<input type="button" id="cmdBuscar" name="cmdBuscar" value="Buscar..." onclick="myBuscar();">
					</td>
				</tr>	
			</table>
			<br>
		</td>
	</tr>
	<tr>	
		<td colspan="2" height="300" valign="top" width="100%">
			<div id="resultados"></div>
		</td>
	</tr>
</table>
<input type="hidden" name="pVisible" id="pVisible" value="<%=MyVisible%>">
<input type="hidden" name="contractsListL" id="contractsListL" value="<%=contractsListL%>">
<input type="hidden" name="clientsListL"   id="clientsListL"   value="<%=clientsListL%>">
</form>
</body>
</html>