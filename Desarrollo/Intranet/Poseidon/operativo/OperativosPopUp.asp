<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientossql.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosfechas.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->
<%
'--------------------------------------------------------------------------------------------------------
Dim Conn
Dim g_strPuerto, g_strSector, countResultados,g_cdProducto
Dim g_rsPesadas, g_rsCaladas, g_rsHumedimetro, g_ctaPorte, g_dtContable,g_dsObservaciones
dim g_fltPromedioHumedad, g_fltPromedioPesoHect, g_fltPromedioTemp, g_fltMaxHumedad, g_fltMinPesoHect, g_fltMaxTemp, g_rsResultados

g_strPuerto = GF_Parametros7("Pto","",6)
g_dtContable = GF_Parametros7("fecha","",6)
g_cartaPorte = GF_Parametros7("cartaPorte","",6)
g_cdOperativo = GF_Parametros7("cdOperativo","",6)
g_cdProducto = GF_Parametros7("cdProducto","",6)


Set rsVag = getVagonesByOperativos(g_cdOperativo,g_dtContable,g_cartaPorte,g_strPuerto)
%>
<HTML>
<HEAD>
<meta http-equiv="X-UA-Compatible" content="IE=9">

	<TITLE>Poseidon - Informacion de Cálidad de Vagon </TITLE>
    
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="../../css/main.css" />	
    <link rel="stylesheet" href="../../css/iwin.css" type="text/css">
	<script type="text/javascript" src="../../scripts/channel.js"></script>
	<script type="text/javascript" src="../../scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript">
		var ch= new channel();
		function verCaladaVagon(pNroVagon, pDtContable){
			var pElement = document.getElementById("divCaladaVagon_" + pNroVagon); 
			if (document.getElementById("trCaladaVagon_" + pNroVagon).className == "troculto") {
				document.getElementById("trCaladaVagon_" + pNroVagon).className = "trvisible";
				var iImgAddArt = document.createElement('img');
				iImgAddArt.id = "loading_"  + pNroVagon;
				iImgAddArt.name = "loading_"  + pNroVagon;
				iImgAddArt.src = "../../images/Loading4.gif";
				iImgAddArt.title = "Agregar Articulo";
				iImgAddArt.setAttribute('style', "cursor:pointer;");
				pElement.align = "center";
				pElement.appendChild(iImgAddArt);				
				ch.bind("operativosInformeAjax.asp?Pto=<%=g_strPuerto%>&nroVagon=" +  pNroVagon + "&dtContable=" + pDtContable +"&cdOperativo=<%=g_cdOperativo%>&cartaporte=<%=g_cartaPorte%>" ,"CallBack_verCaladaVagon("+pNroVagon+")");
				ch.send();
				document.getElementById("imgVerCalada_"+pNroVagon).src = "../../images/Menos.gif"
			}
			else{				
				document.getElementById("trCaladaVagon_" + pNroVagon).className = "troculto";
				removeAllChilds(pElement);
				document.getElementById("imgVerCalada_"+pNroVagon).src = "../../images/Mas.gif"
			}
		}		
		
		function removeAllChilds(a){			
			while(a.hasChildNodes()){
				a.removeChild(a.firstChild);
			}	
		}
		
		function CallBack_verCaladaVagon(pNroVagon){
			var padre = document.getElementById("loading_" + pNroVagon).parentNode;
			padre.removeChild(document.getElementById("loading_" + pNroVagon));			
			var respuesta = ch.response();
			document.getElementById("divCaladaVagon_" + pNroVagon).style.display = "";
			document.getElementById("divCaladaVagon_" + pNroVagon).innerHTML = respuesta;
		}
		function lightOn(tr) {
			tr.className = "reg_Header_navdosHL";
		}
		function lightOff(tr) {
			tr.className = "reg_Header_navdos";
		}
		function abrirNotaRecepcion(pCdOperativo, pCdVagon, pDtContable){		
			window.open("../NotaRecepcionPrint.asp?pto=<%=g_strPuerto%>&cartaPorte="+pCdOperativo+"&cdVagon="+pCdVagon+"&dtContable="+pDtContable);
		}
	</script>
</HEAD>

<BODY>

	<!--<h3> <%=GF_Traducir("Vagones del Operativo")%> </h3> -->

	<div class="col66"></div>

	<INPUT type="hidden" id="Pto" name="Pto" value = <%= g_strPuerto %>>
	<INPUT type="hidden" id="vagon" name="vagon" value = <%= g_idVagon%>>
	
    <div class="tableasidecontent">
    
        <div class="col26 reg_header_navdos"> Operativo </div>
        <div class="col26"> <% if not rsVag.eof then response.Write GF_EDIT_CTAPTE(left(rsVag("Operativo"), 12)) else response.Write g_cdOperativo %>  </div>
        
        <div class="col26 reg_header_navdos"> Carta de Porte </div>
        <div class="col26">  <% if not rsVag.eof then response.Write GF_EDIT_CTAPTE(left(rsVag("CartaPorte"), 12)) else response.Write GF_EDIT_CTAPTE(g_cartaPorte)  %> </div>
        
        <div class="col26 reg_header_navdos"> Fecha </div>
        <div class="col26"> <% =GF_FN2DTE(g_dtContable) %>  </div>
        
        <div class="col26 reg_header_navdos"> Producto </div>
        <div class="col26">  <% =getDsProducto(g_cdProducto) %> </div>
        
    </div>   
        

<div class="col66"></div>

	<%if not rsVag.eof then%>
	<table class="datagrid datagridlv1" width="100%" >
        <thead>
            <tr>
                <th width="3%" align="center"> . </th>
                <th width="10%" align="center"> <%=GF_Traducir("Nro.Vagón")%> </th>
                <th width="15%" align="center"> <%=GF_Traducir("Kilos Netos")%> </th>
                <th width="22%" align="center"> <%=GF_Traducir("Hora")%> </th>
                <th width="15%" align="center" colspan="2"> <%=GF_Traducir("Nota Recepcion")%> </th>
            </tr>
        </thead>
        <tbody>
		<%  while not rsVag.eof %>
			<tr>
				<td align="center"><img style="cursor:pointer;" title="Ver información Calada" id="imgVerCalada_<%=rsVag("cdvagon")%>" onClick="verCaladaVagon('<%=rsVag("cdvagon")%>','<%=g_dtContable%>')" src="../../images/Mas.gif"></td>
				<td align="center"><%=rsVag("cdvagon")%></td>
				<td align="right"><%Response.Write GF_EDIT_DECIMALS(cdbl(rsVag("BRUTO"))-cdbl(rsVag("TARA")),0) & " Kg."%></td>
				<td align="center"><% if (not IsNull(rsVag("DTPESADA"))) then Response.Write Right(GF_FN2DTE(rsVag("DTPESADA")),8)%></td>
				<% if (Cdbl(rsVag("RECIBO")) > 0) then %>
					<td align="center"><% Response.Write rsVag("RECIBO") %></td>					
					<td align="center"><img src="../../images/pdf-16.png" title="Ver Nota de recepción" onclick="abrirNotaRecepcion('<%=g_cdOperativo%>','<%=rsVag("cdvagon")%>','<%=g_dtContable%>')" style="cursor:pointer;"></td>
				<% else %>
					<td colspan="2">
				<% end if %>				
			</tr>
			<tr>
            	<td id="trCaladaVagon_<%=rsVag("cdvagon")%>" name="trCaladaVagon_<%=rsVag("cdvagon")%>" colspan="7" class="troculto">
                	<div id="divCaladaVagon_<%=rsVag("cdvagon")%>"></div>
                </td>
			<tr>
		<%	rsVag.MoveNext()
		wend%>
        </tbody>	
	</table>	    
    <%else%>
    <table class="datagrid datagridlv1" width="100%" >
		<tr><td align="center">No se encontraron vagones</td></tr>
	</table>	    
	<%end if%>

</BODY>
</HTML>
