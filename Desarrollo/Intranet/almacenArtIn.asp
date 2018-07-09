<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->		
<!--#include file="Includes/procedimientosSql.asp"-->		
<!--#include file="Includes/procedimientosVales.asp"-->	
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
dim myColor1, myColor2, myColorP, cont
dim idArticulo, idAlmacen, cdVale, idObra, cdSolicitante, saldo, pIDVS
dim dsArticulo, abrArticulo, saldoArticulo
VS_idArticulo = GF_PARAMETROS7("idArticulo",0 ,6)
VS_idAlmacen = GF_PARAMETROS7("idAlmacen",0 ,6)
VS_cdVale = GF_PARAMETROS7("cdVale","" ,6)
VS_idObra = GF_PARAMETROS7("idObra",0 ,6)
VS_idBudgetArea = GF_PARAMETROS7("idArea",0 ,6)
VS_idBudgetDetalle = GF_PARAMETROS7("idDetalle",0 ,6)
VS_idPedido = GF_PARAMETROS7("idPedido",0 ,6)
VS_FechaSolicitud = GF_PARAMETROS7("fechaSolicitud","" ,6)
VS_idSector = GF_PARAMETROS7("sector",0 ,6)
cdSolicitante = GF_PARAMETROS7("cdSolicitante",0 ,6)
VS_saldo = GF_PARAMETROS7("saldo",0 ,6)
Call GP_ConfigurarMomentos
myColor1 = "#d3d3d3"
myColor2 = "#ffffff"
cont = 0
if VS_cdVale <> "" then
	VS_nroRemito = 0
	call grabarHeaderVale(pIDVS, VS_idPedido)
	call grabarValeDetalle(pIDVS, VS_idPedido)
	call actualizarStock()
	accion = ACCION_CERRAR
end if	
call getArticuloFull(VS_idArticulo, dsArticulo, abrArticulo)
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
<link rel="stylesheet" href="css/iwin.css"			 type="text/css">
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/iwin.js"></script>

<script type="text/javascript">
var refPopUpArtIn;
function artOutOnLoad() {
	refPopUpArtIn = startIWin('popupArt');
	<% if (accion = ACCION_CERRAR) then %>
		//refPopUpArtIn.hide();
	<% end if %>			
}
function grabarDetalleVale(idPedido, cdVale, idAlmacen, idObra, cdSolicitante, fechaSolicitud, idSector, idArea, idDetalle, idArticulo, saldo){
	document.getElementById("cdVale").value = cdVale;
	document.getElementById("idPedido").value = idPedido;
	document.getElementById("idAlmacen").value = idAlmacen;
	document.getElementById("idObra").value = idObra;
	document.getElementById("cdSolicitante").value = cdSolicitante;
	document.getElementById("idArea").value = idArea;
	document.getElementById("fechaSolicitud").value = fechaSolicitud;
	document.getElementById("sector").value = idSector;
	document.getElementById("idDetalle").value = idDetalle;
	document.getElementById("idArticulo").value = idArticulo;
	document.getElementById("saldo").value = document.getElementById(saldo).value;
	document.frmSel.submit();
}

</script>
</head>
<body onLoad="artOutOnLoad()">
<form name="frmSel" method="post" action="almacenArtIn.asp">
	<table class="reg_Header" align="center" width="100%" border="0" >				
		<tr>
			<td colspan="6" align="left"><font class="big"><% =VS_idArticulo & " - " & dsArticulo%></font></td>
		</tr>
		<tr>
			<!--<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Codigo") %></td>
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Descripcion") %></td>-->
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("VMP No") %></td>
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Entregado por") %></td>
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Prestado el") %></td>
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Cantidad") %></td>
			<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("R") %></td>
		</tr>
		<%
		strSQL = " SELECT T1.IDARTICULO, T1.FECHA, T1.IDVALE, T1.CDSOLICITANTE, T1.IDALMACEN, T1.IDOBRA, T1.CANTIDAD1, T2.CANTIDAD2,  T1.IDSECTOR,  T1.IDBUDGETAREA, T1.IDBUDGETDETALLE FROM  " & _
			"(  " & _
			"select C.IDVALE AS IDVALE, C.FECHA AS FECHA, C.CDSOLICITANTE AS CDSOLICITANTE, C.IDALMACEN AS IDALMACEN, C.IDBUDGETAREA AS IDBUDGETAREA,C.IDBUDGETDETALLE AS IDBUDGETDETALLE, " & _
			"       D.IDARTICULO AS IDARTICULO, C.IDOBRA AS IDOBRA, D.CANTIDAD AS CANTIDAD1, C.IDSECTOR AS IDSECTOR from TBLVALESCABECERA C inner join  TBLVALESDETALLE  D on C.IDVALE=D.IDVALE WHERE C.CDVALE='" & CODIGO_VS_PRESTAMO & "'   " & _
			") T1  " & _
			"LEFT join   " & _
			"(  " & _
			"select  " & _ 
			"    C1.PARTIDAPENDIENTE AS PARTIDAPENDIENTE,  " & _
			"    SUM(D1.CANTIDAD) AS CANTIDAD2  " & _ 
			" from TBLVALESCABECERA C1 inner join  TBLVALESDETALLE D1   on C1.IDVALE=D1.IDVALE  WHERE C1.CDVALE='" & CODIGO_VS_DEVOLUCION & "' GROUP BY C1.PARTIDAPENDIENTE " & _
			") T2  " & _
			"ON T1.IDVALE=T2.PARTIDAPENDIENTE  " & _
			"    where   " & _
			"   T1.IDARTICULO=" & VS_idArticulo & " and T1.IDALMACEN=" & VS_idAlmacen & " AND T1.CANTIDAD1 >0   " & _
			"   AND (T2.CANTIDAD2 > 0 OR T2.CANTIDAD2 IS NULL )  " & _
			"    order by T1.IDVALE  " 

		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strsql)
		while not rs.eof
			if isnull(rs("CANTIDAD2")) then
				saldoArticulo = rs("CANTIDAD1")
			else
				saldoArticulo = clng(rs("CANTIDAD1")) - clng(rs("CANTIDAD2"))
			end if	
			if cont mod 2 then
				colorP = myColor1
			else
				colorP = myColor2
			end if		
		%>
			<tr bgcolor="<%=colorP%>">
				<td align="center"><%=rs("IDVALE")%></td>
				<td align="center">
					<%
						VS_cdSolicitante = rs("CDSOLICITANTE")
						VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
						Response.Write VS_dsSolicitante & " - (" & VS_cdSolicitante & ")"
					%>
				</td>
				<td align="center"><%=GF_FN2DTE(rs("FECHA"))%></td>
				<td align='right'>
					<input size="5" style="text-align:right;" type="text" name="saldo_<%=cont%>" id="saldo_<%=cont%>" value="<%=GF_EDIT_DECIMALS(clng(saldoArticulo),0)%>">
				</td>
				<td title="<%=rs("IDVALE") & "-" & CODIGO_VS_DEVOLUCION & "-" & rs("IDALMACEN") & "-" & rs("IDOBRA") & "-" & rs("CDSOLICITANTE") & "-" & VS_idArticulo & "-" & cont%>" align='center'><img title="Recibir" onclick="grabarDetalleVale(<%=rs("IDVALE")%>,'<%=CODIGO_VS_DEVOLUCION%>',<%=rs("IDALMACEN")%>,<%=rs("IDOBRA")%>,'<%=rs("CDSOLICITANTE")%>','<%=GF_FN2DTE(rs("FECHA"))%>',<%=rs("IDSECTOR")%>,<%=rs("IDBUDGETAREA")%>,<%=rs("IDBUDGETDETALLE")%>,<%=VS_idArticulo%>,'saldo_<%=cont%>');" style="cursor:pointer;" src="images/almacenes/arrow_loan-16x16.png"></td>
			</tr>
		<%	
				cont = cont + 1
			rs.movenext
		wend	
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strsql)
		%>
	</table>

	<input type="hidden" name="cdVale" id="cdVale">
	<input type="hidden" name="idPedido" id="idPedido">
	<input type="hidden" name="idAlmacen" id="idAlmacen">
	<input type="hidden" name="idObra" id="idObra">
	<input type="hidden" name="cdSolicitante" id="cdSolicitante">
	<input type="hidden" name="idArticulo" id="idArticulo">
	<input type="hidden" name="saldo" id="saldo">
	<input type="hidden" name="idArea" id="idArea">
	<input type="hidden" name="idDetalle" id="idDetalle">	
	<input type="hidden" name="fechaSolicitud" id="fechaSolicitud">
	<input type="hidden" name="sector" id="sector">
</form>		
</body>
</html>