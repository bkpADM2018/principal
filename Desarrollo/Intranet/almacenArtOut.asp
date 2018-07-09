<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosAlmacenes.asp"-->	
<!--#include file="Includes/procedimientosObras.asp"-->		
<!--#include file="Includes/procedimientosSql.asp"-->		
<!--#include file="Includes/procedimientosVales.asp"-->	
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'***************************************************
'******   COMIENZO DE LA PAGINA
'***************************************************
dim myColor1, myColor2, myColorP, cont
dim idArticulo, idAlmacen, cdVale, idObra, cdSolicitante, saldo, pIDVS
dim dsArticulo, abrArticulo, typeOfView
VS_idArticulo = GF_PARAMETROS7("idArticulo",0 ,6)
VS_idAlmacen = GF_PARAMETROS7("idAlmacen",0 ,6)
VS_cdVale = GF_PARAMETROS7("cdVale","" ,6)
VS_idObra = GF_PARAMETROS7("idObra",0 ,6)
VS_idBudgetArea = GF_PARAMETROS7("idArea",0 ,6)
VS_idBudgetDetalle = GF_PARAMETROS7("idDetalle",0 ,6)
VS_idPedido = GF_PARAMETROS7("idPedido",0 ,6)
VS_cdSolicitante = GF_PARAMETROS7("cdSolicitante",0 ,6)
VS_saldo = GF_PARAMETROS7("saldo",0 ,6)
VS_FechaSolicitud = GF_PARAMETROS7("fechaSolicitud","" ,6)
VS_idSector = GF_PARAMETROS7("sector",0 ,6)
typeOfView = GF_PARAMETROS7("typeOfView","" ,6)
Call GP_ConfigurarMomentos
myColor1 = "#d3d3d3"
myColor2 = "#ffffff"
if VS_cdVale <> "" then
	VS_nroRemito = 0
	VS_secBudget = 0
	call grabarHeaderVale(pIDVS, VS_idPedido)
	call grabarValeDetalle(pIDVS, VS_idPedido)
	call actualizarPMDetalle(VS_idPedido, VS_idArticulo, VS_saldo)
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
var refPopUpArtOut;
function artOutOnLoad() {
	refPopUpArtOut = startIWin('popupArt');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpArtOut.hide();
	<% end if %>			
}
function grabarDetalleVale(idPedido, cdVale, idAlmacen, idObra, fechaSolicitud, sector, cdSolicitante, idArea, idDetalle, idArticulo, saldo){
	document.getElementById("cdVale").value = cdVale;
	document.getElementById("idPedido").value = idPedido;
	document.getElementById("idAlmacen").value = idAlmacen;
	document.getElementById("idObra").value = idObra;
	document.getElementById("cdSolicitante").value = cdSolicitante;
	document.getElementById("idDetalle").value = idDetalle;
	document.getElementById("idArea").value = idArea;		
	document.getElementById("idArticulo").value = idArticulo;
	document.getElementById("fechaSolicitud").value = fechaSolicitud;
	document.getElementById("sector").value = sector;
	document.getElementById("saldo").value = document.getElementById(saldo).value;
	document.frmSel.submit();
}
	function showTableHidden(pTableName){
		var myObj;
		myObj = document.getElementById(pTableName);
		if (myObj.style.visibility == 'visible'){
			myObj.style.visibility = 'hidden';
		}
		else{
			myObj.style.visibility = 'visible';
		}
	}
</script>
</head>
<body onLoad="artOutOnLoad()">
<form name="frmSel" method="post" action="almacenArtOut.asp">
	<table class="reg_Header" align="center" width="100%" border="0" >				
		
					<tr>
						<td colspan="6" align="left"><font class="big"><% =VS_idArticulo & " - " & dsArticulo%></font></td>
					</tr>
					<tr>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Pedido No") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Solicitado por") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Solicitado el") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Cantidad") %></td>
						<td class="reg_Header_nav" align="center">E</td>
					</tr>
					<%					
					strSQL =	"SELECT T1.*, T2.CDVALE FROM " & _
								"	( " & _
								" SELECT C.IDPEDIDO, C.CDSOLICITANTE, C.FECHASOLICITUD, D.SALDO, C.IDBUDGETDETALLE, C.IDBUDGETAREA, C.IDOBRA, C.IDALMACEN, C.IDSECTOR  " & _
								"    from TBLPMDETALLE D inner join TBLPMCABECERA C  " & _
								"            on C.IDPEDIDO=D.IDPEDIDO " & _
								"    where D.SALDO>0 and C.IDALMACEN=" & VS_idAlmacen & " and d.idarticulo=" & VS_idArticulo & _
								"    ) T1 " & _
								" LEFT JOIN  " & _
								"    ( " & _
								"    SELECT PARTIDAPENDIENTE,  CDVALE " & _
								"        FROM TBLVALESCABECERA VC  " & _
								"           where VC.CDVALE IN ('" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_TRANSFERENCIA & "')" & _
								"           GROUP BY PARTIDAPENDIENTE,  CDVALE " & _
								" )T2 " & _
								"       ON T1.IDPEDIDO = T2.PARTIDAPENDIENTE    " & _
								"       ORDER BY T1.IDPEDIDO"
					call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
					while not rs.eof
						if cont mod 2 then
							colorP = myColor1
						else
							colorP = myColor2
						end if		
						%>
						<tr bgcolor="<%=colorP%>">
							<td align="center"><%=rs("IDPEDIDO")%></td>
							<td align="center">
								<%
									VS_cdSolicitante = rs("CDSOLICITANTE")
									VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
									Response.Write VS_dsSolicitante & " - (" & VS_cdSolicitante & ")"
								%>
							</td>
							<td align="center"><%=GF_FN2DTE(rs("FECHASOLICITUD"))%></td>
							<td align='right'>
								<% 
								if rs("CDVALE") = CODIGO_VS_PRESTAMO then 
									myAuxHTML = "<img title='Prestar' onclick=grabarDetalleVale(" & rs("IDPEDIDO") & ",'" & CODIGO_VS_PRESTAMO & "'," & rs("IDALMACEN") & "," & rs("IDOBRA") & ",'" & GF_FN2DTE(rs("FECHASOLICITUD")) & "'," & rs("IDSECTOR") & ",'" & rs("CDSOLICITANTE") & "'," & rs("IDBUDGETAREA") & "," & rs("IDBUDGETDETALLE") & "," & VS_idArticulo & ",'saldo_" & cont & "') style='cursor:pointer;' src='images/almacenes/arrow_loan-16x16.png'>"
								elseif rs("CDVALE") = CODIGO_VS_SALIDA then
									myAuxHTML = "<img title='Entregar' onclick=grabarDetalleVale(" & rs("IDPEDIDO") & ",'" & CODIGO_VS_SALIDA & "'," & rs("IDALMACEN") & "," & rs("IDOBRA") & ",'" & GF_FN2DTE(rs("FECHASOLICITUD")) & "'," & rs("IDSECTOR") & ",'" & rs("CDSOLICITANTE") & "'," & rs("IDBUDGETAREA") & "," & rs("IDBUDGETDETALLE") & "," & VS_idArticulo & ",'saldo_" & cont & "') style='cursor:pointer;' src='images/almacenes/arrow_exit-16x16.png'>"
								elseif rs("CDVALE") = CODIGO_VS_TRANSFERENCIA then
									myAuxHTML = "<img title='Transferir' onclick=grabarDetalleVale(" & rs("IDPEDIDO") & ",'" & CODIGO_VS_TRANSFERENCIA & "'," & rs("IDALMACEN") & "," & rs("IDOBRA") & ",'" & GF_FN2DTE(rs("FECHASOLICITUD")) & "'," & rs("IDSECTOR") & ",'" & rs("CDSOLICITANTE") & "'," & rs("IDBUDGETAREA") & "," & rs("IDBUDGETDETALLE") & "," & VS_idArticulo & ",'saldo_" & cont & "') style='cursor:pointer;' src='images/almacenes/arrow_transfer-16x16.png'>"
								else 
									myAuxHTML = "<img title='Definir Tipo' onclick=showTableHidden('TBL_HDN_" & rs("IDPEDIDO") & "') style='cursor:pointer;' src='images/question.gif'>"
											myAuxHTMLtable = myAuxHTMLtable & "<div>"
											myAuxHTMLtable = myAuxHTMLtable & "<table class='reg_Header' onmouseout=showTableHidden('TBL_HDN_" & rs("IDPEDIDO") & "') id='TBL_HDN_" & rs("IDPEDIDO") & "' style='position:absolute;visibility:hidden;'>"
											myAuxHTMLtable = myAuxHTMLtable & "<tr onclick=grabarDetalleVale(" & rs("IDPEDIDO") & ",'" & CODIGO_VS_SALIDA & "'," & rs("IDALMACEN") & "," & rs("IDOBRA") & ",'" & GF_FN2DTE(rs("FECHASOLICITUD")) & "'," & rs("IDSECTOR") & ",'" & rs("CDSOLICITANTE") & "'," & rs("IDBUDGETAREA") & "," & rs("IDBUDGETDETALLE") & "," & VS_idArticulo & ",'saldo_" & cont & "') style='cursor:pointer;' onmouseover=this.className='TDRESALTE'; onmouseout=this.className='';>"
											myAuxHTMLtable = myAuxHTMLtable & "<td><font class='small'>" & LEYENDA_VS_SALIDA & "</font></td>"
											myAuxHTMLtable = myAuxHTMLtable & "</tr>"
											myAuxHTMLtable = myAuxHTMLtable & "<tr onclick=grabarDetalleVale(" & rs("IDPEDIDO") & ",'" & CODIGO_VS_PRESTAMO & "'," & rs("IDALMACEN") & "," & rs("IDOBRA") & ",'" & GF_FN2DTE(rs("FECHASOLICITUD")) & "'," & rs("IDSECTOR") & ",'" & rs("CDSOLICITANTE") & "'," & VS_idArticulo & ",'saldo_" & cont & "') style='cursor:pointer;' onmouseover=this.className='TDRESALTE'; onmouseout=this.className='';>"
											myAuxHTMLtable = myAuxHTMLtable & "<td><font class='small'>" & LEYENDA_VS_PRESTAMO & "</font></td>"
											myAuxHTMLtable = myAuxHTMLtable & "</tr>"
											myAuxHTMLtable = myAuxHTMLtable & "</table>"
											myAuxHTMLtable = myAuxHTMLtable & "</div>"
								end if 
								Response.Write myAuxHTMLtable
								%>
								<input size="5" style="text-align:right;" type="text" name="saldo_<%=cont%>" id="saldo_<%=cont%>" value="<%=GF_EDIT_DECIMALS(clng(rs("SALDO")),0)%>">
							</td>
							<td align="center" rowspan>
							<%
							Response.Write myAuxHTML
							%>							

							</td>
						</tr>
						<%	
						cont = cont + 1
						rs.movenext
					wend	
					call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
					%>
	</table>

	<input type="hidden" name="cdVale" id="cdVale">
	<input type="hidden" name="idPedido" id="idPedido">
	<input type="hidden" name="idAlmacen" id="idAlmacen">
	<input type="hidden" name="idObra" id="idObra">
	<input type="hidden" name="cdSolicitante" id="cdSolicitante">
	<input type="hidden" name="idArticulo" id="idArticulo">
	<input type="hidden" name="idArea" id="idArea">
	<input type="hidden" name="idDetalle" id="idDetalle">
	<input type="hidden" name="saldo" id="saldo">
	<input type="hidden" name="fechaSolicitud" id="fechaSolicitud">
	<input type="hidden" name="sector" id="sector">
</form>		
</body>
</html>