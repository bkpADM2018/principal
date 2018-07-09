<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'******************************************
'*** COMIENZO DE LA PAGINA
'******************************************
dim pAlmacenOut, myAuxHTML, myAuxHTMLtable
pAlmacenOut = GF_PARAMETROS7("idAlmacen", 0, 6)

dim verSegun1, verSegunVale
verSegun1 = GF_Parametros7("verSegun1", "", 6)
if verSegun1 = "" then verSegun1 = "F"
verSegunVale = GF_Parametros7("verSegunVale", "", 6)
if verSegunVale = "" then verSegunVale = "S"
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/iwin.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<title>Articulos Salida</title>
<script type="text/javascript">
	function loadPopUpArtSalidaPre(idArticulo, idAlmacen, typeOfView, cdVale) {
		parent.loadPopUpArtSalida(idArticulo, idAlmacen, typeOfView, cdVale);
	}
	function loadPopUpAJPPre(idPedido) {
		parent.loadPopUpAJP(idPedido);
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
		//myObj.style.position = 'relative';
	}
</script>

</head>
<body>	
<form id="frmSel" name="frmSel" action="almacenIFartSalida.asp" method="POST">	
<input type="hidden" id="idAlmacen" name="idAlmacen" value="<%=pAlmacenOut%>">
<input type="hidden" name="verSegun1" id="verSegun1" value="<%=verSegun1%>">
<input type="hidden" name="verSegunVale" id="verSegunVale" value="<%=verSegunVale%>">
	<table class="reg_Header round_border_all" valign="top" align="center" width="100%" border="0">
		<%
		if verSegun1 = "F" then
				strSQL = "Select c.idpedido,c.idalmacen,c.idalmacendest, c.cdsolicitante,c.FECHAREQUERIDO,T1.cdvale  from TBLPMCABECERA C " & _
						 "   inner join TBLPMDETALLE D " & _
						 "       on C.idpedido=d.idpedido " & _ 
						 "   LEFT join " & _ 
						 "   ( " & _
						 "   SELECT PARTIDAPENDIENTE, CDVALE  " & _
						 "       FROM TBLVALESCABECERA V " & _ 
						 "           WHERE CDVALE IN ('" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_TRANSFERENCIA & "') and ESTADO=" & ESTADO_ACTIVO & _ 
						 "       GROUP BY PARTIDAPENDIENTE, CDVALE " & _
						 "   ) T1 " & _ 
						 "   ON C.IDPEDIDO = T1.PARTIDAPENDIENTE  " & _
						 "   WHERE C.IDALMACEN=" & pAlmacenOut & " and d.saldo>0 " & _ 
						 "   group by c.idpedido, c.cdsolicitante, c.idalmacen,c.idalmacendest,T1.cdvale,c.FECHAREQUERIDO " & _
						 " order by C.FECHAREQUERIDO"
 				call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				if rs.eof then
					%>
						<tr>
							<td colspan="4" class="TDERROR round_border_all"><%=GF_TRADUCIR("No se encontraron formularios.")%></td>
						</tr>
					<%	
				else
					%>
					<tr>
						<td class="reg_Header_nav round_border_top_left" align="center"><% =GF_TRADUCIR("PM") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Solicitante") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Req. para el") %></td>
						<td class="reg_Header_nav" align="center">.</td>
						<td class="reg_Header_nav round_border_top_right" align="center">.</td>
					</tr>
					<%
					while not rs.eof
					%>
						<tr class="reg_header_navdos">
							<td ALIGN="CENTER"><%=rs("IDPEDIDO")%></td>
							<td align='LEFT'>
							<%
								VS_cdSolicitante = rs("CDSOLICITANTE")
								VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
								Response.Write VS_dsSolicitante & " - (" & VS_cdSolicitante & ")"
							%>
							</td>
							<td align="center">
								<% 
								if rs("CDVALE") = CODIGO_VS_PRESTAMO then 
									myAuxHTML = "<img title='Prestar'	 onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_PRESTAMO & "') style='cursor:pointer;' src='images/almacenes/arrow_loan-16x16.png'>"
								elseif rs("CDVALE") = CODIGO_VS_TRANSFERENCIA then
									myAuxHTML = "<img title='Transferir' onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_TRANSFERENCIA & "') style='cursor:pointer;' src='images/almacenes/arrow_transfer-16x16.png'>"
								elseif rs("CDVALE") = CODIGO_VS_SALIDA then
									myAuxHTML = "<img title='Entregar' onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_SALIDA & "') style='cursor:pointer;' src='images/almacenes/arrow_exit-16x16.png'>"
								elseif isnull(rs("CDVALE")) and rs("IDALMACENDEST")<>0 then
									myAuxHTML = "<img title='Transferir' onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_TRANSFERENCIA & "') style='cursor:pointer;' src='images/almacenes/arrow_transfer-16x16.png'>"
								else 
									myAuxHTML = "<img title='Definir Tipo' onclick=showTableHidden('TBL_HDN_" & rs("IDPEDIDO") & "') style='cursor:pointer;' src='images/question.gif'>"
											myAuxHTMLtable = myAuxHTMLtable & "<div>"
											myAuxHTMLtable = myAuxHTMLtable & "<table class='reg_Header' onmouseout=showTableHidden('TBL_HDN_" & rs("IDPEDIDO") & "') id='TBL_HDN_" & rs("IDPEDIDO") & "' style='position:absolute;visibility:hidden;'>"
											myAuxHTMLtable = myAuxHTMLtable & "<tr onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_SALIDA & "') style='cursor:pointer;' onmouseover=this.className='TDRESALTE'; onmouseout=this.className='';>"
											myAuxHTMLtable = myAuxHTMLtable & "<td><font class='small'>" & LEYENDA_VS_SALIDA & "</font></td>"
											myAuxHTMLtable = myAuxHTMLtable & "</tr>"
											myAuxHTMLtable = myAuxHTMLtable & "<tr onclick=loadPopUpArtSalidaPre(" & rs("IDPEDIDO") & "," & pAlmacenOut & ",'" & verSegun1 & "','" & CODIGO_VS_PRESTAMO & "') style='cursor:pointer;' onmouseover=this.className='TDRESALTE'; onmouseout=this.className='';>"
											myAuxHTMLtable = myAuxHTMLtable & "<td><font class='small'>" & LEYENDA_VS_PRESTAMO & "</font></td>"
											myAuxHTMLtable = myAuxHTMLtable & "</tr>"
											myAuxHTMLtable = myAuxHTMLtable & "</table>"
											myAuxHTMLtable = myAuxHTMLtable & "</div>"
								end if 
								Response.Write myAuxHTMLtable
							%>
							<%=GF_FN2DTE(rs("FECHAREQUERIDO"))%></td>
							<td align="center">
							<%
							Response.Write myAuxHTML
							%>
							</td>
							<td align="center"><img src="images/almacenes/AJP-16x16.png" title="<% =GF_TRADUCIR("Ajustar Pedido")%>" style="cursor:pointer" onClick="loadPopUpAJPPre(<%=rs("IDPEDIDO")%>)"></td>
						</tr>
						<%	
						rs.movenext
					wend
				end if	
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
		else				
				strSQL ="Select SALDO,TG.IDARTICULO AS IDARTICULO, DSARTICULO, STOCK FROM " & _ 
						"(Select SUM(D.SALDO) AS SALDO, A.IDARTICULO AS IDARTICULO, A.DSARTICULO AS DSARTICULO  " & _
						" from TBLPMDETALLE D inner join TBLPMCABECERA C  " & _
						" on C.IDPEDIDO=D.IDPEDIDO inner join TBLARTICULOS A  " & _
						" on D.IDARTICULO=A.IDARTICULO  " & _
						" left join ( " & _
						"        SELECT PARTIDAPENDIENTE FROM TBLVALESCABECERA VC  " & _
						"        WHERE VC.CDVALE IN ('" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_TRANSFERENCIA & "') AND VC.ESTADO=" & ESTADO_ACTIVO & " GROUP BY PARTIDAPENDIENTE) T1 " & _
						" ON C.IDPEDIDO=T1.PARTIDAPENDIENTE " & _
						" where D.SALDO>0 and C.IDALMACEN=" & pAlmacenOut & _
						" GROUP BY A.IDARTICULO, A.DSARTICULO) TG " & _
						" left join (select IDARTICULO, (EXISTENCIA+SOBRANTE) as STOCK  from tblarticulosdatos " & _
						" where idalmacen = " & pAlmacenOut & ")  S on TG.idarticulo = S.idarticulo "
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				if rs.eof then
					%>
						<tr>
							<td colspan="4" class="TDERROR round_border_all"><%=GF_TRADUCIR("No se encontraron articulos pendientes de entrega")%></td>
						</tr>
					<%	
				else
					%>
					<tr>
						<td class="reg_Header_nav round_border_top_left" align="center"><% =GF_TRADUCIR("Cod") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Descripcion") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Stock") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Cant.") %></td>
						<td class="reg_Header_nav round_border_top_right" align="center">.</td>
					</tr>
					<%
					while not rs.eof
					%>
						<tr class="reg_header_navdos">
							<td ALIGN="CENTER"><%=rs("IDARTICULO")%></td>
							<td title="<%=rs("DSARTICULO")%>"><%=left(rs("DSARTICULO"),30)%></td>
							<td align='right'>
								<%	
									if not isNull(rs("STOCK")) then 
										Response.Write GF_EDIT_DECIMALS(CDbl(rs("STOCK"))*1000,3)
									else
										Response.Write 0
									end if
								%>
							</td>
							<td align='right'><%=GF_EDIT_DECIMALS(CDbl(rs("SALDO"))*100,2)%></td>
							<td align='center'><img title="Entregar" onclick="loadPopUpArtSalidaPre(<%=rs("IDARTICULO")%>,<%=pAlmacenOut%>,'<%=verSegun1%>','<%=CODIGO_VS_SALIDA%>')" style="cursor:pointer;" src="images/almacenes/arrow_exit-16x16.png"></td>
						</tr>
					<%	
						rs.movenext
					wend
				end if	
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
		end if
		%>
	</table>
</form>

</body>
</html>