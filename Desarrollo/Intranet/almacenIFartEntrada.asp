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
dim pAlmacenIn
pAlmacenIn = GF_PARAMETROS7("idAlmacen", 0, 6)

dim verSegun2
verSegun2 = GF_Parametros7("verSegun2", "", 6)
if verSegun2 = "" then verSegun2 = "F"
verSegunVale = GF_Parametros7("verSegunVale", "", 6)
if verSegunVale = "" then verSegunVale = "S"
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/iwin.js"></script>
<title>Articulos Entrada</title>
<script type="text/javascript">
	function loadPopUpArtEntradaPre(idArticulo, idAlmacen, typeOfView, cdVale) {
		parent.loadPopUpArtEntrada(idArticulo, idAlmacen, typeOfView, cdVale);
	}	
	function loadPopUpAJUPre(idPedido) {
		parent.loadPopUpAJU(idPedido);
	}
</script>

</head>
<body>	
<form id="frmSel" name="frmSel" action="almacenIFartEntrada.asp" method="POST">	
<input type="hidden" id="idAlmacen" name="idAlmacen" value="<%=pAlmacenIn%>">
<input type="hidden" name="verSegun2" id="verSegun2" value="<%=verSegun2%>">
<input type="hidden" name="verSegunVale" id="verSegunVale" value="<%=verSegunVale%>">
	<table class="reg_Header round_border_all" valign="top" align="left" width="100%" border="0" >				

		<%
		if verSegun2 = "A" then
				strSQL = "SELECT * FROM (select sum(SALDO1) as SALDO, IDARTICULO, DSARTICULO from (Select C.CDVALE, case(C.CDVALE) when '" & CODIGO_VS_DEVOLUCION & "' then SUM(-D.CANTIDAD) when '" & CODIGO_VS_AJUSTE_VALE & "' then SUM(-D.CANTIDAD) when '" & CODIGO_VS_PRESTAMO & "' then SUM(D.CANTIDAD) end " & chr(34) & "SALDO1" & chr(34) & " ,A.IDARTICULO AS IDARTICULO, A.DSARTICULO AS DSARTICULO from TBLVALESDETALLE D inner join TBLVALESCABECERA C on C.IDVALE=D.IDVALE inner join TBLARTICULOS A on D.IDARTICULO=A.IDARTICULO where D.CANTIDAD>0 and C.CDVALE in ('" & CODIGO_VS_DEVOLUCION & "','" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_AJUSTE_VALE & "') and C.IDALMACEN=" & pAlmacenIn & " and C.ESTADO=" & ESTADO_ACTIVO & " GROUP BY A.IDARTICULO, A.DSARTICULO, C.CDVALE) T1 GROUP BY IDARTICULO, DSARTICULO) TAB WHERE TAB.SALDO>0 order by TAB.IDARTICULO"
			
				'response.Write strSQL
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
				if rs.eof then
					%>
						<tr>
							<td colspan="4" class="TDERROR round_border_all"><%=GF_TRADUCIR("No se encontraron articulos pendientes de devolucion")%></td>
						</tr>
					<%	
				else
					%>
					<tr>
						<td class="reg_Header_nav round_border_top_left" align="center"><% =GF_TRADUCIR("Cod") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Descripcion") %></td>
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Cant.") %></td>
						<td class="reg_Header_nav round_border_top_right" align="center">.</td>						
					</tr>

					<%
					while not rs.eof
					%>
						<tr class="reg_header_navdos">
							<td align="center"><%=rs("IDARTICULO")%></td>
							<td title="<%=rs("DSARTICULO")%>"><%=left(rs("DSARTICULO"),34)%></td>
							<td align='right'><%=GF_EDIT_DECIMALS(CDbl(rs("SALDO"))*100,2)%></td>
							<td align='center'><img title="Devolver" onclick="loadPopUpArtEntradaPre(<%=rs("IDARTICULO")%>,<%=pAlmacenIn%>,'<%=verSegun2%>','<%=CODIGO_VS_DEVOLUCION%>')" style="cursor:pointer;" src="images/almacenes/arrow_loan-16x16.png"></td>							
						</tr>
					<%	
						rs.movenext
					wend
				end if	
				call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
		else

				strSQL = "SELECT TG.IDPM, SALDO, CDSOLICITANTE, PM.FECHASOLICITUD AS FECHA FROM " & _
					"	 ( " & _
					"    SELECT PARTIDAPENDIENTE AS IDPM, SUM(SALDO) AS SALDO FROM " & _ 
					"		( " & _
					"		SELECT  C.PARTIDAPENDIENTE as PARTIDAPENDIENTE,  " & _
					"			    C.CDSOLICITANTE AS CDSOLICITANTE, " & _
					"				D.IDARTICULO, C.FECHA, C.CDVALE,  " & _
					"				CASE(C.CDVALE)  WHEN '" & CODIGO_VS_DEVOLUCION & "' THEN SUM(-D.CANTIDAD) " & _
					"					            WHEN '" & CODIGO_VS_PRESTAMO & "' THEN SUM(D.CANTIDAD)  " & _
					"						        WHEN '" & CODIGO_VS_AJUSTE_VALE & "' THEN SUM(-D.CANTIDAD)  " & _
					"							    END  " & chr(34) & "SALDO" & chr(34) & _
					"			FROM TBLVALESCABECERA C  " & _
					"				INNER JOIN TBLVALESDETALLE D   " & _ 
					"					ON C.IDVALE = D.IDVALE WHERE c.cdvale IN ('" & CODIGO_VS_PRESTAMO & "','" & CODIGO_VS_DEVOLUCION & "','" & CODIGO_VS_AJUSTE_VALE & "') " & _
					"					AND C.IDALMACEN=" & pAlmacenIn & " and C.ESTADO=" & ESTADO_ACTIVO & _
					"					GROUP BY C.PARTIDAPENDIENTE, D.IDARTICULO, C.CDVALE, C.CDSOLICITANTE, C.FECHA  " & _
					"		) T1  " & _
					"       GROUP BY T1.PARTIDAPENDIENTE, T1.IDARTICULO HAVING SUM(SALDO) > 0 " & _ 
					"  )TG " & _      
					"    INNER JOIN " & _ 
					"      TBLPMCABECERA PM " & _ 
					" ON TG.IDPM=PM.IDPEDIDO"
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
						<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Requerido el") %></td>
						<td class="reg_Header_nav" align="center">.</td>
						<td class="reg_Header_nav round_border_top_right" align="center">.</td>
					</tr>

					<%
					while not rs.eof
					%>
						<tr class="reg_header_navdos">
							<td align="center"><%=rs("IDPM")%></td>
							<td>
							<%
								VS_cdSolicitante = rs("CDSOLICITANTE")
								VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
								Response.Write VS_dsSolicitante & " - (" & VS_cdSolicitante & ")"
							%>
							</td>
							<td align="center"><%=GF_FN2DTE(rs("FECHA"))%></td>
							<td align='center'>
								<img title="Devolver" onclick="loadPopUpArtEntradaPre(<%=rs("IDPM")%>,<%=pAlmacenIn%>,'<%=verSegun2%>','<%=CODIGO_VS_DEVOLUCION%>')" style="cursor:pointer;" src="images/almacenes/arrow_loan-16x16.png">
							</td>								
							<td align="center"><img src="images/almacenes/AJU-16x16.png" title="<% =GF_TRADUCIR("Ajustar Devolución")%>" style="cursor:pointer" onClick="loadPopUpAJUPre(<%=rs("IDPM")%>)"></td>
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