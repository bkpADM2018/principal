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
dim pAlmacenIn, rsTRANSF, SaldoAux, textAlmacenO_D, dicForms, isWaiting
isWaiting = false
pAlmacenIn = GF_PARAMETROS7("idAlmacen", 0, 6)
set dicForms = Server.CreateObject("Scripting.Dictionary") 
dim verSegun3
verSegun3 = GF_Parametros7("verSegun3", "", 6)
if verSegun3 = "" then verSegun3 = "F"
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/iwin.js"></script>
<title>Articulos Entrada</title>
<script type="text/javascript">
	function loadPopUpArtTransferenciaPre(idArticulo, idAlmacen, typeOfView, cdVale, idAlmacenDest) {
		parent.loadPopUpArtTransferencia(idArticulo, idAlmacen, typeOfView, cdVale, idAlmacenDest);
	}
	function loadPopUpAJTPre(idPedido) {
		parent.loadPopUpAJT(idPedido);
	}
</script>

</head>
<body>	
<form id="frmSel" name="frmSel" action="almacenIFartTransferencia.asp" method="POST">	
<input type="hidden" id="idAlmacen" name="idAlmacen" value="<%=pAlmacenIn%>">
<input type="hidden" name="verSegun3" id="verSegun3" value="<%=verSegun3%>">
	<table class="reg_Header round_border_all" valign="top" align="left" width="100%" border="0" >				
		<%
		strSQL= "SELECT * " & _
		"FROM   ( SELECT PMC.IDPEDIDO, PMC.IDALMACEN, PMC.IDOBRA, PMC.FECHASOLICITUD, PMC.FECHAREQUERIDO, PMC.CDUSUARIO, PMC.MOMENTO, PMC.IDALMACENDEST, PMC.IDBUDGETDETALLE, PMC.IDBUDGETAREA, PMC.COMENTARIOS, PMC.IDSECTOR, PMC.CDSOLICITANTE, " & _
		"               PMD.CANTIDAD, PMD.SALDO                         , " & _
		"               ART.DSARTICULO                          , " & _
		"               TG.PARTIDAPENDIENTE AS PARTPEND, " & _
		"               TG.IDARTICULO                  , " & _
		"               TG.SALDOTOTAL " & _
		"       FROM    ( SELECT  T1.PARTIDAPENDIENTE, " & _
		"                        T1.IDARTICULO       , " & _
		"                        SUM(T1.SALDO) AS SALDOTOTAL " & _
		"               FROM     ( SELECT  VC1.PARTIDAPENDIENTE, " & _
		"                                 VD1.IDARTICULO       , " & _
		"                                 VC1.CDVALE           , " & _
		"                                 CASE(VC1.CDVALE) " & _
		"                                          WHEN '" & CODIGO_VS_TRANSFERENCIA & "' " & _
		"                                          THEN SUM(VD1.CANTIDAD) " & _
		"                                          WHEN '" & CODIGO_VS_RECEPCION & "' " & _
		"                                          THEN SUM(-VD1.CANTIDAD) " & _
		"                                          WHEN '" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "' " & _
		"                                          THEN SUM(-VD1.CANTIDAD) " & _
		"                                 END SALDO " & _
		"                        FROM     TBLVALESCABECERA VC1 " & _
		"                                 INNER JOIN TBLVALESDETALLE VD1 " & _
		"                                 ON       VC1.IDVALE = VD1.IDVALE " & _
		"                        WHERE    VC1.CDVALE IN ('" & CODIGO_VS_TRANSFERENCIA & "', " & _
		"                                                '" & CODIGO_VS_RECEPCION & "', " & _
		"                                                '" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "') " & _
		"                        AND      VC1.ESTADO=" & ESTADO_ACTIVO & _
		"                        GROUP BY VC1.PARTIDAPENDIENTE, " & _
		"                                 VD1.IDARTICULO      , " & _
		"                                 VC1.CDVALE " & _
		"                        ) T1 " & _
		"               GROUP BY T1.PARTIDAPENDIENTE, " & _
		"                        T1.IDARTICULO " & _
		"               HAVING   SUM(T1.SALDO) > 0 " & _
		"               ) TG " & _
		"               INNER JOIN TBLPMCABECERA PMC " & _
		"               ON      TG.PARTIDAPENDIENTE=PMC.IDPEDIDO " & _
		"               INNER JOIN TBLPMDETALLE PMD " & _
		"               ON      PMC.IDPEDIDO  =PMD.IDPEDIDO " & _
		"               AND     PMD.IDARTICULO=TG.IDARTICULO " & _
		"               INNER JOIN TBLARTICULOS ART " & _
		"               ON      ART.IDARTICULO=TG.IDARTICULO " & _
		"       ) TAB " & _
		"WHERE  ( TAB.IDALMACEN    = " & pAlmacenIn & " OR TAB.IDALMACENDEST= " & pAlmacenIn & " )" & _
		"ORDER BY TAB.PARTPEND"
		call executeQueryDb(DBSITE_SQL_INTRA, rsTRANSF, "OPEN", strSQL)
		if verSegun3 = "A" then
			if rsTRANSF.eof then
				%>
				<tr>
					<td colspan="5" class="TDERROR round_border_all"><%=GF_TRADUCIR("No se encontraron articulos pendientes de devolucion")%></td>
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
				while not rsTRANSF.eof
					isWaiting = false
					SaldoAux = 0
					if clng(rsTRANSF("SALDO")) = 0 then 'Se transmitio todo ya
						if clng(rsTRANSF("SALDOTOTAL")) <> 0 then 'Se recibio todo
							SaldoAux = clng(rsTRANSF("SALDOTOTAL"))
							if pAlmacenIn = rsTRANSF("IDALMACEN") then isWaiting = true
						end if	
					else
						if clng(rsTRANSF("SALDOTOTAL")) <> 0 then 'Se recibio todo
							SaldoAux = clng(rsTRANSF("SALDOTOTAL"))
							if pAlmacenIn = rsTRANSF("IDALMACEN") then isWaiting = true
						end if
					end if	


						
							
					if SaldoAux <> 0 then
					%>
					<tr class="reg_header_navdos">
						<td ALIGN="CENTER"><%=rsTRANSF("IDARTICULO")%></td>
						<td title="<%=rsTRANSF("DSARTICULO")%>"><%=left(rsTRANSF("DSARTICULO"),34)%></td>
						<td align='right'><%=GF_EDIT_DECIMALS(CDbl(SaldoAux)*100,2)%></td>
						<td align="center">
						<%	if isWaiting then	%>
								<img title="Aguardando recepcion en destino" style="cursor:pointer;" src="images/almacenes/reception_waiting-16x16.png">
						<%	else %>
								<img title="Recibir" onclick="loadPopUpArtTransferenciaPre(<%=rsTRANSF("IDARTICULO")%>,<%=pAlmacenIn%>,'<%=verSegun3%>','<%=CODIGO_VS_TRANSFERENCIA%>',<%=rsTRANSF("IDALMACENDEST")%>)" style="cursor:pointer;" src="images/almacenes/arrow_reception-16x16.png">
						<%	end if %>
						</td>
					</tr>
				<%	
					end if
					rsTRANSF.movenext
				wend
			end if	
			call executeQueryDb(DBSITE_SQL_INTRA, rsTRANSF, "CLOSE", strSQL)
		else
			if rsTRANSF.eof then
				%>
				<tr>
					<td colspan="5" class="TDERROR round_border_all"><%=GF_TRADUCIR("No se encontraron formularios.")%></td>
				</tr>
				<%	
			else
				%>
				<tr>
					<td class="reg_Header_nav round_border_top_left" align="center"><% =GF_TRADUCIR("PM") %></td>
					<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Solicitante") %></td>
					<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Orig/Dest") %></td>
					<td class="reg_Header_nav" align="center"><% =GF_TRADUCIR("Requerido el") %></td>
					<td class="reg_Header_nav" align="center">.</td>
					<td class="reg_Header_nav round_border_top_right" align="center">.</td>
				</tr>
				<%
				while not rsTRANSF.eof
					isWaiting = false					
					SaldoAux = 0
					if clng(rsTRANSF("SALDO")) = 0 then 'Se transmitio todo ya
						if clng(rsTRANSF("SALDOTOTAL")) <> 0 then 'Se recibio todo
							SaldoAux = clng(rsTRANSF("SALDOTOTAL"))
							if pAlmacenIn = rsTRANSF("IDALMACEN") then isWaiting = true
						end if	
					else
						if clng(rsTRANSF("SALDOTOTAL")) <> 0 then 'Se recibio todo
							SaldoAux = clng(rsTRANSF("SALDOTOTAL"))
							if pAlmacenIn = rsTRANSF("IDALMACEN") then isWaiting = true
						end if
					end if					

					if SaldoAux <> 0 and dicForms.Exists(clng(rsTRANSF("idpedido"))) = false then
						call dicForms.Add (clng(rsTRANSF("idpedido")),rsTRANSF("idpedido"))
					
					%>
					<tr class="reg_header_navdos">
						<td align="center"><%=rsTRANSF("idpedido")%></td>
						<td>
						<%
							VS_cdSolicitante = rsTRANSF("CDSOLICITANTE")
							VS_dsSolicitante = getUserDescription(VS_cdSolicitante)
							Response.Write VS_dsSolicitante
						%>
						</TD>
						<td align="center">
						<%	
							Set rsAlmacenes = obtenerListaAlmacenes(rsTRANSF("IDALMACEN")) 
							if (not rsAlmacenes.eof) then
								textAlmacenO_D = rsAlmacenes("CDALMACEN")
							end if
							Set rsAlmacenes = obtenerListaAlmacenes(rsTRANSF("IDALMACENDEST")) 
							if (not rsAlmacenes.eof) then
								textAlmacenO_D = textAlmacenO_D & "/" & rsAlmacenes("CDALMACEN")
							end if
							Response.Write textAlmacenO_D 
						%>							
						</td>
						<td align="center"><%=GF_FN2DTE(rsTRANSF("FECHAREQUERIDO"))%></td>
						<%	if isWaiting then	%>
								<td align="center">
									<img title="<% =GF_TRADUCIR("Aguardando recepcion en destino") %>" style="cursor:pointer;" src="images/almacenes/reception_waiting-16x16.png">
								</td>
								<td align="center">
									<img title="<% =GF_TRADUCIR("Ajustar Transferencia") %>" onclick="loadPopUpAJTPre(<%=rsTRANSF("IDPEDIDO")%>)" src="images/almacenes/AJT-16x16.png" style="cursor:pointer;">
								</td>
						<%	else %>
								<td align="center">
									<img title="<% =GF_TRADUCIR("Recibir") %>" onclick="loadPopUpArtTransferenciaPre(<%=rsTRANSF("IDPEDIDO")%>,<%=pAlmacenIn%>,'<%=verSegun3%>','<%=CODIGO_VS_RECEPCION%>',<%=rsTRANSF("IDALMACENDEST")%>)" style="cursor:pointer;" src="images/almacenes/arrow_reception-16x16.png">
								</td>
								<td align="center">&nbsp;</td>
						<%	end if %>
					</tr>
				<%
					end if
					rsTRANSF.movenext
				wend
			end if	
			call executeQueryDb(DBSITE_SQL_INTRA, rsTRANSF, "CLOSE", strSQL)
		end if		
		%>
	</table>
</form>
</body>
</html>