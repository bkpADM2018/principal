<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<%
'----------------------------------------------------------------------------------
Function actualizarEstadoPDC(pIdPoliza, pEstado)
	Dim strSQL
	strSQL = "UPDATE TBLPOLIZASCAUCION SET ESTADO = " & pEstado & " WHERE IDPDC = " &	pIdPoliza
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXECUTE", strSQL)
End Function
'----------------------------------------------------------------------------------
Function loadPolizaVencida(pNroPoliza, pIdAseguradora, pTomador, pImporte, pMoneda, pImporteAprox, pFecha, pidDivision, pCdPedido, pIdEstado)
	Dim strSQL, myWhere, pViewPDCvenc		
	Set rs = readPDC(0, pCdPedido, pNroPoliza, pIdAseguradora, pTomador, pIdEstado, pidDivision, "", pImporte, pMoneda, pImporteAprox, pFecha, true)
	if(not rs.Eof)then	%>
		<table class="reg_header" id="TBL1" width="100%" cellpadding=0 cellspacing=0 border=0>		
			<tr class="titleVencido">
				<td align="center"><%=GF_Traducir("Polizas vencidas")%></td>
			</tr>
			<tr>			
				<td>	
					<table width="100%" cellspacing="1" cellpadding="1" align="center" border="0">	
						<tr class="titleVencido">
							<td width="10%" align="center"><%=GF_Traducir("Pedido")%></td>
							<td width="3%"></td>
							<td width="12%" align="center"><%=GF_Traducir("Nro Poliza")%></td>			
							<td width="30%" align="center"><%=GF_Traducir("Aseguradora")%></td>
							<td width="23%" align="center"><%=GF_Traducir("Tomador")%></td>
							<td width="11%" align="center"><%=GF_Traducir("Monto")%></td>
							<td width="8%" align="center"><%=GF_Traducir("Vencimiento")%></td>			
							<td width="3%" align="center">.</td>
						</tr>
					<%	while not rs.Eof
						flagAdmin = isAdmin(rs("IDDIVISION"))
						flagUser  = isUser(rs("IDDIVISION"))	%>			
						<tr class="reg_header_navdos">
							<td align="center"><%=rs("CDPEDIDO")%></td>						
							<td style="text-align: center; cursor:pointer;" ><img onclick="abrirPedido(<% =rs("IDPEDIDO") %>)" src="images/compras/PCT-16X16.png" title="Ver Ficha de Pedido"></td>
							<td align="center"><%=rs("NROPOLIZA")%></td>
							<td align="center"><%=rs("DSASEGURADORA")%></td>						
							<td align="center"><%=Trim(rs("DSEMPRESA"))%></td>						
							<td align="center"><%=getSimboloMoneda(rs("CDMONEDA")) & " " & GF_EDIT_DECIMALS(rs("IMPORTE"),2)%></td>
							<td align="center"><%=GF_FN2DTE(rs("VENCIMIENTO"))%></td>				
							<td align="center">
							<% if((flagAdmin)or(flagUser))then %>
								<IMG style="cursor:pointer;" title="<%=GF_TRADUCIR("Devolvel Poliza")%>" id="devolver_<%=rs("IDPDC")%>" src="images\almacenes\arrow_loan-16x16.png" onclick="devolverPDC(<%=rs("IDPDC")%>, this)">
							<% end if %>							
							</td>
						</tr>
						<% rs.MoveNext
						wend %>
					</table>	
				</td>			
			</tr>				
		</table>		
	<%	
	end if
End Function
'*************************************************
'************** COMIENZO DE LA PAGINA ************
'*************************************************
dim idPoliza, accion, idEstado, nroPoliza, idAseguradora, idTomador, myimporte, cdMoneda, importeAprox, fecha,cdPedido
dim idDivision
Call comprasControlAccesoCM(RES_PDC)
idPoliza = GF_PARAMETROS7("idPoliza", 0, 6)
accion	 = GF_PARAMETROS7("accion", "", 6)
idEstado = GF_PARAMETROS7("idEstado", 0, 6)
nroPoliza = GF_PARAMETROS7("nroPoliza", "", 6)
idAseguradora = GF_PARAMETROS7("idAseguradora", 0, 6)
idTomador = GF_PARAMETROS7("idTomador", 0, 6)
myimporte =  GF_PARAMETROS7("importe", 0, 6)
cdMoneda  = GF_PARAMETROS7("cdMoneda", "", 6)
importeAprox = GF_PARAMETROS7("importeAprox", "", 6)
fecha = GF_PARAMETROS7("fecha", "", 6)
cdPedido = GF_PARAMETROS7("cdPedido", "", 6)
idDivision = GF_PARAMETROS7("idDivision", 0, 6)

if(idPoliza > 0)then
	Call actualizarEstadoPDC(idPoliza,idEstado)
else
	Call loadPolizaVencida(nroPoliza, idAseguradora, idTomador, myimporte, cdMoneda, importeAprox, fecha, idDivision, cdPedido, idEstado)
end if	


Response.End

%>

