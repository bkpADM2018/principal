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
VS_secBudget = 0
VS_idPedido = GF_PARAMETROS7("idPedido",0 ,6)
VS_cdSolicitante = GF_PARAMETROS7("cdSolicitante",0 ,6)
VS_saldo = GF_PARAMETROS7("saldo",0 ,6)
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
	refPopUpArtOut = startIWin('popupArtRec');
	<% if (accion = ACCION_CERRAR) then %>
		refPopUpArtOut.hide();
	<% end if %>			
}
function grabarDetalleVale(idPedido, cdVale, idAlmacen, idObra, cdSolicitante, idArticulo, saldo){
	document.getElementById("cdVale").value = cdVale;
	document.getElementById("idPedido").value = idPedido;
	document.getElementById("idAlmacen").value = idAlmacen;
	document.getElementById("idObra").value = idObra;
	document.getElementById("cdSolicitante").value = cdSolicitante;
	document.getElementById("idArticulo").value = idArticulo;
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
<form name="frmSel" method="post" action="almacenArtRec.asp">
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
				strSQL = "select * from " & _
						"	( " & _
						"	select PMC.IDALMACEN, PMC.IDOBRA, PMC.IDPEDIDO, PMC.CDSOLICITANTE, PMC.FECHASOLICITUD ,PMC.IDALMACENDEST, ART.IDARTICULO, TG.PartidaPendiente as PARTPEND, TG.IdArticulo as NewIdArticulo, TG.SaldoVMT, TG.SaldoVMR, " & _
						"	    case when TG.SaldoVMR IS NULL then TG.SaldoVMT else TG.SaldoVMT- TG.SaldoVMR " & _
						"	        end as SaldoTotal " & _
						"	            from " & _
						"	            ( " & _
						"	            select T1.PartidaPendiente, T1.IdArticulo, sum(CantVMT) as SaldoVMT,  sum(CantVMR) as SaldoVMR " & _
						"	                from " & _
						"	                    ( " & _
						"	                    select VC1.PartidaPendiente, VD1.IdARticulo, sum(VD1.Cantidad) as CantVMT from  " & _
						"	                         TBLVALESCABECERA VC1 inner join TBLVALESDETALLE VD1  " & _
						"	                            on VC1.idVale=VD1.idVale where VC1.cdVale='" & CODIGO_VS_TRANSFERENCIA & "' and VC1.ESTADO=" & ESTADO_ACTIVO & _
						"	                                group by VC1.PartidaPendiente, VD1.IdArticulo " & _
						"	                    ) T1 " & _
						"	                    left join " & _
						"	                    ( " & _
						"	                    select  VC2.PartidaPendiente, VD2.IdArticulo, sum(VD2.Cantidad) as CantVMR from  " & _
						"	                        TBLVALESCABECERA VC2 inner join TBLVALESDETALLE VD2  " & _
						"	                            on VC2.idVale=VD2.idVale where VC2.cdVale='" & CODIGO_VS_RECEPCION & "' and VC2.ESTADO=" & ESTADO_ACTIVO & _
						"			                         group by VC2.PartidaPendiente, VD2.IdArticulo " & _
						"	                    ) T2 " & _
						"	                    on T1.PartidaPendiente = T2.PartidaPendiente and T1.IdArticulo = T2.IdArticulo " & _
						"	                group by T1.PartidaPendiente, T1.IdArticulo " & _
						"	            ) TG " & _
						"	            inner join " & _
						"	                 TBLPMCABECERA PMC  " & _
						"	                    on TG.PARTIDAPENDIENTE=PMC.IDPEDIDO " & _
						"	            inner join  " & _
						"	                 TBLPMDETALLE PMD " & _
						"	                        on PMC.IDPEDIDO=PMD.IDPEDIDO and PMD.IDARTICULO=TG.IDARTICULO " & _
						"	            inner join  " & _
						"	                 TBLARTICULOS ART " & _
						"	                        on ART.IDARTICULO=TG.IDARTICULO " & _
						"	)Tab WHERE (Tab.IDALMACEN = " & VS_idAlmacen & " OR Tab.IDALMACENDEST=" & VS_idAlmacen & ") " & _
						" and TAB.NewIdArticulo = " & VS_idArticulo & " and Tab.SaldoTotal>0 order by IDARTICULO asc"
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
								<input size="5" style="text-align:right;" type="text" name="saldo_<%=cont%>" id="saldo_<%=cont%>" value="<%=GF_EDIT_DECIMALS(clng(rs("SALDOTOTAL")),0)%>">
							</td>
							<td align="center" rowspan>
								<img title='Recibir' onclick=grabarDetalleVale(<%=rs("IDPEDIDO")%>,'<%=CODIGO_VS_RECEPCION%>',<%=rs("IDALMACEN")%>,<%=rs("IDOBRA")%>,'<%=rs("CDSOLICITANTE")%>',<%=VS_idArticulo%>,'saldo_<%=cont%>') style='cursor:pointer;' src='images/almacenes/arrow_reception-16x16.png'>
							</td>
						</tr>
						<%	
						cont = cont + 1
						rs.movenext
					wend	
					call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
					%>
	</table>

	<input type="hidden" name="cdVale" id="cdVale">
	<input type="hidden" name="idPedido" id="idPedido">
	<input type="hidden" name="idAlmacen" id="idAlmacen">
	<input type="hidden" name="idObra" id="idObra">
	<input type="hidden" name="cdSolicitante" id="cdSolicitante">
	<input type="hidden" name="idArticulo" id="idArticulo">
	<input type="hidden" name="saldo" id="saldo">
</form>		
</body>
</html>