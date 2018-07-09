<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/MD5.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
Const RUTA_PLANILLA_COMPARATIVA = 1
const OBRA_NULA = 0

Dim tab, idPedido, c1, c2, c3, c4,myvarRuta
dim myProveedorSeleccionado, rsPCPCabe,rsPCPDet ,MystrSQL, checked,	myOAProDs, myOANroSo, errFirma, rol1, rol2, rol3, rol4, rolDireccion
dim ITnroLinea, ITproveedor, ITproveedorDS, firmaResponsable, firma1, firma2,firma3, firma4, rsFirmas, tienePoliza, flagBoss
dim v_1, v_2,memberDireccionCd, myFirmaTx, myFirmaCd, myFirmaDs, myFirmaRol, rolUsuario, mySecuencia, member1Sec, member2Sec, member3Sec, member4Sec
dim responsableSec, direccionSec, flagBossDireccion, flagBoss1, flagBoss2, flagBoss3, flagBoss4
Dim member1Cd, member2Cd, member3Cd, member4Cd

myvarRuta=RUTA_PLANILLA_COMPARATIVA

idObra = GF_PARAMETROS7("idObra",0,6)
idPedido = GF_PARAMETROS7("idPedido",0,6)
tab = GF_PARAMETROS7("tab",1,6)
checked = ""
Call initHeader(idPedido)

errFirma = GF_PARAMETROS7("errFirma","",6)
if (errFirma <> "") then setError(errFirma)
myProveedorSeleccionado = CLng(pct_idProveedorElegido)

'Determino si el usuario es jefe de sector que le corresponde a la planilla.
flagBoss = isBossOf(session("Usuario"), pct_idSector)

MystrSQL="SELECT * from TBLPCPCABECERA where IDPEDIDO=" & pct_idPedido
Call executeQueryDb(DBSITE_SQL_INTRA, rsPCPCabe, "OPEN", MystrSQL)

	if not rsPCPCabe.eof then
	        rolUsuario = CInt(getRolFirma(session("Usuario"), SEC_SYS_COMPRAS))
			myComentarios = rsPCPCabe("COMENTARIOS")
			Call executeProcedureDb(DBSITE_SQL_INTRA, rsFirmas, "TBLPCPFIRMAS_GET_BY_IDPEDIDO", pct_idPedido)
			if (not rsFirmas.eof) then
	            'Empiezo tomando una firma, si no es la ultima la agrego al primer lugar disponible.
	            'Cuando llegue la ultima, siempre se carga en el lugar de aprobacion general de la planilla fuera del bucle.
	            flagBoss1 = false
	            flagBoss2 = false
	            flagBoss3 = false
	            flagBoss4 = false
	            flagBossDireccion = false
	            myFirmaCd = rsFirmas("CDUSUARIO")
	            myFirmaDs = getUserDescription(rsFirmas("CDUSUARIO"))
	            myFirmaTx = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
	            myFirmaRol = CInt(rsFirmas("IDROL"))
	            mySecuencia = CInt(rsFirmas("SECUENCIA"))
	            rsFirmas.MoveNext()
	            while (not rsFirmas.eof)	    
	                'La coloco en el primer lugar disponible
	                if (responsableCd = "") then
	                    responsableCd = myFirmaCd
	                    firmaResponsable = myFirmaTx
	                    ITResponsable = myFirmaDs
	                    responsableSec = mySecuencia
	                else
	                    if (member1Cd = "") then
	                        member1Cd = myFirmaCd			
		                    member1 = myFirmaDs
		                    firma1 = myFirmaTx
		                    rol1 = myFirmaRol
		                    member1Sec = mySecuencia
		                    if ((flagBoss) and (rol1 = FIRMA_ROL_GTE_SECTOR)) then flagBoss1 = True
	                    else
	                        if (member2Cd = "") then
	                            member2Cd = myFirmaCd			
		                        member2 = myFirmaDs
		                        firma2 = myFirmaTx
		                        rol2 = myFirmaRol
		                        member2Sec = mySecuencia
		                        if ((flagBoss) and (rol2 = FIRMA_ROL_GTE_SECTOR)) then flagBoss2 = True
	                        else
	                            if (member3Cd = "") then
	                                member3Cd = myFirmaCd			
		                            member3 = myFirmaDs
		                            firma3 = myFirmaTx
		                            rol3 = myFirmaRol
		                            member3Sec = mySecuencia
		                            if ((flagBoss) and (rol3 = FIRMA_ROL_GTE_SECTOR)) then flagBoss3 = True
                                else
                                    if (member4Cd = "") then
	                                    member4Cd = myFirmaCd			
		                                member4 = myFirmaDs
		                                firma4 = myFirmaTx
		                                rol4 = myFirmaRol
		                                member4Sec = mySecuencia
		                                if ((flagBoss) and (rol4 = FIRMA_ROL_GTE_SECTOR)) then flagBoss4 = True
	                                end if		                            
	                            end if
                            end if
                        end if
                    end if	                    
	                'Tomo la proxima firma.	    
	                myFirmaCd = rsFirmas("CDUSUARIO")
	                myFirmaDs = getUserDescription(rsFirmas("CDUSUARIO"))
	                myFirmaTx = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))	    
	                myFirmaRol = CInt(rsFirmas("IDROL"))     
	                mySecuencia = CInt(rsFirmas("SECUENCIA"))       
                    rsFirmas.MoveNext()    	    
	            wend
	            'Siempre que salgo del bucle tengo listos los datos de la ultima firma para completar.
	            memberDireccionCd = myFirmaCd
		        firmaDireccion = myFirmaTx
		        memberDireccion = myFirmaDs
		        rolDireccion = myFirmaRol
		        direccionSec = mySecuencia		        
		        if ((flagBoss) and (rolDireccion = FIRMA_ROL_GTE_SECTOR)) then flagBossDireccion = True
	        end if				
	end if	



select case (tab)
	case 1:
		c1="tabbertabdefault"
	case 2:
		c2="tabbertabdefault"
	case 3:
		c3="tabbertabdefault"
	case 4:
		c4="tabbertabdefault"
	case 5:
		c5="tabbertabdefault"
end select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" type="text/css" href="css/main.css">
<link rel="stylesheet" type="text/css" href="css/toolbar.css">
<style type="text/css">
	.titleStyle {
		font-weight: bold;
		font-size: 20px;
	}

	.divOculto {
		display: none;
	}

	option.titulo {
	  font-weight: bold;
	}
	.bordeIframe{
		BORDER-BOTTOM: #F4B800 0px solid;
		BORDER-LEFT: #F4B800 0px solid;
		BORDER-TOP: #F4B800 0px solid;
		BORDER-RIGHT: #F4B800 0px solid;
		text-align: center;		
		-moz-border-radius:5px 5px 5px 5px
	}

	.ocultar{
		display: none;
	}

	.mostrar{
		display: block;
	}
</style>
<title>Sistema de Compras - Comparativo de Ofertas</title>
<script language="javascript" type="text/javascript" src="scripts/tabber.js"></script>
<script type="text/javascript" src="scripts/controles.js">		</script>
<script type="text/javascript" src="scripts/date.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript">
	var link = "comprasFirmarPlanilla.asp?idPedido=<% =pct_idPedido %>&secuencia=";	
	var hkeyR = new Hkey('hkR', link + "<%=responsableSec%>", '<% =HKEY() %>', 'check_callback()');	
	var hkey1 = new Hkey('hk1', link + "<%=member1Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link + "<%=member2Sec%>", '<% =HKEY() %>', 'check_callback()');
	var hkey3 = new Hkey('hk3', link + "<%=member3Sec%>", '<% =HKEY() %>', 'check_callback()');	
	var hkey4 = new Hkey('hk4', link + "<%=member4Sec%>", '<% =HKEY() %>', 'check_callback()');	
	var hkeyD = new Hkey('hkD', link + "<%=direccionSec%>", '<% =HKEY() %>', 'check_callback()');	
	
	
		
	var ch = new channel();	
	
	function bodyOnLoad(){			
		hkeyR.start();	
		hkey1.start();
		hkey2.start();
		hkey3.start();
		hkey4.start();
		hkeyD.start();
		if(document.getElementById("My_IdObra").value != <%=OBRA_NULA%>){		
			document.getElementById("Loading").innerHTML="<table width='100%' height='100%' class='reg_header' cellspacing=2 cellpadding=1><tr><td colspan=4 align='center'><img src='images/loading_blocks_green.gif'></td></tr></table>";		
			document.getElementById("TableroObradetalle").src = "comprasTableroObraDetalle.asp?idobra=<%=idObra%>&ruta=<%=RUTA_PLANILLA_COMPARATIVA%>&areaDetalle=<%= pct_idArea & "," & pct_idDetalle%>"; 	
		}		
	}
		
	function check_callback(resp) {				
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}	
	
	function iFrameOnLoad()
	{
		document.getElementById('Loading').className = "ocultar";
	}	
	function VerCotizacion(idPedido){		
		window.open("comprasFichaPedidoCotizacion.asp?idPedido=" + idPedido + "&tab=2", "_blank", "location=no,scrollbars=yes,menubar=no,statusbar=no,height=500,width=500",false);
	}
	window.onload = bodyOnLoad;
</script>
<script language="javascript" src="scripts/tabber.js"></script>
</head>
<body>
<form method="post" id="frmSel" action="comprasComparativoDeOfertasFirma.asp?idPedido=<%=idPedido%>">
<input type="hidden" id="My_IdObra" name="My_IdObra" value="<%=idObra%>">
<input type="hidden" id="IdPedido_old" name="IdPedido_old" value="<%=pct_idPedido%>">
	<div class="tabber">
		<!--********************************T A B   1 *********************************************************************-->
		<div class="tabbertab <% =c1 %>" title="<% =GF_TRADUCIR("Planilla")%>">

			<table width="100%">
				<tr>
					<td align="center">
						<font class="big">Analisis Comparativo de Ofertas</font>
					</td>			
				</tr>
			</table>
			<div class="tableaside size100"> <% call showErrors() %></div>
			<table class="datagrid" width="95%" align="center">
                <thead>
                    <tr>
                        <th width="25%" align="center"> <% =GF_TRADUCIR("PUERTO") %></td>
                        <th width="25%" align="center"> <% =GF_TRADUCIR("FECHA DE CONCURSO") %></td>
                        <th width="25%" align="center"> <% =GF_TRADUCIR("PEDIDO ") %></td>
                        <th width="25%" align="center"> <% =GF_TRADUCIR("OBRA / TRABAJO") %> </td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
					    <td align="center"><%
						    if ((pct_idObra > 0) and (pct_idObra <> OBRA_GEID)) then		
							    myPuerto = getDivisionObra(pct_idObra)	
						    else
							    myPuerto = pct_dsDivision
						    end if	
						    Response.Write myPuerto	%>
					    </td>
					    <td align="center">Inicio: <%=pct_FechaInicio%> - Cierre: <%=pct_FechaCierre%></td>
					    <td align="center"><%
						    if isnull(pct_cdPedido) then
							    Response.write GF_TRADUCIR("Sin Pedido")
						    else
							    response.write pct_tituloPedido & " (" & pct_cdPedido & ")"
						    end if %>
					    </td>
                        <td>
                       <%   if pct_idObra = 0 then
							    Response.write GF_TRADUCIR("Sin Obra")
                            else							    
                                if pct_idObra = OBRA_GEID then
                                    Response.write OBRA_GECD & "-" & OBRA_GEDS
						        else
							        Set obra = obtenerDescripcionCompletaDetalle(pct_idObra,pct_idArea,pct_idDetalle)
							        strTextoObra = obra("CDOBRA") & "-" & obra("DSOBRA")
							        if (not isNull(obra("IDAREA"))) then
								        strTextoObra = strTextoObra & "<br/> (" & pct_idArea
								        if (not isNull(obra("IDDETALLE"))) then
									        strTextoObra = strTextoObra & "-" & pct_idDetalle & ":" & obra("DSDETALLE")							
								        end if
								        strTextoObra = strTextoObra & ")"
							        end if
							        Response.write strTextoObra	
							    end if
						    end if	 %>
                        </td>
				    </tr>
				</tbody>
			</table>	
		    <br>
            <table class="datagrid" width="95%" align="center">
              <thead>
                  <tr>
                    <th align="center">NºSOBRE</th>
                    <th align="center">PROVEEDOR</th>
                    <th align="center">CARACTERISTICAS</th>
                    <th align="center">PRECIOS</th>
                    <th align="center">CONDICIONES DE PAGO</th>
                    <th align="center">FECHA DE ENTREGA</th>
                    <th align="center">OBSERVACIONES</th>
                    <th align="center">.</th>
                  </tr>
              </thead>
                <tbody>
                    <tr style="display: none;"><td colspan="7"></td></tr>
                <%	    MystrSQL="SELECT * from TBLPCPDETALLE where IDPEDIDO=" & idPedido & " order by NROSOBRE"
						'Response.write strsql
						'Response.end
						Call executeQueryDb(DBSITE_SQL_INTRA, rsPCPDet, "OPEN", MystrSQL)
						while not rsPCPDet.eof
							myClass = ""
							ITproveedor = rsPCPDet("IDPROVEEDOR")
							ITproveedorDS = getDescripcionProveedor(ITproveedor)
							ITcaracteristica = rsPCPDet("CARACTERISTICAS")
							ITimporte = rsPCPDet("IMPORTE")
							ITimporte = Replace(ITimporte,",",".")/100
							ITmoneda = rsPCPDet("CDMONEDA")
							ITcondPago = rsPCPDet("CONDPAGO")
							if ITcondPago = "" then ITcondPago = "No Especificada"
							ITfecEntrega = rsPCPDet("FECENTREGA")
							if len(ITfecEntrega) < 2 then 
								ITfecEntrega = "No Especificada"			
							else
								ITfecEntrega = GF_FN2DTE(ITfecEntrega)
							end if	
							ITnroLinea = rsPCPDet("NROSOBRE")										
							if myProveedorSeleccionado = CLng(rsPCPDet("IDPROVEEDOR")) then								
								myClass = "background:#fef3ab;"
								checked = "Checked"
								myOAProDs = ITproveedorDS
								myOANroSo = ITnroLinea
							end if	
							Set rsAux = getCotizaciones(idPedido, ITproveedor)		
							pct_hayCotizacion = false
							if (not rsAux.eof) then 
								pct_hayCotizacion = true
								pct_pathCotizacion = rsAux("PATHCOTIZACION")
							end if				
							
							if not pct_hayCotizacion or cstr(pct_pathCotizacion) = "NO_COTIZA" then
								myNoCotizaState = "Disabled"
								myNoCotizaValue = "No Cotiza"
								ITcaracteristica = "No Cotiza"
								ITimporte = 0
								ITcondPago = "&nbsp;"
								ITfecEntrega = "&nbsp;"
							else
								myNoCotizaState = ""
								myNoCotizaValue = ""
							end if	%>
                        <tr style="<%=myClass%>">
						    <td align="center"><%=ITnroLinea%></td>
							<td><%=ITproveedorDS%></td>
							<td align="left"><%=ITcaracteristica%></td>
							<td align="right"><%
									if ITcaracteristica = "No Cotiza" then
										Response.write "&nbsp;"
									else
										if(ITmoneda = MONEDA_DOLAR) then 
											response.write getSimboloMoneda(MONEDA_DOLAR) & "&nbsp;"
										else
											Response.Write getSimboloMoneda(MONEDA_PESO) & "&nbsp;"	
										end if
										Response.write GF_EDIT_DECIMALS((ITimporte*100),2) 
									end if %>
							</td>
							<td align="center">
							<%	if ITcondPago = "" then
								    Response.Write "&nbsp;"
								else
								    Response.Write ITcondPago
								end if	 %>
							</td>
							<td align="center"> <%=ITfecEntrega%> </td>
							<td>
			                    <% 'Verifico si alguna de las coptizacion del proveedor fue presentada fuera del palzo.
			                        Set rsCotizaciones = getCotizaciones(pct_idPedido, ITproveedor)
			                        blnFuera = false
			                        while ((not rsCotizaciones.eof) and (not blnFuera))
			                            if  (GF_DTEDIFF(rsCotizaciones("FECHAPRESENTACION"), GF_DTE2FN(pct_FechaCierre), "D") < 0) then 
			                    %>
			                            <label class="reg_header_error round_border_all" title="Cotizacion cargada fuera de termino" style="cursor:pointer; padding:2px;">
			                                Cotizacion cargada fuera de termino
			                            </label>                			            
			                    <%          blnFuera = true
			                            end if 
			                            rsCotizaciones.MoveNExt()
			                        wend
			                    %> 
			                </td>
							<td align="center">
							    <a href="javascript:VerCotizacion('<%=pct_idPedido%>')">
								    <img align="center" id="Ver_Pedido" title="Ver Pedido"  src="images/compras/PCT-16x16.png">		
								</a>
							</td>															
						</tr>
					<%	rsPCPDet.movenext
						wend %>		
                </tbody>
            </table>    
			<br>
			<table width="95%" align="center">
                <tr>
                    <td valign="top" width="70%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th style="border-radius: 8px 8px 0 0"  align="center">COMENTARIOS / SUGERENCIAS T&EacuteCNICAS</td></tr>
                            </thead>
                            <tbody>
                                <tr><td valign="top" align="left"><%=myComentarios%></td></tr>
                            </tbody>
                        </table>
                    </TD>                    
                </tr>
            </table>
			<br>
			<table width="95%" align="center">
                <tr>
                     <td valign="top" width="33%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th colspan="2" style="border-radius: 8px 8px 0 0"  align="center">OFERTA ADJUDICADA</td></tr>
                                <tr><th style="border-radius: 0 0 0 0" align="center">PROVEEDOR</td>
                                    <th style="border-radius: 0 0 0 0" align="center">NRO. SOBRE</td></tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td align="left"><%=myOAProDs%></td>
								    <td align="center"><%=myOANroSo%></td>
                                </tr>
                            </tbody>
                        </table>
                    </TD>
                    <td valign="top" width="33%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th colspan="2" style="border-radius: 8px 8px 0 0"  align="center">P&OacuteLIZA DE CAUCI&OacuteN</td></tr>
                            </thead>
                            <tbody>
                            <%	auxPendiente = 0
				                auxConfirmado = 0
				                auxTotal = 0
				                tienePoliza = "No"
				                strSQL = " SELECT CDMONEDA,IMPORTE, ESTADO  FROM TBLPOLIZASCAUCION WHERE IDPEDIDO = " & idPedido & " AND ESTADO <> " & ESTADO_PDC_ANULADA
				                Call executeQueryDb(DBSITE_SQL_INTRA, rsPDC, "OPEN", strSQL)
				
				                while not rsPDC.eof 
				                    tienePoliza = "Si"
					                auxMoneda = rsPDC("CDMONEDA")
					                if(rsPDC("ESTADO") = ESTADO_PDC_PENDIENTE)then auxPendiente = auxPendiente + Cdbl(rsPDC("IMPORTE"))
					                if(rsPDC("ESTADO") > ESTADO_PDC_PENDIENTE)then auxConfirmado = auxConfirmado + Cdbl(rsPDC("IMPORTE"))
						                 %>
				
				                <%  rsPDC.MoveNext
					             wend   
					             auxTotal = auxConfirmado + auxPendiente %>	
                                <tr>
								    <td width=30%>TIENE P&OacuteLIZA:</td>
								    <td width=70%><%=GF_Traducir(tienePoliza)%></td>
							    </tr>
							    <tr>
								    <td width=30%>PENDIENTE</td>
								    <td width=70%><%= getSimboloMoneda(auxMoneda) & " " & GF_EDIT_DECIMALS(auxPendiente,2)%></td>
							    </tr>
							    <tr>
								    <td width=30%>CONFIRMADO</td>
								    <td width=70%><%= getSimboloMoneda(auxMoneda) & " " & GF_EDIT_DECIMALS(auxConfirmado,2)%></td>
							    </tr>
							    <tr>
								    <td width=30%>TOTAL</td>
								    <td width=70%><%= getSimboloMoneda(auxMoneda) & " " & GF_EDIT_DECIMALS(auxTotal,2)%></td>
							    </tr>
                            </tbody>
                        </table>
                    </TD>
                    <td valign="top" width="33%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th style="border-radius: 8px 8px 0 0"  align="center">RESP. T&EacuteCNICO</td></tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td align="center">
					                <% 	if (firmaResponsable = "") then
										'Todavia no firmo							                
											if (session("Usuario") = responsableCd) then	%>
												<br><div id="hkR"></div><br>
										<%	else	%>
												<br><br><br>
										<%	end if	
                                    	else 
                                            'Ya se firmo.                    		
                                            Response.Write "<img src='images/firmas/" & obtenerFirma(responsableCd) & "'>"
						                    Response.Write firmaResponsable
										end if %>
					                <b><% =ITResponsable %></b>
					                </td>
                                </tr>
                            </tbody>
                        </table>
                    </TD>
                </tr>
            </table>
            <br/>            
            <table width="95%" align="center">
                <tr>
                    <td valign="top" rowspan="2" width="50%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th colspan="2" style="border-radius: 8px 8px 0 0"  align="center">MIEMBROS DEL COMIT&Eacute DE ADJUDICACI&OacuteN</td></tr>
                                <tr><th style="border-radius: 0 0 0 0" width="50%" align="center">NOMBRE</td>
                                    <th style="border-radius: 0 0 0 0" width="50%" align="center">FIRMA</td></tr>
                            </thead>
                            <tbody>
                                <%  if (member1Cd <> "") then %>
                               <tr>	
								    <td align="left" valign="center">&nbsp;<b><%=member1%></b>&nbsp;</td>
								    <td align="center">						
								    <%	if (firma1 = "") then								        
										    if ((session("Usuario") = member1Cd) or ((rolUsuario = rol1) and (rolUsuario <> FIRMA_ROL_GTE_SECTOR)) or (flagBoss1)) then	%>
											    <br><div id="hk1"></div><br>
									    <%	else	%>
										    <br><br><br>
									    <%	end if	
									    else
										    Response.Write "<img src='images/firmas/" & obtenerFirma(member1Cd) & "'>"
										    Response.Write firma1
									    end if
								    %>
								    </td>
							    </tr>
							    <%  end if
							        if (member2Cd <> "") then %>
							    <tr>	
								    <td align="left" valign="center">&nbsp;<b><%=member2%></b>&nbsp;</td>
								    <td align="center">
								    <%	if (firma2 = "") then							
										    if ((session("Usuario") = member2Cd) or ((rolUsuario = rol2) and (rolUsuario <> FIRMA_ROL_GTE_SECTOR)) or (flagBoss2)) then	%>
											    <br><div id="hk2"></div><br>
									    <%	else	%>
										    <br><br><br>
									    <%	end if	
									    else
										    Response.Write "<img src='images/firmas/" & obtenerFirma(member2Cd) & "'>"
										    Response.Write firma2
									    end if
								    %>
								    </td>
							    </tr>
							    <%  end if
							        if (member3Cd <> "") then %>
							    <tr>	
								    <td align="left" valign="center">&nbsp;<b><%=member3%></b>&nbsp;</td>
								    <td align="center">
								    <%	if (firma3 = "") then
										    if ((session("Usuario") = member3Cd) or ((rolUsuario = rol3) and (rolUsuario <> FIRMA_ROL_GTE_SECTOR)) or (flagBoss3)) then	%>
											    <br><div id="hk3"></div><br>
									    <%	else	%>
										    <br><br><br>
									    <%	end if
									    else
										    Response.Write "<img src='images/firmas/" & obtenerFirma(member3Cd) & "'>"
										    Response.Write firma3
									    end if
								    %>
								    </td>
							    </tr>
							    <%  end if
							        if (member4Cd <> "") then %>
							    <tr>	
								    <td align="left" valign="center">&nbsp;<b><%=member4%></b>&nbsp;</td>
								    <td align="center">
								    <%	if (firma4 = "") then								         
										    if ((session("Usuario") = member4Cd) or ((rolUsuario = rol4) and (rolUsuario <> FIRMA_ROL_GTE_SECTOR)) or (flagBoss4)) then	%>
											    <br><div id="hk4"></div><br>
									    <%	else	%>
										    <br><br><br>
									    <%	end if
									    else
										    Response.Write "<img src='images/firmas/" & obtenerFirma(member4Cd) & "'>"
										    Response.Write firma4
									    end if
								    %>
								    </td>
							    </tr>
							    <% end if %>
                            </tbody>
                        </table>
                    </TD>
                    <td valign="top" width="50%">
                        <table class="datagrid" width="100%" align="center">
                            <thead>
                                <tr><th colspan="2" style="border-radius: 8px 8px 0 0"  align="center">APROBACI&OacuteN FINAL DEL PEDIDO</td></tr>
                                <tr><th style="border-radius: 0 0 0 0" width="50%" align="center">NOMBRE</td>
                                    <th style="border-radius: 0 0 0 0" width="50%" align="center">FIRMA</td></tr>
                            </thead>
                            <tbody>
                                 <tr>
								    <td align="left" valign="center" width="50%">&nbsp;
									    <b><% =memberDireccion %></b>&nbsp;
								    </td>
								    <td align="center" width="50%">
								    <%	if (firmaDireccion = "") then										
										    'Si todavia no fue autorizado por la Direccion/Coordinador de Puertos, verifico que el usuario tenga el Rol adecuado para hacerlo
										    if (rolUsuario = rolDireccion) then
								    %>										
											    <br><div id="hkD"></div><br>
									    <%	else %>
										    <br><br><br>
									    <%	end if
									    else																				
										    Response.Write "<img src='images/firmas/" & obtenerFirma(memberDireccionCd) & "'>"
										    Response.Write firmaDireccion
									    end if
								    %>
								    </td>							
							    </tr>
                            </tbody>
                        </table>
                    </TD>
                </tr>
		   </TABLE>           
			<input type="hidden" name="errFirma" id="errFirma">	
		</div>				
		<!--***************************************T A B   2**************************************************************-->		
		<%
		if(idObra = OBRA_NULA)then
			v_1 = c2
			v_2 = c3
			'variables para que cuando no tenga Obra ponga como TAB2 a comprasFichaPCTtab1 y TAB3 a comprasListaAFE
			'dejando de mostrar el TAB de la pagina comprasTableroObraDetalle
		else
			v_1 = c3
			v_2 = c4
			%>
			<div id="TabPartida" class="tabbertab <% =c2 %>" title="<% =GF_TRADUCIR("Partida")%>">		
				<div id="Loading"></div>
				<iframe width="100%" height="500px" class="bordeIframe mostrar" id="TableroObradetalle" name="TableroObradetalle"  ></iframe>						
			</div>	
		<%end if%>
			<!--***************************************T A B   3**************************************************************-->
			<div class="tabbertab <% =v_1 %>" title="<% =GF_TRADUCIR("Pedido")%>"><!--#include file="comprasFichaPCTtab1.asp"--></div>
			<!--***************************************T A B   4**************************************************************-->
			<div class="tabbertab <% =v_2 %>" title="<% =GF_TRADUCIR("Afes")%>"><!--#include file="comprasListaAFE.asp"--></div>
			<!--*****************************************************************************************************-->		
	</div>
</form>
</body>