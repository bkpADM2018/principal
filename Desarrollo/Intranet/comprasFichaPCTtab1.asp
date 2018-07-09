<!--#include file="Includes/procedimientosAFE.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosMG.asp"-->	

<%
'Controla que se den las siguientes condiciones para que pueda aprobar la Apertura de Sobre       
'   - Solo si el estado es menor a Abierto
'   - Si ninguna cotizacion del pedido fue abierta, podra cambiar la confirmacion de la Apertura de Sobre
'   - Si el rol del usuario es Responsable de compras
Function puedeAprobarAperturaDeSobre(pIdPedido,pEstado)
    puedeAprobarAperturaDeSobre = false
    if (Cdbl(pEstado) <= ESTADO_PCT_ABIERTO) then
        if (Cdbl(getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)) =  FIRMA_ROL_GTE_COMPRAS) then
            if (verificarAperturaArchivo(pIdPedido)) then puedeAprobarAperturaDeSobre = true
        end if
    end if
end Function
'---------------------------------------------------------------------------------------------------------------------------------
Dim dsObra, cdObra, rsPCP, pctImage, pMensaje,accion

accion = GF_PARAMETROS7("accion",0,6)

%>
<script type="text/javascript">
    var chn = new channel();

    function abrirAFEs() {
        window.open('comprasPopUpAFE.asp?idObra=<%=pct_idObra %>&idPedido=<% =pct_idPedido %>', '_blank', 'location=no,menubar=no,statusbar=no,height=400,width=500,scrollbars=yes', false);
    }

    function reloadPage() {
        window.location.reload();
    }

    function abrirListaNDA(pIdPedido, pIdCotizacion) {
        window.open("comprasListaNDA.asp?IdPedido=" + pIdPedido + "&IdCotizacion=" + pIdCotizacion, '_blank', 'location=no,menubar=no,statusbar=no,height=400,width=500,scrollbars=yes', false);
    }
    function abrirPedido(id) {
        window.open("comprasPedidoCotizacion.asp?idPedido=" + id , "_blank", "resizable=yes,location=no,menubar=no,scrollbars=yes,scrolling=yes,height=500,width=700", false);
    }
    function aprobarAperturaSobre(e) {
        if (confirm("Desea generar cambios en la apertura de sobre?")) {
            var estado = "<%=ESTADO_BAJA%>";
            if (e.checked) estado = "<%=ESTADO_ACTIVO%>";
            ch.bind("comprasFichaPedidoCotizacion.asp?estado=" + estado + "&idPedido=<%=pct_idPedido%>", "aprobarAperturaSobre_callback(" + estado + ")");
            ch.send();
        }
        else {
            document.getElementById("aprobarApertura").checked = false;
        }
    }
    function aprobarAperturaSobre_callback(pEstado) {
        if (pEstado == "<%=ESTADO_ACTIVO%>") {
            document.getElementById("fechaApertura").innerHTML = '<%="Fecha: "& GF_FN2DTE(Session("MmtoDato")) %>';
            document.getElementById("fechaApertura").style.display = 'block';
            document.getElementById("usuarioApertura").innerHTML = '<%="Usuario: "& getUserDescription(Session("Usuario")) %>';
            document.getElementById("usuarioApertura").style.display = "block";
        }
        else {
            document.getElementById("fechaApertura").style.display = "none";
            document.getElementById("usuarioApertura").style.display = "none";
        }
    }
</script>
<%
pctImage = "PCT-64x64.png"
if (pct_idEstado = ESTADO_PCT_CANCELADO) then pctImage = "PCTR-48x48.png"
%>
<form id="myForm" action="comprasFichaPCTtab1.asp" method="get">
<table width="95%" align="center" border="1" bgcolor="f0f0f0" bordercolor="#999999">
    <tr>
	    <td>
		    <table width="100%" height="100">
			    <tr>
                    <td colspan="5">
                        <table width="100%">
                            <tr>
                                <td width="20%" align="left" rowspan="6">
                                    <img src="images/compras/<% =pctImage %>" alt="<% =GF_TRADUCIR("Ver pedido completo") %>" style="cursor:pointer" onclick="javascript:abrirPedido(<% =pct_idPedido %>)">
                                </td>
                                <td width="25%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top >
                                    <% =GF_TRADUCIR("Pedido") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_cdPedido %></td>
                            </tr>
                            <tr>
                                <td width="20%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top>
                                    <% =GF_TRADUCIR("Titulo") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_tituloPedido %></td>
                            </tr>
                            <tr>
                                <td width="20%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top>
                                    <% =GF_TRADUCIR("Fecha Emision") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_FechaInicio %></td>
                            </tr>
                            <tr>
                                <td width="20%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top>
                                    <% =GF_TRADUCIR("Fecha Cierre") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_FechaCierre %></td>
                            </tr>
                            <tr>
                                <td width="20%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top>
                                    <% =GF_TRADUCIR("Solicitante") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_dsSolicitante %></td>
                            </tr>
                            <tr>
                                <td width="20%" align="right" style="font-weight:bold; color:#000; line-height:15px" valign=top>
                                    <% =GF_TRADUCIR("Division") %>
                                </td>
                                <td width="5%">&nbsp;</td>
                                <td width="55%" align="left" style="color:#000; line-height:15px"><% =pct_dsDivision %></td>
                            </tr>
                        </table>
                    </td>
			        
                </tr>
                <tr>
					<td colspan="5"><hr></td>
				</tr>
				<tr valign="top">
					<td>&nbsp;</td>
					<td align="left" style="font-weight:bold; color:#000; line-height:15px"> <% =GF_TRADUCIR("Archivos") %> </td>
					<td align="right">&nbsp;</td>
					<td>&nbsp;</td>
					<td align="left" style="color:#000; line-height:15px" >&nbsp;</td>
				</tr>
                <tr valign="top">
					<td>&nbsp;</td>
					<td align="left" style="font-weight:bold; color:#000; line-height:15px">
						<div class="reg_header_navdos"><% =GF_TRADUCIR("Especificación técnica") %></div>
					</td>
					<% 	if (hayEspecifTecnica(pct_idPedido)) then	%>
                        <td align="left" valign="middle" colspan="3" style="color:#000; line-height:15px" >&nbsp;
                            <a href="comprasOpenArchivo.asp?idPedido=<% =pct_idPedido %>&fileno=<% =PCT_BINARY_SPECIFICATION %>" target="_blank">
                                <img align="absMiddle" src="images/word-16.png">
                                &nbsp;<% =buildFileName(pct_idPedido, PCT_BINARY_SPECIFICATION, "")%>
                                <img align="absMiddle" src="images/download_b-16x16.png" title="Descargar">
                            </a>
                        </td>
                    <% else %>    
                        <td align="left" colspan="3" valign="middle" style="color:#000; line-height:15px" >&nbsp; <% =GF_TRADUCIR("No hay archivo asociado") %>.</td>
                    <% end if %>
		        </tr>
                 <tr valign="top">
					<td>&nbsp;</td>
					<td align="left" style="font-weight:bold; color:#000; line-height:15px">
						<div class="reg_header_navdos"><% =GF_TRADUCIR("Condiciones Particulares") %></div>
					</td>
                    <%	if (hayCondParticulares(pct_idPedido)) then	%>
                        <td align="left" colspan="3" valign="middle" style="color:#000; line-height:15px" >&nbsp; 
                            <a href="comprasOpenArchivo.asp?idPedido=<% =pct_idPedido %>&fileno=<% =PCT_BINARY_CONDITIONS %>" target="_blank">
                                <img align="absMiddle" src="images/word-16.png">
                                &nbsp;<% =buildFileName(pct_idPedido, PCT_BINARY_CONDITIONS, "")%>
                                <img align="absMiddle" src="images/download_b-16x16.png" title="Descargar">
                            </a>
                        </td>
                    <% else %>
					    <td align="left" colspan="3" valign="middle" style="color:#000; line-height:15px" >&nbsp; <% =GF_TRADUCIR("No hay archivo asociado") %>.</td>
                    <% end if %>
		        </tr>
                <tr valign="top"> <td colspan="5"><hr></td> </tr>
                <% if (pct_dsPedido <> "") then %>
                <tr valign="top">
		            <td>&nbsp;</td>
		            <td align="left" style="font-weight:bold; color:#000; line-height:15px"><% =GF_TRADUCIR("Descripcion") %></td>
		            <td align="right">&nbsp;</td>
		            <td>&nbsp;</td>
		            <td align="left" style="color:#000; line-height:15px" >&nbsp;</td>
		        </tr>
                <tr valign="top">
		            <td>&nbsp;</td>
		            <td align="left" colspan="4" style="font-weight:bold; color:#2e6b4d; line-height:15px" ><%= GF_TRADUCIR(pct_dsPedido)%></td>
		        </tr>
                <% end if %>
	            <tr valign="top">
		            <td colspan="5"><hr /></td>
		        </tr>
                 <tr valign="top">
		            <td>&nbsp;</td>
		            <td align="left" style="font-weight:bold; color:#000; line-height:15px"><% =GF_TRADUCIR("Ptda. Presupuestaria") %></td>
		            <td align="right">&nbsp;</td>
		            <td>&nbsp;</td>
		            <td align="left" style="color:#000; line-height:15px" >&nbsp;</td>
		        </tr>
                <%	Set obra = obtenerDescripcionCompletaDetalle(pct_idObra, pct_idArea, pct_idDetalle) 
				    if (obra.eof) then %>
                        <tr valign="top">
		                    <td colspan="5"><% =GF_TRADUCIR("No se ha encontrado ninguna Partida asociada a este pedido.") %></td>
    		            </tr>
			    <%	else  %>
                    <tr valign="top">
		                <td>&nbsp;</td>
		                <td colspan="4" align="left" style="font-weight:bold; color:#2e6b4d; line-height:15px" >
                            <a href="comprasObras.asp?idObra=<% =pct_idObra %>" target="_blank">
                                <img src="images/compras/Obra_16x16.png" width="16" height="16" alt="En Obra" />                             
                                &nbsp; <% =obra("CDOBRA") %>-<% =GF_TRADUCIR(obra("DSOBRA")) %>
                            </a>
                        </td>
		            </tr>
                    <% if (not isnull(obra("IDAREA"))) then %>
		            <tr valign="top">
		                <td>&nbsp;</td>
		                <td colspan="4" valign="bottom" align="left" style="color:#2e6b4d; line-height:15px" >&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                            <img src="images/compras/Array_fake.png" width="16" height="16" alt="Array-fake" />
                            <a href="comprasObras.asp?idObra=<% =pct_idObra %>" target="_blank"><% = obra("IDAREA") & " - " & obra("DSAREA") %></a>
                        </td>
		            </tr>
                    <% if (not isnull(obra("IDDETALLE"))) then	
						  if (cint(obra("IDDETALLE")) <> 0) then	%>
		                    <tr valign="top">
		                        <td>&nbsp;</td>
		                        <td colspan="4" valign="bottom" align="left" style="color:#2e6b4d; line-height:15px" >&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; 
                                    <img src="images/compras/Array_fake.png" width="16" height="16" alt="Array-fake" />
                                    <a href="comprasObras.asp?idObra=<% =pct_idObra %>" target="_blank">
									    <% = obra("IDDETALLE") & " - " & obra("DSDETALLE") %>
								    </a>
                                </td>
		                    </tr>
                        <% end if %>
                     <% end if %>
                   <% end if %>
                <% end if %>
	            <tr valign="top"><td colspan="5"><hr></td></tr>
                <tr valign="top">
		            <td>&nbsp;</td>
		            <td align="left" colspan="2" style="font-weight:bold; color:#000; line-height:15px"><% =GF_TRADUCIR("Análisis Comparativo") %></td>		            
		            <td>&nbsp;</td>
		            <td align="left" style="font-weight:bold; color:#000; line-height:15px"><% =GF_TRADUCIR("AFEs") %></td>
		        </tr>
                <tr valign="top">
		            <td>&nbsp;</td>
		            <td colspan="2" align="left" style="color:#2e6b4d; line-height:15px" >
                    <%	if (pct_idEstado >= ESTADO_PCT_ABIERTO) then
				            if (puedeModificarPlanilla(pct_idPedido, pct_idEstado)) then			%>
                                <a href="comprasComparativoDeOfertas.asp?idPedido=<% =pct_idPedido %>" target="_blank"><img src="images/edit-16.png"> &nbsp; <% =GF_TRADUCIR("Editar el Analisis Comparativa") %><br>
                         <%	end if
    				        if (pct_idEstado >= ESTADO_PCT_EN_ANALISIS) then	%>
						        <a href="comprasComparativoDeOfertasPrint.asp?idPedido=<% =pct_idPedido %>" target="_blank"><img src="images/print-16.png">&nbsp;  <% =GF_TRADUCIR("Imprimir la Planilla Comparativa") %></a>						 
                        <%  end if
				        else    %>
						    <% =GF_TRADUCIR("No Disponible") %>
				    <%	end if	%>
                    </td>
		            <td>&nbsp;</td>
                    <td align="left" style="color:#2e6b4d; line-height:15px">
                       <span style="cursor:pointer;" onclick="abrirAFEs();"><img src="images/compras/afe-16x16.png" ><% =GF_TRADUCIR("Ver y trabajar con los AFE") %></span>
                    </td>
		        </tr>
                <!--**************************************************************************************-->
                <% if (pct_idEstado >= ESTADO_PCT_COTIZADO) then
                      Set sp_Div = executeProcedureDb(DBSITE_SQL_INTRA, rsApertura, "TBLPCTFIRMASAPERTURA_GET_BY_IDPEDIDO", pct_idPedido)
                      flagApertura= false
                      if (not rsApertura.Eof) then
                         fechaApertura = GF_FN2DTE(rsApertura("FECHAFIRMA"))
                         usuarioApertura = getUserDescription(rsApertura("CDUSUARIO"))
                         flagApertura= true
                      end if %>
                    <tr valign="top"><td colspan="5"><hr></td></tr>
                    <tr valign="top">
		                <td>&nbsp;</td>
		                <td align="left" style="font-weight:bold; color:#000; line-height:15px"><% =GF_TRADUCIR("Apertura de sobre") %></td>
		                <td align="right">&nbsp;</td>
		                <td>&nbsp;</td>
		                <td align="left" style="color:#000; line-height:15px" >&nbsp;</td>
		            </tr>
                    <tr valign="top">
		                <td>&nbsp;</td>
                        <td colspan="2">
                            <% =GF_TRADUCIR("Aprobado: ") %>
                            <% if (puedeAprobarAperturaDeSobre(pct_idPedido,pct_idEstado)) then %>
                                <input type="checkbox" id="aprobarApertura" name="aprobarApertura" <% if (flagApertura) then %> checked <% end if %> onclick="aprobarAperturaSobre(this)" />
                            <% else %>
                                <% if (flagApertura) then 
                                      Response.Write "SI"
                                   else
                                      Response.Write "NO"
                                   end if 
                               end if %>
                        </td>
    		        </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td colspan="2"><div id="fechaApertura" style="width:auto;float:left;"><% if (flagApertura) then Response.Write GF_TRADUCIR("Fecha: ") & fechaApertura end if%></div></td>
                        <td colspan="2"><div id="usuarioApertura" style="width:auto;float:left;"><% if (flagApertura) then Response.Write GF_TRADUCIR("Usuario: ") & usuarioApertura end if%></div></td>
                    </tr>                    
			    <% end if %>
            </table>
        </td>
     </tr>
</table>
</form>
