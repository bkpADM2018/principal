<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->

<%
'-----------------------------------------------------------------------------------------------
    Dim pImporte, pMoneda, picType, pIdPedido, pIdProveedor, dsRolGteSector, dsRolGtePuerto, dsRolGte
    Dim dsRolGteCompras, dsRolCoordPuertos, dsRolDireccion, pDivision, pIdContrato, esExpo, dsRolController, dsRolCtrlGral
    Dim pAccion, pCdSolicitante, rolFirmaSolici
    
    pIdPedido = GF_PARAMETROS7("pedido", 0, 6)
    pIdProveedor = GF_PARAMETROS7("prov", 0, 6)
    pImporte = GF_PARAMETROS7("importe", 2, 6)
    pMoneda = GF_PARAMETROS7("moneda", "", 6)
    pDivision = GF_PARAMETROS7("division", 0, 6)
    pIdContrato = GF_PARAMETROS7("cto", 0, 6)        
    pCdSolicitante = GF_PARAMETROS7("solicitante", "", 6)
    pCdAutorizante = GF_PARAMETROS7("autorizante", "", 6)
    call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_GTE_COMPRAS)
    if (not rs.eof) then dsRolGteCompras = rs("DSROL")
    call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_SUP_PUERTO)    
    if (not rs.eof) then dsRolCoordPuertos = rs("DSROL")
    call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_CONTROLLER)    
    if (not rs.eof) then dsRolController = rs("DSROL")
    call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_DIRECTOR)
    if (not rs.eof) then dsRolDireccion = rs("DSROL")
    
    rolFirmaSolici = getRolFirma(pCdSolicitante, SEC_SYS_COMPRAS)
    
    esExpo = (pDivision = getDivisionID(CODIGO_EXPORTACION))
    'Segun la division se espera el Gte del sector o bien el Gte del puerto.
    if (esExpo) then    
    '    dsRolGte = dsRolGteSector
        dsRolCtrlGral = dsRolController
    else
    '    dsRolGte = dsRolGtePuerto
        dsRolCtrlGral = dsRolCoordPuertos
    end if
    
    picType = getPICAuthorizationType(pIdPedido, pIdContrato, pIdProveedor, pImporte, pMoneda)        
%>        					        
	<table class="datagrid" width="95%" align="right">
   	    <tbody>
    	    <tr>
			    <td align="right"> Solicitante </td>
				<td>&nbsp;&nbsp;
                    <div id="member1"></div>
                </td>
			</tr>
<%
            'Si es SMALL no se coloca al gerente ya que se incorpora como firma final de aprobacion.
            if (picType <> PIC_TYPE_PURCHASE_SMALL) then    
                'Si el solicitante es gerente de compras y la compra es MEDIUM, el autorizante opera directamente como Aprobador de la compra.
                if ((rolFirmaSolici <> FIRMA_ROL_GTE_COMPRAS) or (picType <> PIC_TYPE_PURCHASE_MEDIUM)) then
%>
            <tr>
			    <td align="right"> Autorizante </td>
				<td><% Call dibujarComboGte(pCdSolicitante, pCdAutorizante) %></td>
			</tr>			    
<%              end if
            end if 
            
            if ((picType = PIC_TYPE_PURCHASE_X_MEDIUM) or (picType = PIC_TYPE_PURCHASE_LARGE)) then    
                if (rolFirmaSolici <> FIRMA_ROL_GTE_COMPRAS) then
%>                
                    <tr>
			            <td align="right"> Control Proceso </td>
				        <td>&nbsp;&nbsp;<%  =dsRolGteCompras %></td>
			        </tr>			
    <%          end if 
                if (picType = PIC_TYPE_PURCHASE_LARGE) then %>
                    <tr>
			            <td align="right"> Control Gral. </td>
				        <td>&nbsp;&nbsp;<%  =dsRolCtrlGral %></td>
			        </tr>
    <%              end if  
            end if%>			

            <tr>
			    <td align="right"> Aprobaci&oacute;n </td>
				<td>
                <%  if (picType = PIC_TYPE_PURCHASE_SMALL) then 
                        Call dibujarComboGte(pCdSolicitante, pCdAutorizante)
                    else if (picType = PIC_TYPE_PURCHASE_MEDIUM) then 
                            'PAra este tipo de compras, si el solicitante ya era gte. de compras, entonces el autorizante ya ejerce de aprobador directamente.
                            if (rolFirmaSolici = FIRMA_ROL_GTE_COMPRAS) then
                                Call dibujarComboGte(pCdSolicitante, pCdAutorizante)
                            else
                                response.Write "&nbsp;&nbsp;" & dsRolGteCompras
                            end if    
                         else   if (picType = PIC_TYPE_PURCHASE_X_MEDIUM) then 
                                    response.Write "&nbsp;&nbsp;" & dsRolCtrlGral
                                else
                                    response.Write "&nbsp;&nbsp;" & dsRolDireccion
                                end if                                    
                         end if
                    end if  %>
                </td>
			</tr>
        </tbody>
	</table>
	
	
	