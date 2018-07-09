<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->

<%
    Dim pImporte, pMoneda, pcpType
    Dim dsRolGteSector, dsRolGteCompras, dsRolCoordPuertos, dsRolDireccion
    
    pImporte = GF_PARAMETROS7("importe", 2, 6)
    pMoneda = GF_PARAMETROS7("moneda", "", 6)
    
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_GTE_SECTOR)
    if (not rs.eof) then dsRolGteSector = rs("DSROL")
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_GTE_COMPRAS)
    if (not rs.eof) then dsRolGteCompras = rs("DSROL")
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLROLES_GET_BY_IDROL", FIRMA_ROL_DIRECTOR)
    if (not rs.eof) then dsRolDireccion = rs("DSROL")
     
    pcpType = getPCPAuthorizationType(pImporte, pMoneda)    
%>
    
    <table class="datagrid" width="95%" align="right">
   	    <thead>
            <tr>
          	    <th colspan="3" align="center" style="border-radius: 8px 8px 0 0"> RESPONSABLE T&EacuteCNICO </th>
        	</tr>
    	</thead>
    	<tbody>
    	    <tr>
                <td>&nbsp;</td>
           		<td colspan="3">
                    <div id="responsable"></div>
                </td>
			</tr>
        </tbody>
        <thead>
    	    <tr>
			    <th colspan="3" style="border-radius: 0 0 0 0">MIEMBROS DEL COMIT&Eacute DE ADJUDICACI&OacuteN</th>
  			</tr>
        </thead>
        <tbody>
    	    <tr>
			    <td align="right"> Firma 1 </td>
				<td>&nbsp;&nbsp;
                    <div id="member1"></div>(Opcional)
                </td>
			</tr>
            <tr>
			    <td align="right"> Firma 2 </td>
				<td>&nbsp;&nbsp;
                <%  =dsRolGteSector %>
                </td>
			</tr>
            <tr>
			    <td align="right"> Firma 3 </td>
				<td>&nbsp;&nbsp;
                <%  if (pcpType = PCP_TYPE_PURCHASE_LARGE) then 
                        response.Write dsRolGteCompras
                    end if  %>
                </td>
			</tr>
            <tr>
			    <td align="right"> Aprobaci&oacute;n </td>
				<td>&nbsp;&nbsp;
                <%  if (pcpType <> PCP_TYPE_PURCHASE_LARGE) then 
                        response.Write dsRolGteCompras
                    else  
                        response.write dsRolDireccion                            
                    end if  %>
                </td>
			</tr>
        </tbody>
	</table>