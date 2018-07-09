<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosfechas.asp"-->
<!--#include file="include/procedimientoProducto.asp"-->
<%
'-------------------------------------------------------------------------------------------------------------------
Function DibujarBiotecnologiaDeProducto(pCdProducto)
	Dim index 
    index  = 0
    strSQL = "SELECT IDBIOTECNOLOGIA, " &_
             "       DSBIOTECNOLOGIA, " &_
             "		 CASE WHEN B.CDCLIENTE IS NOT NULL THEN B.CDCLIENTE ELSE 0 END AS CDCLIENTE,"&_
			 "		 CASE WHEN B.DSCLIENTE IS NOT NULL THEN B.DSCLIENTE ELSE '' END AS DSCLIENTE,"&_
			 "		 CASE WHEN A.IDPRODUCTO IS NOT NULL THEN A.IDPRODUCTO ELSE 0 END AS IDPRODUCTO,"&_	
			 "		 CASE WHEN A.HABILITADO IS NOT NULL THEN LTRIM(A.HABILITADO) ELSE '' END AS HABILITADO, "&_
             "       NUSOBRES " &_
			 "FROM TBLBIOTECNOLOGIAS A "&_
			 "LEFT JOIN CLIENTES B ON B.CDCLIENTE = A.IDPROVEEDOR "&_
			 "WHERE A.IDPRODUCTO = " & pCdProducto
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) %>
    <table width='80%' align="center" class="datagrid" id="tblBiotecnologia">
		    <thead>
			    <tr>
				    <th align='center' width="5%"><font size='2'><% =GF_Traducir("Id") %></font></th>
					<th align='center' width="35%"><font size='2'><% =GF_Traducir("Descripcion") %></font></th>
					<th align='center' width="35%"><font size='2'><% =GF_Traducir("Proveedor") %></font></th>
					<th align='center' width="10%"><font size='2'><% =GF_Traducir("Estado") %></font></th>
					<th align='center' width="10%"><font size='2'><% =GF_Traducir("Sobres") %></font></th>
					<th align='center' width="2%">.</th>
					<th align='center' width="2%">.</th>
			   </tr>
			</thead>
            <tbody>
          <% if not rs.Eof then 
                while (not rs.eof) %>		    
			    <tr>
				    <td align='center'>
                        <font size='2'><% =rs("IDBIOTECNOLOGIA") %></font>
                        <input type="hidden" value="<% =rs("IDBIOTECNOLOGIA") %>" id="hidIdBiotecnologia_<%=index %>"/>
                    </td>
					<td align="left">
                        <span id="spanDsBiotecnologia_<%=index %>"><font size='2'><% =rs("DSBIOTECNOLOGIA") %></font></span>
                        <input type="text" id="txtDsBiotecnologia_<%=index %>" value="<%=rs("DSBIOTECNOLOGIA")%>"  style="width:100%;display:none;text-transform:uppercase;" maxlength="250"/>
                    </td>
					<td align="left">
                        <span id="spanDsCliente_<%=index %>"><font size='2'><% =rs("DSCLIENTE") %></font></span>
                        <input type="text"   name="dsCoordinado_<%=index %>" id="dsCoordinado_<%=index %>" value="<%=rs("DSCLIENTE")%>" style="display:none;width:100%;" onblur="controlarProveedor(<%=index %>)">
			            <input type="hidden" name="cdCoordinado_<%=index %>" id="cdCoordinado_<%=index %>" value="<%=rs("CDCLIENTE")%>">
                    </td>
					<td align='center'>
                        <span id="spanEstado_<%=index %>"><font size='2'><% =rs("HABILITADO") %></font></span>
                        <span id="spanHabilitado_<%=index %>" style="display:none;float:left;"><%=GF_TRADUCIR("V:") %><input type="radio" id="rdb_<%=index %>" name="rdb_<%=index %>" <% if(CStr(rs("HABILITADO")) = BIOTEC_ACTIVA)then %> checked <% end if %> value="<% =BIOTEC_ACTIVA %>" title="<%=GF_TRADUCIR("Habilitado") %>" /></span>
                        <span id="spanDeshabilitado_<%=index %>" style="display:none;float:right;"><%=GF_TRADUCIR("F:") %><input type="radio" id="rdb_<%=index %>" name="rdb_<%=index %>" <% if(CStr(rs("HABILITADO")) = BIOTEC_INACTIVA )then %> checked <% end if %> value="<% =BIOTEC_INACTIVA %>" title="<%=GF_TRADUCIR("Deshabilitado") %>" /></span>
                     </td>
                    <td align='center'>
                        <span id="spanNuSobre_<%=index %>"><font size='2'><% =rs("NUSOBRES") %></font></span>
                        <input type="text" id="txtNuSobre_<%=index %>" value="<%=rs("NUSOBRES")%>" style="width:100%;display:none;" maxlength="4" onkeypress="return controlIngreso(this, event, 'N')"/>
                    </td>
	    			<% if(flagPermiso)then %>
                        <td align="center">
                            <img src="../../images/edit-16.png" title="Editar" id="editarBiotecnologia_<%=index %>" style="cursor:pointer;" onclick="EditBiotecnologia(<%=index%>)">
                            <img src="../../images/save-16.png" title="Guardar" id="guardarBiotecnologia_<%=index %>" style="cursor:pointer;display:none;" onclick="SaveBiotecnologia(<%=index%>)">
                         </td>
						<td align="center"><img src="../../images/cross-16.png" title="Eliminar" id="eliminarBiotecnologia_<%=index %>" style="cursor:pointer;" onclick="DeleteBiotecnologia(<%=index%>)"></td>
                    <% else %>
                        <td align="center"></td>
						<td align="center"></td>
                    <% end if %>
				</tr>
				<% rs.MoveNext()
                index = index + 1
			wend  
            else %>
			    <tr><td align='center' colspan="7"><% =GF_TRADUCIR("No se encontraron resultados")%></td></tr>
		   <% end if %>
           <input type="hidden" value="<%=index %>" id="maxRowBiotecnologia" />
           </tbody> 
           <tfoot>
                <tr><td align='left' colspan="7"><div id="msgErrorBiotecnologia" style="display:none;"></div></td></tr>
                <% if(flagPermiso)then %>
                <tr><td colspan="7" align="center"><h3 style=cursor:pointer; onclick=AddBiotecnologia()>Agregar</h3></td></tr>
                <% end if %>
           </tfoot>
	 </table>
	<br></br>
<%		
End Function
'-------------------------------------------------------------------------------------------------------------------
Function ControlarBiotecnologia(p_dsBiotecnologia, p_cdCliente, p_nuSobre, ByRef p_msg )
    if (Trim(p_dsBiotecnologia) <> "") then
        if ( p_cdCliente <> 0 ) then
            if (p_nuSobre < 0) then p_msg = "El numero de sobres es incorrecto"
        else
            p_msg = "Proveedor incorrecto"
        end if
    else
        p_msg = "Descripcion incompleta"
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Dim cdProducto,accion,g_strPuerto,flagPermiso,dsBiotecnologia,cdCliente,nuSobre,estado,idBiotecnologia,msg

idBiotecnologia = GF_Parametros7("id","",6) 
cdProducto      = GF_Parametros7("cdProducto",0,6)
g_strPuerto     = GF_Parametros7("pto","",6)
accion          = GF_Parametros7("accion","",6)
flagPermiso     = GF_Parametros7("permiso","",6)
dsBiotecnologia = GF_Parametros7("descripcion","",6) 
cdCliente       = GF_Parametros7("cliente",0,6)
estado          = GF_Parametros7("estado","",6)
nuSobre         = GF_Parametros7("nuSobre",0,6)

Select Case accion
	Case ACCION_VISUALIZAR
		Call DibujarBiotecnologiaDeProducto( cdProducto )
	Case ACCION_BORRAR
        Call EliminarBiotecnologiaDeProducto( idBiotecnologia, cdProducto, g_strPuerto )
    Case ACCION_GRABAR
        Call ControlarBiotecnologia( dsBiotecnologia, cdCliente, nuSobre, msg )
        if ( msg = "" ) then
            Call GuardarBiotecnologia( idBiotecnologia, UCase(dsBiotecnologia), cdCliente, cdProducto, Ucase(estado), nuSobre, g_strPuerto )
        else
            response.write msg
        end if
End Select


%>

