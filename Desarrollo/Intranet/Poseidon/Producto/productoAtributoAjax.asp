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
Function getComboBoxAtributo(pIndex)
    %>
    <select id="cmbAceptacion_<%=pIndex %>" name="cmbAceptacion_<%=pIndex %>" >
        <option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
	    <%	call GF_BD_Puertos (g_strPuerto, rsAceptacion, "OPEN","SELECT CDACEPTACION, DSACEPTACION FROM dbo.ACEPTACIONCALIDAD ORDER BY CDACEPTACION")
		    while not rsAceptacion.eof %>
			    <option value="<%=rsAceptacion("CDACEPTACION")%>" ><%=rsAceptacion("DSACEPTACION")%></option>
				<%	rsAceptacion.movenext
			wend %>
	</select>
    <%
End Function
'-------------------------------------------------------------------------------------------------------------------
Function drawAtributeByArticulo(pCdProducto)
	Dim index
    index = 0
    strSQL = "SELECT CASE WHEN ATR.ICSTICKER IS NOT NULL THEN ATR.ICSTICKER ELSE -1 END AS STICKER ,"&_
			 "		 CASE WHEN ATR.ICSUPERVISOR IS NOT NULL THEN ATR.ICSUPERVISOR ELSE -1 END AS SUPERVISOR,"&_
			 "		 CASE WHEN ATR.ICCAMARA IS NOT NULL THEN ATR.ICCAMARA ELSE -1 END AS CAMARA ,"&_
			 "		 CASE WHEN ATR.ICMOTIVORECHAZO IS NOT NULL THEN ATR.ICMOTIVORECHAZO ELSE -1 END AS MOTIVORECHAZO,"&_	
			 "		 CASE WHEN ATR.ICGRADO IS NOT NULL THEN ATR.ICGRADO ELSE -1 END AS GRADO ,"&_	
			 "		 CASE WHEN ATR.ICMERMA IS NOT NULL THEN ATR.ICMERMA ELSE -1 END AS MERMA,"&_	
			 "		 CASE WHEN ATR.ICRUBRO IS NOT NULL THEN ATR.ICRUBRO ELSE -1 END AS RUBRO ,"&_	
			 "		 CASE WHEN ATR.ICBALDE IS NOT NULL THEN ATR.ICBALDE ELSE -1 END AS BALDE,"&_
			 "		 CASE WHEN ATR.ICACON IS NOT NULL THEN ATR.ICACON ELSE -1 END AS ACON,"&_
			 "		 CASE WHEN ATR.ICINFORMEINTERNO IS NOT NULL THEN ATR.ICINFORMEINTERNO ELSE -1 END AS INFORMEINTERNO,"&_	
			 "		 ACP.DSACEPTACION, "&_
			 "		 ACP.CDACEPTACION "&_
			 "FROM AtributosDeProducto ATR "&_
			 "INNER JOIN ACEPTACIONCALIDAD ACP ON ACP.CDACEPTACION = ATR.CDACEPTACION "&_
			 "INNER JOIN PRODUCTOS PRO ON PRO.CDPRODUCTO = ATR.CDPRODUCTO "&_
			 "WHERE ATR.CdProducto=" & pCdProducto
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) %>

			<table border=0 width='100%' class="datagrid" id="tblAtributo">
					<thead>
				    <tr>
				        <th align='center'><font size='2'><% =GF_Traducir("Cod. Aceptacion") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Sticker") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Supervisor") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Camara") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Rechazo") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Grado") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Merma") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Rubro") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Balde") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Acond.") %></font></th>
					    <th align='center'><font size='2'><% =GF_Traducir("Inf. Calada") %></font></th>
					    <th align='center'>.</th>
					    <th align='center'>.</th>
				    </tr>
				    </thead>
                    <tbody>
                    <% if not rs.Eof then %>
				    <% while (not rs.eof) %>
						<tr >
						    <td align='left'>
                                <span id="spanDsAceptacion_<%=index %>"><font size='2'><% =rs("DSACEPTACION") %></font></span>
						        <input type="hidden" id="cdAceptacion_<%=index %>" name="cdAceptacion_<%=index %>" value="<%=rs("CDACEPTACION")%>">
                            </td>
						    <td align='center'>
                                <span id="spanSticker_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("STICKER")) %></font></span>
                                <div id="divSticker_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSticker_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>"  <%if(CInt(rs("STICKER")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSticker_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("STICKER")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked"%>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSticker_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("STICKER")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" end if%>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
						    <td align='center'>
                                <span id="spanSupervisor_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("SUPERVISOR")) %></font></span>
                                <div id="divSupervisor_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSupervisor_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("SUPERVISOR")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSupervisor_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("SUPERVISOR")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbSupervisor_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("SUPERVISOR")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
							<td align='center'>
                                <span id="spanCamara_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("CAMARA")) %></font></span>                                
                                <div id="divCamara_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbCamara_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("CAMARA")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbCamara_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("CAMARA")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbCamara_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("CAMARA")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                 </div>           
                            </td>
							<td align='center'>
                                <span id="spanRechazo_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("MOTIVORECHAZO")) %></font></span>
                                <div id="divRechazo_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRechazo_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("MOTIVORECHAZO")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRechazo_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("MOTIVORECHAZO")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRechazo_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("MOTIVORECHAZO")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
							<td align='center'>
                                <span id="spanGrado_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("GRADO")) %></font></span>
                                <div id="divGrado_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbGrado_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("GRADO")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbGrado_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("GRADO")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbGrado_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("GRADO")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
							<td align='center'>
                                <span id="spanMerma_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("MERMA")) %></font></span>
                                <div id="divMerma_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbMerma_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("MERMA")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbMerma_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("MERMA")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbMerma_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("MERMA")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
							<td align='center'>
                                <span id="spanRubo_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("RUBRO")) %></font></span>
                                <div id="divRubro_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRubro_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("RUBRO")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRubro_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("RUBRO")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbRubro_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("RUBRO")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>  
                            </td>
							<td align='center'>
                                <span id="spanBalde_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("BALDE")) %></font></span>
                                <div id="divBalde_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbBalde_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("BALDE")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbBalde_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("BALDE")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbBalde_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("BALDE")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>  
                            </td>
							<td align='center'>
                                <span id="spanAcon_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("ACON")) %></font></span>
                                <div id="divAcon_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbAcon_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("ACON")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbAcon_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("ACON")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbAcon_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("ACON")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
							<td align='center'>
                                <span id="spanInforme_<%=index %>"><font size='2'><% =getDsResultAtribute(rs("INFORMEINTERNO")) %></font></span>
                                <div id="divInforme_<%=index %>" style="display:none;">
                                    <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbInforme_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_AFIRMATIVO%>" <%if(CInt(rs("INFORMEINTERNO")) = VALUE_ATRIBUTE_AFIRMATIVO)then Response.Write "checked" %> /><%=GF_TRADUCIR("Si")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbInforme_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_NEGATIVO%>" <%if(CInt(rs("INFORMEINTERNO")) = VALUE_ATRIBUTE_NEGATIVO)then Response.Write "checked" %>/><%=GF_TRADUCIR("No")%></div>
					                <div style="margin:5px;"><input style="float:left;margin:0;" name="rdbInforme_<%=index %>" type="radio" value="<%=VALUE_ATRIBUTE_OPCIONAL%>" <%if(CInt(rs("INFORMEINTERNO")) = VALUE_ATRIBUTE_OPCIONAL)then Response.Write "checked" %>/><%=GF_TRADUCIR("Opcional")%></div>
                                </div>
                            </td>
	    					<% if(flagPermiso)then %>
                            <td align="center">
                                <img src="../../images/edit-16.png" title="Editar" id="editarAtributo_<%=index %>" style="cursor:pointer;" onclick="editarAtributo(<%=index%>)">
                                <img src="../../images/save-16.png" title="Guardar" id="guardarAtributo_<%=index %>" style="cursor:pointer;display:none;" onclick="guardarAtributo(<%=index%>)">
                            </td>
							<td align="center"><img src="../../images/cross-16.png" title="Eliminar" id="eliminarProd" style="cursor:pointer;" onclick="cargaDel(<% =pCdProducto %>,<%=rs("CDACEPTACION")%>)"></td>
                            <% else %>
                            <td align="center"></td>
							<td align="center"></td>
                            <% end if %>
						</tr>
						<% rs.MoveNext()
                            index= index + 1
				        wend %>				       
                 <% else %>
				    <tr><td align='center' colspan="13">No se encontraron resultados.</td></tr>
		         <% end if %>
                 <input type="hidden" value="<%=index %>" id="maxRowAtributo" />
                  </tbody> 
                   <tfoot>
                        <tr><td align='left' colspan="13"><div id="msgErrorAtributo" style="display:none;"></div></td></tr>
                        <% if(flagPermiso)then %>
                            <tr><td colspan="13" align="center"><h3 style=cursor:pointer; onclick=AddAtributo()>Agregar</h3></td></tr>
                        <% end if %>
                    </tfoot>
				</table>
				<br></br>
<%		
End Function
'-------------------------------------------------------------------------------------------------------------------
Function grabarAtributo(pcdProducto,pcdAceptacion,psticker,psuperv,pcamara,pgrado,pmerma,prubro,pbalde,pacon,pinform,prechaz,pisEdit)
    if (not pisEdit) then
	    Call addAtribute(pcdProducto,pcdAceptacion,psticker,pcamara,prechaz,pgrado,pmerma,prubro,pbalde,pinform,psuperv,pacon,g_strPuerto)
    else
        Call updateAtribute(pcdProducto,pcdAceptacion,psticker,pcamara,prechaz,pgrado,pmerma,prubro,pbalde,pinform,psuperv,pacon,g_strPuerto)
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Function controlarAtributo(pcdProducto,pcdAceptacion,pisEdit,ByRef msg)
    msg = ""
    if (not pisEdit) then
		Set rsAtr = getAtributteProducto(pcdProducto,pcdAceptacion,g_strPuerto)
		if not rsAtr.Eof Then  msg = "El concepto ya existe para el producto"
	end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Dim cdProducto,accion,cdAceptacion,g_strPuerto,flagPermiso,index

cdProducto = GF_Parametros7("cdProducto",0,6)
cdAceptacion = GF_Parametros7("cdAceptacion",0,6)
g_strPuerto = GF_Parametros7("pto","",6)
accion = GF_Parametros7("accion","",6)
flagPermiso = GF_Parametros7("permiso","",6)
sticker = GF_Parametros7("sticker",0,6)
superv = GF_Parametros7("superv",0,6)
camara = GF_Parametros7("camara",0,6)
grado = GF_Parametros7("grado",0,6)
merma = GF_Parametros7("merma",0,6)
rubro = GF_Parametros7("rubro",0,6)
balde = GF_Parametros7("balde",0,6)
acon = GF_Parametros7("acon",0,6)
inform = GF_Parametros7("inform",0,6)
rechaz = GF_Parametros7("rechaz",0,6)
isEdit = GF_Parametros7("isEdit","",6)

Select Case accion
	Case ACCION_VISUALIZAR
		Call drawAtributeByArticulo(cdProducto)
	Case ACCION_BORRAR        
        Call deleteAtributo(cdProducto,cdAceptacion,g_strPuerto)
    Case ACCION_PROCESAR
        'Obtengo el combo Box de los Conceptos
        index = GF_Parametros7("indice",0,6)
        Response.Write getComboBoxAtributo(index)
    Case ACCION_GRABAR
        Call controlarAtributo(cdProducto,cdAceptacion,isEdit,msg)
        if ( msg = "" )  then 
            Call grabarAtributo(cdProducto,cdAceptacion,sticker,superv,camara,grado,merma,rubro,balde,acon,inform,rechaz,isEdit)
        else
            Response.Write msg
        end if    
End Select


%>