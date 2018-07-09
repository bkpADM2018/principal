<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosfechas.asp"-->
<!--#include file="include/procedimientoProducto.asp"-->
<%
Const SIZE_COSECHA = 8 'Es el tamaño de carracteres que debe tener la Cosecha para que sea valida
Function existeCosecha(p_CdProducto,p_CdCosecha)
    existeCosecha = false
    Dim strSQL
    strSQL = "SELECT * FROM dbo.COSECHAS WHERE CDPRODUCTO = "& p_CdProducto &" AND CDCOSECHA = " & p_CdCosecha
    Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) 
    if (not rs.Eof) then existeCosecha = true
End Function
'------------------------------------------------------------------------------------------------------------------
'Verifico si el producto tiene al menos una cosecha habilitada
Function checkCosechaHabilitada(p_CdCosecha,p_CdProducto)
    Dim strSQL
    checkCosechaHabilitada = false
    strSQL = "SELECT * FROM dbo.COSECHAS WHERE CDPRODUCTO = "& p_CdProducto & " AND LTRIM(COSDEF) = '"& ESTADO_ACTIVO  &"' AND CDCOSECHA <> "& p_CdCosecha
    Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)
    if (not rs.Eof) then checkCosechaHabilitada = true
End Function
'-------------------------------------------------------------------------------------------------------------------
'Controla que esten correctos los datos de la cosecha 
Function checkCosecha(p_CdProducto,p_Cosecha1,p_Cosecha2,p_ChkHabilita,p_IsEdit, ByRef msg)
	Dim rsAtr, auxCosecha
    'Controlo que su formato sea correcto
    auxCosecha = replace(Trim(p_Cosecha1) & Trim(p_Cosecha2)," ","")
    if (Len(auxCosecha) = SIZE_COSECHA) then
        if (Cdbl(p_Cosecha1) <= Cdbl(p_Cosecha2)) then
            'Controlo que no esté duplicada la cosecha (Primary Key), si es una modificacion no se controla
	        If ((not existeCosecha(p_CdProducto,auxCosecha))or(p_IsEdit)) then
                'Si asignó a la cosecha como habilitada, se termina el control. Si no asignó a la cosecha como habilitada controlo
                'que por lo menos halla una cosecha habilitada para el producto (si no lo hay debo asignar una si o si)
                if (Cdbl(p_ChkHabilita) <> ESTADO_ACTIVO)and(p_IsEdit) then
                    if (not checkCosechaHabilitada(auxCosecha,p_CdProducto)) then msg = "El producto debe tener una cosecha habilitada"
                end if
	        else
                msg = "Cosecha duplicada"
            end if
        else
            msg = "Error en el periodo de Cosecha"
        end if
    else
        msg = "Error en el formato de la Cosecha"
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Function drawCosechaByArticulo(p_CdProducto)
    Dim strSQL,index
    index = 0
    strSQL = "SELECT TOP 10 CASE WHEN CDCOSECHA IS NULL THEN 0 ELSE CDCOSECHA END AS CDCOSECHA,  " &_
             "       CASE WHEN COSDEF IS NULL THEN "&_
             "          '0' " &_
             "       ELSE  " &_
             "          LTRIM(COSDEF)"&_
             "       END AS DEF " &_
             "FROM COSECHAS WHERE CDPRODUCTO =" & p_CdProducto &_
             " ORDER BY CDCOSECHA DESC "
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) %>	
	    <table border=0 width='50%' class="datagrid" align="center" id="tblCosechas">        
		    <thead>
			    <tr>
				    <th align='center'><font size='2'><% =GF_Traducir("Periodo Inicial") %></font></th>
                    <th align='center'><font size='2'><% =GF_Traducir("Periodo Final") %></font></th>
					<th align='center'><font size='2'><% =GF_Traducir("Habilitado") %></font></th>
					<th align='center'>.</th>
					<th align='center'>.</th>
				</tr>
			</thead>
            <tbody>
            <% if not rs.Eof then %>
			<% while (not rs.eof) %>
			    <tr>
				    <td align='center'>
                        <span id="spanCosecha1_<%=index %>" ><font size='2'><% If (Cdbl(rs("CDCOSECHA")) <> 0 ) then Response.Write left(rs("CDCOSECHA"),4) %></font></span>
                        <input type="hidden" id="hidCosecha1_<%=index %>" value="<%=left(rs("CDCOSECHA"),4) %>" >
                    </td>
					<td align='center'>
                        <span id="spanCosecha2_<%=index %>" ><font size='2'><% If (Cdbl(rs("CDCOSECHA")) <> 0 ) then Response.Write Right(rs("CDCOSECHA"),4) %></font></span>
                        <input type="hidden" id="hidCosecha2_<%=index %>" value="<%=Right(rs("CDCOSECHA"),4) %>" >
                     </td>
					<td align='center'>
                        <span id="spanCosecha3_<%=index %>"><font size='2'><% if(Cdbl(rs("DEF")) = ESTADO_ACTIVO) then response.Write TIPO_AFIRMACION else response.Write TIPO_NEGACION end if %></font></span>
                        <input type="checkbox" style="display:none;" name="chk_<%=index %>" id="chk_<%=index %>" value="<%=rs("DEF") %>" <% if (Cdbl(rs("DEF")) = ESTADO_ACTIVO) then %> checked <% end if %>>
                    </td>
                    <td align="center">
                        <img src="../../images/edit-16.png" id="editCosecha_<%=index %>" title="Editar" style="cursor:pointer;" onclick="editarCosecha(<% =index%>)">
                        <img src="../../images/save-16.png" title="Guardar" id="guardarCosecha_<%=index %>" style="cursor:pointer;display:none;" onclick="guardarCosecha(<%=index%>)">
                    </td>
					<td align="center">
					<% if (CLng(rs("CDCOSECHA")) > 0) then %>
					    <img src="../../images/cross-16.png" title="Eliminar" id="eliminarCosecha_<%=index %>" style="cursor:pointer;" onclick="deleteCosecha(<% =index %>)">
                    <%  end if  %>					    
					</td>
				</tr>
				<% rs.MoveNext()
                    index= index + 1
				wend %>			
            <% else %>
            		<tr><td align='center' colspan="5">No se encontraron resultados.</td></tr>
		    <% end if %>
            <input type="hidden" value="<%=index %>" id="maxRowCosecha" />
            </tbody>
            <tfoot>
                <tr><td align='left' colspan="5"><div id="msgErrorCosecha" style="display:none;"></div></td></tr>
                    <tr><td colspan="5" align="center"><h3 style=cursor:pointer; onclick=AddCosecha()>Agregar</h3></td></tr>
            </tfoot>
		</table>
		<br></br>
<%
End Function
'-------------------------------------------------------------------------------------------------------------------
'Permite saber si la cosecha se puede elimnar, esto se controla si es usada por algun camion
Function puedeEliminarCosecha(p_CdProducto,p_CdCosecha)
    Dim strSQL 
    puedeEliminarCosecha = true
    strSQL = "Select c.IdCamion from dbo.Camiones c,dbo.CamionesDescarga cd "&_
             "where c.cdproducto="&p_CdProducto&" and c.Idcamion=cd.idcamion and cd.cdcosecha="& p_CdCosecha &_
             " union "&_
             "Select c.IdCamion from dbo.HCamiones c, dbo.HCamionesDescarga cd "&_
             "where c.cdproducto="&p_CdProducto&"  and c.Idcamion=cd.idcamion  and c.DtContable=cd.DtContable and cd.cdcosecha="&p_CdCosecha
    Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) 
    If not rs.Eof Then 
        puedeEliminarCosecha = false
    else
        strSQL = "Select c.IdCamion from dbo.Camiones c,dbo.CamionesCarga cd "&_
                 "where c.cdproducto=" & p_CdProducto & " and c.Idcamion=cd.idcamion and cd.cdcosecha=" & p_CdCosecha &_
                 " union " &_
                 "Select c.IdCamion from dbo.HCamiones c, dbo.HCamionesCarga cd "&_
                 "where c.cdproducto=" & p_CdProducto & " and c.Idcamion=cd.idcamion and c.DtContable=cd.DtContable and cd.cdcosecha=" & p_CdCosecha
        Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL) 
        If not rs.Eof Then puedeEliminarCosecha = false
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
Dim cdProducto,accion,cdAceptacion,g_strPuerto,flagPermiso,habilitado,cdCosechaInicio,cdCosechaFin

cdProducto = GF_Parametros7("cdProducto",0,6)
cdCosechaInicio = GF_Parametros7("cosechaI","",6)
cdCosechaFin = GF_Parametros7("cosechaF","",6)
g_strPuerto = GF_Parametros7("pto","",6)
accion = GF_Parametros7("accion","",6)
flagPermiso = GF_Parametros7("permiso","",6)
habilitado = GF_Parametros7("habilitado",0,6)
isEdit = GF_Parametros7("isEdit","",6)


Select Case accion
	Case ACCION_VISUALIZAR
        Call drawCosechaByArticulo(cdProducto)
	Case ACCION_BORRAR
        if (puedeEliminarCosecha(cdProducto,cdCosechaInicio & cdCosechaFin)) then
            Call deleteCosecha(cdProducto, cdCosechaInicio & cdCosechaFin, g_strPuerto)
        else
            Response.Write "No se puede eliminar la cosecha debido a que está siendo utilizada"
        end if
    Case ACCION_GRABAR
        Call checkCosecha(cdProducto,cdCosechaInicio,cdCosechaFin,habilitado,isEdit,msg)
        if ( msg = "" ) then
            if isEdit then
			    Call updateCosecha(cdProducto, cdCosechaInicio & cdCosechaFin, habilitado, g_strPuerto)
		    else
    			Call addCosecha(cdProducto, cdCosechaInicio & cdCosechaFin, habilitado, g_strPuerto)
		    end if
        else
            Response.Write msg
        end if
End Select


%>