<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosTitulos.asp"-->
<!--#include file="../../includes/procedimientosSQL.asp"-->
<%


'----------------------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'----------------------------------------------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'----------------------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto, Conn, params,lineasTotales, mostrar,paginaActual,flagPermiso, flagSearch
dim cmbCdProvincia, txtCdProcedenca

g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)

flagPermiso = true
if (leerPermisos(g_strPuerto, TASK_PRODUCT_USER) = NO_TIENE_PERMISO) then flagPermiso = false

cmbCdProvincia = GF_PARAMETROS7("cmbCdProvincia",0,6)
call addParam("cmbCdProvincia", cmbCdProvincia, params)
txtCdProcedencia = UCASE(GF_PARAMETROS7("txtCdProcedencia","",6))
call addParam("txtCdProcedencia", txtCdProcedencia, params)

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1

if (txtCdProcedencia <> "") then 
	strSQL = "SELECT PROCE.CDPROCEDENCIA, PROCE.DSPROCEDENCIA, PROCE.CDPROCEDENCIACAMARA, PCIA.DSPROVINCIA, PROCE.IDLOCONCCA, L.DSLOC as DSLOCONCCA  " & _
			 "		, LOCC.DSLOCALIDAD AS DSLOCALIDADCAMARA, PCAM.DSPROVINCIA AS DSPROVINCIACAMARA " & _
			 "	FROM dbo.PROCEDENCIAS PROCE " & _
			 "	Left join LOCPROVPART L on L.IDLOC=PROCE.IDLOCONCCA " &_
			 "	LEFT JOIN dbo.PROVINCIAS PCIA ON PROCE.CDPROV=PCIA.CDPROVINCIA " & _			 
			 "	LEFT JOIN dbo.LOCALIDADESCAMARAS LOCC ON CAST(SUBSTRING(PROCE.CDPROCEDENCIACAMARA,1,4) AS INT)=LOCC.CDLOCALIDADCAMARA AND CAST(SUBSTRING(PROCE.CDPROCEDENCIACAMARA,5,3) AS INT)=LOCC.CDLOCALIDADSUBCAMARA" & _			 
			 "	LEFT JOIN dbo.PROVINCIAS PCAM ON LOCC.CDPROVINCIA=PCAM.CDCAMARA " & _
			 " WHERE DSPROCEDENCIA LIKE '%" & txtCdProcedencia & "%' "
			if (cmbCdProvincia <> 0) then strSQL = strSQL & " AND CDPROV = " & cmbCdProvincia
			strSQL = strSQL & " ORDER BY DSPROCEDENCIA "
    call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)	
	Call setupPaginacion(rs, paginaActual, mostrar)
	lineasTotales = rs.recordcount
	flagSearch = true
END IF			
%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administracion de Procedencias </TITLE>
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>
</HEAD>
<script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../../scripts/paginar.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
<script language="javascript">
	var ch = new channel();
	var up1;	
	function onLoadPage(){
		tb = new Toolbar('toolbar', 6,'../../images/');				
		tb.addButton("refresh-16.png", "Refrescar", "submitInfo('<%=ACCION_SUBMITIR%>')");
		<% if(flagPermiso) or 1=1 then %>
        tb.addButton("add-16.png", "Nueva Procedencia", "nuevaProcedencia()");
        <% end if %>
		tb.draw();
		<% if flagSearch then
			if (not rs.eof) then %>
				var pgn = new Paginacion("paginacion");
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "procedenciasAdministrar.asp<% =params %>");
		<%	end if
		end if 	%>	   
		
	}	
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();
	}
	function nuevaProcedencia(){				
		var puw = new winPopUp('popupProcedencia','procedenciasAgregar.asp?pto=<%=g_strPuerto%>','550','400','Agregar nueva procedencia', "submitInfo()");
	}
	function editarProcedencia(pProcedencia){
		var puw = new winPopUp('popupProcedencia','procedenciasAgregar.asp?pto=<%=g_strPuerto%>&cdProcedencia=' + pProcedencia,'550','400','Agregar nueva procedencia', "submitInfo()");	
	}
	
</script>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post>
<div class="tableaside size100"> <!-- BUSCAR -->
	<h3> Filtros</h3>        
	<div id="searchfilter" class="tableasidecontent">
		<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Provincia") %> </div>
	    <div class="col16"> 
			<select id="cmbCdProvincia" name="cmbCdProvincia" >
				<option value="0"><%= GF_TRADUCIR("Todas...")%></option>
				<%
				strSQL = "SELECT CDPROVINCIA, DSPROVINCIA FROM dbo.PROVINCIAS ORDER BY DSPROVINCIA"
				call GF_BD_Puertos (g_strPuerto, rsProvincia, "OPEN",strSQL)										
				while not rsProvincia.eof 
					mySelected = ""
					if cint(cmbCdProvincia) = cint(rsProvincia("CDPROVINCIA")) then mySelected = "SELECTED"
				%>
						<option value="<%=rsProvincia("CDPROVINCIA")%>" <%=mySelected%>><%=rsProvincia("DSPROVINCIA")%></option>
				<%	rsProvincia.movenext
				wend
				%>							
			</select>									
		</div>						
		<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Procedencia") %> </div>
	    <div class="col16"> 
			<input type="text" id="txtCdProcedencia" name="txtCdProcedencia" size="50" value="<%=txtCdProcedencia%>">
		</div>						

		<span class="btnaction"><input type="button" value="Buscar" id=cmdSearch name=cmdSearch onclick="submitInfo('<%=ACCION_SUBMITIR%>');"></span>
	</div>
    
</div><!-- END BUSCAR -->
					
<div class="col66"></div>

		
								
	<table width="90%" class="datagrid" align="center" id="tblResult" name="tblResult">           
<% 	
if flagSearch then
	if (not rs.eof) then %>	
			<thead>
                <tr>
					<th><% =GF_TRADUCIR("Codigo") %>		</th>	
                    <th><% =GF_TRADUCIR("Descripcion") %>	</th>	
				    <th><% =GF_TRADUCIR("Provincia") %>		</th>
				    <th><% =GF_TRADUCIR("Localidad ONCCA") %>	</th>
				    <th><% =GF_TRADUCIR("Proc. C�mara") %>	</th>
				    <th><% =GF_TRADUCIR("Loc. C�mara") %>	</th>
				    <th><% =GF_TRADUCIR("Pcia. C�mara") %>	</th>
				    <th>.</th>
                </tr>
            </thead> 
				<tbody>
		<%		while not rs.EOF and (reg < mostrar)
					reg = reg + 1	    %>
					<tr class="reg_Header_navdos">		
						<td align="CENTER"><font size="2">
							<% =rs("CDPROCEDENCIA")%></font>
						</td>	
						<td align="left"><font size="2">
							<% =rs("DSPROCEDENCIA")%></font>
						</td>	
						<td align="CENTER"><font size="2">
							<% =rs("DSPROVINCIA") %></font>
						</td>
						<td align="CENTER"><font size="2">
							<% if (CLng(rs("IDLOCONCCA")) > 0) then
								response.write rs("IDLOCONCCA") & "-" & rs("DSLOCONCCA") 
								end if %></font>
						</td>
						<td align="CENTER"><font size="2">
							<% =rs("CDPROCEDENCIACAMARA") %></font>
						</td>
						<td align="CENTER"><font size="2">
							<% =rs("DSLOCALIDADCAMARA") %></font>
						</td>						
						<td align="CENTER"><font size="2">
							<% =rs("DSPROVINCIACAMARA") %></font>
						</td>	
                        <% 
                        if (flagPermiso) or 1=1 then %>
							<td align="center"><img src="../../images/edit-16.png" title="Editar" style="cursor:pointer;" onclick="javascript:editarProcedencia(<%=rs("CDPROCEDENCIA")%>)"></td>						
                        <% else %>
							<td></td>
                        <% end if %>
					</tr>
					<tr class="troculto" id="TR_ID_<% =rs("CDPROCEDENCIA") %>">
						<td colspan="3">
							<div align="center" id="TBL_<% =rs("CDPROCEDENCIA") %>" style="position:absolute; "><img src='../../images/Loading4.gif'></div>
						</td>
					</tr>
<%                  rs.movenext()
		        wend %>
			</tbody>			
            				
            <tfoot>
                <td colspan="8"><div id="paginacion"></div></td>
            </tfoot>
	<%else %>
		<tr><td align="center"><%=GF_TRADUCIR("No se encontraron procedencias")%></td></tr>
<%	end if	
else %>
		<tr><td align="center"><%=GF_TRADUCIR("Ingrese la descripci�n de la procedencia a buscar")%></td></tr>
<%	
end if %>
</TABLE>
<input type="hidden" name="accion" id="accion" value="<%= accion %>">
</form>
</BODY>
</HTML>