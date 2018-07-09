<%
Dim cdUnidad, dsUnidad, mkWhereUnidades, verTodosUnidad

cdUnidad = UCase(GF_PARAMETROS7("cdUnidad", "" ,6))
dsUnidad = UCase(GF_PARAMETROS7("dsUnidad", "" ,6))
verTodosUnidad = GF_PARAMETROS7("todos", "" ,6)

if (verTodosUnidad = 0) then Call mkWhere(mkWhereUnidades, "ESTADO", ESTADO_BAJA, "<>", 1)
if (cdUnidad <> "") then Call mkWhere(mkWhereUnidades, "cdUnidad", cdUnidad, "=", 3)
if (dsUnidad <> "") then Call mkWhere(mkWhereUnidades, "dsUnidad", dsUnidad, "LIKE", 3)

strSQL="Select * from TBLUNIDADES " & mkWhereUnidades
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Call setupPaginacion(rs, pagina, regXPag)
'Se procesa el codigo interno de la seccion.	
%>	
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Codigo") %>:</td>
			<td><input type="text" maxlength="10" id="cdUnidad" value="<% =cdUnidad %>" ></td>
		</tr>														
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsUnidad" value="<% =dsUnidad %>"></td>
		</tr>
	</table>	  	
	-#-
	<table id="tableSeccion2" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
		<tr class="reg_header_nav">
			<td align="center">.</td>
			<td width="15%"><% =GF_TRADUCIR("Codigo") %></td>
			<td><% =GF_TRADUCIR("Descripcion") %></td>			
			<td width="24px" align="center">.</td>
			<td width="24px" align="center">.</td>
		</tr>
		<%		
		while ((not rs.eof) and (i < regXPag))			
			i = i+1
			%>
				<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
					<td align="center" width="24px">
						<img src="images/compras/units-16x16.png">
					</td>
					<td align="center" width="10%" onClick="loadPopUpUnidades(<% =rs("IDUNIDAD") %>)"><b><% =rs("CDUNIDAD") %></b></td>
					<td><% =rs("DSUNIDAD") %></td>
					<td align="center">
					<%  if (not isAuditor(SIN_DIVISION)) then
					 		if (rs("ESTADO") <> ESTADO_BAJA) then 	%>
								<img src="images/compras/edit-16x16.png" onClick="loadPopUpUnidades(<% =rs("IDUNIDAD") %>)" style="cursor: pointer" title="Editar Unidad" onclick="deleteElemento('2','<% =rs("IDUNIDAD") %>')"><%								
							end if	
						end if
					%>
					</td>
					<td align="center">
					<%  if (not isAuditor(SIN_DIVISION)) then 	
							if (rs("REFERENCIAS") = 0) then
								if (rs("ESTADO") = ESTADO_BAJA) then 
									%><img src="images/compras/accept-16x16.png" style="cursor: pointer" title="Activar Unidad" onclick="habilitarElemento('2','<% =rs("IDUNIDAD") %>')"><%		
								else 
									%><img src="images/compras/cancel-16x16.png" style="cursor: pointer" title="Borrar Unidad" onclick="deleteElemento('2','<% =rs("IDUNIDAD") %>')"><%		
								end if
							end if	
						end if
					%>
					</td>
				</tr>
			<%			
			rs.MoveNext()
		wend
		if (i = 0) then		
			%>			
				<tr>
					<td class="TDNOHAY" colspan="5"><% =GF_TRADUCIR("No existen unidades registradas") %></td>
				</tr>
			<%		
		end if 
		%>	
	</table>