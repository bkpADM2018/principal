<%
dsPresupuestoDetalle = GF_PARAMETROS7("dsPresupuestoDetalle", "" ,6)
if (verTodos = 0) then Call mkWhere(mkWhereCategorias, "IDESTADO", ESTADO_BAJA, "<>", 1)
if (dsPresupuestoDetalle <> "") then Call mkWhere(mkWhereCategorias, "DSDETALLE", ucase(dsPresupuestoDetalle), "=", 3)

strSQL="Select * from TBLBUDGETDETALLES " & mkWhereCategorias
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
'Response.write strSQL

Call setupPaginacion(rs, pagina, regXPag)
'Se procesa el codigo interno de la seccion.	
%>	
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsPresupuestoDetalle" value="<% =dsPresupuestoDetalle %>"></td>
		</tr>																
	</table>	  	
	-#-
	<table align="center" id="tableSeccion5" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1" background=file://C:\luna-y-mar-wallpaper-.jpg>
		<tr class="reg_header_nav">
			<td align="center">.</td>
			<td width="5%" align="center"><% =GF_TRADUCIR("Item") %></td>
			<td><% =GF_TRADUCIR("Descripcion") %></td>
			<td width="24px" align="center">.</td>
			<td width="24px" align="center">.</td>
		</tr>
		<%		
		'Response.Write "ACA" & dsPresupuestoDetalle
		'Response.write strSQL
		while ((not rs.eof)	and (i < regXPag))
				i = i + 1			
%>
			<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
				<td align="center" width="24px"><img id="imgCategoria<%= i %>" src="images/compras/Budget_Item-16x16.png"></td>
				<td align="center" width="10%"><b><% =rs("IDDETALLE") %></b></td>
				<td><% =rs("DSDETALLE") %></td>				
				<td align="center" width="24px">
					<%  if (not isAuditor(SIN_DIVISION)) then
								if (rs("IDESTADO") <> ESTADO_BAJA) then	%>
									<img src="images/compras/edit-16x16.png" onClick="loadPopUpPresupuestos(<% =rs("IDDETALLE") %>,2)" style="cursor: pointer" title="Editar Detalle">
						<% 		end if
					end if %>					
					
				</td>
				<td align="center" width="24px">
					<%  if (not isAuditor(SIN_DIVISION)) then
								if (rs("IDESTADO") = ESTADO_BAJA) then	%>
									<img onclick="habilitarElemento('6','<% =rs("IDDETALLE") %>')" src="images/compras/accept-16x16.png" style="cursor: pointer" title="Activar Detalle">
						<%		else  %>	
									<img onclick="deleteElemento('6','<% =rs("IDDETALLE") %>')" src="images/compras/cancel-16x16.png" style="cursor: pointer" title="Borrar Detalle">
						<% 		end if
						end if	%>					
				</td>
			</tr>
<%			rs.MoveNext()
		wend
		if (i = 0) then		
%>			
		<tr>
			<td class="TDNOHAY" colspan="4"><% =GF_TRADUCIR("No existen areas registradas") %></td>
		</tr>
<%		end if %>	
	</table>
