<%
Dim cdCategoria, dsCategoria, mkWhereCategorias, verTodosCategoria

cdCategoria = UCase(GF_PARAMETROS7("cdCategoria", "" ,6))
dsCategoria = GF_PARAMETROS7("dsCategoria", "" ,6)
verTodosCategoria = GF_PARAMETROS7("todos", "" ,6)

if (verTodosCategoria = 0) then Call mkWhere(mkWhereCategorias, "ESTADO", ESTADO_BAJA, "<>", 1)
if (cdCategoria <> "") then Call mkWhere(mkWhereCategorias, "CDCATEGORIA", cdCategoria, "=", 3)
if (dsCategoria <> "") then Call mkWhere(mkWhereCategorias, "DSCATEGORIA", dsCategoria, "LIKE", 3)

strSQL="Select * from TBLARTCATEGORIAS " & mkWhereCategorias & " order by CDCATEGORIA"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
'Response.write strSQL
Call setupPaginacion(rs, pagina, regXPag)

'Se procesa el codigo interno de la seccion.	
%>	
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Nombre") %>:</td>
			<td><input type="text" maxlength="6" id="cdCategoria" value="<% =cdCategoria %>" ></td>
		</tr>														
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsCategoria" value="<% =dsCategoria %>"></td>
		</tr>																
	</table>	  	
	-#-
	<table id="tableSeccion1" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
		<tr class="reg_header_nav">
			<td align="center">.</td>
			<td width="15%"><% =GF_TRADUCIR("Nombre") %></td>
			<td><% =GF_TRADUCIR("Descripcion") %></td>
			<td width="24px" align="center">.</td>
			<td width="24px" align="center">.</td>
		</tr>
<%			while ((not rs.eof)	and (i < regXPag))
			i = i+1			
%>
			<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
				<td align="center" width="24px"><img id="imgCategoria<%= i %>" src="images/compras/categories-16x16.png"></td>
				<td align="center" width="10%"><b><% =rs("CDCATEGORIA") %></b></td>
				<td><% =rs("DSCATEGORIA") %></td>				
				<td align="center" width="24px">
					<%  if (not isAuditor(SIN_DIVISION)) then
							if ( rs("REFERENCIAS") = 0) then
								if (rs("ESTADO") <> ESTADO_BAJA) then	%>
									<img src="images/compras/edit-16x16.png" onClick="loadPopUpCategorias(<% =rs("IDCATEGORIA") %>)" style="cursor: pointer" title="Editar Categoria">
						<% 		end if
							end if 
						
					end if %>					
					
				</td>
				<td align="center" width="24px">
					<%  if (not isAuditor(SIN_DIVISION)) then
						 	if ( rs("REFERENCIAS") = 0) then
								if (rs("ESTADO") = ESTADO_BAJA) then	%>
									<img onclick="deleteElemento('1','<% =rs("IDCATEGORIA") %>')" src="images/compras/accept-16x16.png" style="cursor: pointer" title="Activar Categoria">
						<%		else  %>	
									<img onclick="deleteElemento('1','<% =rs("IDCATEGORIA") %>')" src="images/compras/cancel-16x16.png" style="cursor: pointer" title="Borrar Categoria">
						<% 		end if
							end if 
						end if	%>					
				</td>
			</tr>
<%			rs.MoveNext()
		wend
		if (i = 0) then		
%>			
		<tr>
			<td class="TDNOHAY" colspan="5"><% =GF_TRADUCIR("No existen categorias registradas") %></td>
		</tr>
<%		end if %>	
	</table>