<%
Dim cdAlmacen, dsAlmacen, dsDivision, mkWhereAlmacenes, verTodosAlmacen

cdAlmacen = UCase(GF_PARAMETROS7("cdAlmacen", "" ,6))
dsAlmacen = GF_PARAMETROS7("dsAlmacen", "" ,6)
dsDivision = GF_PARAMETROS7("dsDivision", "" ,6)
verTodosAlmacen = GF_PARAMETROS7("todos", "" ,6)
if (verTodosAlmacen = 0) then Call mkWhere(mkWhereAlmacenes, "ESTADO", ESTADO_BAJA, "<>", 1)
if (cdAlmacen <> "") then Call mkWhere(mkWhereAlmacenes, "cdAlmacen", cdAlmacen, "=", 3)
if (dsAlmacen <> "") then Call mkWhere(mkWhereAlmacenes, "dsAlmacen", dsAlmacen, "LIKE", 3)
if (dsDivision <> "") then Call mkWhere(mkWhereAlmacenes, "dsDivision", dsDivision, "LIKE", 3)

strSQL="Select * from TBLALMACENES A inner join TBLDIVISIONES D ON A.IDDIVISION=D.IDDIVISION " & mkWhereAlmacenes & " order by cdalmacen"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Call setupPaginacion(rs, pagina, regXPag)

'Se procesa el codigo interno de la seccion.	
%>	
<br>
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Nombre") %>:</td>
			<td><input type="text" maxlength="6" id="cdAlmacen" value="<% =cdAlmacen %>" ></td>
		</tr>														
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsAlmacen" value="<% =dsAlmacen %>"></td>
		</tr>																
		<tr>
			<td align="right"><% = GF_TRADUCIR("Division") %>:</td>
			<td>
						<%
						strSQL="Select * from TBLDIVISIONES"
						Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
						%>
							<select id="dsDivision" name="dsDivision">
								<option value="" selected="true"><% =GF_TRADUCIR("Seleccionar...") %>
						<%		while (not rsDivision.eof) 	%>										
									<option value="<% =rsDivision("DSDIVISION") %>" <% if (dsDivision = rsDivision("DSDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
						<%			rsDivision.MoveNext()
								wend	
						%>
							</select>
						<%
						Call executeQueryDB(DBSITE_SQL_INTRA, rsDivision, "CLOSE", strSQL)
						%>
			<!--<input type="text" id="dsDivision" value="<% =dsDivision %>"></td>-->
		</tr>	
	</table>	  	
	-#-
	<table id="tableSeccion1" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
		<tr class="reg_header_nav">
			<td width="5%" align="center">.</td>
			<td width="12%"><% =GF_TRADUCIR("Nombre") %></td>
			<td width="20%"><% =GF_TRADUCIR("Descripcion") %></td>
			<td><% =GF_TRADUCIR("Division") %></td>
			<td width="5%" align="center">.</td>
			<td width="5%" align="center">.</td>
		</tr>
<%			while ((not rs.eof)	and (i < regXPag))
			i = i+1			
%>
			<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
				<td align="center" onClick="loadPopUpAlmacenes(<% =rs("IDALMACEN") %>)"><img id="imgCategoria<%= i %>" src="images/almacenes/warehouses-16x16.png"></td>
				<td align="center" onClick="loadPopUpAlmacenes(<% =rs("IDALMACEN") %>)"><b><% =rs("cdAlmacen") %></b></td>
				<td onClick="loadPopUpAlmacenes(<% =rs("IDALMACEN") %>)"><% =ucase(rs("dsAlmacen")) %></td>				
				<td onClick="loadPopUpAlmacenes(<% =rs("IDALMACEN") %>)"><% =ucase(rs("DSDIVISION")) %></td>				
				<td align="center"><img src='images/almacenes/campana-16x16.png' title='Alertas' style='cursor:pointer' onclick='loadPopUpAlertasAlmacenes(<% =rs("IDALMACEN") %>)'></td>
				<td align="center">
					<%	if (rs("ESTADO") = ESTADO_BAJA) then	%>
								<img src="images/icon_ok.gif" style="cursor: pointer" title="Activar Almacen" onclick="habilitarElemento('0','<% =rs("IDALMACEN") %>')">
					<%	else  %>	
								<img src="images/icon_del.gif" style="cursor: pointer" title="Borrar Almacen" onclick="deleteElemento('0','<% =rs("IDALMACEN") %>')">
					<%	end if	%>					
				</td>
			</tr>
<%			rs.MoveNext()

		wend
		if (i = 0) then		
%>			
		<tr>
			<td class="TDNOHAY" colspan="6"><% =GF_TRADUCIR("No existen almacenes registrados") %></td>
		</tr>
<%		end if %>	
	</table>
