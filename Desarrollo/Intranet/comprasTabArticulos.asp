
<%
'Funcion responsable por determinar si el artículo especificado tiene stock en alguna almacen de la empresa.
Function hayStock(pIdArticulo)
	Dim strSQL, rs, cant, ret
	
	ret = true
	
	strSQL="Select sum(EXISTENCIA+SOBRANTE) CANT from TBLARTICULOSDATOS where IDARTICULO=" & pIdArticulo
	Call executeQueryDb(DBSITE_SQL_INTRA,rs, "OPEN", strSQL)
	if (rs.eof)	then 
		'No hay stock ni nunca hubo.
		ret=false
	else
		'Actualmente no hay stock?.
		if (isNull(rs("CANT"))) then
			cant=0	
		else
			cant = rs("CANT")
		end if
		if (CDbl(cant) = 0) then ret=false
	end if
	hayStock = ret
End Function
'----------------------------------------------------------------
'Funcion responsablbe de detarminar si hay compras actuvas que involucren al artículo indicdo.
Function hayComprasPendientes(pIdArticulo)

	Dim strSQL, rs, cant, ret
	
	ret = true
	
	strSQL= "Select count(*) PICS from	TBLCTZCABECERA C " &_
			"					inner join TBLCTZDETALLE D on C.IDCOTIZACION=D.IDCOTIZACION " &_
			" where ESTADO in ('"& CTZ_PENDIENTE & "', '" & CTZ_EN_FIRMA & "', '" & CTZ_FIRMADA & "', '" & CTZ_EN_AJUSTE & "') and IDARTICULO =" & pIdArticulo		
	Call executeQueryDb(DBSITE_SQL_INTRA,rs, "OPEN", strSQL)
	if (rs.eof)	then 
		'No hay Compras.
		ret=false
	else
		if (isNull(rs("PICS"))) then
			cant=0	
		else
			cant = rs("PICS")
		end if
		if (CDbl(cant) = 0) then ret=false
	end if
	hayComprasPendientes = ret		
	
End Function
'********************************************************
'***	INICIO DE LA PAGINA
'********************************************************
Dim idArticulo, dsArticulo, mkWhereArticulos, rsUnidad, connUnidad, rsStock, cdArtCategoria, todos

idArticulo = UCase(GF_PARAMETROS7("idArticulo", "" ,6))
dsArticulo = ucase(GF_PARAMETROS7("dsArticulo", "" ,6))
cdArtCategoria = UCase(GF_PARAMETROS7("cdArtCategoria", "" ,6))
todos = GF_PARAMETROS7("todos", "" ,6)

if (todos <> "1")		then Call mkWhere(mkWhereArticulos, "A.ESTADO", ESTADO_BAJA, "<>", 1)
if (idArticulo <> "")	then Call mkWhere(mkWhereArticulos, "A.idArticulo", idArticulo, "=", 1)
if (dsArticulo <> "")	then Call mkWhere(mkWhereArticulos, "A.dsArticulo", dsArticulo, "LIKE", 3)
if (cdArtCategoria <> "")	then Call mkWhere(mkWhereArticulos, "B.CDCATEGORIA", cdArtCategoria, "=", 3)
strSQL="Select A.*, B.CDCATEGORIA, B.DSCATEGORIA  from TBLARTICULOS A "
strSQL = strSQL & " left join TBLARTCATEGORIAS B on A.IDCATEGORIA = B.IDCATEGORIA "
strSQL = strSQL & mkWhereArticulos & " order by IDARTICULO"
Call executeQueryDb(DBSITE_SQL_INTRA,rs, "OPEN", strSQL)
Call setupPaginacion(rs, pagina, regXPag)

'Se procesa el codigo interno de la seccion.	
%>	

	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Codigo") %>:</td>
			<td><input type="text" maxlength="6" id="idArticulo" value="<% =idArticulo %>" ></td>
		</tr>
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsArticulo" value="<% =dsArticulo %>"></td>
		</tr>
		<tr>
			<td align="right"><% = GF_TRADUCIR("Categoria") %>:</td>
			<td><input type="text" id="cdArtCategoria" value="<% =cdArtCategoria %>"></td>
		</tr>			
	</table>	  	
	-#-
	<table id="tableSeccion3" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
		<tr class="reg_header_nav">
			<td align="center">.</td>
			<td width="8%"><% =GF_TRADUCIR("Codigo") %></td>
			<td width="40%"><% =GF_TRADUCIR("Descripcion") %></td>
			<td width="10%"><% =GF_TRADUCIR("Unidad") %></td>
			<td width="30%"><% =GF_TRADUCIR("Categoria") %></td>
			<td width="24px" align="center">.</td>
			<td width="24px" align="center">.</td>
		</tr>
	<%		i=0
		while ((not rs.eof) and (i < regXPag))			
			strSQL="Select * from TBLUNIDADES where IDUNIDAD=" & rs("IDUNIDAD")
			Call executeQueryDb(DBSITE_SQL_INTRA,rsUnidad, "OPEN", strSQL)
			i = i+1
	%>
				<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
					<td align="center" width="24px"><img src="images/compras/items-16x16.png"></td>
					<td align="center" width="8%"><b><% =rs("IDARTICULO") %></b></td>
					<td width="40%"><% =rs("DSARTICULO") %></td>
					<td align="center" width="10%"><b><% =rsUnidad("DSUNIDAD") %></b></td>
					<td width="30%"><b><% =rs("CDCATEGORIA") & " - " & rs("DSCATEGORIA") %></b></td>
					<td width="24px" align="center">
						<%  if (not isAuditor(SIN_DIVISION)) then
								 if (rs("ESTADO") <> ESTADO_BAJA) then %>
								<img src="images/compras/edit-16x16.png" style="cursor: pointer" title="Editar Articulo" onClick="loadPopUpArticulos(<% =rs("IDARTICULO") %>)">
								<% end if 
							end if%>
					</td>
					<td width="24px" align="center">
					<%  flags = ""
						if (not isAuditor(SIN_DIVISION)) then
							if (rs("ESTADO") <> ESTADO_BAJA) then 
								if (hayStock(rs("IDARTICULO"))) then
									flags= "S"									
									descFlags = "Hay Stock en algún almacen"
								end if
								if (hayComprasPendientes(rs("IDARTICULO"))) then
									flags= flags & "C"	
									descFlags = descFlags & " | Hay Compras Activas."
								end if	
								if (flags <> "") then	%>
								<span style="cursor: pointer" title="<% =descFlags %>"><% =flags %></span>
						<%		else	%>													
								<img onclick="deleteElemento('3','<% =rs("IDARTICULO") %>')" src="images/compras/cancel-16x16.png" style="cursor: pointer" title="Bloquear Articulo">
						<%		end if
							else  
								if (rs("IDREEMPLAZO") = 0) then %>
								<img onclick="habilitarElemento('3','<% =rs("IDARTICULO") %>')" src="images/compras/accept-16x16.png" style="cursor: pointer" title="Activar Articulo">
						<%		else	%>
								<img onclick="alert('El articulo fue reemplazado por el <% =rs("IDREEMPLAZO") %>')" src="images/compras/warning-16x16.png" style="cursor: pointer" title="El articulo fue reemplazado por el <% =rs("IDREEMPLAZO") %>">
						<%		end if
							end if 
						end if	%>
					</td>
				</tr>
				
	<%			rs.MoveNext()
		wend
		if (i = 0) then		
	%>			
		<tr>
			<td class="TDNOHAY" colspan="7"><% =GF_TRADUCIR("No existen articulos registrados") %></td>
		</tr>
	<%		end if %>
	</table>