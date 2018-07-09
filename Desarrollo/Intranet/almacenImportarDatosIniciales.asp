<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%

Function importPrices(idDivision)
	Dim filename, fso, line, data, strSQL, conn, rs, contador, proximaLinea, i
	Dim tc, vlPesos, vlDolares, idArroyo, idPiedrabuena, idTransito
	idArroyo = getDivisionID(CODIGO_ARROYO)
	idPiedrabuena = getDivisionID(CODIGO_PIEDRABUENA)
	idTransito = getDivisionID(CODIGO_TRANSITO)
	proximaLinea = GF_Parametros7("proximaLinea",0 ,6)
	filename = Server.MapPath("Documentos\prices.csv")
	tc = getTipoCambio(MONEDA_DOLAR, session("MmtoSistema"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(filename)) then
		'Se importan los datos del archivo
		Set file=fso.OpenTextFile(filename,1)
		for i = 0 to proximaLinea - 2
			if not file.AtEndOfStream then file.SkipLine
		next
		while (not file.AtEndOfStream)
			contador = contador + 1
			line = file.ReadLine
			data = Split(line, ";")
			if isnumeric(data(0)) then
				strSQL="SELECT * FROM TBLARTICULOSPRECIOS WHERE IDARTICULO=" & data(0) & " AND MMTOPRECIO LIKE '" & left(session("MmtoSistema"),8) & "%'"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "OPEN", strSQL)	
				if rs.eof then
					vlPesos = CDbl(data(6))
					vlDolares = vlPesos / tc
					'Se arma la SQL para insertar
					strSQL="Insert into TBLARTICULOSPRECIOS values(" & session("MmtoSistema") & ", " & idArroyo & ", " & data(0) & ", " & CLng(vlPesos*100) & ", " & CLng(vlDolares*100) & ", " & tc & ",null, null, null, null)"
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
					strSQL="Insert into TBLARTICULOSPRECIOS values(" & session("MmtoSistema") & ", " & idPiedrabuena & ", " & data(0) & ", " & CLng(vlPesos*100) & ", " & CLng(vlDolares*100) & ", " & tc & ",null, null, null, null)"
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
					strSQL="Insert into TBLARTICULOSPRECIOS values(" & session("MmtoSistema") & ", " & idTransito & ", " & data(0) & ", " & CLng(vlPesos*100) & ", " & CLng(vlDolares*100) & ", " & tc & ",null, null, null, null)"
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
				end if
			end if
			if contador = 1000 then
				proximaLinea = file.Line
				Response.Redirect "almacenImportarDatosIniciales.asp?accion=1&proximaLinea= " & proximaLinea
			end if
		wend
		file.Close
		Set file=Nothing
	else
		Response.Write "El archivo prices.csv no existe!"
	end if
	Set fso=Nothing
End Function
'----------------------------------------------------------------------------------------------------------------------
Function ajustar(idAlmacen, pIdVale)
	Dim filename, fso, line, data, strSQL, conn, rs
	Dim ajusteE, ajusteS, ajuste, stockE, stockS
	
	'Se crea el vale de ajuste y se toma su ID para utilizarlo en el proceso.
	if (pIdVale = 0) then		
		'Se crea un nuevo vale de ajuste.
		Call clearHeaderVale()
		VS_cdVale = "AJS"
		VS_idAlmacen = idAlmacen
		VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoSistema"), 8))		
		Call GF_MGKS("SG", "FAR", VS_idSolicitante, VS_dsSolicitante)
		Call grabarHeaderVale(pIdVale, 0)
		Call grabarComentarioVale(pIdVale, "AJUSTE POR STOCKS INICIALES RELEVADOS")
	else
		strSQL="Delete from TBLVALESDETALLE where IDVALE=" & pIdVale
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
	end if
	'Se lee el stock actual
	strSQL="Select A.IDARTICULO IDARTICULO, A.EXISTENCIA EXISTENCIA, A.SOBRANTE SOBRANTE, B.EXISTENCIA EX, B.SOBRANTE SO from TBLARTICULOSDATOS2 A left join TBLARTICULOSDATOS B on A.IDALMACEN=B.IDALMACEN and A.IDARTICULO=B.IDARTICULO where A.IDALMACEN=" & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "OPEN", strSQL)	
	while (not rs.eof)
		'Se calcula el ajuste a realizar.				
		if (isNull(rs("EX"))) then			
			'No hay data anterior.
			ajusteE = CDbl(rs("EXISTENCIA"))				
			ajusteS = CDbl(rs("SOBRANTE"))	
		else
			ajusteE = CDbl(rs("EXISTENCIA")) - CDbl(rs("EX"))				
			ajusteS = CDbl(rs("SOBRANTE"))  - CDbl(rs("SO"))			
		end if
		ajuste = ajusteE + ajusteS			
		'Se graba el ajuste.				
		if (ajuste <> 0) then
			strSQL="Insert into TBLVALESDETALLE values(" & pIdVale & ", " & rs("IDARTICULO") & " , " & ajuste & " , " & ajusteE & " , " & ajusteS & ", 0, 0)"					
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
		end if
		rs.MoveNext()		
	wend
		
End Function
'----------------------------------------------------------------------------------------------------------------------
Function blanquearStock(idAlmacen)
	strSQL="Delete from TBLARTICULOSDATOS2"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL="Insert into TBLARTICULOSDATOS2 Select * from TBLARTICULOSDATOS where IDALMACEN=" & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
	strSQL="Update TBLARTICULOSDATOS2 set EXISTENCIA=0, SOBRANTE=0"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
End Function
'----------------------------------------------------------------------------------------------------------------------
Function importStocks(idAlmacen, fileno)
	Dim filename, fso, line, data, strSQL, conn, rs
	Dim ajusteE, ajusteS, ajuste, stockE, stockS
	
	filename = Server.MapPath("Documentos\stocks_" & fileno & ".csv")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(filename)) then
		'Se importan los datos del archivo
		Set file=fso.OpenTextFile(filename,1)
		while (not file.AtEndOfStream)
			line = file.ReadLine
			data = Split(line, ";")
			'Se lee el stock actual
			strSQL="Select * from TBLARTICULOSDATOS2 where IDARTICULO=" & data(0)
			Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "OPEN", strSQL)	
			stockE = CDbl(data(2))
			stockS = CDbl(data(3))			
			if (not rs.eof) then
				'Se calcula el ajuste a realizar.				
				stockE = stockE + CDbl(rs("EXISTENCIA"))				
				stockS =  stockS + CDbl(rs("SOBRANTE"))								
				'Se actualiza el stock					
				strSQL="Update TBLARTICULOSDATOS2 set EXISTENCIA=" & stockE & ", SOBRANTE=" & stockS & " where IDALMACEN=" & idAlmacen & " and IDARTICULO=" & data(0)
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
			else				
				'Se graba el stock
				strSQL="Insert into TBLARTICULOSDATOS2 values(" & data(0) & ", " & stockE & ", " & stockS & ", 'SYNC', " & session("MmtoSistema") & ", " & idAlmacen & ", 0, 0, 0, 0, '')"
				Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
			end if
		wend
	else
		Response.Write "El archivo stocks_" & fileno & ".csv no existe!"
	end if
	Set fso=Nothing
	
	
End Function
'----------------------------------------------------------------------------------------------------------------------
Function createItems()
	filename = Server.MapPath("Documentos\Items_sn.csv")
	filename2 = Server.MapPath("Documentos\Items_added.csv")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(filename)) then
		'Se importan los datos del archivo
		Set file=fso.OpenTextFile(filename,1)
		Set file2=fso.OpenTextFile(filename2,2, true)
		while (not file.AtEndOfStream)
			line = file.ReadLine
			data = Split(line, ";")
			idArti = grabarArticulo(0, data(4), data(5), data(1), "", "N", "", "")
			line = replace(line, "ALTA", idArti)
			file2.WriteLine(line)			
		wend
		file.close
		file2.close
	end if
	Set file=Nothing
	Set file2=Nothing
	Set fso=Nothing
End Function
'----------------------------------------------------------------------------------------------------------------------
Function grabarArticulo(idArticulo, dsArticulo, idCategoria, idUnidad, cdCuenta, bienUso, cdCuentaGastos, cCosto)
	Dim strSQL, rs, conn
		
	'Es una unidad nueva
	strSQL="Insert into TBLARTICULOS(DSARTICULO, IDCATEGORIA, IDUNIDAD, CDCUENTA, BIENUSO, ESTADO, CDCUENTAGASTOS, CCOSTOS, CDUSUARIO, MOMENTO)"
	strSQL= strSQL & " values('" & UCase(dsArticulo) & "', " & idCategoria & ", " & idUnidad & ",'" & cdCuenta & "', '" & bienUso & "', " & ESTADO_ACTIVO & ", '" & cdCuentaGastos & "','" & cCosto & "','SYNC', " & session("MmtoSistema") & ")"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
	strSQL = "Select MAX(IDARTICULO) as IDARTICULO from TBLARTICULOS"
	'Response.Write strSQL & "<br>"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "OPEN", strSQL)	
	grabarArticulo = rs("IDARTICULO")
	
End Function
'----------------------------------------------------------------------------------------------------------------------
Function createFiles()
	index = 0
	filename = Server.MapPath("Documentos\stocks.csv")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists(filename)) then
		'Se importan los datos del archivo
		Set file=fso.OpenTextFile(filename,1)
		while (not file.AtEndOfStream)
			index = index + 1
			filename2 = Server.MapPath("Documentos\stocks_" & index & ".csv")
			Set file2=fso.OpenTextFile(filename2,2, true)
			count = 0
			while (not file.AtEndOfStream) and (count < 500)
				line = file.ReadLine()
				file2.WriteLine(line)
				count = count + 1
			wend
			file2.Close
			Set file2=Nothing
		wend
		Set file=Nothing	
	end if
	Set fso=Nothing
	createFiles = index
End Function
'----------------------------------------------------------------------------------------------------------------------
Function unificar(idAlmacen)
	strSQL="Delete from TBLARTICULOSDATOS where IDALMACEN=" & idAlmacen
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
	strSQL="Insert into TBLARTICULOSDATOS Select * from TBLARTICULOSDATOS2"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "EXEC", strSQL)	
End Function
'*********************************************************************
Dim id, accion, rsDivision, rsAlmacenes

'Se reciben los parametros
accion = GF_PARAMETROS7("accion", 0, 6)
id = GF_PARAMETROS7("id", 0, 6)
pIdVale = GF_PARAMETROS7("vale", 0, 6)
fileno = GF_PARAMETROS7("fileno", 0, 6)
filesCount=GF_PARAMETROS7("filesCount", 0, 6)

'Se inicializan los momentos
Call GP_ConfigurarMomentos()

'Se procesa la acción
Select case (accion)
	case 1:
		Call importPrices(id)
		response.write "Proceso 1 Completo!!<br>"
	case 2:
		Call importStocks(id, fileno)
		response.write "Proceso 2 Completo!!<br>"
	case 3:
		Call createItems()
		response.write "Proceso 3 Completo!!<br>"
	case 4:
		filesCount =  createFiles()
		response.write "Proceso 4 Completo!!<br>"
	case 5:
		Call blanquearStock(id)
	case 6:
		Call ajustar(id, pIdVale)
	case 7:
		Call unificar(id)
End Select
%>
<html>
	<head>
		<script type="text/javascript">
			function runPrices() {
				var cmb = document.getElementById("idDivision")
				location.href = "almacenImportarDatosIniciales.asp?accion=1&id=" + cmb.options[cmb.selectedIndex].value;
			}
			function runStocks(fn, fc) {
				var cmb = document.getElementById("idAlmacen")				
				location.href = "almacenImportarDatosIniciales.asp?accion=2&fileno=" + fn + "&filesCount=" + fc + "&id=" + cmb.options[cmb.selectedIndex].value;
			}
			function runAjuste() {
				var cmb = document.getElementById("idAlmacen")
				var vale = document.getElementById("idVale")
				location.href = "almacenImportarDatosIniciales.asp?accion=6&id=" + cmb.options[cmb.selectedIndex].value + "&vale=" + vale.value;
			}
			function runClear() {
				var cmb = document.getElementById("idAlmacen")				
				location.href = "almacenImportarDatosIniciales.asp?accion=5&id=" + cmb.options[cmb.selectedIndex].value;
			}	
			function runNew() {
				location.href = "almacenImportarDatosIniciales.asp?accion=3";
			}
			function runFiles() {
				location.href = "almacenImportarDatosIniciales.asp?accion=4";
			}
			function runUnify() {
				var cmb = document.getElementById("idAlmacen");				
				location.href = "almacenImportarDatosIniciales.asp?accion=7&id=" + cmb.options[cmb.selectedIndex].value;
			}		
		</script>
	</head>
	<body>
		<%	Set rsAlmacenes = obtenerListaAlmacenes(0)	%>
		<select id="idAlmacen">
			<%	while (not rsAlmacenes.eof)	
					if (id = CInt(rsAlmacenes("IDALMACEN"))) then %>
					<option value="<% =rsAlmacenes("IDALMACEN") %>" selected><% =rsAlmacenes("DSALMACEN") %>
			<%		else %>
					<option value="<% =rsAlmacenes("IDALMACEN") %>"><% =rsAlmacenes("DSALMACEN") %>					
			<%		end if
					rsAlmacenes.MoveNext()
				wend		%>
		</select>				
		<table>
			<% paso = 1%>
			<tr>
				<td colspan="2">
					<% =paso %>º - Determinar categorias de articulos a crear y exportar los datos al archivo Items_sn.csv
				</td>
			</tr>
			<% paso = paso +1 %>			
			<tr>
				<td>
					<% =paso %>º - Crear Articulos nuevos
				</td>
				<td>
					<input type="button" value="Crear Articulos" onClick="javascript:runNew()" id=button1 name=button1>						
				</td>
			</tr>			
			<% paso = paso +1 %>
			<tr>
				<td colspan="2">
					<% =paso %>º - Levantar en Excel el archivo Items_added.csv en una hoja nueva, unificar las columnas correspondientes en la hoja donde estan todos los stock registrados. Exportar la hoja final al archivo stocks.csv
				</td>
			</tr>			
			<% paso = paso +1 %>						
			<tr>
				<td>
					<% =paso %>º - Blanquear Stocks
				</td>
				<td>
					<input type="button" value="Clear" onClick="javascript:runClear()" id=button1 name=button1>						
				</td>
			</tr>
			<% paso = paso +1 %>
			<tr>
				<td>
					<% =paso %>º - Crear Archivos trabajo
				</td>
				<td>
					<input type="button" value="Crear Archivos" onClick="javascript:runFiles()" id=button1 name=button1>						
				</td>
			</tr>							
			<% paso = paso +1 %>					
			<tr>
				<td><% =paso %>º - Importar Archivos										
				</td>
				<td>
					<%	xxx = 1
						while (xxx <= filesCount)
							if (xxx <= fileno) then
								acc = "alert('Ya fue importado')"
							else
								acc = "javascript:runStocks(" & xxx & ", " & filesCount & ")"
							end if
					%>
					
						<input type="button" value="Impotar Archivo <% =xxx %>" onClick="<% =acc %>"><br>
					<%		xxx = xxx + 1
						wend
					%>
				</td>
			</tr>			
			<% paso = paso +1 %>					
			<tr>
				<td>
					<% =paso %>º - Crear Ajuste
				</td>
				<td>
					Vale de Ajuste a Reutilizar: <input type="text" id="idVale" value="0">
					<input type="button" value="Ajustar" onClick="javascript:runAjuste()">						
				</td>
			</tr>		
			<% paso = paso +1 %>						
			<tr>
				<td>
					<% =paso %>º - Unificar Stocks
				</td>
				<td>
					<input type="button" value="Unificar" onClick="javascript:runUnify()" id=button1 name=button1>						
				</td>
			</tr>									
			<tr>
				<td colspan="2"><hr></td>
			</tr>
			<tr>
				<td>
				<%	strSQL="Select * from TBLDIVISIONES"
					Call executeQueryDB(DBSITE_SQL_INTRA, rs, oConn, "OPEN", strSQL)	
				%>
					<select id="idDivision">
					<%	while (not rsDivision.eof)	%>
							<option value="<% =rsDivision("IDDIVISION") %>"><% =rsDivision("DSDIVISION") %>
					<%		rsDivision.MoveNext()
						wend		%>
					</select>
				</td>
				<td>
					<input type="button" value="Impotar Precios" onClick="javascript:runPrices()" id=button2 name=button2>
				</td>
			</tr>
			
		</table>		
	</body>
</html>