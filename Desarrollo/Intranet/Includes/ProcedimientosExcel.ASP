<% 
dim fso 'objeto fileSystem
dim act 'objeto de Archivo
dim xls_mode
'--> Variables de uso interno del objeto - NO USAR FUERA DE ESTE ARCHIVO!!
Dim xls_DataTitleOpen, xls_DataTableOpen, xls_DataRowOpen
'<--

const XLS_FILE_MODE = 0
const XLS_STREAM_MODE = 1
const XLS_HTML_MODE = 2

xls_mode = XLS_STREAM_MODE 'valor por defecto

'----------------------------------------------------------------------------------------
Function GF_createXLS(filename)
        Dim xlsPath
        
        'Garantizo que tenga la extencion
        if (Ucase(Right(filename, 4)) <> ".XLS") then filename = filename & ".xls"
        xlsPath = getTempFilePath() & "\" & filename
        
		if xls_mode = XLS_STREAM_MODE then
			'inicia el archivo para que abra del lado del cliente
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "filename=" & filename 
		end if
		if xls_mode = XLS_FILE_MODE then
			'inicia el archivo para que se grabe en el servidor			
			Set fso = createobject("scripting.filesystemobject")
			Set act = fso.CreateTextFile(xlsPath, true)
		end if
		GF_createXLS = xlsPath
End Function
'-----------------------------------------------------------------------------------------------------------
Function getTempFilePath()
	Dim path, baseBath, found, fso, curr, folder
	
	folder      = "\Temp"  ' nombre de la carpeta contenedora
	
	Set fso = CreateObject("Scripting.FileSystemobject")
	
	'Path de la raiz del sitio web, si se llega a este y no se encontr� debe abandonar la b�squeda y devolver error.
	basePath = Server.MapPath("/") & folder
	path = Server.MapPath(".") & folder
	found = false
	while ((not found) and (path <> basePath))
		if (fso.FolderExists(path)) Then
			found = true
		else
			curr = curr & "../"
			path = Server.MapPath(curr) & folder
		end if
	wend
	getTempFilePath = path 
End Function
'----------------------------------------------------------------------------------------
Function GF_setXLSMode(xls_accion)
	xls_mode = xls_accion 'setea el modo	
End Function
'----------------------------------------------------------------------------------------
Sub writeXLS(stringToWrite)
	if (xls_mode = XLS_STREAM_MODE) or (xls_mode = XLS_HTML_MODE) then 
		Response.write stringToWrite
	else
		act.WriteLine(stringToWrite)
	end if
End sub
'----------------------------------------------------------------------------------------
Sub closeXLS()
	if xls_mode = XLS_FILE_MODE then
		Set act = Nothing 
		Set fso = Nothing 
	end if
End sub
'----------------------------------------------------------------------------------------
Function XLS_Open(filename)
	'Crea el archivo.
	XLS_Open = GF_createXLS(filename)
	Call writeXLS("<html>")
	Call writeXLS("<head>")
	Call writeXLS("	<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />")
	Call writeXLS("	<style>")
	Call writeXLS("		body		{ background-color:#ffffff; border:none; }	")		
	Call writeXLS("		table							{ background-color:#fff; font-family:Arial, Helvetica, sans-serif; font-style:normal; font-weight:normal; border-collapse:collapse;}")
	Call writeXLS("		table thead th 					{ background-color:#396E8F; border-left:1px solid #fff; padding:5px 10px; color:#fff; font-size:11px; font-weight:bold;}")
	Call writeXLS("		table thead th:first-child 		{ border:none; -webkit-border-radius:8px 0 0 0; -moz-border-radius:8px 0 0 0; border-radius:8px 0 0 0; }")
	Call writeXLS("		table thead th:last-child 		{ -webkit-border-radius:0 8px 0 0; -moz-border-radius:0 8px 0 0; border-radius:0 8px 0 0; }")
	Call writeXLS("		table thead td 					{ background-color:#396E8F; border-left:1px solid #fff; border-top:1px solid #fff; padding:5px 10px; color:#fff; font-size:11px; font-weight:bold; }")
	Call writeXLS("		table tbody td 		{ font-size:11px; } ")
	Call writeXLS("		.tabletxtfijo		{ font-size:11px; font-style:italic; color:#9b9b9b; text-align:right; padding-right:10px; }")
	Call writeXLS("		table tbody .rtotal	td		{ color:#fff; font-size:12px; font-weight:bold; }")
	Call writeXLS("		.tableaside h3 a,")
	Call writeXLS("		.tableaside h3	 		{ font-size:12px; color:#333; text-transform:uppercase; font-weight:bold !important; }")
	Call writeXLS("		.tableaside h4 			{ font-size:11px; color:#666; text-transform:capitalize; }")
	Call writeXLS("		.tableasidecontent		{ font-family:Arial, Helvetica, sans-serif; font-size:11px; color:#666; }")
	Call writeXLS("		.tableasidecontent span	{ font-size:10px; font-style:italic; }")
	Call writeXLS("		.datagrid		{ border-collapse:collapse;	text-align:left; margin-top:20px; }")
	Call writeXLS("		.datagrid td,")
	Call writeXLS("		.datagrid th 	{ padding:2px 5px; }")
	Call writeXLS("		.datagrid tbody td 						{ border-left:1px solid #CCCCFF; }")
	Call writeXLS("		.datagrid tbody td:first-child 			{ border-left:none; }")
	Call writeXLS("		.datagrid tbody tr:nth-child(odd) 		{ background-color:#F4F4F4; } ")
	Call writeXLS("		.datagridlv1 tbody tr:nth-child(odd) 	{ background-color:#fff !important; } ")
	Call writeXLS("		.datagridlv1 tbody tr:nth-child(4n+1) 	{ background-color:#F4F4F4 !important; } ")
	Call writeXLS("		.datagridlv1 tbody tr:nth-child(4n+2) 	{ background-color:#F4F4F4 !important; } ")
	Call writeXLS("		.datagrid tbody td 						{ border-bottom:1px solid #CCCCFF; } ")
	Call writeXLS("		.datagrid tbody tr:hover 					{ color: #FFFFFF; background-color:#80A2B7 !important; } ")
	Call writeXLS("		.datagridlv2 td							{ background-color:#fff !important; } ")
	Call writeXLS("	</style>")
	Call writeXLS("</head>")	
	Call writeXLS("<body>")	
	xls_DataTableOpen = 0
End Function
'----------------------------------------------------------------------------------------
'Agrega la linea de titulos de la tabla.
' Nomenclaruta esperada:
'	pTitlesVec ::= <valor><pSeparator><valor><pSeparator>...
'	pSeparator default ','
Function XLS_AddDataTitle(pTitlesVec, pSeparator)
	Dim mySep, titulos, i
	
	if (xls_DataTableOpen = 0) then 
		Call writeXLS("<table class='datagrid' width='100%' align='center'>	")
		xls_DataTableOpen = 1
	end if
	mySep = ","
	if (pSeparator <> "") then mySep = pSeparator	
	titulos = Split(pTitlesVec, mySep)
	
	Call writeXLS("<thead>")
	Call writeXLS("<tr>")
	For each i in titulos
		Call writeXLS("<th>" & Trim(i) & "</th>")		
	Next
	Call writeXLS("</tr>")
	Call writeXLS("</thead>")
End Function
'----------------------------------------------------------------------------------------
'Agrega una linea en la tabla.
' Nomenclaruta esperada:
'	pData ::= <valor>[:alineacion]<pSeparator><valor>[:alineacion]<pSeparator>...
'	Si no se incluye alineacion, se asume 'Left'
'	pSeparator default ','
Function XLS_AddDataRow(pData, pSeparator)
	Dim mySep, datos, i, myAlign, data, pos
	
	if (xls_DataTableOpen = 0) then 
		Call writeXLS("<table class='datagrid' width='100%' align='center'>	")
		xls_DataTableOpen = 1
	end if
	if (xls_DataRowOpen = 0) then 
		Call writeXLS("	<tbody>")
		xls_DataRowOpen = 1	
	end if
	mySep = ","
	if (pSeparator <> "") then mySep = pSeparator	
	datos = Split(pData, mySep)
		
	Call writeXLS("<tr>")
	For each i in datos		
		pos = InStrRev(i, ":")
		if (pos > 0) then
			data = Trim(Left(i, pos-1))
			myAlign = Trim(Right(i, Len(i)-pos))			
		else
			data = i
			myAlign = "left"
		end if				
		Call writeXLS("	<td align='" & myAlign & "'>" & data & "</td>")		
	Next
	Call writeXLS("</tr>")
	
End Function
'----------------------------------------------------------------------------------------
Sub XLS_closeXLS()	
	'Si se abrio una tabla de datos, se cierran todas las estructuras.
	if (xls_DataRowOpen = 1) then Call writeXLS("<tbody>")
	if (xls_DataTableOpen = 1) then Call writeXLS("</table>")		
	Call writeXLS("</body>")
	Call writeXLS("</html>")
	Call closeXLS()	
End sub
%>