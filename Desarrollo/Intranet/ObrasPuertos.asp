<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientossql.asp"--> 
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientospaginacion.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->


<%ProcedimientoControl "ObrasPuert"
dim objFS, baseFolder, subFolder

set objFS = Server.createObject("Scripting.FileSystemObject")
set baseFolder = objFS.getFolder(server.MapPath(".") & "\Obras Puertos")

%>
<html>
<head>
	<title>Obras En Puertos</title>
	<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>
<body>
<%=GF_TITULO("ute.jpg","Fotos  de Obras en Puertos")%>
<table align=center cellpadding=0 cellspacing=2>
<%if baseFolder.subFolders.Count > 0 then
for each subFolder in baseFolder.subFolders%>
	<tr>	
		<td align="left"><a href="FotosObras.asp?strFolderName=<%=subFolder.name%>"><img src="images/image8.gif" width=10 align="absmiddle">&nbsp;<%=subFolder.name%></a><br></td>
	</tr>
<%next
else%>
	<br><div align=center">No hay carpetas para explorar</div>	
<%end if%>
</body>
</html>