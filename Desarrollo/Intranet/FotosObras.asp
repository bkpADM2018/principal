<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientossql.asp"--> 
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->

<%
dim objFS
dim BaseFolder
dim strFolderName, Page
dim File

strFolderName = GF_Parametros7("strFolderName","",6)
Page = GF_Parametros7("Page",0,6)
if Page = 0 then Page = 1

set objFS = Server.createObject("Scripting.FileSystemObject")
set baseFolder = objFS.getFolder(server.MapPath(".") & "\Obras Puertos\" & strFolderName)

%>
<html>
<head>
	<title>Fotos en <%=strFolderName%></title>
	<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>
<body>
<%=GF_TITULO("ute.jpg","Fotos en " & baseFolder.Name)%>
<table align=center cellpadding=0 cellspacing=2>
<%if baseFolder.Files.Count > 0 then%>
<TABLE width=100%>
	<tr>
		<td align=right>
			Paginas:
			<%if baseFolder.Files.Count mod 4 > 0 then 
				cantPaginas = baseFolder.Files.Count/4 + 1
			else
				cantPaginas = baseFolder.Files.Count/4
			end if
			for k = 1 to cantPaginas
				if k <> Page then%>
					<a href="FotosObras.asp?strFolderName=<%=strFolderName%>&Page=<%=k%>"><%=k%></a>
				<%else%>
					<font color=red>[<%=k%>]</font>
				<%end if
			next%>
		</td>
	</tr>
<table>
<br>
<table align=center cellpadding=0 cellspacing=2 width="650">
	<tr>
		<td align=center>
			<%indice = 1
			for each Foto in baseFolder.Files
				if indice > 4 * (Page - 1) and indice <= 4 * Page then%>
					<img src="obras Puertos/<%=baseFolder.name%>/<%=Foto.name%>" width=300>
				<%end if	
				indice = indice + 1
			next%>
		</td>
	</tr>
</table>
<%else%>
	<br><div align=center>No hay imagenes para mostrar</div>
<%end if%>
</body>
</html>