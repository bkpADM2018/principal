<!--#include file="Includes/procedimientosMG.asp"-->
<%

Dim folder, theFileName, pathFinal, resp, fs, fo, x
if session("UM_ERROR") <> "" then
	resp = "ERROR|" & session("UM_ERROR")
	session("UM_ERROR") = ""
else
	accion=GF_PARAMETROS7("accion","",6)
	theFileName = GF_PARAMETROS7("file","",6)
	folder=GF_PARAMETROS7("folder","",6)
	folder = Replace(folder, "$", "\")
	if (Right(folder,1) <> "\") then folder = folder & "\"

	if ((InStr(folder, ":\") = 0) and (InStr(folder, "\\") = 0)) then
		'Es un path relativo!!
		if (Left(folder,1) <> "\") then folder = "\" & folder
		pathFinal = server.mappath(".") & folder
	else
		'Es un path absoluto.
		pathFinal = folder
	end if
	
	resp = "STILL WORKING|"
	Set fs = CreateObject("Scripting.FileSystemObject")
	if (fs.FileExists(pathFinal & theFileName)) then				
		if (accion = "upload") then resp = "DONE|"	
	else
		if (accion = "remove") then resp = "DONE|"
	end if
end if	
'response.write pathFinal & theFileName & "|" & accion & "|" & fs.FileExists(pathFinal & theFileName) & "|"
response.write resp
%>