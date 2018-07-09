<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->

<%

Const UP_ACTION_COUNT = "count"
Const UP_ACTION_FILES = "files"
'--------------------------------------------------------------------------------------------------
Function processFolder(pFolder,pAction)
	dim rtrn,objFS,baseFolder
	
	if (pAction = UP_ACTION_COUNT) then rtrn = 0
	if (pAction = UP_ACTION_FILES) then rtrn = ""
	
	set objFS = Server.createObject("Scripting.FileSystemObject")
	if (objFS.FolderExists(pFolder)) then 
		'Obtengo los nombres de los archivos de la carpeta
		set baseFolder = objFS.getFolder(pFolder)
		
		for each myFile in baseFolder.Files
			if (extensionHabilitada(objFS.GetExtensionName(Server.MapPath("/")&myFile.name))) then
				select case pAction
					case UP_ACTION_COUNT
						rtrn = rtrn +1
					case UP_ACTION_FILES
						rtrn = rtrn & myFile.name & ","
				end select
			end if
		next	
	end if
	
	if (pAction = UP_ACTION_FILES and len(rtrn)>0) then rtrn = left(rtrn,len(rtrn)-1)
	
	processFolder = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function getCantFiles(pFolder)
	getCantFiles = processFolder(pFolder,UP_ACTION_COUNT)
End Function
'--------------------------------------------------------------------------------------------------
Function getFiles(pFolder)
	getFiles = processFolder(pFolder,UP_ACTION_FILES)
End Function
'--------------------------------------------------------------------------------------------------

Function extensionHabilitada(ext)
	extensionHabilitada = true
	ext= LCase(ext)
	if 	(ext <> "doc") and (ext <> "DOC") and _		
		(ext <> "pdf") and (ext <> "PDF") and _
		(ext <> "xls") and (ext <> "XLS") and _
		(ext <> "txt") and (ext <> "TXT") and _
		(ext <> "gif") and (ext <> "GIF") and _
		(ext <> "jpg") and (ext <> "JPG") and _
		(ext <> "png") and (ext <> "PNG") and _
		(ext <> "tif") and (ext <> "TIF") and _
		(ext <> "zip") and (ext <> "ZIP") and _
		(ext <> "rar") and (ext <> "RAR") and _		
		(ext <> "msg") and (ext <> "MSG") and _	
		(ext <> "xml") and (ext <> "XML") and _
		(ext <> "csv") and (ext <> "CSV") and _
		(ext <> "xlsx") and (ext <> "XLSX") and _
		(ext <> "docx") and (ext <> "DOCX") and _		
		(ext <> "rtf") and (ext <> "RTF") then
		extensionHabilitada = false
	end if
End Function

Function subirArchivo(pathDestino) 
	Dim success
	
	On Error Resume Next

	success = true
	
	Call logInfo("Subiendo Archivo...")
	
	forWriting = 2
	adLongVarChar = 201
	lngNumberUploaded = 0
	
	Call logInfo("Get binary data from form")
	noBytes = Request.TotalBytes		
	binData = Request.BinaryRead(noBytes)
	Call logInfo("Bytes Recibidos:" & noBytes)
	Call logInfo("convert the binary data to a string")
	Set RST = CreateObject("ADODB.Recordset")
	LenBinary = LenB(binData)
	if LenBinary > 0 Then
		RST.Fields.Append "myBinary", adLongVarChar, LenBinary
		RST.Open
		RST.AddNew
		RST("myBinary").AppendChunk BinData
		RST.Update
		strDataWhole = RST("myBinary")		
	End if
	Call logInfo("Bytes convertidos a String OK!")
	strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
	lngBoundryPos = instr(1,strBoundry,"boundary=") + 8
	strBoundry = "--" & right(strBoundry,len(strBoundry)-lngBoundryPos)
	Call logInfo("Get first file boundry positions.")
	lngCurrentBegin = instr(1,strDataWhole,strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
	Do While lngCurrentEnd > 0
		Call logInfo("Get the data between current boundry and remove it from the whole.")
		strData = mid(strDataWhole,lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
		strDataWhole = replace(strDataWhole,strData,"")
		Call logInfo("Get the full path of the current file.")
		lngBeginFileName = instr(1,strdata,"filename=") + 10
		lngEndFileName = instr(lngBeginFileName,strData,chr(34))
		Call logInfo("Make sure they selected at least one file.")
		if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then
			msg = "Debe seleccionar un archivo"			
			exit do
		End if
		Call logInfo("There could be one or more empty file boxes.")
		if lngBeginFileName <> lngEndFileName Then
			strFilename = mid(strData,lngBeginFileName,lngEndFileName - lngBeginFileName)
			Call logInfo("Loose the path information and keep just the file name.")
			tmpLng = instr(1,strFilename,"\")
			Do While tmpLng > 0
				PrevPos = tmpLng
				tmpLng = instr(PrevPos + 1,strFilename,"\")
			Loop

			FileName = right(strFilename,len(strFileName) - PrevPos)			
			Call logInfo("Get the begining position of the file data sent.")
			'if the file type is registered with the
			' browser then there will be a Content-Type
			lngCT = instr(1,strData,"Content-Type:")

			if lngCT > 0 Then
				lngBeginPos = instr(lngCT,strData,chr(13) & chr(10)) + 4
			Else
				lngBeginPos = lngEndFileName
			End if
			Call logInfo("Get the ending position of the file data sent.")
			lngEndPos = len(strData)

			Call logInfo("Calculate the file size.")
			lngDataLenth = lngEndPos - lngBeginPos
					
			dim vNombre, vExtension, vLenExt
			Call logInfo("Nombre Archivo:" & FileName)
			vNombre = left(FileName, InStrRev(FileName,".")-1)
			vLenExt = Len(FileName) - Len(vNombre) - 1 'Extencion = Todo - Nombre - punto
			vExtension = right(FileName, vLenExt)
			strNombreTodosArchivos = strNombreTodosArchivos & " " &  FileName
			if (len(vNombre) < 1) then
			   msg = "El nombre del archivo '" & FileName & "' no es valido"
			else
				if (not extensionHabilitada(vExtension)) then
					msg = "La extension del archivo '" & FileName & "' no es valida."
				end if
			end if
			
			Call logInfo("Controlar longitud de archivo")
			if lngDataLenth = 0 then
				msg = msjError & "El archivo " & FileName & " no existe o su tamaño es nulo"
			end if
			'Si hubo errores, sale de ciclo
			if msg <> "" then
				exit do
			end if
			
			Call logInfo("Get the file data")
			strFileData = mid(strData,lngBeginPos,lngDataLenth)
			Call logInfo("Create the file.")
			Set fso = CreateObject("Scripting.FileSystemObject")
			'response.write server.mappath("..") & PATH_DESTINO & FileName
			If (not fso.FolderExists(pathDestino)) Then
				'Se crea el directorio destino

				fso.CreateFolder(pathDestino)
			end if
			Set f = fso.OpenTextFile(pathDestino & FileName, forWriting, True)
			f.Write strFileData
			Set f = nothing
			Set fso = nothing

			lngNumberUploaded = lngNumberUploaded + 1

		End if

		Call logInfo("Get then next boundry postitions if any.")
		lngCurrentBegin = instr(1,strDataWhole,strBoundry)
		lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
	loop
	Call logInfo("Se devuelve el link virtual")
	'subirArchivo = "http://" & Request.ServerVariables("SERVER_NAME") & PATH_WEB & FileName
	'subirArchivo = server.mappath(".") & PATH_DESTINO & FileName	
	subirArchivo = pathDestino & FileName	
			
	fileType = CHR(034)&"type"&CHR(034)&":"&CHR(034)&vExtension&CHR(034)
	fileSize = CHR(034)&"size"&CHR(034)&":"&CHR(034)&CLng( lngDataLenth ) &CHR(034)
	fileName = CHR(034)&"name"&CHR(034)&":"&CHR(034)&fileName&CHR(034)
	if (msg = "") then
		result = CHR(034)&"success"&CHR(034)&":"&CHR(034)&"true"&CHR(034)	
	else
		result = CHR(034)&"error"&CHR(034)&":"&CHR(034)&msg&CHR(034)
	end if
	'response.write "{success:true,fileName:"&fileName&"}" '"success"  " '"{"&fileName&","&fileType&","&fileSize&"}"
	If Err.Number <> 0 Then
  
	  result = CHR(034)&"error"&CHR(034)&":"&CHR(034)&Err.Description&CHR(034)
	End If
	Call logInfo("Return:" & "{"&result&","&fileName&","&fileType&","&fileSize&"}")
	response.write "{"&result&","&fileName&","&fileType&","&fileSize&"}"
	response.end
End Function
'**************************************************
'*****	COMIENZO DE PAGINA

Call startLog(HND_FILE+HND_VIEW,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)

Dim theFileName, folder, dltFile, fs, accion, pathFinal, pathFinalRemove
dim msg

Call logInfo("*********** SE INICIA UPLOAD ***********")

accion=GF_PARAMETROS7("accion","",6)
folder=GF_PARAMETROS7("folder","",6)
folder = Replace(folder, "$", "\")

if (Right(folder,1) <> "\") then folder = folder & "\"


if ((InStr(folder, ":\") = 0) and (InStr(folder, "\\") = 0) or ((left(folder,1) = "\") or (InStr(folder, "..") <> 0))) then
	'Es un path relativo!!	
	if ((InStr(folder, "..") <> 0) or (left(folder,1) = "\")) then
		pathFinal = server.mappath(folder)
		if (Right(pathFinal,1) <> "\") then pathFinal = pathFinal & "\"
		
	else
		if (Left(folder,1) <> "\") then folder = "\" & folder
		pathFinal = server.mappath(".") & folder
		
	end if

else
	'Es un path absoluto.
	pathFinal = folder
end if

Call logInfo("Operacion: " & accion)
Call logInfo("Directorio Destino:" & pathFinal)
Set fs = CreateObject("Scripting.FileSystemObject")
'set file= fs.CreateTextFile(server.MapPath(".") & "\LOGUPLOAD.TXT", true)
if (accion = "upload") then	
	'Se sube el archivo a una carpeta exclusiva para este pedido.	
	serverPath = subirArchivo(pathFinal)		
else	
	if (accion = "delete") then
		theFileName = GF_PARAMETROS7("file","",6)		
		'file.WriteLine(pathFinal & theFileName)
		
		if fs.FileExists(pathFinal & theFileName) then
			fs.DeleteFile(pathFinal & theFileName)		
		end if	
	end if
	
	if (accion = "cant")  then response.write getCantFiles(pathFinal)
	if (accion = "files") then response.write getFiles(pathFinal)
	
end if
Call logInfo("*********** FIN UPLOAD ***********")
session("UM_ERROR") = msg
'file.Close()
'Set file = nothing
Set fs = Nothing
%>