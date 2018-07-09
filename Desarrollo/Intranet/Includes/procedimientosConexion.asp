<%
'*********************************************************************************************************************
'**********************************************   ENCRIPTADOR RC4  ***************************************************
'*********************************************************************************************************************
Function algorithmRC4(inp , Key) 
    Dim S(255), K(255),i 
    Dim j , temp, y , t , x 
    Dim Outp 
    
    For i = 0 To 255
        S(i) = i
    Next
    
    j = 1
    For i = 0 To 255
        If j > Len(Key) Then j = 1
        K(i) = Asc(Mid(Key, j, 1))
        j = j + 1
    Next     
    j = 0
    For i = 0 To 255
        j = (j + S(i) + K(i)) Mod 256
        temp = S(i)
        S(i) = S(j)
        S(j) = temp
    Next
    i = 0
    j = 0
    For x = 1 To Len(inp)
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        temp = S(i)
        S(i) = S(j)
        S(j) = temp
        t = (S(i) + (S(j) Mod 256)) Mod 256
        y = S(t)
        Outp = Outp & Chr(Asc(Mid(inp, x, 1)) Xor y)
    Next
    algorithmRC4 = Outp
End Function
'--------------------------------------------------------------------------------------------
Function encodeRC4(pValue)	
    Dim i, singleValue, str
    For i = 1 To Len(pValue)
        singleValue = Hex(Asc(Mid(pValue, i, 1)))
        If(Len(singleValue) = 1)Then singleValue = "0" & singleValue
        str = str & singleValue        
    Next
    encodeRC4 = str
End Function
'--------------------------------------------------------------------------------------------
' Función:	getConfigFilePath
' Autor: 	JAS - Javier A. Scalisi
' Fecha: 	10/05/13
' Objetivo:	
'			A partir del directorio en el que se encuentra la página llamadora, busca el archivo de configuración solicitado.
'			Si no esta en un directorio, busca en el inmediato superior. Cuando lelga al directorio raiz para y devuelve ese path.
' Parametros:
'			filename 	[string] 	Raiz de donde proviene la conexion.
'--------------------------------------------------------------------------------------------
Function getPassFilePath()
	Dim folderName, path, baseBath, found, fso, curr, folder
	
	folder      = "\ToepferPass"  ' nombre de la carpeta contenedora
	
	Set fso = CreateObject("Scripting.FileSystemobject")
	
	'Path de la raiz del sitio web, si se llega a este y no se encontró debe abandonar la búsqueda y devolver error.
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

	getPassFilePath = path & folderName
	
End Function
'--------------------------------------------------------------------------------------------
' Función:	loadConfigFile
' Autor: 	CNA - Nahuel Ajaya
' Fecha: 	10/05/13
' Objetivo:	
'			Lee desde el archivo txt los datos del Usuario,Clave y Alias. Luego los carga en una variable de session
' Parametros:
'			pConnection 	[string] 	Raiz de donde proviene la conexion
' Devuelve:
'			En caso de no encontrar el archivo , lo informa
'--------------------------------------------------------------------------------------------
Function loadConfigFile(pConnection)
	Dim fso, f, sText, z, file ,pathAaccess
	Set fso = CreateObject("Scripting.FileSystemobject")			
	
	file        = getPassFilePath() & "\ToepferPass" & pConnection & ".txt"	' nombre del archivo	
	'Response.Write file
	If fso.FileExists(file) Then
		Set f = fso.OpenTextFile(file, 1)
		if not f.AtEndOfStream then
			line = f.ReadLine
			str = Chr(210) & Chr(182) & Chr(115) & Chr(116) & Chr(126) & Chr(39) & Chr(7) & Chr(75) & Chr(130) & Chr(191) & Chr(140) & Chr(6) & Chr(175) & Chr(94) & Chr(129) & Chr(249)			
			aux = algorithmRC4(DecodeRC4(Trim(line)), str)			
			'Se genera la variable para la nueva modalidad donde se guarda el connection string directamente sin ODBC
			session("conn" & pConnection &  "cs")  = aux
			'Variables para el viejo sistema con ODBC.
			strCadena = Split(aux,";")			
			session("conn" & pConnection &  "User")  = strCadena(0)
			session("conn" & pConnection &  "Key")   = strCadena(1)
			session("conn" & pConnection &  "Alias") = strCadena(2)
		End If
		f.Close
		Set f = Nothing
		Set fso = Nothing
	else
		Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR EN EL ARCHIVO DE CONEXION</font><b><br>POR FAVOR REVISE SI EXISTE EL ARCHIVO: " & file & "<Hr>" 
		err.Clear
		response.end
	end if
End Function
'---------------------------------------------------------------------------------------------------
Function DecodeRC4(pValue)
    Dim i, rtrn    
    For i = 1 To Len(pValue) Step 2		
        rtrn = rtrn & Chr(CLng("&H" & Mid(pValue, i, 2)))
    Next    
	DecodeRC4 = rtrn
End Function
'----------------------------------------------------------------
function GF_BD_Control(byref rs,byref con,P_oprc,byref P_strSQL)
'Está función genera la conexión con la base de datos 
on error resume next
GF_BD_Control = false
call executeQueryDb(DBSITE_SQL_INTRA, rs, P_oprc, P_strSQL)
	   GF_BD_CONTROL = true
end function
'--------------------------------------------------------------------------------------
'Funcion executeSP_Puertos
'Descripción:   Función que ejecuta un store procedure para la base de datos SQL Server - Puertos.
'Parametros :
'           pRecordset      : Recordset donde se devuelven los datos de la consulta
'           pNameSP         : Nombre del procedimiento a ejecutar
'           pParametersInput: Lista de parametros
'                                  pParametersInput ::= <PARAM_IN>[$$<PARAM_OUT>]
'                                  PARAM_IN = valor1||valor2||...||valorN
'                                  PARAM_OUT = key1||key2||...||keyN
'                             Los parametros de salida son opcionales, se devuelven en un diccionario cuyas key son las indicadas.
'                             El diccionario se carga con dos valores por defecto que son el id y descripción del error.
'Devuelve   : El recordset cargado con los registros y un diccionario con los parametros de salida.(SP_IDERROR y SP_DSERROR)
function executeSP_Puertos(byref pRecordset, byVal pPto, byVal pNameSP, byVal pParametersInput)
    On Error Resume Next
    Dim params, index, rtrn, size, idx, outParams, inParams

    if(IsEmpty(session("connPUERTO" & pPto &  "CS"))) then	Call loadConfigFile("PUERTO" & pPto)
    
    Set pRecordset = server.CreateObject("ADODB.Recordset")
    pRecordset.CursorType = 3 'adOpenStatic
    pRecordset.LockType = 3 'adLockOptimistic

    Set con = server.CreateObject("ADODB.connection")
    con.CursorLocation = 3 'adUseClient
    
    con.open session("connPUERTO" & pPto &  "CS")

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = con
    cmd.CommandText = pNameSP
    cmd.CommandType = 4 'adCmdStoredProc
    cmd.Parameters.Refresh
    if pParametersInput <> "" then    
        params = split(pParametersInput,"$$")	
        inParams = split(params(0),"||")	
        if (uBound(params) = 1) then outParams = split(params(1),"||")	
	    for index=0 to ubound(inParams)
            'Para SQL Server el primer indice debe arrancar en 1
            cmd.Parameters(Cint(index) + 1) = CStr(inParams(index))
	    next 
    end if
    Set pRecordset = cmd.Execute    
    Set rtrn = Server.CreateObject("Scripting.Dictionary")
    'Si se esperan parametros de salida se reciben y se cargan al diccionario.    
    if (isArray(outParams)) then
        For idx = LBound(outParams) to UBound(outParams)        
            rtrn.Add outParams(idx), cmd.Parameters(index)        
            index = index + 1
        Next
    end if    
    rtrn.Add SP_IDERROR, cmd.Parameters(index)
    rtrn.Add SP_DSERROR, cmd.Parameters(index + 1)
    if err.number <> 0 and err.number <> 424 and err.number<>3265 then
	    Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>POR FAVOR REVISE EL SQL<Hr>" & pNameSP & "<hr>Error:" & err.number   & err.Description  & "<HR>idError: " & rtrn(SP_IDERROR) & "<BR>dsError: " & rtrn(SP_DSERROR) & "<HR>"
	    err.Clear 
	    response.end 
    end if
    Set executeSP_Puertos = rtrn
end function
'--------------------------------------------------------------------------------------
function executeQueryDb(pDbSite, byref rs, P_oprc, P_strSQL)
    Dim con
    'Está función genera la conexión con la base de datos 
    on error resume next
    'response.Write "<BR>sql-->" & P_strSQL & " OP(" & P_oprc & ")"
    executeQueryDb = false
    session("strSQL")=P_strSQL
    if instr(cstr(p_strsql),"TOEPFERDB")>0 then 
		Response.Write "sql(" & p_strsql & ")"
		Response.End 
	end if	
    if (IsEmpty(session("conn" & pDbSite & "CS"))) then Call loadConfigFile(pDbSite)
	if p_oprc = "CLOSE" THEN
		rs.close
		'con.close
		set con = nothing
		executeQueryDb = true
	end if
	if P_OPRC = "OPEN" or P_OPRC = "UPDATE" then
		set con = server.CreateObject("ADODB.connection")
		set rs = server.CreateObject("ADODB.Recordset")						
		con.open session("conn" & pDbSite & "CS")		
		if (con.State = 0) then
			Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>La conexión no esta abierta! (Operacion OPEN - State = " & con.State & ")"
			response.end 
		end if
		rs.CursorLocation = 3
		rs.Open p_strSQL, con, 2, 3, 1
					
		'Response.Write rs.eof
		executeQueryDb = true
	end if
'Se ejecuta una sentencia strSQL sobre la base. 
if ((P_OPRC = "EXECUTE") or (P_OPRC = "EXEC")) and (P_strSQL <> "") then 
	   Set con = server.CreateObject("ADODB.connection")                              
       con.CommandTimeout = 500
       con.open session("conn" & pDbSite & "CS")	   
	   if (con.State = 0) then
			Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>La conexión no esta abierta! (Operacion: EXEC - State = " & con.State & ")"
			response.end 
		end if
       con.execute P_strSQL
   con.close
   executeQueryDb = true
end if  
	if err.number <> 0 and err.number <> 424 then
		Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>POR FAVOR REVISE EL SQL<Hr>" & P_strSQL & "<hr>Error:" & err.number   & err.Description 
		err.Clear 
		response.end 
	end if
end function
'--------------------------------------------------------------------------------------
function executeProcedureDb(pDbSite, byref rs, pNameSP, pParametersInput)
    On Error Resume Next
    Dim params, index, rtrn, size, idx, outParams, inParams

    if(IsEmpty(session("conn" & pDbSite &  "CS"))) then	Call loadConfigFile(pDbSite)
    
    Set rs = server.CreateObject("ADODB.Recordset")
    rs.CursorType = 3 'adOpenStatic
    rs.LockType = 3 'adLockOptimistic

    Set con = server.CreateObject("ADODB.connection")
    con.CursorLocation = 3 'adUseClient
    
    con.open session("conn" & pDbSite &  "CS")

	if (con.State = 0) then
		Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>La conexión no esta abierta! (Operacion: SP - State = " & con.State & ")"
		response.end 
	end if
		
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = con    
    cmd.CommandText = pNameSP
    cmd.CommandType = 4 'adCmdStoredProc
    cmd.Parameters.Refresh
    index = 1
    if pParametersInput <> "" then    
        params = split(pParametersInput,"$$")	
        inParams = split(params(0),"||")	
        if (uBound(params) = 1) then outParams = split(params(1),"||")	        
	    for idx=0 to ubound(inParams)	        
            cmd.Parameters(Cint(index)) = CStr(inParams(idx))
            index = index + 1
	    next 	    
    end if
    Set rs = cmd.Execute    
    Set rtrn = Server.CreateObject("Scripting.Dictionary")
    'Si se esperan parametros de salida se reciben y se cargan al diccionario.    
    if (isArray(outParams)) then
        For idx = LBound(outParams) to UBound(outParams)        
            rtrn.Add outParams(idx), cmd.Parameters(index)        
            index = index + 1
        Next
    end if    
    rtrn.Add SP_IDERROR, cmd.Parameters(index)
    rtrn.Add SP_DSERROR, cmd.Parameters(index + 1)
    if err.number <> 0 and err.number <> 424 and err.number<>3265 then
	    Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>POR FAVOR REVISE EL SQL<Hr>" & pNameSP & "<hr>Error:" & err.number   & err.Description  & "<HR>idError: " & rtrn(SP_IDERROR) & "<BR>dsError: " & rtrn(SP_DSERROR) & "<HR>"
	    err.Clear 
	    response.end 
    end if
    Set executeProcedureDb = rtrn
end function
'--------------------------------------------------------------------------------------
%>