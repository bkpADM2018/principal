<!--#include file="procedimientosMG.asp"-->
<!--#include file="procedimientosConexion.asp"-->
<%
Const DATA_NAME 		= 10
Const DATA_TELEPHONE 	= 9
Const DATA_MOBILE 		= 8
Const DATA_MAIL 		= 7
Const DATA_COMPANY 		= 6
Const DATA_TITLE 		= 5
Const DATA_DEPARTMENT 	= 4
Const DATA_USER 		= 3
Const DATA_SN			= 2
Const CONEXION_USER = "USER"
'-----------------------------------------------------------------------------------------
Function getUserDescription(username)
	dim rtrn, rs, strSQL
		
	rtrn = ""
	if (username <> "") then
		rtrn = "#ERROR_USUARIO#"
		strSQL="Select Apellido, Nombre from Profesionales Pro inner join Personas Per on Pro.IdProfesional=Per.IdPersona where Pro.CDUsuario='" & Trim(username) & "'"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
		if (not rs.eof) then
		    rtrn = Trim(rs("Apellido")) & ", " & Trim(rs("Nombre"))
		else
			rtrn = UCase(username)			
		end if	
	end if
	getUserDescription = rtrn	
End Function
'-----------------------------------------------------------------------------------------
Function getUserHKey(pUser)
	Dim strSQL,rs,rtrn
	strSQL = "Select * from TOEPFERDB.TBLREGISTROFIRMAS where CDUSUARIO='" & UCase(pUser) & "'"
	call executeQuery(rs, "OPEN", strSQL)
	
	rtrn = ""
	if (not rs.EoF) then rtrn = rs("HKEY")
	
	getUserHKey = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function getUserMail(username)
	
	On Error Resume Next
	
	err.clear	
	rtrn = ""
	if (username <> "") then	
		data = getData(username)
		if (ubound(data)>0) then			
			rtrn = replace(data(0,DATA_MAIL), "'", "")			
		end if
	end if
	
	if (err.Number > 0) then rtrn = ""
	
	getUserMail = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function getData(pUserName)
	Dim objDomain,objADsPath,objConn,objCom,users,myData,vData(),i,j

	redim vData(0,0)
	
	users = pUserName
	if (users = "") then users = "*"
	if(IsEmpty(session("conn" & CONEXION_USER &  "Alias")))then Call loadConfigFile(CONEXION_USER)	
	Set objDomain = GetObject ("GC://RootDSE")
	objADsPath = objDomain.Get("defaultNamingContext")
	Set objDomain = Nothing
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.provider = session("conn" & CONEXION_USER &  "Alias")	
	objConn.Properties("User ID") = session("conn" & CONEXION_USER &  "User")	
	objConn.Properties("Password") = session("conn" & CONEXION_USER &  "Key")
	objConn.open "Active Directory Provider"
	Set objCom = CreateObject("ADODB.Command")
	Set objCom.ActiveConnection = objConn
	objCom.CommandText ="select name,telephonenumber,mobile,mail,company,title,department,sAMAccountName,sn,userAccountControl,msexchhidefromaddresslists FROM 'GC://"+objADsPath+"' where sAMAccountname='"&users&"' ORDER by sAMAccountname"

	'=======Executre queury on LDAP for all accounts=========
	Set myData = objCom.Execute
	j = 0
	
	while not myData.EoF
		redim vData(myData.recordCount,myData.fields.count-1)
		
		for i = 0 to myData.fields.count-1
			vData(j,i) = myData.fields(i)
		next
		
		j = j + 1
		myData.MoveNext
	wend
	
	getData = vData
	
End Function
'-----------------------------------------------------------------------------------------
Function isUserInGroup(pUser,pGroup)

	strUsername = pUser
	strUserName = Right(strUserName, Len(strUserName) - InStrRev(strUserName, "\"))
	if(IsEmpty(session("conn" & CONEXION_USER &  "Alias")))then Call loadConfigFile(CONEXION_USER)			
	Set objDomain = GetObject ("GC://rootDSE")
	objADsPath = objDomain.Get("defaultNamingContext")
	Set objDomain = Nothing
	Set con = Server.CreateObject("ADODB.Connection")
	con.provider = session("conn" & CONEXION_USER &  "Alias")	
	con.Properties("User ID") = session("conn" & CONEXION_USER &  "User")	
	con.Properties("Password") = session("conn" & CONEXION_USER &  "Key")
	con.open "Active Directory Provider"
	Set Com = CreateObject("ADODB.Command")
	Set Com.ActiveConnection = con
	Com.CommandText ="select memberof FROM 'GC://"+objADsPath+"' where sAMAccountname='"+strUsername+"'"
	Set rs = Com.Execute
	rtrn = false

	if (not rs.EoF) then
		
		membership=rs("memberof")
		rs.Close
		con.Close
		Set rs = Nothing
		Set con = Nothing

		if (not isnull(membership)) then
			For each group in membership
			 newgroup=split(group,"=")
			 myGroup = left(newgroup(1), len(newgroup(1))-3)
			 if (ucase(myGroup) = ucase(pGroup)) then rtrn = true
			Next
		end if
	end if
	isUserInGroup = rtrn
End Function
'-----------------------------------------------------------------------------------------
Function getUserDivision(pCdUser)
    Dim mySector
    
    mySector = getUserSector(pCdUser)
    
    'ESTA FUNCION DEBE CAMBIARSE PARA NO HARCODEAR LOS ID!!
	getUserDivision = 1
    Select case CLng(mySector)
		Case 38 '--ARROYO
			getUserDivision = 2
		Case 36 '--Transito
			getUserDivision = 4			
		Case 37 '--Bahia Blanca
			getUserDivision = 3
	end Select 	
	
End Function
'-----------------------------------------------------------------------------------------
Function getUserSector(pCdUser)
    Dim rs, strSQL, conn, ret
        
    strSQL= "Select SECTORKR from WFPROFESIONAL where CDUSUARIO='" & pCdUser & "'"
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	ret = 0
	if (not rs.eof) then ret = CLng(rs("SECTORKR"))
	getUserSector = ret
	
End Function
'-----------------------------------------------------------------------------------------
Function getListBossOf(pCdUser)
    
    Dim rs, strSQL, conn, ret, myRelKR
    
    'ESTA FUNCION DEBE CAMBIARSE PARA DEJAR DE UTILIZAR EL MG!!
    
    'Tomo el KR de la relacion ESJEFEDE.    
    Call GF_MGC("SR", "ESJEFEDE", myRelKR, "")
    'Obtengo los sectores buscados.
    strSQL="Select * from RelacionesConsulta where SRO1KR=" & myRelKR & " and SRVALOR<>'*' and SRO2KM='SG' and SRO2KC='" & pCdUser & "'"    
    'response.Write strsql
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    ret = "-1"
    while (not rs.eof)
        ret= ret & ", " & rs("SRO3KR")
        rs.MoveNext()
    wend    
    getListBossOf = ret
    
End Function
'-----------------------------------------------------------------------------------------
Function isBossOf(pCdUser, pIdSector)
    
    Dim rs, strSQL, conn, myRelKR
    
    'ESTA FUNCION DEBE CAMBIARSE PARA DEJAR DE UTILIZAR EL MG!!
    
    'Tomo el KR de la relacion ESJEFEDE.    
    Call GF_MGC("SR", "ESJEFEDE", myRelKR, "")
    'Obtengo los sectores buscados.
    strSQL="Select * from RelacionesConsulta where SRO1KR=" & myRelKR & " and SRVALOR<>'*' and SRO2KM='SG' and SRO2KC='" & pCdUser & "' and SRO3KR=" & pIdSector
    'Response.Write strSQL
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    isBossOf = false
    if (not rs.eof) then isBossOf = true
    
End Function
%>
