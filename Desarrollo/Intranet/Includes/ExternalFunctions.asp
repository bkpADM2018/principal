<%

Const ID_PROV_NO_EXISTE = -1
Const DS_PROV_NO_EXISTE = "-"
Const ID_PROV_MAX = 100000
'#############################################
' Function: GetCDEnterprise
' Params:	UserName
' Porpouse: Recover Enterprise Code from Database
'#############################################

Function GetCDEnterprise(UserName)
	Dim strSQL,rs
	Dim rtrn
	
	strSQL = ""
	strSQL = strSQL & " SELECT * "
	strSQL = strSQL & " FROM usuario "
	strSQL = strSQL & " WHERE cdUserName = '" & UserName & "'"

	GF_BD_Puertos "WEB", rs, "OPEN",strSQL
	
	if rs.eof then
		rtrn = ID_PROV_NO_EXISTE
	else
		rtrn = rs("cdEmpresa")
	end if
	
	GetCDEnterprise = rtrn
End Function

'#############################################
' Function: GetCDEnterprise
' Params:	UsecuitrName
' Porpouse: Recover Enterprise Code from Database
'#############################################

Function GetCDEnterprise3(cuit)
	Dim strSQL,rs,con
	Dim rtrn
	
	if isnumeric(cuit) then
		strSQL = ""
		strSQL = strSQL & " SELECT * "
		strSQL = strSQL & " FROM toepferdb.vwempresas "
		strSQL = strSQL & " WHERE CUIT = '" & cuit & "'"
		GF_BD_AS400_2 rs,con, "OPEN",strSQL
		
		if rs.eof then
			rtrn = ID_PROV_NO_EXISTE
		else
			rtrn = CDbl(rs("IdEmpresa"))
		end if
	else
		rtrn = ID_PROV_NO_EXISTE
	end if
	
	GetCDEnterprise3 = rtrn
End Function

'#############################################
' Function: GetDSEnterprise
' Params:	UserName
' Porpouse: Recover Enterprise Description from Database
'#############################################

Function GetDSEnterprise(UserName)
	Dim strSQL,rs,con
	Dim rtrn,idempresa
	
	idempresa = GetCDEnterprise(UserName)
	if idempresa <> -1 then
		strSQL = ""
		strSQL = strSQL & " SELECT * "
		strSQL = strSQL & " FROM vwempresas "
		strSQL = strSQL & " WHERE IdEmpresa = '" & idempresa & "'"
		'response.write strSQL
		'response.end
		GF_BD_AS400_2 rs,con, "OPEN",strSQL
		
		if rs.eof then
			rtrn = DS_PROV_NO_EXISTE
		else
			rtrn = rs("DSEmpresa")
		end if
	else
		rtrn = idempresa
	end if
	
	GetDSEnterprise = rtrn
End Function

'#############################################
' Function: GetCDEnterprise
' Params:	ID Empresa
' Porpouse: Recover Enterprise Code from Database
'#############################################

Function GetDsEnterprise2(CdEmpresa)

	Dim strSQL,rs,con
	Dim rtrn
	
	if (CLng(CdEmpresa) <> ID_PROV_NO_EXISTE) then
		strSQL = ""
		strSQL = strSQL & " SELECT * "
		strSQL = strSQL & " FROM toepferdb.vwempresas "
		strSQL = strSQL & " WHERE IdEmpresa = '" & CdEmpresa & "'"
		'response.write strSQL
		'response.end
		GF_BD_AS400_2 rs,con, "OPEN",strSQL
		
		if rs.eof then
			rtrn = DS_PROV_NO_EXISTE
		else
			rtrn = trim(rs("DSEmpresa"))
		end if
	else
		rtrn = Cdempresa
	end if
	
	GetDsEnterprise2 = rtrn

End Function

'#############################################
' Function: GetDsEnterprise3
' Params:	CUIT
' Porpouse: Recover Enterprise Description from Database
'#############################################

Function GetDsEnterprise3(cuit)

	Dim strSQL,rs,con
	Dim rtrn
	
	if isnumeric(cuit) then
		strSQL = ""
		strSQL = strSQL & " SELECT * "
		strSQL = strSQL & " FROM toepferdb.vwempresas "
		strSQL = strSQL & " WHERE CUIT = '" & cuit & "'"
		GF_BD_AS400_2 rs,con, "OPEN",strSQL
		
		if rs.eof then
			rtrn = DS_PROV_NO_EXISTE
		else
			rtrn = trim(rs("DSEmpresa"))
		end if
	else
		rtrn = DS_PROV_NO_EXISTE
	end if
	
	GetDsEnterprise3 = rtrn

End Function
%>