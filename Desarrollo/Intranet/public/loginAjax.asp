<!--#include file="loginController.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosJSON.asp"-->
<%
'-------------------------------------------------------------------------------------------------------
Function checkLogin(pUser, pPass, pToken)
	Dim strSQL,rs,user,pass,rtrn
	
	rtrn = false
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", "SELECT * FROM TBLUSUARIOS WHERE NombreUsuario = '" & pUser & "' and Clave <> '' and ESTADO=" & ESTADO_ACTIVO)	
	if (not rs.eof) then
		'response.write "(2 Clave) " & Trim(rs("Clave")) & "<br>"
		'response.write "(3 Hash) " & UCase(MD5(Trim(rs("Clave")) & pToken)) & "<br>"		
		if (UCase(MD5(Trim(rs("Clave")) & pToken)) = pPass) then rtrn = true		
	end if
	checkLogin = rtrn	
End Function
'-------------------------------------------------------------------------------------------------------

Dim lang, userName, password, myToken, ajaxMsg, ip, mensaje, aux, repass, oJson
Dim rs
userName = UCASE(GF_PARAMETROS7("user","",6))
password = UCASE(GF_PARAMETROS7("pass","",6))
pass = UCASE(GF_PARAMETROS7("p","",6))
ll = UCASE(GF_PARAMETROS7("ll","",6))

'response.write "Llave: " & ll & "<br>"
'response.write "MD5 Clave: " & pass & "<br>"
'response.write "Hash: " & password & "<br>"

Set oJson = jsObject()
ip = Request.ServerVariables("REMOTE_ADDR") 'recupero la IP de la pc

Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", "Select * from TBLLOGINTOKENS where TOKENOWNER='" & ip & "' and ISSUEDATE <= GETDATE() and VALIDTO >= GETDATE()")
myToken = ""
if (not rs.eof) then myToken = Trim(rs("TOKEN"))
'response.write "(1 Llave) " & myToken & "<br>"
if(checkLogin(userName, password, myToken)) then
	'response.write "(4)<br>"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", "Update TBLLOGINTOKENS Set USEROWNER= '" & userName & "' where TOKENOWNER='" & ip & "'")
	oJson("error") = ""			
	oJson("url") = "Home.asp"
else
	oJson("error") = errMessage(USUARIO_PASS_INCORRECTO)
	'oJson("llave") = generarLlave()
end if
oJson.Flush
%>
