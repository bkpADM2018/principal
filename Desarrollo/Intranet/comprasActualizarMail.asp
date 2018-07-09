<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosValidacion.asp"-->
<%
'----------------------------------------------------------------------------
'************************************************
'**********   COMIENZO DE LA PAGINA   ***********
'************************************************
'----------------------------------------------------------------------------
Dim idEmpresa, mail, rs, conn, strSQL, rtrn, auxMails, control

idEmpresa = GF_PARAMETROS7("idEmpresa", 0, 6)
mail = GF_PARAMETROS7("mail", "", 6)

if ((idEmpresa > 0) and (mail <> "")) then
	control = true
	auxMails = split(mail,";")
	for i=0	to uBound(auxMails)
		if (auxMails(i) = "") then control = false
		if (control) then control = GF_CONTROL_EMAIL(auxMails(i))
	next
	if (control) then
		strSQL = "Select * from TBLMAILSCOMPRAS where IDEMPRESA=" & idEmpresa
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (rs.eof) then
			strSQL="Insert into TBLMAILSCOMPRAS values (" & idEmpresa & ", '" & mail & "')" 
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		else			
			strSQL = "Update TBLMAILSCOMPRAS Set EMAIL = '" & mail & "' Where IDEMPRESA = " & idEmpresa
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		end if
		rtrn = mail
	else
		rtrn = ""
	end if
end if

Response.Write rtrn
%>