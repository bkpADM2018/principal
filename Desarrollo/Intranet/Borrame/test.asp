<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%

strSQL="Select * from TRANSPORTISTAS where codPos is Null and NUDOCUMENTO <> ''"
Call executeQueryDb(DBSITE_BAHIA, rs, "OPEN", strSQL)
while (not rs.eof)
	strSQL="Select * from MET001A where nrodoc=" & rs("NUDOCUMENTO")
	Call executeQueryDb(DBSITE_SQL_MAGIC, rs2, "OPEN", strSQL)
	altura = 0
	codpos= 0
	domici=""
	if (rs("DSDOMICILIO") <> "") then domici = replace(rs("DSDOMICILIO"), "'", "''")
	if (not rs2.eof) then
		altura= rs2("numero")
		codpos= rs2("codpos")
		domici= Left(replace(rs2("domemp"), "'", "''"), 50)
	end if
	strSQL="Update TRANSPORTISTAS SEt altura=" & altura & ", DSDOMICILIO='" & domici & "', codpos=" & codpos & " where CDTRANSPORTISTA=" & rs("CDTRANSPORTISTA")
	Call executeQueryDb(DBSITE_BAHIA, rsX, "EXEC", strSQL)
	rs.MoveNext()
wend


response.write "Listo!"
%>