<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<%
dim dtContable, nuCtaPte, CTG, nuCtaPteAnt, CTGAnt, strSQL, rs, pOption
'----------------------------------------------------------------------------------------------------
pPto = GF_Parametros7("pto","",6)
pOption = GF_Parametros7("option","",6) 

if pOption = "UPD" then
	dtContable = GF_Parametros7("dtContable","",6)
    if (cdate(dtcontable) < cdate("01-04-2014")) then
        Response.Write "ERROR en fecha, debe ser mayor a 01-04-2014"
		Response.end
	end if
    dtContable = GF_DTE2FN(dtContable)
    dtContable = left(dtContable,4) & "-" & mid(dtContable,5,2) & "-" & mid(dtContable,7,2)
	nuCtaPte = GF_Parametros7("cartaPorte","",6)
	if instr(nuCtaPte,"-") <> 5 then
		Response.Write "ERROR en formato de Carta de Porte. Formato esperado= XXXX-XXXXXXXX"
		Response.end
	else
		if len(mid(nuCtaPte,1,instr(nuCtaPte,"-")-1)) <> 4 then
			Response.Write "ERROR en formato de Carta de Porte. Formato esperado= XXXX-XXXXXXXX"
			Response.end
		end if	
		if len(nuCtaPte) - instrrev(nuCtaPte,"-") <> 8 then
			Response.Write "ERROR en formato de Carta de Porte. Formato esperado= XXXX-XXXXXXXX"
			Response.end
		end if	

	end if
	nuCtaPte = replace(nuCtaPte,"-","")
	CTG = GF_Parametros7("CTG","",6)
	if len(CTG)<> 8 then
		Response.Write "ERROR en formato de CTG. Formato esperado= XXXXXXXX"
		Response.end
	end if
end if
nuCtaPteAnt = GF_Parametros7("cartaPorteAnt","",6)
nuCtaPteAnt = left(replace(nuCtaPteAnt,"-",""),12)
CTGAnt = GF_Parametros7("CTGAnt","",6)

'Controles
rtrn = "ERR"
strSQL = "SELECT * FROM dbo.WSCTG_CAMIONES WHERE NUCARTAPORTE='" & nuCtaPteAnt & "' AND CTG=" & CTGAnt
Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
'Response.Write strSQL
If Not rs.EOF Then
	if pOption = "UPD" then
		strSQL = "UPDATE dbo.WSCTG_CAMIONES SET DTCONTABLE='" & dtContable & "', NUCARTAPORTE='" & nuCtaPte & "', CTG=" & CTG & " WHERE NUCARTAPORTE='" & nuCtaPteAnt & "' AND CTG=" & CTGAnt
	elseif pOption = "DEL" then
		strSQL = "UPDATE dbo.WSCTG_CAMIONES SET MMTOCONFIRMACION=" & GF_DTE2FN(now()) & ", CDUSERCONFIRMACION='ASP', ESTADOCONFIRMACION=" & WSCTG_QUITADO & ", MMTODESVIO=" & GF_DTE2FN(now()) & ", CDUSERDESVIO='ASP', ESTADODESVIO=" & WSCTG_QUITADO & ", MMTORECHAZO=" & GF_DTE2FN(now()) & ", CDUSERRECHAZO='ASP', ESTADORECHAZO=" & WSCTG_QUITADO & " WHERE NUCARTAPORTE='" & nuCtaPteAnt & "' AND CTG=" & CTGAnt
	end if
    'Response.Write strSQL
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
	if err.number = 0 then 
		rtrn = "OK"
	else
		rtrn = err.Description 
	end if	
End If
Response.Write rtrn	

%>