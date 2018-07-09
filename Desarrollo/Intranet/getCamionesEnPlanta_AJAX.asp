<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<%	
	dim strSql
	dim query
	dim auxText
	dim puertos(3)	
	dim portCx(3)
	
	portCx(1) = DBSITE_ARROYO
	portCx(2) = DBSITE_TRANSITO
	portCx(3) = DBSITE_BAHIA
	
	puertos(1)= "Arroyo"
	puertos(2)= "Pto. San Martín"				 
	puertos(3)= "Bahía Blanca"

	strSQL = ""
	strSQL = strSQL & "SELECT   "
	strSQL = strSQL & "         COUNT(*) AS cantidad "
	strSQL = strSQL & "FROM     camiones c inner join "
	strSQL = strSQL & "( "
	strSQL = strSQL & "		Select IDCAMION, CDCLIENTE, CDVENDEDOR, CDCORREDOR from CAMIONESCARGA "
	strSQL = strSQL & "		union "
	strSQL = strSQL & "		Select IDCAMION, CDCLIENTE, CDVENDEDOR, CDCORREDOR from camionesdescarga "
	strSQL = strSQL & ") CD on CD.IDCAMION=C.IDCAMION "
	if (not IsToepfer(session("KCOrganizacion"))) then
		strSQL = strSQL & " WHERE     "
		strSQL = strSQL & "		cdestado <> " & CAMIONES_ESTADO_BAJA
		strSQL = strSQL & "		and (cdcliente in (Select CDCLIENTE from clientes where NUCUIT = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & "     OR cdvendedor in (Select CDVENDEDOR from VENDEDORES where NUDOCUMENTO = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & "     OR cdcorredor in (Select CDCORREDOR from CORREDORES where NUCUIT = '" & session("CuitOrganizacion") & "')) "
	end if

    for i = 1 to 3
		if (portCx(i) <> "") then
			if (executeQueryDb(portCx(i), query, "OPEN", strSql)) then
				if not query.eof then			
					auxText = auxText & puertos(i) & " (" & query("cantidad") & ")|"
				else	
					auxText = auxText & puertos(i) & " (0)|"
				end if
			else
				auxText = auxText & puertos(i) & " (0)|"
			end if
		else
			auxText = auxText & puertos(i) & " (Fuera de Linea)|"
		end if
	next	
	if len(auxText) > 0 then Response.Write left(auxText,len(auxText)-1)
	
%>
