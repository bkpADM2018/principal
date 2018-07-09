<!--#include file="../includes/procedimientosMG.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientossql.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<%
'------------------------------------------------------------------------------------------------------
'Function checkKilos : chequea que los kilos sean correctos, validando con la Balanza y el Draft Survey
'Parametros: pKilos = kilos ingresados 
'	 	  pKilosBza = kilos de la Balanza actual 
'		pKilosDraft	= kilos del Draft Survey
Function checkKilos(pKilos, pKilosBza, pKilosDraft)	
	Dim myKilos
	if(pKilosDraft > 0)then
		'Si tiene Draft Survey, la prioridad tiene los kilos del Draft
		myKilos = pKilos + pKilosBza
		if (myKilos > pKilosDraft) Then msg = "<li>Los kilos no deben superar el Draft Survey.</li>" & msg
	else
		'Si no tiene Draft Survey, se valida con los kilos de la Balanza
		if (pKilos > pKilosBza) Then msg = "<li>Los kilos no deben superar el total de la Balanza.</li>" & msg
	end if
End Function
'------------------------------------------------------------------------------------------------------
dim strPuerto, rs, cdProducto, cdAviso, kilos, permiso, cosecha, msg, accion, kilosDraft
accion = GF_Parametros7("accion","",6)
strPuerto = GF_Parametros7("Pto","",6)
cdProducto = GF_Parametros7("producto",0,6)
cdAviso = GF_Parametros7("aviso",0,6)
kilos = GF_Parametros7("kilos",2,6)
permiso = GF_Parametros7("permiso","",6)
cosecha = GF_Parametros7("cosecha",0,6)
kilosBza = GF_Parametros7("kilosBza",2,6) 
kilosDraft = GF_Parametros7("kilosDraft",2,6)

msg = ""
if accion = "DEL" then
	strSQL = "SELECT * FROM EMBARQUESDATOS WHERE CDAVISO=" & cdAviso & " AND CDPRODUCTO=" & cdProducto & " AND CDCOSECHA=" & cosecha 
	call GF_BD_Puertos (strPuerto, rs, "OPEN",strSql)
	'Response.Write "<hr>SQL: " & strSQL
	if not rs.eof then
		strSQL = "DELETE FROM EMBARQUESDATOS WHERE CDAVISO=" & cdAviso & " AND CDPRODUCTO=" & cdProducto & " AND CDCOSECHA=" & cosecha 
		'Response.Write "<hr>SQL: " & strSQL
		response.write "OK"
		'msg = "OK" '"<font color='green'>Carga dada de baja exitosamente!</font>"
		call GF_BD_Puertos (strPuerto, rs, "EXEC",strSql)
	end if	
else	
	'Control de parametros
	if kilos = 0 then msg = "<li>Debe ingresar los kilos.</li>"
	Call checkKilos(kilos, kilosBza, kilosDraft)
	if cosecha = 0 then 
		msg = "<li>Debe ingresar la cosecha.</li>" & msg
	elseif len(cosecha)<> 8 then
		msg = "<li>La cosecha no es valida.</li>" & msg
	elseif (left(cosecha,4) < 2008 or left(cosecha,4) > 2080) or (right(cosecha,4) < 2008 or right(cosecha,4) > 2080) or (left(cosecha,4) > right(cosecha,4)) or ((cint(left(cosecha,4))-cint(right(cosecha,4)) <> -1)) then
		msg = "<li>La cosecha no es valida.</li>" & msg
	end if	
	if msg = "" then
		strSQL = "SELECT * FROM EMBARQUESDATOS WHERE CDAVISO=" & cdAviso & " AND CDPRODUCTO=" & cdProducto & " AND CDCOSECHA=" & cosecha 
		call GF_BD_Puertos (strPuerto, rs, "OPEN",strSql)
		'Response.Write "<hr>SQL: " & strSQL		
		if not rs.eof then
			strSQL = "UPDATE EMBARQUESDATOS SET KILOS=" & kilos & ", PERMISO='" & permiso & "' WHERE CDAVISO=" & cdAviso  & " AND CDPRODUCTO=" & cdProducto & " AND CDCOSECHA=" & cosecha  
		else
			strSQL = "INSERT INTO EMBARQUESDATOS VALUES(" & cdAviso & "," & cdProducto & "," & cosecha  & "," & kilos & ",'" & permiso & "')" 
		end if
		'Response.Write "<hr>SQL: " & strSQL
		call GF_BD_Puertos (strPuerto, rs, "EXEC",strSql)
		response.write "OK"
	else
		msg = "<u>ATENCION:</u><ul>" & msg & "</ul>"
		Response.Write msg
	end if
end if
%>
