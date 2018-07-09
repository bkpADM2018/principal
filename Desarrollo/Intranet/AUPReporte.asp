<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 07/01/2007                         -->
'<!--               Pagina     : AUPReporte.ASP                     -->
'<!--               Descripcion: Listado de cambios solicitados     -->
'<!------------------------------------------------------------------->
'ProcedimientoControl "AUPREP"
dim FrmDic
dim strSQL, rsPersonas, oConn, myApellido, myNombre, myNroLegajo
dim mySRO1KR, myO1KRPS, mySRO1KR_PRVS
dim cn, rs, sql, myList, myTextOP
mySRO1KR_PRVS = ""
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString 
   FrmDic.Add  i,Request.QueryString(i).item
Next
'Obtener datos del empleado seleccionado
strSQL = "select Apellido, Nombre, NroLegajo from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional where P.idprofesional=" & FrmDic("pIdPersona") & " and EgresoValido='F'"
call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
if not rsPersonas.eof then
	myApellido = rsPersonas("Apellido")
	myNombre = rsPersonas("Nombre")
	myNroLegajo = decrypt(rsPersonas("NroLegajo"))
end if

call GF_MGC("SR","TSTT",mySRO1KR,"")
call GF_MGC("SR","PRTT",myO1KRPS,"")

'Cargar lista de tareas asignadas al usuario
sql= "Select * from RelacionesConsulta where sro1kr=" & myO1KRPS & " and sro2kr=" & FrmDic("pIdPersona") & " and srvalor<>'*'"
call GF_BD_CONTROL (rs,cn,"OPEN",sql)
while not rs.eof
		myList = myList & "," & rs("sro3kr")
	rs.movenext
wend
myList = myList & ","
call GF_BD_CONTROL (rs,cn,"CLOSE",sql)

'Operacion que se esta realizando
select case FrmDic("pOP")
		case "A"
			myTextOP = "Alta"
		case "D"
			myTextOP = "Baja"
		case "M"
			myTextOP = "Modificación"
		case else
			myTextOP = "ABM"
end select 			

function LF_MostrarElemento(pKR, pDS)
dim rtrn, myImg
myImg = "icon_unchecked.gif"
	if instr(myList,"," & pKR & ",") <> 0 then myImg = "icon_checked.gif"
	rtrn = "<img src='images/" & myImg & "'>&nbsp;" & GF_Traducir(pDS)
	LF_MostrarElemento = rtrn
end function
%>
<html>
<head>
<title>Intranet ActiSA - Reporte</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body style="font-family: arial;" onload="print();">
<form name="frmMain" action="AUPReporte.asp" method="post">
	<table border="2" cellspacing="0" cellpadding="0" style="height:950px;" width="100%" align="center">
		<tr style="height:30px;">
			<td>
				<table width="100%">
					<tr>
						<td colspan="3" align="middle"><font size="+1"><b><%=ucase(GF_Traducir("Solicitud de " & myTextOP & " de Perfiles de Usuarios"))%></b></font></td>
					</tr>
				</table>	
			</td>
		</tr>						
		<tr style="height:60px;">
			<td>
				<table width="100%">
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td width="15%"><b>Fecha:</b></td>
						<td><%=now()%></td>
						<td align="right" valign="top" rowspan="4"><IMG width="194" height="49" src="images/logo1.gif"></td>
					</tr>
					<tr>
						<td><b>Nombres:</b></td>
						<td><%=myNombre%></td>
					</tr>
					<tr>
						<td><b>Apellidos:</b></td>
						<td><%=myApellido%></td>
					</tr>
					<tr>
						<td><b>Legajo:</b></td>
						<td><%=myNroLegajo%></td>
					</tr>
					<tr><td colspan="2">&nbsp;</td></tr>
				</table>
			</td>
		</tr>		
		<%
		strSQL= "select * from RelacionesConsulta where sro1kr=" & mySRO1KR & " and srvalor<>'*' order by sro2ds, sro3ds"
		call GF_BD_CONTROL (rsSistemas,oConn,"OPEN",strSQL)
		while not rsSistemas.eof 
			if mySRO1KR_PRVS <> rsSistemas("sro2kr") then
				if mySRO1KR_PRVS <> "" then Response.Write "</table></td></tr>"
				mySRO1KR_PRVS = rsSistemas("sro2kr")
				%>
				<tr valign="top">
					<td>
						<table width="100%">
							<tr>
								<td colspan="3"><b><%=GF_Traducir(rsSistemas("sro2ds"))%></b></td>
							</tr>
							<tr><td colspan="3">&nbsp;</td></tr>
			<% 
			end if 
			%>
							<tr>
								<td width="15%">&nbsp;</td>
								<td width="40%">
									<%=LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"))%>
								</td>	
								<% 
								if not rsSistemas.eof then rsSistemas.movenext 
								if not rsSistemas.eof then
									if mySRO1KR_PRVS <> rsSistemas("sro2kr") then
										rsSistemas.movePrevious
									else	
										%>
										<td width="45%">
											<%=LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"))%>
										</td>
								<%  end if 
								end if
								%>
							</tr>	
							<%	
							if not rsSistemas.eof then rsSistemas.movenext
					wend	
		call GF_BD_CONTROL (rsSistemas,oConn,"CLOSE",strSQL)
		%>
						</table>	
					</td>
				</tr>	
				<tr style="height:20px;">
					<td>
						<table width="100%">
							<tr>
								<td><b>&nbsp;</b></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align="right"><%=GF_Traducir("---------------------------------")%></td>
							</tr>
							<tr>
								<td align="right"><font size="-1"><%=GF_Traducir("Firma del Gerente del Sector")%></font></td>
							</tr>
						</table>	
					</td>
				</tr>									
		</table>
</form>
</body>
</html>