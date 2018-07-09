<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 10/01/2007                         -->
'<!--               Pagina     : AUPReporte2.ASP                    -->
'<!--               Descripcion: Listado de cambios solicitados     -->
'<!------------------------------------------------------------------->
'ProcedimientoControl "AUPREP2"
dim FrmDic
dim strSQL, rsPersonas, oConn, myApellido, myNombre, myNroLegajo, myUser, myEmail, myNroCredencial
dim mySRO1KR, myO1KRPS, mySRO1KR_PRVS
dim cn, rs, sql, myList, myTextOP
dim WordApp, wordDoc
mySRO1KR_PRVS = ""
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString 
   FrmDic.Add  i,Request.QueryString(i).item
Next
'Obtener datos del empleado seleccionado
strSQL = "select MG.mg_kc as PUser, Email,	Apellido, Nombre, NroLegajo, NroCredencial from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional inner join MG on Pe.IdPersona=mg.mg_kr where P.idprofesional=" & FrmDic("pIdPersona") & " and EgresoValido='F'"
call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
if not rsPersonas.eof then
	myApellido = rsPersonas("Apellido")
	myNombre = rsPersonas("Nombre")
	myNroLegajo = decrypt(rsPersonas("NroLegajo"))
	myUser = rsPersonas("PUser")
	myEmail = rsPersonas("Email")
	myNroCredencial = rsPersonas("NroCredencial")
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

<script>
function printReport(){
	window.print();
	if (confirm("Desea imprimir las normas y Politicas?")){
		window.open ('documentos/Politicas de Internet y Correo de Toepfer.doc',"NORMAS","toolbar=yes,menubar=yes,type=fullWndow,resizable=yes,scrollbars=1");
		//window.print();
	}
}
</script>

<html>
<head>
<title>Intranet ActiSA - Reporte</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body style="font-family: arial;" onload="javascript:printReport();">
<form name="frmMain" action="AUPReporte.asp" method="post">
	<table border="2" cellspacing="0" cellpadding="0" style="height:950px;" width="100%" align="center">
		<tr>
			<td>
				<table width="100%">
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
				</table>
			</td>
		</tr>	
		<tr>
			<td>
				<table width="100%">
					<tr>
						<td width="30%"><b>Usuario I-Series:</b></td>
						<td><%=myUser%></td>
					</tr>
					<tr>
						<td><b>Usuario Windows NT:</b></td>
						<td><%=myUser%></td>
					</tr>
					<tr>
						<td><b>Usuario E-Mail Internet:</b></td>
						<td><%=myApellido & left(myNombre,1) & "@toepfer.com"%></td>
					</tr>
					<tr>
						<td colspan="2"><font size="-2"><%=GF_Traducir("*La contraseña suministrada por la empresa, se deberá cambiar en el primer ingreso del usuario al sistema.")%></font></td>
						<td>&nbsp;</td>
					</tr>

				</table>	
			</td>
		</tr>						
		<%
		strSQL= "select * from RelacionesConsulta where sro1kr=" & mySRO1KR & " and srvalor<>'*' order by sro2ds, sro3ds"
		'Response.Write strsql
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
					<% 
						if rsSistemas("sro2kc") = "AAB" then 
						%>
							<tr>
								<td>
									<tr><td colspan="3"><%=GF_Traducir("Numero de Credencial")%>:&nbsp;<%=myNroCredencial%></td></tr>
						<%
						end if
					end if		
				%>
								<tr>
									<td width="4%">&nbsp;</td>
									<td width="48%">
										<%=LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"))%>
									</td>	
									<% 
									if not rsSistemas.eof then rsSistemas.movenext 
									if not rsSistemas.eof then
										if mySRO1KR_PRVS <> rsSistemas("sro2kr") then
										%>
											<td width="48%">
												&nbsp
											</td>
									<%
											rsSistemas.movePrevious
										else	
											%>
											<td width="48%">
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
					
		<tr>
			<td>
				<table width="100%">
					<tr>
						<td><font size="-2"><p align="justify"><%=GF_Traducir("Informamos a los usuarios de los sistemas informáticos de 'Alfred C. Toepfer International S.R.L.' que la contraseña suministrada es confidencial e intransferible, la misma posee una vigencia de un (1) mes, luego de transcurrido este lapso el sistema pedirá el cambio de la misma. A su vez comunicamos que la información y/o elementos suministrados por 'Alfred C. Toepfer International S.R.L.' son para uso exclusivamente laboral y la venta y/o divulgación de información esta penada por ley. A su vez informamos que todos los elementos teleinformaticos (software y/o hardware) utilizados por los usuarios pertenecen a 'Alfred C. Toepfer International S.R.L.' sin excepcion.")%></p></font></td>
					</tr>

				</table>	
			</td>
		</tr>								
						
						
		<tr>
			<td>
				<table width="100%">
					<tr>
						<td colspan="2"><b><%=GF_Traducir("Con mi firma confirmo:")%></b></td>
					</tr>
					<tr><td colspan="2"><font size="-2"><%=GF_Traducir("- Haber leído y entendido las políticas de Internet y correo electrónico de Alfred C. Toepfer International S.R.L.")%></font></td></tr>
					<tr><td colspan="2"><font size="-2"><%=GF_Traducir("- Haber aceptado todas las condiciones y/o restricciones contenidas en las políticas de Alfred C. Toepfer International S.R.L.")%></font></td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td width="40%"> <%=GF_Traducir("Firma y aclaración del asociado")%></td>
						<td align="left"><%=GF_Traducir("_____________________________________")%></td>
					</tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td>             <%=GF_Traducir("Firma del Gerente del Sector")%></td>
						<td align="left"><%=GF_Traducir("_____________________________________")%></td>
					</tr>
				</table>	
			</td>
		</tr>									
	</table>

<font size="-2">
	Alfred C. Toepfer International S.R.L.<br>
	Alicia Moreau de Justo 2050/2090 Piso 2º (1107) - Bs.As. Argentina - Tel:(11) 4317-0000 - Fax:(11) 4312-8268 - TLX:7305/06/07 
</font>
</form>
</body>
</html>