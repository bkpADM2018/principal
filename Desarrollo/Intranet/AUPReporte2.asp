<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
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
dim strSQL, rsPersonas, oConn, myApellido, myEmpresaKM, myEmpresaKR, myEmpresaKC, myEmpresaDS, myNombre, myNroLegajo, myUser, myEmail, myNroCredencial,myFecha
dim mySRO1KR, myO1KRPS, mySRO2KR_PRVS
dim cn, rs, sql, myList, myTextOP
dim WordApp, wordDoc, prevSC
dim dicContabilidad, entro, queImprime
set dicContabilidad = CreateObject ("Scripting.Dictionary")
dim elementA, elementC, elementN, index, index2
mySRO2KR_PRVS = ""
queImprime = ""
entro = false
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString 
   FrmDic.Add  i,Request.QueryString(i).item
Next
'Obtener datos del empleado seleccionado
'strSQL = "select MG.mg_kc as PUser, Empresa, Email,	Apellido, Nombre, NroLegajo, NroCredencial from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional inner join MG on Pe.IdPersona=mg.mg_kr where P.idprofesional=" & FrmDic("pIdPersona") & " and EgresoValido='F'"
strSQL = ""
strSQL = strSQL & "SELECT mg.mg_kc AS puser, " 						& vbCrLf
strSQL = strSQL & "       empresa          , "						& vbCrLf
strSQL = strSQL & "       email            , "						& vbCrLf
strSQL = strSQL & "       apellido         , "						& vbCrLf
strSQL = strSQL & "       nombre           , "						& vbCrLf
strSQL = strSQL & "       nrolegajo        , "						& vbCrLf
strSQL = strSQL & "       nrocredencial    , "						& vbCrLf
strSQL = strSQL & "       fechaingreso " 							& vbCrLf
strSQL = strSQL & "FROM   profesionales p " 						& vbCrLf
strSQL = strSQL & "       INNER JOIN personas pe " 					& vbCrLf
strSQL = strSQL & "         ON pe.idpersona = p.idprofesional " 	& vbCrLf
strSQL = strSQL & "       INNER JOIN mg " 							& vbCrLf
strSQL = strSQL & "         ON pe.idpersona = mg.mg_kr "            & vbCrLf
strSQL = strSQL & "WHERE  p.idprofesional = " & FrmDic("pIdPersona")& vbCrLf
strSQL = strSQL & "       AND egresovalido = 'F'"

'Response.Write strSQL

call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
if not rsPersonas.eof then
	myApellido = rsPersonas("Apellido")
	myNombre = rsPersonas("Nombre")
	myNroLegajo = rsPersonas("NroLegajo")
	myUser = rsPersonas("PUser")
	myEmail = rsPersonas("Email")
	myNroCredencial = rsPersonas("NroCredencial")
	myEmpresaKR = rsPersonas("Empresa")
	myFecha = left(GF_FN2DTE(rsPersonas("fechaingreso")),10)
end if

call GF_MGC("SR","TSTT",mySRO1KR,"")
call GF_MGC("SR","PRTT",myO1KRPS,"")
call GF_MGC("","",trim(myEmpresaKR),myEmpresaDS)
'Cargar lista de tareas asignadas al usuario
sql= "Select * from RelacionesConsulta where sro1kr=" & myO1KRPS & " and sro2kr=" & FrmDic("pIdPersona") & " and srvalor<>'*'"
'Response.Write sql
call GF_BD_CONTROL (rs,cn,"OPEN",sql)
while not rs.eof
		myList = myList & "," & rs("sro3kr") & "=" & rs("srValor")
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
'-------------------------------------------------------------------------------
function getPreparedWord(pWord)
	dim rtrn
	rtrn = replace(pWord,"Arroyo - ","")
	rtrn = replace(rtrn,"Transito - ","")
	rtrn = replace(rtrn,"Exportacion - ","")
	'rtrn = replace(rtrn,"Nivel ","")
	getPreparedWord = rtrn
end function
'-------------------------------------------------------------------------------
function PrepareWord(pWord)
dim auxWord
	auxWord = replace(pWord,"á","a")
	auxWord = replace(auxWord,"é","e")
	auxWord = replace(auxWord,"í","i")
	auxWord = replace(auxWord,"ó","o")
	auxWord = replace(auxWord,"ú","u")
	auxWord = replace(auxWord,"ñ","n")
	auxWord = replace(auxWord,"(","{")
	auxWord = replace(auxWord,")","}")
	PrepareWord = auxWord
end function
'-------------------------------------------------------------------------------
function getPreparedWord(pWord)
dim rtrn
rtrn = replace(pWord,"Arroyo - ","")
rtrn = replace(rtrn,"Transito - ","")
rtrn = replace(rtrn,"Exportacion - ","")
rtrn = replace(rtrn,"Piedrabuena - ","")
'rtrn = replace(rtrn,"Nivel ","")
getPreparedWord = rtrn
end function
'-------------------------------------------------------------------------------
function LF_MostrarElemento(pKR, pDS, pValue)
dim rtrn, myImg, myTxt, myTxtFinal, myTxtAux, myIndexFin
myImg = "icon_unchecked.gif"
	if instr(myList,"," & pKR & "=") <> 0 then 
		myImg = "icon_checked.gif"
		if pValue then
			myTxtAux = mid(myList,instr(myList,"," & pKR & "=")+1,len(myList))
			myIndexFin = instr(myTxtAux,"=") + 1
			myTxtFinal = mid(myTxtAux,myIndexFin,1)
			if myTxtFinal = "A" then
				myTxt = "&nbsp;&nbsp;<font size='-1'><b>(" & GF_Traducir("Admin") & ")</b></font>"			
			elseif myTxtFinal = "Y" then
				myTxt = "&nbsp;&nbsp;<font size='-1'><b>(" & GF_Traducir("Auditor") & ")</b></font>"
			elseif myTxtFinal = "U" then
				myTxt = "&nbsp;&nbsp;<font size='-1'><b>(" & GF_Traducir("Usuario") & ")</b></font>"
			end if
		end if
	end if
	rtrn = "<img src='images/" & myImg & "'>&nbsp;" & GF_Traducir(getPreparedWord(pDS))
	rtrn = rtrn & myTxt
	LF_MostrarElemento = rtrn
end function
%>

<script>
function printReport(){
	/*
	window.print();
	if (confirm("Desea imprimir las normas y Politicas?")){
		window.open ('documentos/Politicas de Internet y Correo de Toepfer.doc',"NORMAS","toolbar=yes,menubar=yes,type=fullWndow,resizable=yes,scrollbars=1");
		window.print();
	}
	//location.reload() 
	*/
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
						<td><%=myFecha%></td>
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
					<tr>
						<td><b>Empresa:</b></td>
						<td><%=myEmpresaDS%></td>
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
		strSQL= "select * from RelacionesConsulta where sro1kr=" & mySRO1KR & " and srvalor<>'*' order by sro2kc, sro3kc"
		'Response.write strSQL
		call GF_BD_CONTROL (rsSistemas,oConn,"OPEN",strSQL)
		while not rsSistemas.eof		
				if mySRO2KR_PRVS <> rsSistemas("sro2kr") then
					if printThen then
						if queImprime = "SC" then
							elementAA = split(dicContabilidad("TK_AA"),"///")
							elementEA = split(dicContabilidad("TK_EA"),"///")
							elementTA = split(dicContabilidad("TK_TA"),"///")
							elementPA = split(dicContabilidad("TK_PA"),"///")
							elementN  = split(dicContabilidad("SC_N_"),"///")
							while index < ubound(elementAA)+1
								Response.Write "<tr class=reg_header_navdos>"
								if index = 0 then Response.Write "<td width='4%'>&nbsp;</td><td class=reg_header_nav><i>Arroyo</i></td><td class=reg_header_nav><i>Piedrabuena</i></td><td class=reg_header_nav><i>Transito</i></td><td class=reg_header_nav><i>Exportacion</i></td></tr><tr class=reg_header_navdos>"
								Response.Write "<td>&nbsp;</td>"
								Response.write elementAA(index)
								Response.write elementEA(index)
								Response.write elementTA(index)
								Response.write elementPA(index)							
								index = index + 1
							wend	
							Response.Write "</tr>"
							Response.Write "<tr class=reg_header_nav><td>&nbsp;</td><td colspan=4 class=reg_header_nav><i>Niveles</i></td></tr>"
							Response.Write "<tr><td>&nbsp;</td><td colspan=4>"
							Response.Write "<table width=95% border=0 class=reg_header_navdos>"
							Response.Write "<tr class=reg_header_navdos>"
							index = 0
							while index < ubound(elementN)+1
								Response.write elementN(index)
								index = index + 1
							wend
							Response.Write "</tr>"
							Response.Write "</table></td></tr>"
							printThen = false
							dicContabilidad.RemoveAll 
						elseif queImprime = "SAW" then
							elementAA = split(dicContabilidad("TK_AA"),"///")
							elementEA = split(dicContabilidad("TK_EA"),"///")
							elementTA = split(dicContabilidad("TK_TA"),"///")
							elementPA = split(dicContabilidad("TK_PA"),"///")
							elementN  = split(dicContabilidad("SC_N_"),"///")
							while index < ubound(elementAA)+1
								Response.Write "<tr class=reg_header_navdos>"
								if index = 0 then Response.Write "<td width='4%'>&nbsp;</td><td class=reg_header_nav><i>Arroyo</i></td><td class=reg_header_nav><i>Exportacion</i></td><td class=reg_header_nav><i>Transito</i></td><td class=reg_header_nav><i>Piedrabuena</i></td></tr><tr class=reg_header_navdos>"
								Response.Write "<td>&nbsp;</td>"
								Response.write elementAA(index)
								Response.write elementEA(index)
								Response.write elementTA(index)
								Response.write elementPA(index)							
								index = index + 1
							wend	
							Response.Write "</tr>"
							index = 0
							while index < ubound(elementN)+1
								Response.write elementN(index)
								index = index + 1
							wend
							printThen = false
							dicContabilidad.RemoveAll 						
						elseif queImprime = "SCW" then
							index = 0
							elementAA = split(dicContabilidad("TK_AA"),"///")
							elementEA = split(dicContabilidad("TK_EA"),"///")
							elementTA = split(dicContabilidad("TK_TA"),"///")
							elementPA = split(dicContabilidad("TK_PA"),"///")
							elementN  = split(dicContabilidad("SCW"),"///")
							while index < ubound(elementAA)+1
								Response.Write "<tr class=reg_header_navdos>"
								if index = 0 then Response.Write "<td width='4%'>&nbsp;</td><td class=reg_header_nav><i>Arroyo</i></td><td class=reg_header_nav><i>Exportacion</i></td><td class=reg_header_nav><i>Transito</i></td><td class=reg_header_nav><i>Piedrabuena</i></td></tr><tr class=reg_header_navdos>"
								Response.Write "<td>&nbsp;</td>"
								Response.write elementAA(index)
								Response.write elementEA(index)
								Response.write elementTA(index)
								Response.write elementPA(index)							
								index = index + 1
							wend	
							Response.Write "</tr>"
							Response.Write "<tr class=reg_header_nav><td>&nbsp;</td><td colspan=4 class=reg_header_nav><i>Firmas</i></td></tr>"
							Response.Write "<tr><td>&nbsp;</td><td colspan=4>"
							Response.Write "<table width=95% border=0 class=reg_header_navdos>"
							Response.Write "<tr class=reg_header_navdos>"
							index = 0
							while index < ubound(elementN)+1
								Response.write elementN(index)
								index = index + 1
							wend
							Response.Write "</tr>"
							Response.Write "</table></td></tr>"
							printThen = false
							dicContabilidad.RemoveAll 
						end if
					end if
					if mySRO2KR_PRVS <> "" then Response.Write "</table></td></tr>"
					mySRO2KR_PRVS = rsSistemas("sro2kr")
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
									
					if rsSistemas("sro2kc") = "SC" then
							myKey = left(rsSistemas("sro3kc"),5)
							myNameFunction = "cargarLista"
							myOutput = "<td nowrap align=left><font>"
							myOutput = myOutput & LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"),false) & "</font></td>"
							if dicContabilidad.Exists(myKey) then
								dicContabilidad(myKey) = dicContabilidad(myKey) & "///" & myOutput
							else
								dicContabilidad.Add myKey, myOutput
							end if
							printThen = true
							queImprime = "SC"
					elseif rsSistemas("sro2kc") = "SAW" then
							myKey = left(rsSistemas("sro3kc"),5)
							myNameFunction = "cargarLista"
							myOutput = "<td nowrap align=left><font>"
							myOutput = myOutput & LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"),false) & "</font></td>"
							if dicContabilidad.Exists(myKey) then
								dicContabilidad(myKey) = dicContabilidad(myKey) & "///" & myOutput
							else
								dicContabilidad.Add myKey, myOutput
							end if
							printThen = true
							queImprime = "SAW"							
					elseif  rsSistemas("sro2kc") = "SCW" then
							if left(rsSistemas("sro3kc"),3) = "SCW"  then
								myKey = left(rsSistemas("sro3kc"),3)
							else
								myKey = left(rsSistemas("sro3kc"),5)
							end if
								if mid(rsSistemas("sro3kc"),8,1) = "1" then
									myOutput = "<td colspan='5' align=left><hr></td><tr><td nowrap align=left><font>"
								else
									myOutput = "<td nowrap align=left><font>"
								end if
							myOutput = "<td nowrap align=left><font>"
							myNameFunction = "cargarLista"
							myOutput = myOutput & LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"), true) & "</font></td>"
							if dicContabilidad.Exists(myKey) then
								dicContabilidad(myKey) = dicContabilidad(myKey) & "///" & myOutput
							else
								dicContabilidad.Add myKey, myOutput
							end if
							myOutput = ""
							printThen = true
							queImprime = "SCW"
					else
								%>	
								<tr>
									<td width="4%">&nbsp;</td>
									<td width="48%">
										<%=LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"),false)%>
									</td>	
									<% 
									if not rsSistemas.eof then rsSistemas.movenext 
									if not rsSistemas.eof then
										if mySRO2KR_PRVS <> rsSistemas("sro2kr") then
										%>
											<td width="48%">
												&nbsp
											</td>
									<%
											rsSistemas.movePrevious
										else	
											%>
											<td width="48%">
												<%=LF_MostrarElemento(rsSistemas("sr3okr"), rsSistemas("sro3ds"),false)%>
											</td>
									<%  end if 
									end if
									%>
								</tr>	
							<%	
					end if		
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
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td>             <%=GF_Traducir("Firma del Gerente de IT")%></td>
						<td align="left"><%=GF_Traducir("_____________________________________")%></td>
					</tr>
				</table>	
			</td>
		</tr>									
	</table>

<font size="-2">
	Alfred C. Toepfer International S.R.L.<br>
	Av. Del Libertador 350 10º piso.  (B1638BEP) - Vicente Lopez - Buenos Aires - República Argentina
</font>
</form>
</body>
</html>