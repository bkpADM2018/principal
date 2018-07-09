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
'<!--               Fecha      : 22/01/2008                         -->
'<!--               Pagina     : AUPAuditoria.ASP                   -->
'<!--               Descripcion: Listado para Auditoria             -->
'<!--               Modificacion: Henzel Pavlo              -->
'<!--               Fecha      : 22/01/2008                         -->
'<!------------------------------------------------------------------->

call GP_CONFIGURARMOMENTOS()

dim myEgresoValido, myFechaIngreso, myFechaEgreso
dim strSQL, rsAuditoria, rsSistemas, oConn
dim myO1KRPS, myUsuANT, mySisANT, myTarANT,pDesSec
dim myTTKRPS, frmDic, myWhere, myColor1, myColorIndex1
dim myColor2, myColorIndex2, myCantLineas, myTope, MyClass1, MyClass2, MyClass3
dim myLinea1, myLinea2, myLinea3, myLinea4, myLinea5, myFlag, MyEnding, myImg
set FrmDic= CreateObject ("Scripting.Dictionary") 
dim strAccion, usrDS, usrKC, dtConf, pDesUsr, myAuxTitle
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
myColor2 = "#fffaf0"
'51
dim nameOfView, textOfView 
nameOfView = "RelacionesConsulta"
myWhere = " where p.EgresoValido <> 'V' "


myTope = 45
strAccion=GF_Parametros7("p_accion","",6)
strResponsable = GF_Parametros7("p_responsable","",6)
call GF_MGC("SR","PRTT",myO1KRPS,"")
call GF_MGC("","",frmDic("pSector"),pDesSec)
call GF_MGC("","",frmDic("pUser"),pDesUsr)


'Esto es solo para consultas hitoricas!!!!!!!
if frmDic("pHistorica") = "S" then 
	myWhere = ""
	nameOfView = "RelacionesConsulta2"
	'frmDic("pHistorica") = ""



	strSQL = "Select * from Profesionales where idProfesional=" & frmDic("pUser")
	'Response.Write strsql
	call GF_BD_CONTROL (rsAuditoria,oConn,"OPEN",strSQL)
	if not rsAuditoria.eof then
		myFechaIngreso = rsAuditoria("FechaIngreso")
		myEgresoValido = rsAuditoria("EGRESOVALIDO")
		if trim(myEgresoValido) = "V" then
			myFechaEgreso = rsAuditoria("FechaEgreso")
		end if	
	end if
	call GF_BD_CONTROL (rsAuditoria,oConn,"CLOSE",strSQL)
	'Response.Write myEgresoValido & "-" & myFechaEgreso 
end if	
'-----------------------------------------------------------------------------------------------
sub PrintEncabezados()
%>
	<tr class="reg_header_nav">
		<td width="18%" class="MarcoMiddle" align="center"><b><%=GF_Traducir("Usuario")%>		</b></td>
		<td width="30%" class="MarcoMiddle" align="center"><b><%=GF_Traducir("Sistema")%>		</b></td>
		<td width="35%" class="MarcoMiddle" align="center"><b><%=GF_Traducir("Tarea")%>			</b></td>
		<td width="2%"  class="MarcoMiddle" align="center"><b><%=GF_Traducir("Res.")%>			</b></td>
		<td width="15%" class="MarcoMiddle" align="center"><b><%=GF_Traducir("Fecha")%>			</b></td>				
	</tr>
<%
end sub
'-----------------------------------------------------------------------------------------------
function PrintLine(pTexto,pMax)
	'if len(pTexto) > pMax then 
		'myCantLineas = myCantLineas + 1 
		'Response.Write "T(" & pTexto & ") Max(" & pMax & ")(" & mid(pTexto,pMax) & ")"
	'end if
	
PrintLine = left(pTexto,pMax)
end function
'-----------------------------------------------------------------------------------------------
sub ControlarFinPagina()
if myCantLineas = myTope-1 then
	MyEnding = "style='BORDER-BOTTOM: #000000 1px solid;'"
elseif myCantLineas = myTope then 
	'myTope = 59
	myTope = 51
	call PrintEncabezados	
	call ReLoadData		
	myCantLineas = 1
end if	
end sub
'-----------------------------------------------------------------------------------------------
sub ReLoadData()
				myLinea1 = left(UCASE(rsAuditoria("UserDS")),23) & "(" & UCASE(rsAuditoria("UserKC")) & ")"
				myLinea2 = rsAuditoria("Sistemas")
				myLinea3 = rsAuditoria("Tareas")
				myLinea4 = rsAuditoria("Usuario")
				myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))
		if myLinea5 <> "&nbsp;" then 
				if rsAuditoria("Valor") = "*" then
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/arrow_minus.gif>"
				else
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/arrow_plus.gif>"					
				end if			
				if rsAuditoria("Momento") < dtConf then
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src='images/checked2.gif'>"
				else
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src='images/unchecked2.gif'>"
				end if	
		end if						
end sub
'-----------------------------------------------------------------------------------------------
sub CargarDatosCabecera()
dim rs, cn, sql
if (frmDic("pSector")<>"") then
	sql = "select M.mg_kc as USERKC, m.mg_ds as USERDS, c.mmtoconf as DATECONF from ConfirmacionesPermisos C inner join mg M on C.krultimousuario=m.mg_kr where c.krSector=" & frmDic("pSector") & " order by mmtoconf desc"
	'Response.Write sql
	call GF_BD_Control (rs, cn, "OPEN", sql)
		if not rs.eof then
			usrKC = rs("USERKC")
			usrDS = rs("USERDS")
			dtConf = rs("DATECONF")
		end if
	call GF_BD_Control (rs, cn, "CLOSE", sql)
end if	
end sub

'-----------------------------------------------------------------------------------------------

%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<title>Intranet ActiSA - Reporte de Situación Actual de Permisos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript">
	function confirmar(p_sector, p_responsable){
		var link
		link = 'confirmacionPermisos.asp?p_sector=' + p_sector + "&p_responsable=" + p_responsable;
		document.location.href = link;
	}
</script>
</head>
<body class="print" onload="alert('Recuerde configurar la página de la siguiente manera: \n -Orientación: Horizontal\n -Encabezado y Pie de página: En Blanco\n -Bordes: 1 cm.');">
<form name="frmMain" method="post">
<%

if frmDic("pSector") <> "" then
	myAuxTitle = " sector <b>" & pDesSec & "</b>"
	myWhere = myWhere & " and P.Sector=" & frmDic("pSector")
else
	if frmDic("pUser") <> "" then
		myAuxTitle = " usuario <b>" & pDesUsr & "</b>" 
		if frmDic("pHistorica") = "S" then myAuxTitle = myAuxTitle & "<br>&nbsp;Fecha de ingreso al sistema: " & mid(myFechaIngreso,7,2) & "/" & mid(myFechaIngreso,5,2) & "/" & left(myFechaIngreso,4) & " " & mid(myFechaIngreso,9,2) & ":" & mid(myFechaIngreso,11,2) & ":" & mid(myFechaIngreso,13,2)
		myWhere = myWhere & " and P.idProfesional=" & frmDic("pUser")
	end if
end if
strSQL =" select case when C1.Usuario is null then C1.Valor else C1.Usuario end Usuario, C1.Valor as Valor,M.mg_kc as UserKC, M.mg_ds as UserDS, p.Sector as SuSector, M.mg_kr as UserKR, C1.TSDS as Sistemas,C1.TTKR as TareaKR, C1.TTDS as Tareas, C1.MMTO as Momento " & _
		"	from Profesionales P inner join mg M on p.IdProfesional=m.mg_kr " & _
		"	left join " & _
		"		(select  rc1.srValor as Valor, rc1.sruser as Usuario, RC1.sro2KR as UserKR, RC1.sro2kc as UserKC, RC1.sro2ds as UserDS, RC2.sro2ds as TSDS, RC2.sro3kr as TTKR,RC2.sro3ds as TTDS, rc1.srmmdt  as MMTO " & _
		"			from " & nameOfView & " RC1 inner join " & nameOfView & "  RC2 on RC1.sro3kr=RC2.sr3okr " & _		
		"				where RC1.sro1kr=" & myO1KRPS & _ 
		"			group by rc1.srValor , rc1.sruser , RC1.sro2KR , RC1.sro2kc, RC1.sro2ds, RC2.sro2ds, RC2.sro3kr, RC2.sro3ds, rc1.srmmdt) C1 " & _
		"	on P.idProfesional=C1.UserKR " & _ 
		"	inner join MG as M1 on P.IdProfesional = M1.MG_KR " & myWhere & _
		"	order by userds, sistemas, tareas, Momento desc" 
		'Response.Write "SQL(" & strSQL & ")"
		'Response.end
call GF_BD_CONTROL (rsAuditoria,oConn,"OPEN",strSQL)
	if not rsAuditoria.eof then
		if frmDic("pSector") = "" then	frmDic("pSector") = rsAuditoria("SuSector")
	end if	
	CargarDatosCabecera

if frmDic("pHistorica") = "S" then 
	Response.Write GF_TITULO("Usuarios.gif","Reporte de Situación Historica de Permisos de Usuarios para el " & myAuxTitle & "<br><font size=-2>&nbsp;*La información abajo listada ha sido confirmada el dia " & GF_FN2DTE(dtConf) & " por " & usrDS & " (" & ucase(usrKC) & ")</font>")
else
	Response.Write GF_TITULO("Usuarios.gif","Reporte de Situación Actual de Permisos de Usuarios para el " & myAuxTitle & "<br><font size=-2>&nbsp;*La información abajo listada ha sido confirmada el dia " & GF_FN2DTE(dtConf) & " por " & usrDS & " (" & ucase(usrKC) & ")</font>")
end if	
'Response.Write GF_TITULO("Usuarios.gif","Reporte de Situación Actual de Permisos de Usuarios para el " & myAuxTitle & "<br><font size=-2>*La información abajo listada ha sido confirmada el dia " & GF_FN2DTE(dtConf) & " por " & usrDS & " (" & usrKC & ")</font>")
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
<%	
	call PrintEncabezados()

	while not rsAuditoria.eof
		myLinea1 = "&nbsp;"
		myLinea2 = "&nbsp;"
		myLinea3 = "&nbsp;"
		myLinea4 = "&nbsp;"
		myLinea5 = "&nbsp;"
		MyEnding = ""
		if rsAuditoria("UserKr") <> myUsuANT then
			if not isnull(rsAuditoria("Sistemas")) then
				myLinea1 = left(UCASE(rsAuditoria("UserDS")),23) & "(" & UCASE(rsAuditoria("UserKC")) & ")"
				myLinea2 = rsAuditoria("Sistemas")
				myLinea3 = rsAuditoria("Tareas")
				myLinea4 = rsAuditoria("Usuario")
				myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))
				myUsuANT = rsAuditoria("UserKr")
				mySisANT = rsAuditoria("Sistemas")			
				myTarANT = rsAuditoria("Tareas")
				myFlag = true
				myColorIndex1 = myColorIndex1 + 1
				if myColorIndex1 mod 2 = 0 then
					myColor1 = "#f5fffa"
				else
					myColor1 = "#ffffff"
				end if	
				myColorIndex2 = myColorIndex2 + 1
				if myColorIndex2 mod 2 = 0 then
					myColor2 = "#ffffff"
				else
					myColor2 = "#dcf7dc"
				end if	
			end if
		else
			if rsAuditoria("Sistemas") <> mySisANT then
				myColorIndex2 = myColorIndex2 + 1
				if myColorIndex2 mod 2 = 0 then
					myColor2 = "#ffffff"
				else
					myColor2 = "#dcf7dc"
				end if	
				
				myLinea2 = rsAuditoria("Sistemas")

				if rsAuditoria("Tareas") <> myTarANT then
					myLinea3 = rsAuditoria("Tareas")
					myLinea4 = rsAuditoria("Usuario")
					myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))	
					myTarANT = rsAuditoria("Tareas")
				else
					myLinea4 = rsAuditoria("Usuario")
					myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))	
				end if
				mySisANT = rsAuditoria("Sistemas")
			else
				'Response.Write "L3(" & myLinea3 & ")"
				if rsAuditoria("Tareas") <> myTarANT then
					myLinea3 = rsAuditoria("Tareas")
					myLinea4 = rsAuditoria("Usuario")
					myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))	
					myTarANT = rsAuditoria("Tareas")
				else
					myLinea4 = rsAuditoria("Usuario")
					myLinea5 = GF_FN2DTE(rsAuditoria("Momento"))	
				end if
			end if
		end if
		if myLinea5 <> "&nbsp;" then 
				if rsAuditoria("Valor") = "*" then
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/arrow_minus.gif>"
				else
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/arrow_plus.gif>"					
				end if

				if rsAuditoria("Momento") < dtConf then
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/checked2.gif>"
				else
					myLinea5 = myLinea5 & "&nbsp;<img align='absmiddle' src=images/unchecked2.gif>"
				end if	
		end if				
		fclass = ""
		if (strAccion <> "CONFIRMA") then
			fclass = "courier4"
		end if
		if myLinea5 <> "&nbsp;" then
			myCantLineas = MyCantLineas + 1
			'Response.Write "<hr>ACA(" & MyCantLineas & ")(" & left(myLinea3,20) & ")"
			call ControlarFinPagina
			MyClass1 = "MarcoL"
			MyClass2 = "MarcoL"
			MyClass3 = "MarcoL"			
			if myLinea1 <> "&nbsp;" then MyClass1 = "MarcoTL"
			if myLinea2 <> "&nbsp;" then MyClass2 = "MarcoMiddle"
			if myLinea3 <> "&nbsp;" then MyClass3 = "MarcoMiddle"
			%>		
			<tr class="<%=MyClass1%>">
				<td class="<%=MyClass1%>"	<%=MyEnding%>>									 <b><font class="<%= fclass %>"><%=PrintLine(myLinea1,28)%></font></b>	</td> 
				<td class="<%=MyClass2%>"	bgcolor="<%=myColor2%>" <%=MyEnding%>>				<font class="<%= fclass %>"><%=PrintLine(myLinea2,50)%></font>		</td>
				<td class="<%=MyClass3%>"	bgcolor="<%=myColor2%>" <%=MyEnding%>>				<font class="<%= fclass %>"><%=PrintLine(myLinea3,67)%></font>		</td>
				<td class="MarcoMiddle"		bgcolor="<%=myColor2%>" align="center" <%=MyEnding%>>	<font class="<%= fclass %>"><%=PrintLine(myLinea4,3) %></font>		</td>
				<td class="MarcoTR"			bgcolor="<%=myColor2%>" align="center" <%=MyEnding%>>	<font class="<%= fclass %>"><%=PrintLine(myLinea5,160)%></font>		</td>
			</tr>
			<%
		end if
		'call ControlarFinPagina
		rsAuditoria.movenext		
	wend
		%>
	<%
if strAccion = "CONFIRMA" and strResponsable <> "" then
%>
	<tr>
		<td colspan="5" style="BORDER-TOP: #000000 1px solid;" align="center"><input type="button" value="Confirmar" onclick="javascript:confirmar(<%=frmDic("pSector")%>, <%=strResponsable%>);"></td>
	</tr>
<%
end if

%>	

	<tr>
		<td colspan="5" style="BORDER-TOP: #000000 1px solid;">&nbsp;</td>
	</tr>
</table>
<%
if trim(myEgresoValido) = "V" then
%>
<table width="100%">	
	<tr>
		<td colspan="5" class="TDERROR"><font class="big"><%=GF_Traducir("Fecha de baja:") & "&nbsp;" & mid(myFechaEgreso,7,2) & "/" & mid(myFechaEgreso,5,2) & "/" & left(myFechaEgreso,4) & " " & mid(myFechaEgreso,9,2) & ":" & mid(myFechaEgreso,11,2) & ":" & mid(myFechaEgreso,13,2)%></font></td>
	</tr>
</table>	
<%
end if
%>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/arrow_plus.gif"> indican el momento en que ha sido otorgado el acceso para ese Sistema-Tarea.</font>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/arrow_minus.gif"> indican el momento en que ha sido denegado el acceso para ese Sistema-Tarea.</font>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/unchecked2.gif"> indican que estas autorizaciones se encuentran pendientes de ser aprobadas ya que su fecha de actualización es posterior a la fecha de confirmación.</font>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/checked2.gif"> indican que estas autorizaciones ya han sido aprobadas.</font>



<% call GF_BD_CONTROL (rsAuditoria,oConn,"ClOSE",strSQL) %>
</form>
</body>
</html>
