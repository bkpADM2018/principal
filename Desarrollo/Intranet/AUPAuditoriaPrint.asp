<!--#include file="Includes/procedimientosMG.asp"-->
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
dim rsAuditoria, oConn, strSQL, myColorIndex, myColor
Dim Sector,Desde,Hasta, myWhere,primero,FechaDesde,FechaHasta,mySectoresDS
dim CantLineas
CantLineas=11

Sector   = GF_Parametros7("sector","",6)
Desde    = GF_Parametros7("Desde","",6)
Hasta    = GF_Parametros7("Hasta","",6)
Detalles = GF_Parametros7("Detalles","",6)

FechaDesde  = GF_DTE2FN(Desde)
FechaHasta  = GF_DTE2FN(Hasta)

if inStr(Sector,"-1") <> 0 then
	dim rsaux
	strSQL = ""
	strSQL = strSQL & "SELECT * FROM mg WHERE MG_KM = 'SS'"
	'call mostrarSQL(strsql,false)

	call GF_BD_CONTROL (rsAux,oConn,"OPEN",strSQL)
	Sector = ""
	while not rsAux.eof
		Sector = Sector & rsAux("mg_kr") & ","
		rsAux.movenext
	wend
	Sector = left(Sector,len(Sector)-1)
end if

%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<link rel="stylesheet" type="text/css" media="all" href="CSS/calendar-win2k-2.css" title="win2k-2" />
<title>Reporte de Situación Actual de Confirmacion de Permisos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<form name="frmMain" method="post">
<%
'strSQL = "SELECT M.MG_KC AS USERKC, M.MG_DS AS USERDS, C.MMTOCONF as DATECONF, M2.MG_KR AS SECTORKR, M2.MG_DS AS SECTORDS FROM CONFIRMACIONESPERMISOS C INNER JOIN MG M ON C.KRULTIMOUSUARIO=M.MG_KR INNER JOIN MG M2 ON C.KRSECTOR=M2.MG_KR ORDER BY M2.MG_DS"
if Sector <> "" then
		call AgregarWhere(myWhere,"m2.mg_kr in (" & Sector & ")")
		mySectoresDs = SectoresDS()
end if
If Desde <> "" then
		call AgregarWhere(myWhere,"c.mmtoconf >= '" & FechaDesde & "'")
end if
If Hasta <> "" then
		call AgregarWhere(myWhere,"c.mmtoconf <= '" & FechaHasta & "'")
end if

strSQL = ""
strSQL = strSQL & "SELECT   m.mg_kc    AS userkc, " 
strSQL = strSQL & "         m.mg_kr    AS userkr, " 
strSQL = strSQL & "         m.mg_ds    AS userds, " 
strSQL = strSQL & "         c.mmtoconf AS dateconf, " 
strSQL = strSQL & "         m2.mg_kr   AS sectorkr, " 
strSQL = strSQL & "         m2.mg_ds   AS sectords "
strSQL = strSQL & "FROM     confirmacionespermisos c " 
strSQL = strSQL & "         inner join mg m on  m.mg_kr=c.KrUltimoUsuario" 
strSQL = strSQL & "         inner join mg m2 on c.krsector = m2.mg_kr" 
strSQL = strSQL & myWhere
strSQL = strSQL & "ORDER BY c.krsector,c.mmtoconf desc"

'response.Write strsql & "<br>"

call GF_BD_CONTROL (rsAuditoria,oConn,"OPEN",strSQL)
'---------------------------------------------------
function AgregarWhere(byref p_where,p_agregar)

    if p_where = "" then
        p_where = " WHERE " & p_agregar 
    else
        p_where = p_where & " AND " & p_agregar  
    end if

end function
'--------------------------------------------
function SectoresDS()
	dim strSQL,rs,rtrn,i
	
	strSQL = "Select * from mg where mg_kr in (" & Sector & ") order by mg_ds asc"
	call GF_BD_CONTROL (rs,oConn,"OPEN",strSQL)
	rtrn = "<TABLE border=0 cellspacing=1 cellpadding=1><TR>"
	i=1
	while not rs.eof
		
		if rtrn = "" then
			rtrn = rs("MG_DS")
		else
			rtrn = rtrn & "<TD><B>" & rs("MG_DS") & "</B></TD>"
			if i mod 3 = 0 then 
				rtrn = rtrn & "</TR><TR>"
				CantLineas = CantLineas +1
			end if
		end if
		i=i+1
		
		rs.movenext
	wend
	rtrn = rtrn & "</TR></TABLE>"
	SectoresDS = rtrn
	
end Function

%>
<TABLE align='left' border='0'>
	<TR>
		<TD colspan='2'>
			<U><B>Parametros de Busqueda</B></U>
		</TD>
	</TR>
	<%if Sector <> "" then%>
	<TR>
		<TD valign='top'>
			<label for="Sector"  >Sector</label><BR>
		</TD>	
		<TD>
			<label for="Sector2" ><B><%=mySectoresDs%></B></label>
		</TD>
	</TR>
	<%end if
	if Desde <> "" or Hasta <> "" then%>
	<TR>
		<TD valign='top'>
			<label for="Periodo"  >Periodo</label><BR>
		</TD>	
		<TD>
			<table cellspacing=1 cellpadding=1>
			<label for="Periodo2"  ><B>
			<%
			if Desde <> "" then 
				response.write "<TR><TD> Desde </TD><TD><B>" & Desde & "</B></TD></TR>" 
				CantLineas = CantLineas +1
			end if
			if Hasta <> "" then 
				response.write "<TR><TD> Hasta </TD><TD><B>" & Hasta & "</B></TD></TR>" 
				CantLineas = CantLineas +1
			end if
			%></B>
			</label>
			</Table>
		</TD>
	</TR>
	<%end if%>
	<TR>
		<TD>
			<label for="Detalles"  >Mostrar Detalles</label><BR>
		</TD>	
		<TD>
			<label for="Detalles2"  ><B><%if Detalles = "true" then response.write "Si" else response.write "No" end if %></B></label><BR>
		</TD>
	</TR>
</TABLE>
<BR>
<table border="1" cellspacing="0" cellpadding="0" width="100%" align="center">
	<%
	CantLineas = CantLineas +1
	if rsAuditoria.eof then%>
		<tr class="reg_header_nav">
			<td align='center'> no se encontraron resultados </td>
		</tr>
	<%else%>
		<tr class="reg_header_nav">
			<td><font class="courier3"><B><%=GF_TRADUCIR("Sector")%></B></font></td>
			<td><font class="courier3"><B><%=GF_TRADUCIR("Responsable del Sector")%></B></font></td>
			<td align="center"><font class="courier3"><B><%=GF_TRADUCIR("Fecha Confirmación")%></B></font></td>

		</tr>	
	<%
	CantLineas = CantLineas +1
	end if%>
<%
Dim Blanco,Verde
Blanco = "#ffffff"
Verde  = "#dcf7dc"
while not rsAuditoria.eof
	myColorIndex = myColorIndex + 1
	'if myColorIndex mod 2 = 0 then
	'	myColor = "#dcf7dc"
	'else
	'	myColor = "#ffffff"
	'end if	
%>
	
	<tr bgcolor="<%=Verde%>">
		<%CantLineas = CantLineas +1%>
		<td><font class="courier3"><B><%=rsAuditoria("SECTORDS") & " (" & rsAuditoria("SECTORKR") & ")"%></B></font></td>
		<td><font class="courier3"><B><%=ucase(rsAuditoria("USERDS")) & " (" & ucase(rsAuditoria("USERKC")) & ")"%></B></font></td>
		<td align="center"><font class="courier3"><%
			if isnull(rsAuditoria("DATECONF")) then
				response.write "<B>Sin Confirmar</B>"
			else
				response.write "<B>" & GF_FN2DTE(rsAuditoria("DATECONF")) & "</B>"
			end if
			%></font>
		</td>
		<%if detalles then
			call imprimirDetalle(rsAuditoria("sectorkr"),rsAuditoria("dateconf"))
		end if%>
	</tr>

<% 
	rsAuditoria.movenext
	if not rsAuditoria.eof then%>
	<tr class="reg_header_nav">
			<td><font class="courier3"><B><%=GF_TRADUCIR("Sector")%></B></font></td>
			<td><font class="courier3"><B><%=GF_TRADUCIR("Responsable del Sector")%></B></font></td>
			<td align="center"><font class="courier3"><B><%=GF_TRADUCIR("Fecha Confirmación")%></B></font></td>
	</tr>	
	<%end if
wend	
call GF_BD_CONTROL (rsAuditoria,oConn,"ClOSE",strSQL) %>
</table>

</form>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/arrow_plus.gif"> indican el momento en que ha sido otorgado el acceso para ese Sistema-Tarea.</font>
<br>
<font class="courier3">Las fechas seguidas por <img align="absmiddle" src="images/arrow_minus.gif"> indican el momento en que ha sido denegado el acceso para ese Sistema-Tarea.</font>

</body>
</html>
<%

'-----------------------------------------------------------------------------------------------
sub PrintEncabezados()
%>
	<tr>
		<TD width='3%'>&nbsp </TD>
		<td class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B><%=GF_Traducir("Usuario")%>		</b></font></td>
		<td class="MarcoMiddle" align="center" bgcolor="#517B4A" ><font color='WHITE' class="courier3"><b><%=GF_Traducir("Sistema")%>		</b></font></td>
		<td class="MarcoMiddle" align="center" bgcolor="#517B4A" ><font color='WHITE' class="courier3"><b><%=GF_Traducir("Tarea"  )%>		</b></font></td>
		<td class="MarcoMiddle" align="center" bgcolor="#517B4A" ><font color='WHITE' class="courier3"><b><%=GF_Traducir("Res."   )%>		</b></font></td>
		<td class="MarcoMiddle" align="center" bgcolor="#517B4A" ><font color='WHITE' class="courier3"><b><%=GF_Traducir("Fecha"  )%>   	</b></font></td>				
	</tr>
<%
end sub
'---------------------------------------------------------------------------------------------
function ControlarPagina()
	dim rtrn
	rtrn = false
	if CantLineas => 55 then
		CantLineas= 0
		printEncabezados()
		rtrn = true
	end if
	ControlarPagina = rtrn
end Function
'---------------------------------------------------------------------------------------------
Function imprimirDetalle(p_krsector,mmto)
	Dim strSQL, rs
	dim ultimoU,ultimoS
	
	strSQL = ""
	strSQL = strSQL & "SELECT   CASE  " 
	strSQL = strSQL & "           WHEN c1.usuario IS NULL " 
	strSQL = strSQL & "           THEN c1.valor " 
	strSQL = strSQL & "           ELSE c1.usuario " 
	strSQL = strSQL & "         END 'Usuario', " 
	strSQL = strSQL & "         c1.valor AS valor, " 
	strSQL = strSQL & "         m.mg_kc  AS userkc, " 
	strSQL = strSQL & "         m.mg_ds  AS userds, " 
	strSQL = strSQL & "         p.sector AS susector, " 
	strSQL = strSQL & "         m.mg_kr  AS userkr, " 
	strSQL = strSQL & "         c1.tsds  AS sistemas, " 
	strSQL = strSQL & "         c1.ttkr  AS tareakr, " 
	strSQL = strSQL & "         c1.ttds  AS tareas, " 
	strSQL = strSQL & "         c1.mmto  AS momento " 
	strSQL = strSQL & "FROM     profesionales p " 
	strSQL = strSQL & "         INNER JOIN mg m " 
	strSQL = strSQL & "           ON p.idprofesional = m.mg_kr " 
	strSQL = strSQL & "         LEFT JOIN (SELECT   rc1.srvalor AS valor, " 
	strSQL = strSQL & "                             rc1.sruser  AS usuario, " 
	strSQL = strSQL & "                             rc1.sro2kr  AS userkr, " 
	strSQL = strSQL & "                             rc1.sro2kc  AS userkc, " 
	strSQL = strSQL & "                             rc1.sro2ds  AS userds, " 
	strSQL = strSQL & "                             rc2.sro2ds  AS tsds, " 
	strSQL = strSQL & "                             rc2.sro3kr  AS ttkr, " 
	strSQL = strSQL & "                             rc2.sro3ds  AS ttds, " 
	strSQL = strSQL & "                             max(rc1.srmmdt)  AS mmto " 
	strSQL = strSQL & "                    FROM     relacionesconsulta rc1 " 
	strSQL = strSQL & "                             INNER JOIN relacionesconsulta rc2 " 
	strSQL = strSQL & "                               ON rc1.sro3kr = rc2.sr3okr " 
	strSQL = strSQL & "                             INNER JOIN mg " 
	strSQL = strSQL & "                               ON rc1.srvalor = mg.mg_kc " 
	strSQL = strSQL & "                                   OR rc1.srvalor = '*' " 
	strSQL = strSQL & "                    WHERE    rc1.sro1kr = 20806 " 
	strSQL = strSQL & "                             AND mg_km = 'UP' " 
	strSQL = strSQL & "                    GROUP BY rc1.srvalor, " 
	strSQL = strSQL & "                             rc1.sruser, " 
	strSQL = strSQL & "                             rc1.sro2kr, " 
	strSQL = strSQL & "                             rc1.sro2kc, " 
	strSQL = strSQL & "                             rc1.sro2ds, " 
	strSQL = strSQL & "                             rc2.sro2ds, " 
	strSQL = strSQL & "                             rc2.sro3kr, " 
	strSQL = strSQL & "                             rc2.sro3ds, " 
	strSQL = strSQL & "                             rc1.srmmdt) c1 " 
	strSQL = strSQL & "           ON p.idprofesional = c1.userkr " 
	strSQL = strSQL & "         INNER JOIN mg AS m1 " 
	strSQL = strSQL & "           ON p.idprofesional = m1.mg_kr " 
	strSQL = strSQL & "WHERE    p.egresovalido <> 'V' " 
	strSQL = strSQL & "         AND p.sector = " & p_krsector 
	strSQL = strSQL & "         AND c1.mmto < '" & mmto & "'" 
	strSQL = strSQL & "ORDER BY userds, " 
	strSQL = strSQL & "         sistemas, " 
	strSQL = strSQL & "         tareas, " 
	strSQL = strSQL & "         momento DESC"
	
	'call mostrarSQL(strsql,false)
	
	call GF_BD_CONTROL (rs,oConn,"OPEN",strSQL)
	
	if rs.eof then%>
		<TR>
			<TD colspan='3'>
				<Table  border=0 cellspacing=0 cellpadding=0 width='100%'>
						<TR>
							<TD width='3%'>&nbsp </TD>
							<TD class="MarcoMiddle" colspan='5' align='center' bgcolor="#517B4A"><font color='WHITE' class="courier3"><B>Detalles del sector</B></font></TD>
						</TR>
						<TR >
							<TD width='3%'>&nbsp </TD>
							<TD colspan='5' class="MarcoMiddle" align='center'><font class="courier3"><B>No se encontraron resultados</B></font></TD>
							
						</TR>
						
				</TABLE>
			</TD>
		</TR>
	<%else
	
	%>

			<TR>
				<TD colspan='3'>
					<Table  border=0 cellspacing=0 cellpadding=0 width='100%'>
						<TR>
							<TD width='3%'>&nbsp </TD>
							<TD class="MarcoMiddle" colspan='5' align='center' bgcolor="#517B4A"><font color='WHITE' class="courier3"><B>Detalles del sector</B></font></TD>
						</TR>
						<TR >
							<TD width='3%'>&nbsp </TD>
							<TD class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B>Usuario</B></font></TD>
							<TD class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B>Sistema</B></font></TD>
							<TD class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B>Tarea</B></font></TD>
							<TD class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B>Res.</B></font></TD>
							<TD class="MarcoMiddle" align='center' bgcolor="#517B4A" ><font color='WHITE' class="courier3"><B>Fecha</B></font></TD>
						</TR>
						
						<%
						dim aux
						dim aux2
						while not rs.eof
							aux2 = ControlarPagina()
							if aux2 = true then ultimoU = ""
							CantLineas = CantLineas +1%>
						<TR >						
							<TD width='3%'>&nbsp</TD>
							
							
							<%if ultimoU <> rs("userds") then 
								%><TD class="MarcoMiddle"><font class="courier3">&nbsp <%response.write rs("userds") & "(" & rs("userkc") & ")"%></font></TD><%
								ultimoU = rs("userds")
								aux = true
							else 
								%><TD class="MarcoL">&nbsp</TD><%
							end if%>
							
							
							
							<%if ultimoS <> rs("sistemas") or aux = true then 
								%><TD class="MarcoMiddle"><font class="courier3">&nbsp <%=rs("sistemas")%></font></TD><%
								ultimoS = rs("sistemas")
								aux = false
							else 
								%><TD class="MarcoL">&nbsp</TD><%
							end if%>
							

							
							<TD class="MarcoMiddle"><font class="courier3">&nbsp <%=rs("tareas")%></font></TD>
							<TD class="MarcoMiddle"><font class="courier3">&nbsp <%=rs("usuario")%></font></TD>
							<TD class="MarcoMiddle"><font class="courier3">&nbsp <%=GF_FN2DTE(rs("momento"))%>
							<%	if rs("Valor") = "*" then
									response.write "&nbsp<img align='absmiddle' src=images/arrow_minus.gif>"
								else
									response.write "&nbsp<img align='absmiddle' src=images/arrow_plus.gif>"					
								end if		
							
								
							%></font>
							</TD>
						</TR>
						<%
						rs.movenext
						wend%>
							
						
						
					</TABLE>
				</TD>
				
			</TR>
		
<%
	end if
end Function
%>