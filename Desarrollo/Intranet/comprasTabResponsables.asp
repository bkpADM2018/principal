<%
Dim ds, cdResponsable, dsResponsable, hkResponsable, verEmpleados
'-------------------------------------------------------------
Function puedeParticiparAjuste(valor)
	Dim ret 
	ret =""
	if (valor=1) then
		ret ="<img src='images/compras/accept-16x16.png'>"
	end if
	puedeParticiparAjuste = ret
End Function
'-------------------------------------------------------------
Function puedeModificarArticulos(valor)
	Dim ret 
	ret =""
	if (valor=1) then
		ret ="<img src='images/compras/accept-16x16.png'>"
	end if
	puedeModificarArticulos = ret
End Function
'-------------------------------------------------------------
Function puedeConfirmarContratos(valor)
	Dim ret 
	ret =""
	if (valor=1) then
		ret ="<img src='images/compras/accept-16x16.png'>"
	end if
	puedeConfirmarContratos = ret
End Function
'-------------------------------------------------------------


'**************************************************
cdResponsable = UCase(GF_PARAMETROS7("cdResponsable", "" ,6))
dsResponsable = GF_PARAMETROS7("dsResponsable", "",6)
hkResponsable = GF_PARAMETROS7("hkResponsable", "",6)
verEmpleados = GF_PARAMETROS7("verEmpleados", 0,6)
if (verEmpleados = 0) then verEmpleados = ESTADO_ACTIVO

strSQL= " Select P.*, " & _
		" 	case when RF.AJEXPORTACION is Null then 0 else RF.AJEXPORTACION end AJEXPORTACION, " & _
		"	case when RF.AJARROYO is Null then 0 else RF.AJARROYO end AJARROYO, " & _
		"	case when RF.AJPIEDRABUENA is Null then 0 else RF.AJPIEDRABUENA end AJPIEDRABUENA, " & _
		"	case when RF.AJTRANSITO is Null then 0 else RF.AJTRANSITO end AJTRANSITO, " & _
		"	case when RF.ASEXPORTACION is Null then 0 else RF.ASEXPORTACION end ASEXPORTACION, " & _
		"	case when RF.ASARROYO is Null then 0 else RF.ASARROYO end ASARROYO, " & _
		"	case when RF.ASPIEDRABUENA is Null then 0 else RF.ASPIEDRABUENA end ASPIEDRABUENA, " & _
		"	case when RF.ASTRANSITO is Null then 0 else RF.ASTRANSITO end ASTRANSITO, " & _
		"	case when RF.MODIFICAARTICULOS is Null then 0 else RF.MODIFICAARTICULOS end MODIFICAARTICULOS, " & _
		"	case when RF.CONFIRMACONTRATOS is Null then 0 else RF.CONFIRMACONTRATOS end CONFIRMACONTRATOS, " & _
		"	case when RF.HKEY is Null then '' else RF.HKEY end HKEY, " & _
		"	case when RF.AJPTOEXPORTACION is Null then 0 else RF.AJPTOEXPORTACION end AJPTOEXPORTACION, " & _
		"	case when RF.AJPTOARROYO is Null then 0 else RF.AJPTOARROYO end AJPTOARROYO, " & _
		"	case when RF.AJPTOPIEDRABUENA is Null then 0 else RF.AJPTOPIEDRABUENA end AJPTOPIEDRABUENA, " & _
		"	case when RF.AJPTOTRANSITO is Null then 0 else RF.AJPTOTRANSITO end AJPTOTRANSITO " & _
		" from WFPROFESIONAL P left join TBLREGISTROFIRMAS RF on P.CDUSUARIO=RF.CDUSUARIO where EGRESOVALIDO in (" 
if (verEmpleados = ESTADO_ACTIVO) then 
	strSQL  = strSQL & "'F'"
else
	strSQL  = strSQL & "'V'"
end if
strSQL  = strSQL & ")"
if (cdResponsable <> "") then strSQL= strSQL & " and P.CDUSUARIO='" & cdResponsable & "'"
if (dsResponsable <> "") then strSQL= strSQL & " and NOMBRE LIKE '%" & dsResponsable & "%'"
if (hkResponsable <> "") then 
	'Primero armo la lista
	strSQL2= "Select CDUSUARIO from TBLREGISTROFIRMAS where HKEY LIKE '%" & hkResponsable & "%'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL2)
	listaHk = ""
	while (not rs.eof) 		
		listaHk = listaHk & "'" & rs("CDUSUARIO") & "',"
		rs.MoveNext()
	wend
	'Agrego los usuarios que cumplen la condicion a la SQL principal.
	if (Len(listaHk) > 0) then
		listaHk = left(listaHk, Len(listaHk)-1)
		strSQL= strSQL & " and CDUSUARIO IN (" & listaHk & ")"
	else	
		'No hay nadie con esa llave!
		strSQL= strSQL & " and IDPROFESIONAL=0"
	end if
end if
strSQL= strSQL & " order by NOMBRE"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Call setupPaginacion(rs, pagina, regXPag) 
'Se procesa el codigo interno de la seccion.	
%>	
	
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Usuario") %>:</td>
			<td><input type="text" maxlength="10" id="cdResponsable" value="<% =cdResponsable %>" ></td>
		</tr>														
		<tr>
			<td align="right"><% = GF_TRADUCIR("Nombre") %>:</td>
			<td><input type="text" id="dsResponsable" value="<% =dsResponsable %>"></td>
		</tr>																
		<tr>
			<td align="right"><% = GF_TRADUCIR("Llave (HK)") %>:</td>
			<td><input type="text" id="hkResponsable" value="<% =hkResponsable %>"></td>
		</tr>																
		<tr>
			<td align="right"><% = GF_TRADUCIR("EStado") %>:</td>
			<td>
				<input type="radio" id="verEmpleados1" name="verEmpleados" value="<% =ESTADO_ACTIVO %>" <% if (verEmpleados = ESTADO_ACTIVO) then  Response.Write "checked" %>> Activos
				<input type="radio" id="verEmpleados2" name="verEmpleados" value="<% =ESTADO_BAJA %>"   <% if (verEmpleados <> ESTADO_ACTIVO)	then  Response.Write "checked" %>> Ex-Empleados
			</td>
		</tr>																
	</table>	  	
	-#-	
	<table id="tableSeccion3" align="center" width="80%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
	<tr class="reg_header_nav">
		<td align="center" rowspan="2">.</td>
		<td width="75" rowspan="2"><% =GF_TRADUCIR("Nombre") %></td>
		<td width="361" rowspan="2"><% =GF_TRADUCIR("Descripcion") %></td>
		<td width="200" rowspan="2"><% =GF_TRADUCIR("Llave (HK)") %></td>		
		<td align="center" colspan="4"><% =GF_TRADUCIR("Ajuste de Stock") %></td>
		<td align="center" colspan="4"><% =GF_TRADUCIR("Asientos") %></td>
		<td align="center" rowspan="2"><% =GF_TRADUCIR("Articulos") %></td>
		<td align="center" rowspan="2"><% =GF_TRADUCIR("Contratos") %></td>
		<td align="center" colspan="4"><% =GF_TRADUCIR("Ajuste de Puertos") %></td>
		<td align="center" rowspan="2" width="40"><% =GF_TRADUCIR(".") %></td>
		<td align="center" rowspan="2" width="40"><% =GF_TRADUCIR(".") %></td>
		<td align="center" rowspan="2" width="40"><% =GF_TRADUCIR(".") %></td>
	</tr>
	<tr class="reg_header_nav">
	  <td width="24" align="center"><% =GF_TRADUCIR("Exp") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Arr") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Pie") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Tra") %></td>
      <td width="24" align="center"><% =GF_TRADUCIR("Exp") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Arr") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Pie") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Tra") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Exp") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Arr") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Pie") %></td>
	  <td width="24" align="center"><% =GF_TRADUCIR("Tra") %></td>
	<tr>
	<%		i=0
	while ((not rs.eof)	and (i < regXPag))
		i = i+1					
	%>
			<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
				<td align="center" width="33"><img src="images/compras/users-16x16.png"></td>
			  <td align="center" width="75"><b>
		      <% =UCase(rs("CDUSUARIO")) %></b></td>
			  <td><% =Trim(rs("Nombre")) %></td>
				<td align="center"><% =rs("HKEY") %></td>				
				<td align="center"><% =puedeParticiparAjuste(rs("AJEXPORTACION")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJARROYO")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJPIEDRABUENA")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJTRANSITO")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("ASEXPORTACION")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("ASARROYO")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("ASPIEDRABUENA")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("ASTRANSITO")) %></td>
				<td align="center"><% =puedeModificarArticulos(rs("MODIFICAARTICULOS")) %></td>
				<td align="center"><% =puedeConfirmarContratos(rs("CONFIRMACONTRATOS")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJPTOEXPORTACION")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJPTOARROYO")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJPTOPIEDRABUENA")) %></td>
				<td align="center"><% =puedeParticiparAjuste(rs("AJPTOTRANSITO")) %></td>
				<td align="center">
				    <img src="images/access-16.png" onclick="loadPopUpResponsablesRoles('<% =rs("CDUSUARIO") %>')" title="Roles de Firma">
				</td>
				<%  if (not isAuditor(SIN_DIVISION)) then %>
				<td align="center">
					<img src="images/edit-16.png" onclick="loadPopUpResponsablesApertura(<% =rs("IDPROFESIONAL") %>)" title="Propiedades">
				</td>
				<td align="center">
					<img src="images/lock-16.png" onclick="loadPopUpResponsablesAccesos('<% =rs("CDUSUARIO") %>')" title="Accesos">
				</td>
				<%	else %>
				<td align="center">.</td>
				<td align="center">.</td>
				<%	end if %>
			</tr>
			<%			
		rs.MoveNext()
	wend
	if (i = 0) then		
	%>			
	<tr>
		<td class="TDNOHAY" colspan="23"><% =GF_TRADUCIR("No existe personal registrado") %></td>
	</tr>
	<%		end if %>
	</table>