<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->

<%
Const DIV_EXPORTACION = "1"
Const DIV_ARROYO      = "2"
Const DIV_PIEDRABUENA = "3"
Const DIV_TRANSITO    = "4"
'------------------------------------------------------------------------------------------------------
Function getValuacionCompleta(pPrecio,pExistencia,pSobrante)
	dim auxValCompleta
	auxValCompleta = ""
	if not isNull(pPrecio)then 
		if not isNull(pSobrante)then		
			if not isNull(pExistencia)then	
				auxValCompleta = (cdbl(pExistencia) + cdbl(pSobrante)) * cdbl(pPrecio)
				totValCompleta  = auxValCompleta + totValCompleta
				auxValCompleta = GF_EDIT_DECIMALS(auxValCompleta,2)				
			end if
		end if
	end if		
	getValuacionCompleta = auxValCompleta
End function
'------------------------------------------------------------------------------------------------------
Function getValuacionOperativa(pPrecio,pExistencia)
	dim auxValOperacion
	auxValOperacion	= ""
	if not isNull(pPrecio)then
		if not isNull(pExistencia)then
			auxValOperacion = cdbl(pExistencia) * cdbl(pPrecio)
			totValOperativa = auxValOperacion + totValOperativa 
			auxValOperacion = GF_EDIT_DECIMALS(auxValOperacion,2)
		end if	
	end if	
	getValuacionOperativa = auxValOperacion 
End function
'------------------------------------------------------------------------------------------------------
Function cargarFirmas(pIdVale)
	Dim rsFirmas, connFirmas, strSQL
	strSQL = "Select * from TBLVALESFIRMAS where IDVALE=" & pIdVale & " order by SECUENCIA"
	Call executeQueryDB(DBSITE_SQL_INTRA, rsFirmas, "OPEN", strSQL)
	while not rsFirmas.eof
		select case cint(rsFirmas("SECUENCIA"))
			case VS_FIRMA_RESPONSABLE
				member1Cd = rsFirmas("CDUSUARIO")
				member1 = getUserDescription(member1Cd)				
				if (rsFirmas("HKEY") <> "") then member1Firma = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			case VS_FIRMA_GERENTE
				member2Cd = rsFirmas("CDUSUARIO")
				member2 = getUserDescription(member2Cd)				
				if (rsFirmas("HKEY") <> "") then member2Firma = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			case VS_FIRMA_COORD_AUDIT
				member3Cd = rsFirmas("CDUSUARIO")
				member3 = getUserDescription(member3Cd)				
				if (rsFirmas("HKEY") <> "") then member3Firma = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))
			case VS_FIRMA_DIRECTOR				
				member4Cd = rsFirmas("CDUSUARIO")
				member4 = getUserDescription(member4Cd)				
				if (rsFirmas("HKEY") <> "") then member4Firma = armarTextoFirma(rsFirmas("HKEY"), rsFirmas("FECHAFIRMA"))	
		end select
		rsFirmas.movenext
	wend
	Call executeQueryDB(DBSITE_SQL_INTRA, rsFirmas, "CLOSE", strSQL)
End Function
'------------------------------------------------------------------------------------------------------
Function getCdObra(pIdObra) 
	Dim strSQL, rs, rtrn
	rtrn = ""
	strSQL = "Select CDOBRA from TBLDATOSOBRAS where idObra = " & pIdObra
	'Response.Write strsql
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then rtrn = rs("CDOBRA")
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getCdObra = rtrn 	
End Function
'------------------------------------------------------------------------------------------------------
Function getArticuloDatosArticulo (idAlmacen, idArticulo, ByRef dsArticulo, ByRef abrrArticulo, ByRef cdInterno)
	Dim strSQL, rs, conn
	
	call getArticuloFull (idArticulo, dsArticulo, abrrArticulo)
	'Se trae el codigo interno
	strSQL = "Select * from TBLARTICULOSDATOS where IDALMACEN=" & idAlmacen & " and  idArticulo=" & idArticulo
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then cdInterno = rs("CDINTERNO")
End Function
'------------------------------------------------------------------------------------------------------
Function puedeFirmarAjs(pCdUsuario,pFirma,pIdVale, usuarioEspecial)
	Dim rol,rtrn,rsRegistros,usrKr,aux1,divVale

	rtrn = false
	'rol = getRolFirma(pCdUsuario)	
	
	strSQL = "select * from tblregistrofirmas where cdusuario = '" & session("Usuario") & "'"
	Call executeQueryDB(DBSITE_SQL_INTRA, rsRegistros, "OPEN", strSQL)
	divVale = obtenerDivisionVale(pIdVale)
		
	rtrn = false
	if ( not rsRegistros.EoF) then
		if (rsRegistros("AJARROYO") = 1 and divVale = DIV_ARROYO) then rtrn = true
		if (rsRegistros("AJTRANSITO")= 1 and divVale = DIV_TRANSITO) then rtrn = true
		if (rsRegistros("AJPIEDRABUENA") = 1 and divVale = DIV_PIEDRABUENA) then rtrn = true
		if (rsRegistros("AJEXPORTACION") = 1 and divVale = DIV_EXPORTACION) then rtrn = true
	end if
	    
	'--------------
	if (rtrn) then
	    rtrn=false	    
		strSQL = "Select * from TBLVALESFIRMAS where IDVALE=" & pIdVale & " and SECUENCIA=" & pFirma				
		Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then		
			if ((rs("CDUSUARIO") = pCdUsuario) or (rs("CDUSUARIO")=usuarioEspecial)) then
				rtrn = true
			end if
		end if
	end if
	'--------------
	
	puedeFirmarAjs = rtrn
End Function
'****************************************************************************
'							INICIO DE LA PAGINA
'****************************************************************************
	Dim tipo,myIdvale 
	Dim strSQL, rsVale,rsValeDet, conn
	Dim dsArticulo, unidad,cdArticulo, rol, usuarioEspecial
	Dim member1Cd,member2Cd,member3Cd,member1,member2,member3,member1Firma,member2Firma, member3Firma,member4Cd,member4,member4Firma
	Dim almacenActual,totValOperativa,totValCompleta
	totValCompleta = 0
	totValOperativa = 0	
		
	tipo   = GF_PARAMETROS7("tipo","",6)
	myIdvale = GF_PARAMETROS7("idvale",0,6)
	errFirma = GF_PARAMETROS7("errFirma","",6)
	
	'Se preparan los dataos para ver si puede firmar o no.
	usuarioEspecial = ""
	rol = getRolFirma(session("Usuario"), SEC_SYS_ALMACENES)
	if (rol = FIRMA_ROL_RESP_PUERTO) then usuarioEspecial = VS_NO_USER 
	if (rol = FIRMA_ROL_AUDITOR) then usuarioEspecial = VS_AUDIT_USER
	if (rol = FIRMA_ROL_SUP_PUERTO) then usuarioEspecial = VS_PORT_SUPERVISOR_USER	
	if (rol = FIRMA_ROL_DIRECTOR) then usuarioEspecial = DIRECTOR_USER			    
	
	if (errFirma <> "") then Call setError(errFirma)

	Call cargarFirmas(myIdVale)
	
	strSQL = "select c.*,com.COMENTARIO,a.DSALMACEN from TBLVALESCABECERA c "
	strSQL = strSQL & " left join TBLVALESCOMENTARIOS com on com.IDVALE = c.IDVALE "
	strSQL = strSQL & " left join TBLALMACENES a on a.IDALMACEN = c.IDALMACEN "
	strSQL = strSQL & " where c.IDVALE = " & myIdvale 
	Call executeQueryDB(DBSITE_SQL_INTRA, rsVale, "OPEN", strSQL)
	
	strSQL = "select * from TBLVALESDETALLE where IDVALE = " & myIdvale 
	Call executeQueryDB(DBSITE_SQL_INTRA, rsValeDet, "OPEN", strSQL)
	
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
	
	var link = "almacenFirmarVales.asp?idVale=<% =myIdVale %>&secuencia=";
	var hkey0 = new Hkey('hk0', link + "<%=VS_FIRMA_RESPONSABLE%>"	, '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', link + "<%=VS_FIRMA_GERENTE    %>"	, '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', link + "<%=VS_FIRMA_COORD_AUDIT%>"	, '<% =HKEY() %>', 'check_callback()');
	<% if member4Cd <> "" then %>
		var hkey3 = new Hkey('hk3', link + "<%=VS_FIRMA_DIRECTOR%>"	, '<% =HKEY() %>', 'check_callback()');
	<% end if%>
	function check_callback(resp) {
		if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;		
		document.getElementById("frmSel").submit();
	}
	function bodyOnLoad(){
		hkey0.start();
		hkey1.start();
		hkey2.start();
		<% if member4Cd <> "" then %>
			hkey3.start();
		<% end if %>	
	}
</script>

</head>

<body onLoad="bodyOnLoad()">
<% 
select case tipo
	case CODIGO_VS_AJUSTE_STOCK
		call GF_TITULO2("kogge64.gif","Autorizaciones - Ajuste de Stock")
	case CODIGO_VS_AJUSTE_STOCK_X
		call GF_TITULO2("kogge64.gif","Autorizaciones - Anulacion de Ajuste de Stock")
	case CODIGO_VS_RECLASIFICACION_STOCK
		call GF_TITULO2("kogge64.gif","Autorizaciones - Reclasificacion de Stock")
	case CODIGO_VS_RECLASIFICACION_STOCK_X
		call GF_TITULO2("kogge64.gif","Autorizaciones - Anulacion de Reclasificacion de Stock")
end select

%>

<form name="frmSel" id="frmSel" method="POST" action="almacenValesFirma.asp?idVale=<% =myIdVale %>&tipo=<%=tipo%>">
<table width="70%" align="center" class="reg_header">
<tr>
	<td><% call showErrors() %></td>
</tr>
<tr>	
        <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="reg_header">
          <tr>
            <td colspan="2" align="center" class="reg_header_nav round_border_top_left"><%=GF_TRADUCIR("Almacen")%></td>
            <td colspan="3" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Part. Presup./Budget | Sector")%></td>
          </tr>
          <tr>
            <td colspan="2" align="center" class="recuadro reg_header_navdos"><%=rsVale("DSALMACEN")%></td>
            <td colspan="3" align="center" class="recuadro reg_header_navdos">
				<%
					if (cstr(rsVale("IDOBRA"))="0") then
						response.write "-"
					else
						response.write getCdObra(rsVale("IDOBRA")) &"-"& rsVale("IDBUDGETAREA") &"-"& rsVale("IDBUDGETDETALLE") 
					end if
				%>            </td>
          </tr>
          <tr>
            <td width="19%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Solicitado el")%></td>
            <td width="20%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Requerido para el")%></td>
            <td width="39%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Solicitante")%></td>
            <td colspan="2" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Ajusto")%></td>
          </tr>
          <tr>
            <td align="center" class="recuadro reg_header_navdos round_border_bottom_left"><%=GF_FN2DTE(rsVale("FECHA"))%></td>
            <td align="center" class="recuadro reg_header_navdos"><%=GF_FN2DTE(rsVale("FECHA"))%></td>
            <td align="center" class="recuadro reg_header_navdos">
				<%
					dsUsuario = getUserDescription(cstr(rsVale("CDSOLICITANTE")))
					response.write dsUsuario
				%>
			</td>
            <td colspan="2" align="center" class="recuadro reg_header_navdos round_border_bottom_right">
				<%
					dsUsuario = getUserDescription(cstr(rsVale("CDUSUARIO")))
					response.write dsUsuario
				%>
			</td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellpadding="2" cellspacing="1" class="reg_header">
          <tr>
            <td colspan="4" align="center" class="reg_header_nav round_border_top"><%=GF_TRADUCIR("Articulos")%></td>
            <td colspan="2" align="center" class="reg_header_nav round_border_top"><%=GF_TRADUCIR("Valuación")%></td>
          </tr>
          <tr>
            <td width="5%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Codigo")%></td>
            <td width="55%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Descripcion")%></td>
            <td width="14%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("C.Interno")%></td>
            <td width="10%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Cantidad")%></td>            
			<td width="8%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Valor Contable")%></td>
            <td width="8%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Valor Fisico")%></td>
          </tr>
          <% while not rsValeDet.EoF %>
          <tr>
          	<% Call getArticuloDatosArticulo (cstr(rsVale("idAlmacen")), rsValeDet("IDARTICULO"), dsArticulo, unidad, cdArticulo) %>
            <td align="center" class="reg_header_navdos"><%=rsValeDet("IDARTICULO")%></td>
            <td class="reg_header_navdos"><%=dsArticulo%></td>
            <td align="center" class="reg_header_navdos"><%=cdArticulo%></td>
            <td align="right" class="reg_header_navdos"><%=rsValeDet("CANTIDAD") & " " & unidad%></td>			
            <td align="right">$ <% =getValuacionOperativa(rsValeDet("VLUPESOS"),rsValeDet("EXISTENCIA")) %></td>			
            <td align="right">$ <% =getValuacionCompleta(rsValeDet("VLUPESOS"),rsValeDet("EXISTENCIA"),rsValeDet("SOBRANTE")) %></td>		
          </tr>
          <%
		  	rsValeDet.MoveNext
		  wend
		  %>
		  <tr>
		    <td colspan="4"></td>
		    <td colspan="2"><hr /></td>
		  </tr>
		  <tr>
			<td  align="right" colspan="4" class="reg_header_navdos">
				<font size="14">
					<b><%=GF_TRADUCIR("TOTAL")%></b>
				</font>
			</td>				
			<td align="right" ><b>$ <%=GF_EDIT_DECIMALS(totValOperativa,2) %></b></td>
			<td align="right" ><b>$ <%=GF_EDIT_DECIMALS(totValCompleta,2) %></b></td>						
		  </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="reg_header">
          <tr>
            <td colspan="<% if member4Cd <> "" then Response.Write 4 else Response.Write 3 end if%>" align="center" class="reg_header_nav round_border_top"><%=GF_TRADUCIR("Observaciones")%></td>
          </tr>
          <tr>
            <td colspan="<% if member4Cd <> "" then Response.Write 4 else Response.Write 3 end if%>" class="recuadro">            
				<% if (isnull(rsVale("COMENTARIO")) or (rsVale("COMENTARIO")="")) then
					response.write "&nbsp;"
                else
					response.write rsVale("COMENTARIO")
                end if%>
            </td>
          </tr>
          <tr>
            <td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Responsable")%></td>
            <td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Aprobacion Gerente de Planta")%></td>
            <td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Aprobacion Coordinador / Auditor")%></td>
            <% if member4Cd <> "" then %>            
            <td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Aprobacion Director")%></td>
            <% end if%>
          </tr>
          <tr>
            <td height="100" align="center" class="recuadro round_border_bottom_left">
					<%	
						if (member1Firma  <> "" ) then %>
                        <img src="images/firmas/<% =obtenerFirma(member1Cd) %>"><br>
                        <% =member1Firma %>
                    <%	else
                            if (member1Cd = session("Usuario")) then	%>
                                <br><div id="hk0"></div><br>
                        <%	else	%>
                                <br><br><br>
                        <%	end if	
                    end if	%>
                    ________________________________________<br />
                  <%if (member1Firma <> "") then%>
					  	<%=member1%>
                  <%else%>
						<br />
                  <%end if%>
            </td>
            <td align="center" class="recuadro" >
					<%	if (member2Firma  <> "") then %>
                        <img src="images/firmas/<% =obtenerFirma(member2Cd) %>"><br>
                        <% =member2Firma %>
                    <%	else	
							
                            if (puedeFirmarAjs(session("Usuario"),VS_FIRMA_GERENTE,myIdvale, usuarioEspecial) ) then	
                                   if (member3Cd <> session("Usuario")) then%>
  		                                <br><div id="hk1"></div><br>
                                   <%else
										response.write GF_TRADUCIR("Usted ya ha firmado como Auditor.")
                					end if%>
                            <%	else	%>
                                    <br><br><br>
                            <%	end if	
                        end if	%>
                    ________________________________________<br />
                	<%if (member2Firma <> "") then%>
                    	<%=member2%>
                    <%end if%>
            </td>
            <td align="center" class="recuadro" >
					<%	if (member3Firma  <> "") then %>
                        <img src="images/firmas/<% =obtenerFirma(member3Cd) %>"><br>
                        <% =member3Firma %>
                    <%	else								
                            if (puedeFirmarAjs(session("Usuario"),VS_FIRMA_COORD_AUDIT,myIdvale, usuarioEspecial) ) then	                            
                                   if (member2Cd <> session("Usuario") or member2Firma  = "" ) then%>                                   
  		                                <br><div id="hk2"></div><br>
                                   <%else
										response.write GF_TRADUCIR("Usted ya ha firmado como Gerente de Planta.")
                					end if%>
                            <%	else	%>
                                    <br><br><br>
                            <%	end if	
                        end if	%>
                    ________________________________________<br />
                	<%if (member3Firma <> "") then%>
                    	<%=member3%>
                    <%end if%>
            </td>
	        <% if member4Cd <> "" then %>	        
				<td align="center" class="recuadro" >
					<%	if (member4Firma  <> "") then %>
                        <img src="images/firmas/<% =obtenerFirma(member4Cd) %>"><br>
                        <% =member4Firma %>
                    <%	else	
							
                            if (puedeFirmarAjs(session("Usuario"),VS_FIRMA_DIRECTOR,myIdvale, usuarioEspecial) ) then
                                   if (member4Cd <> session("Usuario")) then%>
  		                                <br><div id="hk3"></div><br>
                                   <%else
										response.write GF_TRADUCIR("Usted ya ha firmado como Director.")
                					end if%>
                            <%	else	%>
                                    <br><br><br>
                            <%	end if	
                        end if	%>
                    ________________________________________<br />
                	<%if (member4Firma <> "") then%>
                    	<%=member4%>
                    <%end if%>
				</td>
            <% end if%>            
          </tr>
          
        </table></td>
    </tr>
    </table>
  <input type="hidden" name="errFirma" id="errFirma">
</form>
</body>
</html>
