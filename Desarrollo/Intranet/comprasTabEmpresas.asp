<!--#include file="Includes/procedimientosFormato.asp"-->
<%
Dim whereEmpresas, mkWhereEmpresas, myFetch, rsCount, listaNegra

idEmpresa = GF_PARAMETROS7("idEmpresa", "" ,6)
if (idEmpresa <> "") then Call mkWhere(mkWhereEmpresas, "A.NROEMP", idEmpresa, "=", 1)
dsEmpresa = UCase(GF_PARAMETROS7("dsEmpresa", "" ,6))
if (dsEmpresa <> "") then Call mkWhere(mkWhereEmpresas, "A.NOMEMP", dsEmpresa, "LIKE", 3)
cuit = GF_PARAMETROS7("cuit", "" ,6)
if (cuit <> "") then Call mkWhere(mkWhereEmpresas, "A.NRODOC", cuit, "=", 1)
listaNegra = GF_PARAMETROS7("listaNegra", "" ,6)
if (listaNegra) then
	Call mkWhere(mkWhereEmpresas, "B.ESTADO", 0, ">", 1)
end if

strSQL = "Select A.NROEMP AS IDEMPRESA, A.NOMEMP AS DSEMPRESA, A.NRODOC AS CUIT, B.ESTADO from [Database].[dbo].met001a A left join TBLESTADOEMPRESAS B "
strSQL = strSQL & "on A.NRODOC = B.CUIT "
strSQL = strSQL & mkWhereEmpresas & " order by A.NROEMP "
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
Call setupPaginacion(rs, pagina, regXPag)

'Se procesa el codigo interno de la seccion.	
%>	
	<% =rs.RecordCount %>-h-
	<table width="100%" align="center" border="0">
		<tr>
			<td align="right"><% = GF_TRADUCIR("Codigo") %>:</td>
			<td><input type="text" maxlength="6" id="idEmpresa" value="<% =idEmpresa %>"  onkeypress="return controlIngreso(this, event, 'N')"></td>
		</tr>														
		<tr>
			<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
			<td><input type="text" id="dsEmpresa" value="<% =dsEmpresa %>"></td>
		</tr>																
		<tr>
			<td align="right"><% = GF_TRADUCIR("C.U.I.T.") %>:</td>
			<td><input type="text" id="cuit" maxlength="11" value="<% =cuit %>"></td>
		</tr>
		<tr>
			<td align="right"><% = GF_TRADUCIR("Lista Negra") %>:</td>
			<td><input type="checkbox" name="listaNegra" id="listaNegra" <% if (listaNegra) then Response.Write " CHECKED " %>></td>
		</tr>
	</table>	  	
	-#-

	<table id="tableSeccion5" width="100%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">
		<tr class="reg_header_nav">
			<td align="center">.</td>
			<td width="10%" align="center"><% =GF_TRADUCIR("Codigo") %></td>
			<td width="10%" align="center"><% =GF_TRADUCIR("CUIT") %></td>
			<td><% =GF_TRADUCIR("Descripcion") %></td>
			<td width="24px" align="center"><% =GF_TRADUCIR(".") %></td>
			<td align="center" width="24px">.</td>
		</tr>
<%		i=0
		while ((not rs.eof)	and (i < regXPag))		
			i = i+1			
%>
				<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
					<td align="center" width="16px"><img src="images/compras/Company-16x16.png"></td>
					<td align="center"><b><% =rs("IDEMPRESA") %></b></td>
					<td align="center"><b><% =GF_STR2CUIT(rs("CUIT")) %></b></td>					
					<td><% =rs("DSEMPRESA") %></td>
					<td align="center"><% =getEstado(rs("ESTADO")) %></td>
					<td align="center" width="24px">
						<%  if (not isAuditor(SIN_DIVISION)) then %>
						<img src="images/compras/edit-16x16.png" style="cursor: pointer" onClick="loadPopUpEmpresas(<% =rs("IDEMPRESA") %>)">
						<%  end if  %>
					</td>
				</tr>
<%			rs.MoveNext()
		wend
		if (i = 0) then		
%>			
		<tr>
			<td class="TDNOHAY" colspan="6"><% =GF_TRADUCIR("No existen empresas registradas.") %></td>
		</tr>
<%		end if %>
	</table>
</body>
</html>
<%
'-------------------------------------------
Function getEstado(estado)
	if ((estado="") or (isnull(estado))) then estado=0
	Select case (cInt(estado))
		case PROV_ACTIVO:
			getEstado = "<img alt='Activa' title='Activa' src='images/compras/action_ok-16x16.png'>"
		case PROV_PROHIBIDO_PEDIDOS:
			getEstado = "<img alt='No puede recibir pedidos' title='No puede recibir pedidos' src='images/compras/action_warning-16x16.png'>"
		case PROV_PROHIBIDO_PAGOS:
			getEstado = "<img alt='No puede recibir pedidos ni pagos' title='No puede recibir pedidos ni pagos' src='images/compras/action_error-16x16.png'>"
	end Select
End Function
%>