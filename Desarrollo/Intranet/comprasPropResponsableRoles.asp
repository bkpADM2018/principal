<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)

'Const TOTAL_RECURSOS = 6 'SE CUENTA EL 0 - TOTAL 7
'MODIFICACION: SE AGREGO UN NUEVO RECURSO, EL DE POLIZA DE CAUCION, SU ID VA A ALTERAR EL ORDEN YA QUE ES 13
Const TOTAL_RECURSOS = 7 'SE CUENTA EL 0 - TOTAL 8
Const TOTAL_DIVISIONES = 4 ' NO CUNATA EL 0 - TOTAL 4

'--------------------------------------------------------------------------------------
Function writeCombo(pIdProducto, pRsRoles, pIdRol)	
	%>
	<Select id="P<% =pIdProducto %>" name="P<% =pIdProducto %>">
	    <option value="<% =FIRMA_ROL_NINGUNO %>" ><% =GF_TRADUCIR("Sin rol específico") %></option>
	<%  while (not pRsRoles.eof) %>
		<option value="<% =pRsRoles("IDROL")%>" <% if (CInt(pIdRol) = CInt(pRsRoles("IDROL")))  then Response.Write "selected" %>><% =GF_Traducir(pRsRoles("DSROL")) %></option>
    <%      pRsRoles.MoveNext()
        wend %>		
	</select>
	<%
End Function
'----------------------------------------------------------------------------------------
Function drawSistema(pIdProducto, pDsProducto, pCdResponsable)
    Dim rsRolUsuario, rsRoles, idRolUsuario
    Call executeProcedureDb(DBSITE_SQL_INTRA, rsRoles, "TBLROLES_GET_BY_IDPRODUCTO", pIdProducto)
    Call executeProcedureDb(DBSITE_SQL_INTRA, rsRolUsuario, "TBLROLESUSUARIOS_GET_BY_CDUSER_IDSISTEMA", pCdResponsable & "||" & pIdProducto)
    idRolUsuario = FIRMA_ROL_NINGUNO
    if (not rsRolUsuario.eof) then idRolUsuario = CInt(rsRolUsuario("IDROL"))
%>
	<tr>
	    <td><% =pDsProducto %></td>
        <td>
            <% Call writeCombo(pIdProducto, rsRoles, idRolUsuario) %>
        </td>
    </tr>
<%
End Function
'--------------------------------------------------------------------------------------
Function saveRolAsignado(pCdResponsable, pIdProducto, pIdRol)
    Dim rsRolUsuario
    
    if (pIdRol = FIRMA_ROL_NINGUNO) then
    'No se asignó ningun rol. O se quiere borrar una asignación previa.
		Call executeProcedureDb(DBSITE_SQL_INTRA, rsRolUsuario, "TBLROLESUSUARIOS_DEL", pCdResponsable & "||" & pIdProducto)
    else
    'Se graba el rol.
        Call executeProcedureDb(DBSITE_SQL_INTRA, rsRolUsuario, "TBLROLESUSUARIOS_INS", pCdResponsable & "||" & pIdProducto & "||" & pIdRol)
    end if
End Function
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'----------------------				COMIENZO DE PAGINA				-------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
Dim accion, cdResponsable, dsResponsable, strSQL
Dim idRecurso, rsProductos, idRolAsignado

cdResponsable = GF_PARAMETROS7("cdResponsable","",6)
accion = GF_PARAMETROS7("accion","",6)
Call executeProcedureDb(DBSITE_SQL_INTRA, rsProductos, "TBLSYSPRODUCTOS_GET", "")
if (accion = ACCION_GRABAR) then
    'Se leen los parametros de permisos para cada producto y se graban.
    while (not rsProductos.eof)
        idRolAsignado = GF_PARAMETROS7("P" & rsProductos("IDPRODUCTO"),0,6)
        Call saveRolAsignado(cdResponsable, rsProductos("IDPRODUCTO"), idRolAsignado)
        rsProductos.MoveNext()
    wend
    rsProductos.MoveFirst()
end if
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css">
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="Scripts/botoneraPopUp.js"></script>
<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">
	var botones = new botonera("botones");
	function responsableOnLoad() {
		var refPopUpResponsableRoles = getObjPopUp('popupResponsableRoles');
		<% if ((accion = ACCION_CERRAR) or (accion = ACCION_GRABAR)) then %>
			refPopUpResponsableRoles.hide();
		<% end if 
		   if (not isAuditor(SIN_DIVISION)) then %>
			botones.addbutton('<%=GF_TRADUCIR("Aceptar")%>','submitir()');
			botones.show();
		<% end if %>
	}
	function submitir()
	{
		$("#frmSel").submit();
	}
</script>
</head>
<body onLoad="responsableOnLoad()">
<form name="frmSel" id="frmSel" method="post" action="comprasPropResponsableRoles.asp">
<table align="center" border=0 width="100%">
	<tr>
		<td class="title_sec_section" align="left" colspan="2">
			<img align="absMiddle" src="images/access-50.png">
			<b><% =UCase(cdResponsable) %> - <% =UCase(getUserDescription(cdResponsable)) %></b>
		</td>
	</tr>	
	<%  
	while (not rsProductos.eof)
        Call drawSistema(rsProductos("IDPRODUCTO"), rsProductos("DSPRODUCTO"), cdResponsable)
        rsProductos.MoveNext()
    wend    
    %>    
	<tr><td colspan="2"></td></tr>
	<tr>	
		<td colspan="2">
			<table align="right">	
				<tr>
					<td>
					<%  if (not isAuditor(SIN_DIVISION)) then %>
                    	<div id="botones"></div>
					<%	end if %>
					</td>
				</tr>
			</table>
		</td>		
	</tr>	
	<tr><td>&nbsp;</td></tr>
	<tr><td>&nbsp;</td></tr>
</table>
<input type="hidden" name="accion" value="<% =ACCION_GRABAR %>">
<input type="hidden" name="cdResponsable" value="<% =cdResponsable %>">
</form>
</body>
</html>