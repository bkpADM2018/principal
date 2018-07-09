<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_ADM)


Const TOTAL_DIVISIONES = 4 ' NO CUNATA EL 0 - TOTAL 4

'--------------------------------------------------------------------------------------
Function writeCombo(sistema, usr, recurso, division)
	dim SelectName, strSQL, rsPermiso, con, auxpermiso

	auxpermiso = SEC_X
	SelectName = "Select_" & recurso & "_" & division & "_" & sistema
	
	strSQL= "Select * from TBLUSUARIOPERMISOS where CDUSUARIO='" & usr & "' "
	strSQL= strSQL & " and IDDIVISION = " & division & " and IDRECURSO = " & recurso & " and IDSISTEMA=" & sistema
	Call executeQueryDB(DBSITE_SQL_INTRA, rsPermiso, "OPEN", strSQL)
	if (not rsPermiso.eof) then	auxpermiso = cInt(rsPermiso("PERMISO"))

	%>
	<Select id="<% =SelectName %>" name="<% =SelectName %>">
		<option value="<% =SEC_X %>" <% if (auxpermiso = SEC_X)  then Response.Write "selected" %>><% =GF_Traducir("Denegado") %>
		<option value="<% =SEC_U %>" <% if (auxpermiso = SEC_U)  then Response.Write "selected" %>><% =GF_Traducir("Usuario") %>
		<option value="<% =SEC_Y %>" <% if (auxpermiso = SEC_Y)  then Response.Write "selected" %>><% =GF_Traducir("Auditor") %>
		<option value="<% =SEC_A %>" <% if (auxpermiso = SEC_A)  then Response.Write "selected" %>><% =GF_Traducir("Admin") %>
	</select>
	<%
End Function
'--------------------------------------------------------------------------------------
Function existeRegistro(p_usr, p_division, p_recurso, p_sistema)
	dim str, cn, rs
	existeRegistro = false
	str= "Select * from TBLUSUARIOPERMISOS where CDUSUARIO='" & p_usr & "' "
	str= str & " and IDDIVISION = " & p_division & " and IDRECURSO = " & p_recurso & " and IDSISTEMA=" & p_sistema
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", str)
	if (not rs.eof) then existeRegistro = true
End Function
'----------------------------------------------------------------------------------------
Function loadRecursoDicc()
	'------ COMPRAS ------
	if(not oDiccIndexRecurso.Exists(RES_CC))then oDiccIndexRecurso.Add RES_CC, "C. Concursos"
	if(not oDiccIndexRecurso.Exists(RES_CD))then oDiccIndexRecurso.Add RES_CD, "C. Directa"
	if(not oDiccIndexRecurso.Exists(RES_OBR))then oDiccIndexRecurso.Add RES_OBR, "Obras"
	if(not oDiccIndexRecurso.Exists(RES_AFE))then oDiccIndexRecurso.Add RES_AFE, "Afes"
	if(not oDiccIndexRecurso.Exists(RES_AUD))then oDiccIndexRecurso.Add RES_AUD, "Auditoria"
	if(not oDiccIndexRecurso.Exists(RES_ADM))then oDiccIndexRecurso.Add RES_ADM, "Administración"
	if(not oDiccIndexRecurso.Exists(RES_ACC_AL))then oDiccIndexRecurso.Add RES_ACC_AL, "Contabilidad"
	if(not oDiccIndexRecurso.Exists(RES_PDC))then oDiccIndexRecurso.Add RES_PDC, "Poliza Caucion"
	'------ PROVEEDORES ------
	if(not oDiccIndexRecursoProv.Exists(RES_PRV_MASTER))then oDiccIndexRecursoProv.Add RES_PRV_MASTER, "Maestro de Proveedores"
	'------ FACTURACION ------
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_MERCADERIAS))then oDiccIndexRecursoFac.Add RES_FAC_MER, "Mercaderias"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_EJL))then oDiccIndexRecursoFac.Add RES_FAC_EJL, "Ej. Intl. Local"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_EJE))then oDiccIndexRecursoFac.Add RES_FAC_EJE, "Ej. Intl. Expo"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_CG))then oDiccIndexRecursoFac.Add RES_FAC_CG, "Contaduria"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_TRA))then oDiccIndexRecursoFac.Add RES_FAC_TRA, "El Transito"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_ARR))then oDiccIndexRecursoFac.Add RES_FAC_ARR, "Arroyo"
	if(not oDiccIndexRecursoFac.Exists(RES_FAC_LPB))then oDiccIndexRecursoFac.Add RES_FAC_LPB, "Bahia Blanca"
	'------ MERCADERIAS ------
	if(not oDiccIndexRecursoMer.Exists(RES_MER_ANALISIS))then oDiccIndexRecursoMer.Add RES_MER_ANALISIS, "Analisis"
	if(not oDiccIndexRecursoMer.Exists(RES_MER_RECEPCOND))then oDiccIndexRecursoMer.Add RES_MER_RECEPCOND, "Recep. Garantias"
    if(not oDiccIndexRecursoMer.Exists(RES_MER_GTACONTRATO_BSAS))then oDiccIndexRecursoMer.Add RES_MER_GTACONTRATO_BSAS, "Garantias - Bs As"
	if(not oDiccIndexRecursoMer.Exists(RES_MER_GTACONTRATO_ROS))then oDiccIndexRecursoMer.Add RES_MER_GTACONTRATO_ROS, "Garantias - Rosario"
	'------ POSEIDON ------
	if(not oDiccIndexPoseidon.Exists(RES_PSD_ANALISIS))then oDiccIndexPoseidon.Add RES_PSD_ANALISIS, "Info Analisis"
End Function
'----------------------------------------------------------------------------------------
'Modificacion: Solo para el sistema de Mercaderias, recurso Contrato de Garantias (3 y 4) no mostrará todas las divisiones, solo exportacion
'              debido a que el mismo recurso genera el vinculo con la division
Function drawSistema(sistema, titulo, dicc, resp)
    Dim idDivision
%>
    <tr><td>&nbsp;</td></tr>
	<tr><td>
        <table class="reg_header">
        <tr>
        <th class="reg_header_nav round_border_top"><% =titulo %></th>
        </tr>
        <tr>
        <td>
            <table class="reg_header" align="center">
                <tr class="reg_header_nav">
                    <td>&nbsp;</td>
                    <td align="center"><% =GF_Traducir("Exportación") %></td>
                    <td align="center"><% =GF_Traducir("Arroyo") %></td>
                    <td align="center"><% =GF_Traducir("Piedrabuena") %></td>
                    <td align="center"><% =GF_Traducir("Transito") %></td>
                </tr>
                <% For each recurso in dicc	%>
                    <tr>
                        <td class="reg_header_navdos"><% =GF_TRADUCIR(dicc(recurso)) %></td>
                        <% for idDivision = 1 to TOTAL_DIVISIONES %>                        
                            <td><% =writeCombo(sistema, resp, recurso, idDivision) %></td>
                        <% next %>
                    </tr>
                <% next %>
            </table>
        </td>
        </tr>        
        </table>
    </td></tr>
<%
End Function
'--------------------------------------------------------------------------------------
Function savePermisos(dicc, resp)
    Dim idDivision, recurso, rsRegistro, arr
    
    For each myKey in dicc.Keys				
        arr = Split(myKey, "|")
        idDivision = arr(0)
        idSistema = arr(1)
        recurso = arr(2)        
        if (existeRegistro(resp, idDivision, recurso, idSistema)) then
	        strSQL="Update TBLUSUARIOPERMISOS set PERMISO=" & dicc(myKey) & " where CDUSUARIO='" & UCase(resp) & "' and IDDIVISION =" & idDivision & " and IDRECURSO =" & recurso	& " and IDSISTEMA=" & idSistema
        else	            
	        strSQL="Insert into TBLUSUARIOPERMISOS values (" & idSistema & ", '" & UCase(resp) & "'," & idDivision & ", " & recurso & ", " & dicc(myKey) & ") "
        end if
        Call executeQueryDB(DBSITE_SQL_INTRA, rsRegistro, "EXEC", strSQL)
    next
    
End Function
'--------------------------------------------------------------------------------------
'----------------------				COMIENZO DE PAGINA				-------------------
'--------------------------------------------------------------------------------------
Dim accion, cdResponsable, dsResponsable, strSQL, oDiccIndexPoseidon
Dim idRecurso, idDivision, oDiccIndexRecurso, myKey, oDiccPermiso, oDiccIndexRecursoProv, oDiccIndexRecursoFac, oDiccIndexRecursoMer

Set oDiccIndexRecurso  = createObject("Scripting.Dictionary")
Set oDiccIndexRecursoProv  = createObject("Scripting.Dictionary")
Set oDiccIndexRecursoFac  = createObject("Scripting.Dictionary")
Set oDiccIndexRecursoMer  = createObject("Scripting.Dictionary")
Set oDiccPermiso	   = createObject("Scripting.Dictionary")
Set oDiccIndexPoseidon	   = createObject("Scripting.Dictionary")

Call loadRecursoDicc()

cdResponsable = GF_PARAMETROS7("cdResponsable", "",6)
accion = GF_PARAMETROS7("accion","",6)

'obtengo el permiso del usuario enviados en el form, segun recurso y division
for idDivision = 1 to TOTAL_DIVISIONES
	For each recurso in oDiccIndexRecurso
		if(not oDiccPermiso.Exists(idDivision & "|" & SEC_SYS_UNASIGNED & "|" & recurso))then oDiccPermiso.Add idDivision & "|" & SEC_SYS_UNASIGNED & "|" & recurso, GF_PARAMETROS7("Select_"&recurso&"_"&idDivision&"_"&SEC_SYS_UNASIGNED,0,6)	    
	next
	For each recurso in oDiccIndexRecursoProv		
		if(not oDiccPermiso.Exists(idDivision & "|" & SEC_SYS_UNASIGNED & "|" & recurso))then oDiccPermiso.Add idDivision & "|" & SEC_SYS_UNASIGNED & "|" & recurso, GF_PARAMETROS7("Select_"&recurso&"_"&idDivision&"_"&SEC_SYS_UNASIGNED,0,6)
	next
	For each recurso in oDiccIndexRecursoFac		
		if(not oDiccPermiso.Exists(idDivision & "|" & SEC_SYS_FACTURACION & "|" & recurso))then oDiccPermiso.Add idDivision & "|" & SEC_SYS_FACTURACION & "|" & recurso, GF_PARAMETROS7("Select_"&recurso&"_"&idDivision&"_"&SEC_SYS_FACTURACION,0,6)
	next
	For each recurso in oDiccIndexRecursoMer		
		if(not oDiccPermiso.Exists(idDivision & "|" & SEC_SYS_MERCADERIAS & "|" & recurso))then oDiccPermiso.Add idDivision & "|" & SEC_SYS_MERCADERIAS & "|" & recurso, GF_PARAMETROS7("Select_"&recurso&"_"&idDivision&"_"&SEC_SYS_MERCADERIAS,0,6)
	next
	For each recurso in oDiccIndexPoseidon		
		if(not oDiccPermiso.Exists(idDivision & "|" & SEC_SYS_POSEIDON & "|" & recurso))then oDiccPermiso.Add idDivision & "|" & SEC_SYS_POSEIDON & "|" & recurso, GF_PARAMETROS7("Select_"&recurso&"_"&idDivision&"_"&SEC_SYS_POSEIDON,0,6)
	next
next

'Obtener datos profesionales
dsResponsable = getUserDescription(cdResponsable)

if (accion = ACCION_GRABAR) then
	if (cdResponsable <> "") then
	    Call savePermisos(oDiccPermiso, cdResponsable)		
	end if
end if
%>
<html>
<head>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" type="text/css">
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="Scripts/botoneraPopUp.js"></script>
<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">
	var botones = new botonera("botones");
	function responsableOnLoad() {
		refPopUpResponsableAccesos = getObjPopUp('popupResponsableAccesos');
		<% if ((accion = ACCION_CERRAR)or (accion = ACCION_GRABAR)) then %>
			refPopUpResponsableAccesos.hide();
		<% end if %>
		pngfix();
		<%  if (not isAuditor(SIN_DIVISION)) then %>
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
<form name="frmSel" id="frmSel" method="post" action="comprasPropResponsableAccesos.asp">
<table align="center" border=0 width="100%">
	<tr>
		<td class="title_sec_section" align="left">
			<img align="absMiddle" src="images/compras/users-32x32.png">
			<b><% =UCase(cdResponsable) %> - <% =UCase(dsResponsable) %></b>
		</td>
	</tr>	
	<tr>
		<td><% call showErrors() %></td>
	</tr>	
    <%  
    Call drawSistema(SEC_SYS_UNASIGNED, "Sistema de Compras", oDiccIndexRecurso, cdResponsable)
    Call drawSistema(SEC_SYS_UNASIGNED, "Sistema de Proveedores", oDiccIndexRecursoProv, cdResponsable)
    Call drawSistema(SEC_SYS_FACTURACION, "Sistema de Facturacion", oDiccIndexRecursoFac, cdResponsable)
    Call drawSistema(SEC_SYS_MERCADERIAS, "Sistema de Mercaderias", oDiccIndexRecursoMer, cdResponsable)
    Call drawSistema(SEC_SYS_POSEIDON, "Sistema de Administraci&oacute;n Portuaria", oDiccIndexPoseidon, cdResponsable)
    %>    
	<tr><td></td></tr>
	<tr>	
		<td>
			<table align="right">	
				<tr>
					<td>
					<%  if (not isAuditor(SIN_DIVISION)) then %>
                    	<div id="botones"></div>
						<!--<input type="submit" id="aceptar" name="aceptar" value="<% =GF_TRADUCIR("Aceptar") %>" <% if (idResponsable = 0) then response.write "disabled=true" %>>					-->
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