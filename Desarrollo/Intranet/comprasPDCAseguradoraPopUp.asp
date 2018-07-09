<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosvalidacion.asp"-->
<% 
'---------------------------------------------------------------------------------------------
Function addAseguradora(pDsAseguradora, pCuit)
	Dim strSQl, rs 	
	strSQL = "INSERT INTO TBLPDCASEGURADORAS(DSASEGURADORA,CUIT) VALUES ('" & Trim(Ucase(pDsAseguradora)) & "', " & pCuit & ")" 
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End function	
'---------------------------------------------------------------------------------------------
Function getAseguradoraDuplicada(pCuit)
	Dim strSQl, rs
	rtrn = false	
	strSQL = "SELECT CUIT AS CANT FROM TBLPDCASEGURADORAS WHERE CUIT = " & pCuit		
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if(rs.Eof)then rtrn = true	
	getAseguradoraDuplicada = rtrn
End function	
'---------------------------------------------------------------------------------------------------
Function controlarAseguradora(dsAseguradora, pCuit)
	Dim flagGuardar
	flagGuardar = false
	if(dsAseguradora <> "")then
		if (GF_CONTROL_CUIT(pCuit)) then			
			if(getAseguradoraDuplicada(pCuit))then				
				flagGuardar = true			
			else
				Call setError(EMPRESA_EXISTE)	
			end if	
		else
			Call setError(CUIT_ERRONEO)
		end if	
	else
		Call setError(DESCRIPCION_VACIA)
	end if	
	controlarAseguradora = flagGuardar
End Function
'*********************************************************************************************'
'********************************	INICIO PAGINA  *******************************************'
'*********************************************************************************************'
Dim idAseguradora, dsAseguradora, monto, cuit1, cuit2, cuit3, cuitEmpresa

accion		  = GF_PARAMETROS7("accion","",6)
dsAseguradora = Trim(Ucase(GF_PARAMETROS7("dsAseguradora", "", 6)))
cuit1 = GF_PARAMETROS7("cuit1","",6)
cuit2 = GF_PARAMETROS7("cuit2","",6)
cuit3 = GF_PARAMETROS7("cuit3","",6)

if(accion = ACCION_GRABAR)then
	cuitEmpresa = cuit1 & cuit2 & cuit3
	if(controlarAseguradora(dsAseguradora, cuitEmpresa)) then Call addAseguradora(dsAseguradora, cuitEmpresa)
end if

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Agregar Aseguradora") %></title>
<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<SCRIPT type="text/javascript" src="scripts/controles.js"></SCRIPT>
<script type="text/javascript" src="Scripts/botoneraPopUp.js"></script>
<script type="text/javascript">
	var botones = new botonera("botones");
	
	function onLoadPage(){
		botones.addbutton('<%=GF_Traducir("Guardar")%>','submitInfo()');
		botones.show();	
		document.getElementById("msgGuardado").innerHTML = "";
		<% if(flagGuardar)then %>
			document.getElementById("msgGuardado").innerHTML="<% =GF_TRADUCIR("Se guardo correctamente") %>";
			document.getElementById("msgGuardado").className = "TDSUCCESS";	
		<% end if %>	
	}
	
	
	function submitInfo(){
		document.getElementById("myForm").submit();
	}
	
	
	
	</SCRIPT>
</HEAD>
<BODY onload="onLoadPage();">
	<FORM id="myForm" name="myForm" action="comprasPDCAseguradoraPopUp.asp" method="post">
	<INPUT type='hidden' name='accion' id='accion' value='<%=ACCION_GRABAR%>'>
	<% call showErrors() %>
	<TABLE  width="100%">
		<TR>
			<TD colspan="2">
				<DIV id="msgGuardado" align="center" class="TDBAJAS"></DIV>
			</TD>						
		</TR>
		<TR>			
			<TD width="30%" class="reg_header">
				<% =GF_TRADUCIR("Descripcion") %>
			</TD>
			<TD>
				<input type="text" id="dsAseguradora" name="dsAseguradora" value="<%=dsAseguradora%>" size="40">
			<TD>
		</TR>		
		<TR>			
			<TD width="30%" class="reg_header">
				<% =GF_TRADUCIR("CUIT") %>
			</TD>
			<TD>
				<INPUT type="text" id="cuit1" name="cuit1" maxlength="2" size="2" value="<% =cuit1 %>" onKeyPress="return controlDatos(this, event, 'N')">-
				<INPUT type="text" id="cuit2" name="cuit2" maxlength="8" size="8" value="<% =cuit2 %>" onKeyPress="return controlDatos(this, event, 'N')">-
				<INPUT type="text" id="cuit3" name="cuit3" maxlength="1" size="1" value="<% =cuit3 %>" onKeyPress="return controlDatos(this, event, 'N')">
			<TD>
		</TR>
	</TABLE>
	<div id="botones"></div>
	</FORM>
</BODY>
</HTML>

