<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<%
Call comprasControlAccesoCM(RES_AUD)

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'----------------------				COMIENZO DE PAGINA				-------------------
'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
Dim accion, idProveedor, rsProv


accion = GF_PARAMETROS7("accion","",6)
idProveedor = GF_PARAMETROS7("idProveedor","",6)

if (accion = ACCION_GRABAR) then Call executeProcedureDb(DBSITE_SQL_INTRA, rsProv, "TBLPROVEEDORESCD_INS", idProveedor)        
if (accion = ACCION_BORRAR) then Call executeProcedureDb(DBSITE_SQL_INTRA, rsProv, "TBLPROVEEDORESCD_DEL", idProveedor)        

Call executeProcedureDb(DBSITE_SQL_INTRA, rsProv, "TBLPROVEEDORESCD_GET", "")

%>
<html>
<head>
<title>Proveedores Pre-autorizados para Compras Directas</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" >
    function bodyOnLoad() {
		var tb = new Toolbar('toolbar', 5, "");	
		tb.addButton("toolbar-home", "Home", "irA('almacenAuditoria.asp')");		
		tb.draw();				
	}
	
	function irA(pLink) {
		location.href = pLink;
	}
	
    function baja(id) {
        document.getElementById("accion").value='<% =ACCION_BORRAR %>';
        document.getElementById("idProveedor").value= id;
        document.getElementById("frmSel").submit();
    }
    function alta() {
        document.getElementById("accion").value='<% =ACCION_GRABAR %>';
        document.getElementById("idProveedor").value=document.getElementById("prove").value;
        document.getElementById("frmSel").submit();
    }
    
</script>
</head>
<body onload="javascript:bodyOnLoad()">
<div id="toolbar"></div>
<form name="frmSel" id="frmSel" method="post" action="comprasProveedoresCD.asp">
<table align="center" border=0 width="100%">
	<tr>
	    <td>
	        <table class="datagrid" align="center" width="30%">
	            <thead>
                    <th>ID Proveedor</th>
                    <th>Raz&oacute;n Social</th>
                    <th>.</th>
	            </thead>
	            <tbody>	            
	        <%  
	        while (not rsProv.eof)
	        %>   
	            <tr>
                    <td><% =rsProv("IDPROVEEDOR") %></td>
                    <td><% =rsProv("NOMEMP") %></td>
                    <td align="center"><img style="cursor:pointer" src="images/cross-16.png" onclick="javascript:baja(<% =rsProv("IDPROVEEDOR") %>)" /></td>
                </tr> 
            <%
                rsProv.MoveNext()
            wend    
            %>   
                <tr>
                    <td><input id="prove" /></td>
                    <td></td>
                    <td align="center"><img style="cursor:pointer" src="images/plus-16.png" onclick="javascript:alta()" /></td>
                </tr>
                </tbody>                
            </table>
        </td> 	
    </tr>    
</table>
<input type="hidden" id="accion" name="accion" value="">
<input type="hidden" id="idProveedor" name="idProveedor" value="">
</form>
</body>
</html>