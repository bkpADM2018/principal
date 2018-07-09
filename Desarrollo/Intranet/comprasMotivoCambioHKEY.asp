<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
Dim idResponsable,cdResponsable,dsResponsable,accion,momento

idResponsable 	= GF_PARAMETROS7("idResponsable",0,6)
myHkeyNueva 	= GF_PARAMETROS7("HkeyNew","",6)
myHkeyVieja 	= GF_PARAMETROS7("HkeyOld","",6)
accion 			= GF_PARAMETROS7("accion","",6)
momento 		= GF_PARAMETROS7("momento","",6)
motivo	 		= GF_PARAMETROS7("motivo","",6)



strSQL="Select * from WFPROFESIONAL where EGRESOVALIDO = 'F' and IdProfesional=" & idResponsable
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	cdResponsable = rs("CDUSUARIO")
	dsResponsable = rs("NOMBRE")
end if

Call GP_ConfigurarMomentos()

if (accion = ACCION_GRABAR) then
	if (motivo <> "") then
		strSQL = "update TBLLOGHKEY set motivo = '"&motivo&"' where MMTOLOG= " & momento & " and CDUSUARIO = '"&cdResponsable&"' and CDUSRLOG = '"&session("usuario")&"'" 
		 Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	end if
else
	momento = session("mmtosistema")
	strSQL = "insert into TBLLOGHKEY (HKEY,HKEYOLD,MOTIVO,CDUSUARIO,CDUSRLOG,MMTOLOG) values('"&myhkeyNueva&"','"&myHkeyVieja&"','El usuario no ha dado motivos','"&cdResponsable&"','"&session("usuario")&"',"&momento&")"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
end if

if (myHkeyNueva = "")  then myhKeyNueva = "Ninguna"
if (myHkeyVieja = "")  then myHkeyVieja = "Ninguna"

%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Justificacion cambio HKEY</title>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript">
	refPopUpJustificacion = getObjPopUp('popupResponsable');
	refPopUpJustificacion.resize(500,300);

	function justificacionHKEYOnLoad(){
		<% if (accion = ACCION_GRABAR) then %>
			refPopUpJustificacion.hide();
		<% end if %>	
		document.getElementById("motivo").focus();
	}
</script>

</head>

<body onload="justificacionHKEYOnLoad()">
<br>
<form action="comprasMotivoCambioHKEY.asp" method="post">
    <table width="450" border="0" align="center" cellpadding="0" cellspacing="0" class="reg_header round_border_all">
      <tr align="center">
        <td colspan="3" class="reg_header_nav round_border_top">Justifique el cambio de HKEY</td>
      </tr>
      <tr>
        <td width="20px">&nbsp;</td>
        <td width="410" align="right" >
        <br />
        <label><%=GF_FN2DTE(momento)%></label><input type="hidden" id="momento" name="momento" value="<%=momento%>">
        <br /><br />
        <table width="100%" border="0" cellpadding="1" cellspacing="1" class="reg_header">
          <tr>
            <td colspan="4" align="center" class="round_border_top reg_header_nav">HKEY</td>
            </tr>
          <tr>
            <td width="100" align="right" class="reg_header_navdos">Nuevo&nbsp;</td>
            <td width="100" align="center" >
				<%=myHkeyNueva%>
				<input type="hidden" id="hkeynew" name="hkeynew" value="<%=myHkeyNueva%>">
			</td>
            <td width="100" align="right" class="reg_header_navdos" >Anterior&nbsp;</td>
            <td width="100" align="center" >
				<%=myHkeyVieja%>
				<input type="hidden" id="hkeyold" name="hkeyold" value="<%=myHkeyVieja%>">
			</td>
          </tr>
          <tr>
            <td colspan="4" align="center" class="reg_header_nav">Usuarios</td>
            </tr>
          <tr>
            <td align="right" class="reg_header_navdos">Responsable&nbsp;</td>
            <td align="center">
				<%=cdResponsable%>
				<input type="hidden" id="idResponsable" name="idResponsable" value="<%=idResponsable%>">
			</td>
            <td align="right" class="reg_header_navdos">Cambio&nbsp;</td>
            <td align="center"><%=session("Usuario")%></td>
          </tr>
          <tr>
            <td colspan="4" align="left" class="reg_header_nav">Motivo</td>
            </tr>
          <tr>
            <td colspan="4" align="left" valign="top"><textarea name="motivo" id="motivo" cols="65" rows="5"></textarea></td>
            </tr>
    
          <tr>
            <td colspan="4" align="right"><input type="submit" name="aceptar" id="aceptar" value="Aceptar" class="round_border_bottom_right" /></td>
            </tr>
        </table></td>
        <td width="20px">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<input type="hidden" id="accion" name="accion" value="<%=ACCION_GRABAR%>">
</form>
</body>
</html>
