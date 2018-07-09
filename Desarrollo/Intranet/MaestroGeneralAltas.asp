<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<% ProcedimientoControl "MG240" %> 
<% 
dim p_km,P_Mensaje, habil
p_km = GF_ParametrosForm("p_km")
P_mensaje = GF_ParametrosForm("p_mensaje")
%>
<ScRIPT LANGUAGE="javascript">
<!--
function form1_onsubmit()
{
  if (document.form1.FORM_KC.value == "" || document.form1.FORM_KC.value == "?")
   {
    alert ("Debe ingresar el codigo");
	return false;
   }
  if (document.form1.FORM_DS.value == "" || document.form1.FORM_DS.value == "?")
   {
    alert ("Debe ingresar la descripcion");
	return false;
   }
   return true
}
//-->   
</SCRIPT>

<html>
<head>
<title>Altas de Registros</title> <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"> 
<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<br><br><br><br>
<table class="reg_header" border="0" ALIGN="CENTER" CELLSPACING="1" CELLPADDING="2"> 
<form name="form1" method="post" action="MaestroGeneralAltas2.asp" onsubmit = "return form1_onsubmit()"> 
 <tr> 
  <td colspan="2" > 
   <b><%=GF_TRADUCIR("Actualizar Maestro")%></b> 
  </td>
 </tr> 
 <tr> 
  <td class="reg_header_nav">
    <DIV ALIGN="RIGHT"><%=GF_TRADUCIR("Maestro")%></DIV>
  </td>
  <td width="53%" height="33" class="reg_header_navdos"> 
    <input type="text" name="FORM_KM" VALUE="<%=p_km%>" > 
  </td>
 </tr> 
 <tr> 
  <td class="reg_header_nav">
    <DIV ALIGN="RIGHT"><% =GF_TRADUCIR("Codigo") %></DIV>
  </td>
  <td width="53%" height="33" class="reg_header_navdos"> 
    <input type="text" name="FORM_KC" VALUE="?" maxlength="10">
  </td>
 </tr> 
 <tr> 
  <td width="47%" class="reg_header_nav"> 
    <DIV ALIGN="RIGHT"><% =GF_TRADUCIR("Descripcion") %></DIV>
  </td>
  <td width="53%" height="33" class="reg_header_navdos"> 
    <input type="text" name="FORM_DS" VALUE="?"> 
  </td>
 </tr>
 <tr> 
  <td width="53%" HEIGHT="33"></td>
 </tr> 
 <tr> 
  <td width="47%"> 
  </td>
 </tr> 
 <tr> 
  <td colspan="2"> 
    <div align="center"><font color="#CC0033">
    <input type="submit" name="Submit2" value="<% =GF_TRADUCIR("Agregar") %>"> 
    <INPUT TYPE="button" NAME="BotonCancelar" VALUE="<% =GF_TRADUCIR("Cancelar") %>" onclick="document.location.href='javascript:history.back(2)'">	
  </td>
 </tr> 
</form>
</table>
</body>
</html>
