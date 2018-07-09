<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<% ProcedimientoControl "MG300" %>
<%
'Cambiar Codigos de registro..
dim rs, cn, sql, sqlupd, My_Sql
dim p_km1,p_kc1,p_kr1,p_ds1, p_km1ds
dim p_km2,p_kc2,p_kr2,p_ds2, p_km2ds
dim My_Reg2MSG, My_Reg1MSG, My_BtnCambiarEstado, p_Cambiar
My_Reg2MSG = ""
My_BtnCambiarEstado = ""
Obtener_Valores
Comprobar_Valores
if p_Cambiar = "OK" then
   Grabar_Valores
end if 
'--------------------------------------------------------------------------------------------
sub Grabar_Valores
  sql = "Select *from MG where mg_kr=" & p_kr1
  GF_BD_CONTROL rs, cn, "OPEN", sql
    sqlupd = "Update MG SET mg_km='" & p_km2 & "', mg_kc='" & p_kc2 & "' where mg_kr=" & p_kr1 
	GF_BD_CONTROL rs, cn, "EXEC", sql
	
  GF_BD_CONTROL rs, cn, "CLOSE", sql
  p_km1 = p_km2
  p_kc1 = p_kc2
  My_BtnCambiarEstado = "DISABLED"
end sub
'--------------------------------------------------------------------------------------------
sub Obtener_Valores
p_km1 = gf_parametros("P_KM1","p_km1")
p_km1 = UCASE (p_km1) 
p_kC1 = gf_parametros("P_Kc1","p_kc1")
p_kc1 = UCASE (p_kc1) 
p_kr1 = gf_parametros("P_Kr1","p_kr1")
p_km2 = gf_parametros("P_KM2","p_km2")
p_km2 = UCASE (p_km2) 
p_kC2 = gf_parametros("P_Kc2","p_kc2")
p_kc2 = UCASE (p_kc2) 
p_kr2 = gf_parametros("P_Kr2","p_kr2")
p_cambiar = gf_parametros("p_Cambiar", "p_Cambiar")
end sub
'--------------------------------------------------------------------------------------------
sub Comprobar_Valores
if not GF_MGKS(p_km1,p_kc1,p_kr1,p_ds1) then My_Reg1MSG = "Registro Actual, inexistente"
if GF_MGKS(p_km2,p_kc2,p_kr2,p_ds2) then My_Reg2MSG = "Ya existe el Registro Nuevo"
if p_kc2 = "" then My_Reg2MSG = "Registro Nuevo, inexistente"
if My_Reg2MSG = "" and My_Reg1MSG = "" then
   if p_km1 = p_km2 then
        My_Reg2MSG = "Preparado para realizar el cambio."
   elseif p_km1 = "NA" then
           My_Reg2MSG = "No se controlan movimientos."  
   end if
else
      My_BtnCambiarEstado = "DISABLED"
end if 
end sub
'--------------------------------------------------------------------------------------------
%>

<html>
<head>
<title>Cambiar Codigo de registros.</title>	
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
</head>
<body>
<% =GF_TITULO("TablaMG.gif","Cambiar Registros de Maestro") %>
<FORM NAME="formCambiar" METHOD="post" ACTION="MG300.ASP"> 
  <%=session("titulostandard")%>
  <table class="reg_Header" align="center" width=90% border=1 CELLSPACING="1" cellpadding=2 rules="rows">
     <tr>
         <td colspan=6 align=center>
              <strong>Modificar el Maestro y el C&oacute;digo de un registro</strong>
         </td>
     </tr> 
	 <tr>	 
         <td colspan=6 align=center class="reg_Header_nav">
              <FONT SIZE="+1"><strong>Registro Actual</strong></FONT>
         </td>
     </tr> 
     <tr class="reg_Header_navdos">
     	 <td align=center>Maestro</td>
		 <td align=center>
			  <INPUT TYPE="text" NAME="P_KM1" OnFocus=OnChange() VALUE="<%=P_km1%>" size=2> 
		 </td>
		 <td align=center>C&oacute;digo</td>
		 <td align=center>
			  <INPUT TYPE="text" NAME="P_Kc1" OnFocus=OnChange() VALUE="<%=P_kc1%>" size=10> 
		 </td>
		 <td align=center><a href="MG210.ASP?P_KR=<%=p_kr1%>"><%=P_Ds1%></a></td>
 		 <td align=center><strong><%=P_Kr1%></strong></td>
	 </tr>
	 <tr class="TDERROR">
		 <td colspan=6 align=center><%=My_Reg1MSG%></td>
	 </tr>
	 <tr>
		 <td colspan=6 align=center class="reg_Header_nav"> 
			 <strong>Registro Nuevo</strong>
		 </td>
	 </tr>
	 <tr class="reg_Header_navdos"> 
		 <td align=center>Maestro</td>
		 <td align=center>
			 <INPUT TYPE="text" NAME="P_KM2" OnFocus=OnChange() VALUE="<%=P_km2%>" size=2> 
		 </td>
		 <td align=center>C&oacute;digo</td>
		 <td align=center>
		    <INPUT TYPE="text" NAME="P_Kc2" OnFocus=OnChange() VALUE="<%=P_kc2%>" size=10> 
		 </td> 
		 <td align=center>La descripci&oacute;n no se modifica.</td>
 		 <td align=center><strong>&nbsp;</strong></td>
	 </tr>
	 <tr class="TDERROR">
		 <td colspan=6 align=center><%=My_Reg2MSG%></td>
	 </tr>
	 <tr>
		 <td align=center colspan=6>
 			<INPUT TYPE="submit" NAME="Cambiar" onclick="Cambiar_onclick()" VALUE="Cambiar" <%=My_BtnCambiarEstado%>>&nbsp;&nbsp;&nbsp;
 			<INPUT TYPE="submit" NAME="Submit" VALUE="Controlar">&nbsp;&nbsp;&nbsp;
			<INPUT TYPE="button" NAME="Submit3" onclick="document.location.href='javascript:history.back()'" VALUE="Cancelar">
		 </td>
	 </tr>
 </table>
</FORM> 
</body>
</html>
<SCRIPT LANGUAGE = "javascript">
<!--
function OnChange()
  {
   document.formCambiar.Cambiar.disabled = true
  }
function Cambiar_onclick()
  {
   document.formCambiar.action = "MG300.ASP?p_Cambiar=OK"
  }
-->
</SCRIPT>
