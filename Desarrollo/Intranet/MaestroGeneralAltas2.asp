<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<% ProcedimientoControl "MG245"
dim mensaje
dim p_km, p_kc, p_ds, p_kr
dim p_dsmtro
dim p_krmtro
dim my_error
p_km = GF_ParametrosForm("FORM_KM")
p_kc = UCASE(GF_ParametrosForm("FORM_KC"))
p_Kc = GF_ControlarInputKc(p_kc)
my_error = false
GF_MGKS "SM", p_km, p_dsmtro, p_krmtro ' Obtener datos del maestro
	my_error=false'GF_MGKS(p_km,p_kc,p_DS,p_KR) 
	if not (my_error) then
   	     p_ds = GF_ParametrosForm("FORM_DS") 'Obtener la nueva descripción
		 GF_MGADD  p_km, p_kc, p_ds, p_kr          
		 Mensaje = "Se agrego correctamente, [<a href=" & chr(39) & "MGBKS.asp?P_KC=" & P_KM & chr(39) & ">Volver</a>] <br>"
         Mensaje = Mensaje & "Agregar Atributos, [<a href=" & chr(39) & "MG210.asp?p_kr=" & p_kr & chr(39) & ">Click aqui.</a>] <br>"		 		 
    else
		 Mensaje = "El codigo ya existe, [<a href=" & chr(39) & "javascript:window.history.back()" & chr(39) & ">Volver</a>]"
    end if   	   
%>
<html>
<head>
<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<table width="100%" height="100%">
  <tr>
    <td> 
      <table class="reg_Header" align="center" border="0" cellspacing="1" cellpadding="2" bordercolor="#00AACA" bgcolor="#FFFFFF">
        <tr>
          <td><%=mensaje%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>

