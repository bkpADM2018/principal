<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/GF_ARMAR_LISTA_CARGOS.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->
<html>
	<head>
	<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<%
Dim P_strUsuario,intKRusuario,strDS, P_strPassword , strdbPassword
Dim strErrorMsg,blnValido,intORKR,intKR,intIdiomaOLD,strORKC
if session("Usuario") <> "" and session("Usuario") <> "GUEST" then 
	session.Contents.RemoveAll()
	%>
	<script>
		window.parent.location.href = "../ActisaIntra/";
	</script>
	<%
END IF	
session.LCID=11274
'Se leen  los parametros
P_strPassword = Request.Form("strPassword")
P_strUsuario = UCASE(Request.Form("strUsuario"))
strErrorMsg=""
blnValido=false
if (P_strPassword <> "") or (P_strUsuario <> "") then
   GP_CONFIGURARMOMENTOS	
   IF not GF_MGC ( "UP",(P_strUsuario),intKRusuario,strDS ) then 
      strErrorMsg=GF_TRADUCIR("El usuario no posee autorizacion para ingresar al sistema") & "."
   else	  
      strdbPassword = GF_DT1("READ","UPPSWR","","","UP",P_strUsuario)
      if (strdbPassword <> P_strPassword) then strErrorMsg=GF_TRADUCIR("El password es incorrecto") & "."
   end if	  
   if (strErrorMsg = "") then
	  intIDIOMAOLD= GF_GET_IDIOMA()      
      Call GF_SET_IDIOMA(intIdiomaOLD)
	  'Todos los datos son correctos, se cargan las variables de session necesarias
	  session("Usuario") = P_strUsuario
      	  session.timeout = 1000
	  'Genero la lista de cargos.
	  GF_ARMAR_LISTA_CARGOS GF_SESSIONKR("UP",session("Usuario"))
	  'Seteo el reloj en automatico
	  session("MG940/Prmtr/__CLK__") ="ON"
	  'Obtengo la organizacion en la que trabaja el usuario
	  GF_MGC "SR","TRABAJAEN",intKR,""
	  intKRusuario=""
	  GF_MGC "SG",session("Usuario"),intKRusuario,""
	  GF_MGSR_EXISTE intKR,intKRusuario,intORKR
	  GF_MGC "",strORKC,intORKR,""
	  session("KCOrganizacion") = strORKC
	  GP_CONFIGURARMOMENTOS	
	  blnValido=true
%>
<script language="javascript">
   window.parent.location.href = "../ActisaIntra/";
</script>

<%     
   end if
end if
%>
		<meta name="GENERATOR" Content="Microsoft Visual Studio.NET 7.0">
	</head>
	<body>
		<br><br><br><br><br><br>
		<table align=center width=70%><tr>
		<td align=center> 
		<img src="images/kogge256.gif"><br><br>
		<font color="#527B42" size=4><% =GF_TRADUCIR("Acceda a todos los") %> <br> <% =GF_TRADUCIR("servicios de nuestra Empresa") %></font>
		</td>
		<td width="10%"></td>
		<td>
		<table align=Right cellspacing=5 cellpadding=5 border=1 width="160">
		<tr><td>
		<form method="POST" name="frmLogin" action="INTcontrolLogin.asp"> 
           <div align="center"><font color="#000000" size=2><u><% =GF_TRADUCIR("Usuarios Registrados") %></u></font></div>
		   <br>
		   <div style="font:11px verdana;"><font color="#527B42"><span class="lblUsuario"><b><% =GF_TRADUCIR("USUARIO") & ":" %></b></span></font></div>
		   <div align="center"><input Type="Text" Name="strUsuario" size="15" maxlength="10"> </div>
		   <div style="font:11px verdana;"><font color="#527B42"><span class="lblPassword"><b><% =GF_TRADUCIR("CONTRASEÑA") & ":" %></b></span></font></div>
		   <div align="center"> <input Type="password" Name="strPassword" size="15" maxlength="10"> </div>
		   <br>
		   <div align="Right">
		   <input type="Submit" name="btnLogin" value="<% =GF_TRADUCIR("Entrar") %>">
		   </div>
	     </form>
		 </td></tr>
         </table>
		 </td></tr>
		 <tr>
    <td width="70%"></td>
	<td width="10%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td align="center" width="30%"><font color="RED">
      <% =strErrorMsg %>
      </font></td>
  </tr> 
    	 </table><br>
		 <table align= center border=0 width=160>
		    <tr align="center"><td>
			  <table border="1" cellspacing="0" bordercolor="#EFEFEF" bgcolor="#FFFFFF">
                <tr>
                  <td class="cargo" bgcolor="#EFEFEF"><% =GF_TRADUCIR("Nuevo usuario") %></td>
                </tr>
                <tr>
                  <td><% =GF_TRADUCIR("Si no esta registrado") %><br>
                      <% =GF_TRADUCIR("haga click") %><A href="INTAltaExternos.asp" target="RightFrame"><b> <% =GF_TRADUCIR("Aqui!") %></b></a>.</td>
                </tr>
              </table>
			</td></tr>
		    <tr><td><div align="center"><a href="INTPasswordRecovery.asp"><% =GF_TRADUCIR("Olvido su contraseña?") %></a></div></td></tr>
		 </table>
   </body>
</html>
