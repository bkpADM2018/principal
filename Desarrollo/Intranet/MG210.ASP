<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<% ProcedimientoControl "MG210" %> 
<%
' Mostrar los datos de un registro.
dim con, rs, Sql, rs1, cn1, sql1, sqlins, sqlupd
dim contador,MY_ERROR, MY_VERSION, P_MMD, P_MMS,MY_DATOFECHA, SQL_DATO, CON_DATO, RS_DATO
dim  i,V
dim perm,strDatos,strRelaciones
dim prim
dim myNombreCompleto,My_RegCodAcceso, My_AccesoDato, My_Valor, My_Name,chg_valor,My_KrAnterior,My_txt
dim MyUsuario
dim P_KM,P_KC,P_DS,P_KR,P_SMDS,P_SMKR , p_kmacceso, p_kcacceso
dim MY_dato,mg_km,mg_kc,mg_ds,Mg_kr	
dim MG_con
dim MG_rs
dim MG_Sql 
dim MomentoSys,  my_accesoRegistro, P_SMACCESO
dim p_ACCION
dim P_Relaciones
dim P_DATOS
dim My_pasadas
MomentoSys= session("MomentoSistema")
MY_ERROR = ""
P_DS = ""
Sub TomarParametros
my_kranterior = request.form("P_kr")
P_KR = request.querystring("P_KR")
if not isnumeric(p_kr) then p_kr = 0
IF P_KR > 0 THEN 
   GF_MGC P_KM,P_KC,P_KR,P_DS
   ELSE
   P_KR = 0
   P_KM = GF_PARAMETROSFORM("P_KM")
   P_KC = GF_PARAMETROSFORM("P_KC")
END IF
P_ACCION = ucase(GF_PARAMETROS("P_ACCION",""))
' Definir si se muestran o no los datos
P_DATOS = GF_PARAMETROS7("P_DATOS","",7)
IF P_Accion = "P_DATOS" or p_Datos = "" THEN DefinirSiNo p_Datos
' Definir si se muestran o no las relaciones
P_Relaciones = gf_parametros7("P_Relaciones","",7)
IF P_Accion = "RELACIONES" or p_Relaciones = ""  THEN DefinirSiNo p_Relaciones

end sub

SUB DefinirSiNo(byref p_parametro)
   IF P_Parametro = "NO" then
      P_Parametro = "SI"
	  else
	  P_Parametro = "NO"
   end if
end sub  

TomarParametros

' Verificar el control de acceso al registro.
my_accesoRegistro = GF_MG_Acceso(p_km,p_kc,p_kr,p_ds,P_SMKR,P_SMDS,My_RegCodAcceso,MY_ERROR)   
my_txt  = "MGBKS.ASP?P_KC=" & P_KM


%>

<html>
<head>
<title>Codigo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<% =GF_TITULO("Lupa.gif","Detalle de un Registro")%>
<form name="form1" method="post" action="MG210.asp" > 
<input type="hidden" name="P_DATOS" value="<%=P_DATOS%>" >
<input type="hidden" name="P_Relaciones" value="<%=P_Relaciones%>" >

<INPUT TYPE="hidden" NAME="P_KR" value="<%=P_KR%>">
<table align="center" width="95%" border="0" cellspacing=0> 
<TR>
 <TD>
      <INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Aceptar.gif" NAME="Submit" alt="<%=gf_traducir("Aceptar")%>" tabindex="0">&nbsp;<% =GF_TRADUCIR("Aceptar") %> |
 	  <INPUT align=absMiddle class="NOBORDER" TYPE="image" src="images/Anterior.gif" NAME="B_ANTERIOR"  alt="<% =GF_TRADUCIR("Anterior") %>" onclick="B_Anterior_onclick()">&nbsp;<% =GF_TRADUCIR("Anterior") %> |
 	  <INPUT align=absMiddle class="NOBORDER" TYPE="image"  src="images/Siguiente.gif" NAME="B_Siguiente" alt="<% =GF_TRADUCIR("Siguiente") %>" onclick="B_Siguiente_onclick()">&nbsp;<% =GF_TRADUCIR("Siguiente") %> |  	      
  <% if MY_ERROR = "" AND MY_ACCESOREGISTRO > "2" then %>
       <INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Agregar.gif" NAME="Agregar" alt="<%=gf_traducir("Agregar")%>" onclick=document.form1.action="MAESTROGENERALALTAS.ASP?P_KM=<%=P_KM %>" >&nbsp;<% =GF_TRADUCIR("Agregar") %> |<br><br>
  <% end if %>
  <% strDatos= "No" 
     if (p_Datos = "SI") then strDatos=""
     strRelaciones= "No" 
     if (p_relaciones = "SI") then strRelaciones=""     
   %>   
       <INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/datos.gif" NAME="Datos" style="font-size:8pt;" onclick="Datos_onclick()">&nbsp;<%=GF_TRADUCIR(strDatos & " Ver Datos") %> |
	   <INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Relacionados.gif" NAME="Relaciones" style="font-size:8pt;" onclick="Relaciones_onclick()">&nbsp;<%=GF_TRADUCIR(strRelaciones & " Ver Relaciones") %> |	   
 </td>
</TR> 
</table>
<table class="TABLE2" align="center" width="100%" cellspacing=0 rules="rows" border=0> 
  <tr class=reg_header_nav>
    <td align="LEFT"> 
      <%=GF_TRADUCIR("Maestro: ")%>&nbsp;
	  <INPUT TYPE="text" NAME="P_KM" VALUE="<%=P_KM%>" size="2"> 
	</td>
    <td align="left" colspan=2> 
      <DIV><%=p_smds%></DIV>
	</td>
	<td align="center">&nbsp;</td>
	<td align="center">&nbsp</td>
  </tr>
  <tr class=reg_header_nav>	
	<td align="LEFT"> 
      <%=GF_TRADUCIR("Codigo: ")%>&nbsp;
      <INPUT TYPE="text" NAME="P_KC" VALUE="<%=P_KC%>" size="8"> 
	</td>
	<td colspan=2 align="left">
      <DIV><%=p_ds%></DIV>
	</td>
	<td align="left" colspan=2> 
      <DIV><%=p_KR%></DIV>
	</td>
  </tr> 
</table>
<table align="center" width="100%" rules="rows" cellspacing=0 border=0> 
<% if MY_ERROR <> "" then %>
 <tr> 
    <td bgcolor="white" align=center>
  	  <font color="red"><b><%=MY_ERROR%></b></font> 
	</td>
 </tr>
<% end if %>
<% 
IF MY_ERROR = "" THEN 
%>
</table>

  <%IF P_DATOS = "NO" THEN %>
  <br>
  <table align="center" width="100%" class=reg_header rules="cols" cellspacing=1 cellpadding=1 border=0> 
<TR>
    <td align="center"  ><% =GF_TRADUCIR("DATO") %></td>
    <td align="center" ><% =GF_TRADUCIR("DESCRIPCION") %></td>
    <td align="center"  class="TDENCABEZADO"><% =GF_TRADUCIR("VALOR") %></td>
    <td align="center" ><% =GF_TRADUCIR("LOG") %></td>
 </TR> 
<%
MY_PASADAS = 0
WHILE MY_PASADAS < 2
  MY_PASADAS = MY_PASADAS + 1
  IF MY_PASADAS = 1 THEN
  sql = "SELECT * FROM RELACIONESCONSULTA WHERE SRO1KM  = 'SR' AND SRO1KC = 'SMSD' "
  SQL = SQL & " AND SRO2KM = 'SM' AND SRO2KC = '" & P_KM & "' and srvalor <> '*' "
  sql = sql & " AND SRO3KM = 'SD' ORDER BY SRO3KC " 
  END IF
  IF MY_PASADAS = 2 THEN
  sql = "SELECT * FROM RELACIONESCONSULTA WHERE SRO1KM  = 'SR' AND SRO1KC = 'XXSD' "
  SQL = SQL & " AND SRO2KR = " & P_KR & " and srvalor <> '*' "
  sql = sql & " AND SRO3KM = 'SD' ORDER BY SRO3KC " 
  END IF
  'RESPONSE.WRITE SQL
  GF_BD_CONTROL rs,con,"OPEN",SQL
  MG_KR  = -1 
  while not rs.eof 'leyendo los datos del registro
  
    IF MG_KR <> RS("SRO3KR") THEN
	
	   MG_KR =  RS("SRO3KR")
	   MG_KC =  RS("SRO3KC")
	   MG_DS =  RS("SRO3DS")
	   my_accesoDato = GF_controlAccesoKS("ACCESO","SD",mg_kc,mg_kr,mg_ds)
	   'RESPONSE.WRITE "Acceso a DATO(" & MG_KC & " Valor:" & my_accesoDato & ")<br>"
		If My_RegCodAcceso < My_accesoDato then My_AccesoDato = My_RegCodAcceso
		if my_accesoDato > "0" then
			' Leer el dato vigente y verificar si tiene mas de una version para indicarlo
			P_MMD = session("MomentoDato")
			P_MMS = session("MomentoSistema")
			sql_DATO = "SELECT TOP 2 DT_MMMV,DT_VALOR FROM MGDT WHERE DT_KO = " & P_KR & " AND DT_Objetos = 1 AND DT_KR = " & mg_KR & " and dt_mmmv <= " & P_MMD & " and dt_MmSy <= " & P_MMS & " ORDER BY dt_mmmv DESC, dt_mmSY DESC"
			'Response.write "gf_dt1-Busca>(" & SQL_DATO & ")"
			gf_bd_control rs_DATO,con_DATO,"OPEN",SQL_DATO
			MY_VERSION = 1
			mY_DATOFECHA = "__/__/____"
			MY_VALOR = "?"
			while not rs_DATO.eof  
			   if my_version = 1 then
				    my_valor = rs_DATO("DT_Valor") 
					my_version = 2
					mY_DATOFECHA = GF_FN2DTE(rs_DATO("dt_mmMV"))
				  ELSE 
				    'mY_DATOFECHA = mY_DATOFECHA & "<FONT COLOR=RED><b>&nbsp;<marquee width='5%' direction=left>&nbsp;<STRONG><</STRONG>&nbsp;</marquee></b></FONT>"	
				    mY_DATOFECHA = mY_DATOFECHA & "<FONT COLOR=RED><b>&nbsp;<img src='images/add.gif'></b></FONT>"	
			    END IF
				  RS_DATO.MOVENEXT
			wend	  
			GF_BD_CONTROL rs_DATO,con_DATO,"CLOSE",SQL_DATO
			   ' MY_VALOR = MY_DT1(MG_KC,"","",P_KM,P_KC)
			    my_name = "Valor_" & mg_kc %>
				<tr bgcolor="#FFFFFF">
					<td>
						<A HREF="mg210.ASP?P_KR=<%=mg_kR%>"><%=MG_kc%></A>
					</td>
					<td><%=MG_ds%></td>
					<%
					if my_accesoDato = "1" then %>
						<td><%=my_valor%></td>
					<% 
   					else   
						'Grabar el valor modificado si es el mismo registro
						IF (P_KR+0) = (My_KrAnterior+0) THEN 
							chg_valor = request.form(my_name)
							if chg_valor <> "" and chg_valor <> my_valor then
								my_valor = chg_valor
								gf_dt1w mg_kc,P_KM,P_KC,my_valor,""
							end if 
						END IF 
					%>  
						<td> 	
				        <% dim Pass_Kr
						GF_MGC "SD", "UPPSWR",Pass_Kr,""
						if mg_kr = Pass_Kr and session("Usuario")<>"ADMIN" then %>
							<font size="4"><INPUT TYPE="password" size=40 NAME="<%=my_name%>" VALUE="<%=My_valor%>"> </font></td>
						<% else %>
							<font size="4"><INPUT TYPE="text" size=50 NAME="<%=my_name%>" VALUE="<%=My_valor%>"> </font></td>
						<% end if
					end if 
					if mg_kr <> Pass_Kr OR session("Usuario")="ADMIN" then 
						V = "MGSD1.ASP?P_SMKR=" & P_KR & "&P_SDKR=" & MG_KR
					%>
						<td align="center"> <A HREF="<%=V%>"> <%=mY_DATOFECHA%></A> </td>
					<%
					end if 
					%>
				</tr>
		<% 
		end if 
		%>
    <% end if %>
    <% rs.movenext
  wend
 GF_BD_Control  rs, con, "CLOSE", sql 
WEND 
END IF 
%>

</table>
<%END IF%>
<%if p_relaciones = "NO" then%>
<br>
<table align="center" width="100%" class=reg_header rules="cols" cellspacing=1 border=0>
<TR >
    <td  align="center"><% =GF_TRADUCIR("RELACION ESTABLECIDA") %></td>
    <td  align="center"><% =GF_TRADUCIR("RELACIONANTE") %></td>
    <td  align="center"><% =GF_TRADUCIR("RELACIONADO") %></td>
	<td  align="center"><% =GF_TRADUCIR("VALOR") %></td>
</TR> 
<% my_pasadas = 0 
WHILE MY_PASADAS < 3 
   my_pasadas = my_pasadas + 1
   ' Buscar el registro como relacion
                        sql = "SELECT * FROM RELACIONESCONSULTA WHERE"
   if my_pasadas = 1 then  sql = sql & " SRO2KR = " 
   if my_pasadas = 2 then  sql = sql & " SRO3KR = " 
   if my_pasadas = 3 then  sql = sql & " SRO1KR = "
                           sql = sql & P_KR & " and srvalor <> '*' "
   gf_bd_control rs,con,"OPEN",SQL
   'RESPONSE.WRITE SQL
   while not rs.eof 'leyendo los datos del registro
%>
     <TR><td bgcolor="#FFFFFF" align="left"> <A HREF=<%="mg210.asp?P_KR=" & rs("sro1kr")%>> <%=rs("sro1ds")%></A></td>
     <td bgcolor="#FFFFFF" align="left"> <A HREF=<%="mg210.asp?P_KR=" & rs("sro2kr")%>> <%=rs("sro2ds")%></A></td>
     <td bgcolor="#FFFFFF" align="left"> <A HREF=<%="mg210.asp?P_KR=" & rs("sro3kr")%>> <%=rs("sro3ds")%></A></td>
	 <td bgcolor="#FFFFFF" align="center"> <A HREF=<%="mg235.asp?P_KR1=" & rs("sro1kr") & "&P_KR2=" & RS("SRO2KR") & "&P_KR3=" & RS("SRO3KR")    %>> <%=rs("srVALOR")%></A></td>

     </TR>  
   <%
     rs.movenext
     wend
     GF_BD_Control  rs, con, "CLOSE", sql %>
<%
wend
%>
</table>  
<%
END IF
%>
</form>
</body>
</html>
<SCRIPT LANGUAGE = "javascript">
function B_Siguiente_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=Siguiente"
}
function B_Anterior_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=Anterior"
}
function Add_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=Add"
}
function NextO1_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=NextO1"
}
function Datos_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=P_Datos"
}
function Relaciones_onclick()
{
document.form1.action="MG210.ASP?P_ACCION=Relaciones"
}


</SCRIPT>
