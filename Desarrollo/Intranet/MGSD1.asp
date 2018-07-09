<!--#include file="../ActiSAIntra/Includes/procedimientosMG.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"-->
<!--#include file="../ActiSAIntra/Includes/procedimientosfechas.asp"-->
<% ProcedimientoControl "MGSD1" %> 
<% 
 
DIM con, sql, rs
dim p_smkm, p_smkc, p_smkr, p_smds, p_sdkc, p_sdkr, p_sdds, p_sdvl, p_sdvl_act
DIM P_Mensaje, P_CONF,G_KM,G_KC,G_DS,G_KR
dim dia, mes, anio, My_mmmv
FUNCTION MY_KC(P_KR)
IF ISNULL(P_KR) OR P_KR = 0 THEN
   G_KC = "?"
ELSE   
   GF_MGC G_KM,G_KC,P_KR,G_DS
END IF   
MY_KC = G_KC 
END FUNCTION
SUB P_M(T)
IF P_MENSAJE = "" THEN P_MENSAJE = T
END SUB

'Recuperar valores de maestro
p_smkm = GF_PARAMETROS("p_smkm","my_smkm")
p_smkc = GF_PARAMETROS("p_smkc","my_smkc")
p_smkr = GF_PARAMETROS("p_smkr","P_SMKR")
GF_MGC p_smkm,p_smkc,p_smkr,p_smds

'Recuperar valores de datos
p_sdkc = GF_PARAMETROS("p_sdkc","my_sdkc") 
p_sdkr = GF_PARAMETROS("p_sdkr","P_SDKR")
GF_MGC "SD",p_sdkc,p_sdkr,p_sdds
'Recuperar el nuevo valor 
p_sdvl = GF_PARAMETROS("p_sdvl","my_sdvl")

'Recuperar valor de confirmacion
P_CONF = GF_PARAMETROS("P_CONF","P_CONF")
'Setear valores
P_mensaje = ""
p_smkr = 0
'p_smkm = "SM"

'Recuperar valores de fecha
dia = request.form("txted") 
mes = request.form("txtem") 
anio = request.form("txtea") 
My_MmMv = "CONVERT(datetime,'" & mes & "/" & dia & "/" & anio  & " " & time() & "',120)"

'Comprobar que exista el maestro
IF NOT GF_MGKS("SM",p_smkm,p_smkr,P_smds) THEN P_M "ERROR DE MAESTRO"
'Comprobar que exista el registro
IF NOT GF_MGKS(p_smkm,p_smkc,p_smkr,p_smds) THEN P_M "NO EXISTE REGISTRO"
'Comprobar que exista el dato
IF NOT GF_MGKS("SD",p_sdkc,p_sdkr, p_sdds) THEN P_M "DEFINIR DATO"

IF p_sdvl_act = "" AND p_MENSAJE = "" THEN p_sdvl_act = GF_DT1("READ",p_sdkc,"","",p_smkm,p_smkc)
'Si tengo la confirmacion y no hay errores grabo
IF p_MENSAJE = "" AND P_CONF = "S" THEN 
    'Generar el insert
	if p_sdvl = "" or dia = "" or mes = "" or anio = "" then
	   p_Mensaje = "Problemas al insertar. debe completar los campos."
    else
	if (GF_CONTROL_FECHA(dia,mes,anio)) then
	   if isnumeric(p_sdvl) then p_sdvl = replace(p_sdvl,".",",")
	      if GF_DT1W(p_sdkc,p_smkm,p_smkc,p_sdvl,My_MmMv) = true then
	   	     p_sdvl_act = p_sdvl
	      else
	         p_Mensaje = "Problemas al insertar. compruebe que la fecha no exista para ese dato."
	      end if
	else
	   p_Mensaje="La fecha es incorrecta."
	   end if
	end if   
end if
%>
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"> 
<LINK REL="stylesheet" HREF="../css/global.css" TYPE="text/css">
<script language="vbscript" src="../ActiSAIntra/Scripts/script_fechas.vbs"></script>
</head>

<body bgcolor="#ffffff" text="#000000">
<% =GF_TITULO("lupa.gif","Acceso a Datos") %>
<br>
<form name="frmsel" method="post" action="MGSD1.asp"> 
<table align="center" width="100%" border="0" cellspacing=0> 
<TR>
   <TD align="left"> 
      <INPUT align="absMiddle" class="NOBORDER" TYPE="image" src ="images/Aceptar.gif" name="Submit2">&nbsp;<% =GF_TRADUCIR("Aceptar")%> |
      <a href="MG210.asp?P_KR=<% =p_smkr %>"><img align="absMiddle" src ="images/Anterior.gif" border="0"></a>&nbsp;<%=gf_traducir("Volver")%> |
      <INPUT align="absMiddle" class="NOBORDER" TYPE="image" src ="images/Agregar.gif" name="Submit22" onClick=document.frmsel.action="MGSD1.asp?P_CONF=S">&nbsp;<%=gf_traducir("Agregar")%> |
   </td>
</TR> 
</table>
<table bordercolor="#cccccc" border="1" ALIGN="center" CELLSPACING="0" CELLPADDING="1" width="100%">
  <tr> 
   <td class="TDNOHAY"><%=GF_TRADUCIR("Maestro")%></td>
   <td ><input name="my_smkm" VALUE="<%=p_smkm%>" size=3></td>
   <td class="TDNOHAY"><%=GF_TRADUCIR("Codigo") %></td>
   <td ><input name="my_smkc" VALUE="<%=p_smkc%>" size=10></td>
   <td class="TDNOHAY"><%=GF_TRADUCIR("Descripcion") %></td>
   <td ><%=p_smds%></td>
   <td class="TDNOHAY">Kr</td>
   <td ><%=p_smkr%></td>
</tr>
<tr> 
   <td class="TDNOHAY"><%=GF_TRADUCIR("Dato") %></td>
   <td ><input name="my_sdkc" size=10 VALUE="<%=p_sdkc%>"></td>
   <td class="TDNOHAY"><%=GF_TRADUCIR("Descripcion") %></td>
   <td> <%=p_sdds%></td>
   <td class="TDNOHAY"><%=GF_TRADUCIR("Valor actual") %></td>
   <td> <%=p_sdvl_act%> </td>
   <td class="TDNOHAY">Kr</td>
   <td> <%=p_sdkr%></td>
</tr>
<tr> 
   <td class="TDNOHAY" ><%=GF_TRADUCIR("Nuevo Valor") %></td>
   <td colspan=3> <input name="my_sdvl" size=40 value=<%=p_sdvl%> ></td>
   <td class="TDNOHAY"><%=GF_TRADUCIR("Fecha") %></td>
   <td colspan=3>
     <INPUT TYPE="text" NAME="txted" VALUE="<%=dia%>" size="2" maxlength="2" > /
     <INPUT TYPE="text" NAME="txtem" VALUE="<%=mes%>" size="2" maxlength="2" > / 
     <INPUT TYPE="text" NAME="txtea" VALUE="<%=anio%>" size="4" maxlength="4">
   </td>
</tr> 
<tr>
<% if p_mensaje <> "" then %> 
  <td colspan=6 align=center><font color="#cc0033"><B><%=GF_TRADUCIR(p_mensaje)%></B></font></td>
<% end if %> 

</tr>
<tr>
</tr> 
</TABLE>
</form>
<%
if P_MENSAJE = "" THEN 
	SQL = "SELECT DT_MMMV,DT_MMSY,DT_VALOR,DT_USER FROM MGDT WHERE DT_KO=" & p_smkr & " AND DT_OBJETOS = 1 AND DT_KR = " & p_sdkr & " ORDER BY DT_MMMV desc, DT_MMSY desc"
	gf_bd_control rs,con,"OPEN",SQL
if not rs.eof then%>
<table bordercolor="#cccccc" border="1" ALIGN="center" CELLSPACING="0" CELLPADDING="1" width="100%">

   <TR class="TDNOHAY">
       <TD width="30%" align=center colspan=1><% =GF_TRADUCIR("Momento del dato") %></FONT></TD>
       <TD width="30%" align=center colspan=1><% =GF_TRADUCIR("Momento del sistema") %></FONT></TD>
       <TD width="10%" align=center colspan=1><% =GF_TRADUCIR("Usuario") %></FONT></TD>
       <TD width="30%" align=center colspan=1><% =GF_TRADUCIR("Valor") %></FONT></TD>
   </TR>
<%end if%>
<%while not rs.eof%>

<TR>
	<TD width="30%" align="RIGHT" colspan=1><%=rs("DT_MMMV")%></TD>
	<TD width="30%" align="RIGHT" colspan=1><%=rs("DT_MMSY")%></TD>
	<TD width="10%" align="RIGHT" colspan=1><%=MY_KC(rs("DT_USER"))%></TD>
	<TD width="30%" align="LEFT" colspan=1><%=rs("DT_VALOR")%></TD>
</TR>
<%rs.movenext
  wend
  gf_bd_control rs,con,"CLOSE",SQL
%>
</TABLE>
<% 
end if
%>
</body>
</html>
