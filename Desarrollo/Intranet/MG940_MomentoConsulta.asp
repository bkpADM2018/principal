<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->
<% ProcedimientoControl "MG940" %> 
<%
Dim mensaje,CONDED,CONDEM,CONDEA
Dim CONDVD,CONDVM,CONDVA
Dim CONDVHOR,CONDVMIN,CONDVSEG
Dim CONDEHOR,CONDEMIN,CONDESEG
Dim ACCION,P_set,P_chk,CONDEAMPM,CONDVAMPM
Dim MmtoConsultaActual,MmtoSistemaActual
'--------------------------------------------------------------------------------------------
function ControlarFecha(ByRef P_CONDED,ByRef P_CONDEM,ByRef P_CONDEA,ByRef P_CONDEHOR,ByRef P_CONDEMIN,ByRef P_CONDESEG)
 dim msg
  msg=""
  if not(GF_CONTROL_FECHA(P_CONDED,P_CONDEM,P_CONDEA)) then 
	msg="Fecha no valida"
  else	  	
  	if (P_CONDEHOR > 23) or (P_CONDEHOR < 0) then msg="Hora no valida"
  	if (P_CONDEMIN > 59) or (P_CONDEMIN < 0) then msg="Minutos no valida"
  	if (P_CONDESEG > 59) or (P_CONDESEG < 0) then msg="Segundos no valida"
  	call GF_STANDARIZAR_MM(P_CONDEHOR,P_CONDEMIN,P_CONDESEG)  	
  end if	  
  ControlarFecha=msg
end function
'--------------------------------------------------------------------------------------------
%>
<SCRIPT LANGUAGE=VBscript SRC="../latyd/script_fechas.vbs"></script>
<%
'Recupero los parametros
ACCION=GF_PARAMETROS("","P_ACCION")
CONDED=GF_PARAMETROS("","txted")
CONDEM=GF_PARAMETROS("","txtem")
CONDEA=GF_PARAMETROS("","txtea")
CONDEHOR=GF_PARAMETROS("","txtehor")
CONDEMIN=GF_PARAMETROS("","txtemin")
CONDESEG=GF_PARAMETROS("","txteseg")
CONDEAMPM=GF_PARAMETROS("","txteampm")
if (CONDEAMPM <> "AM") or ((CONDEAMPM <> "PM")) then CONDEAMPM=""
CONDVD=GF_PARAMETROS("","txtvd")
CONDVM=GF_PARAMETROS("","txtvm")
CONDVA=GF_PARAMETROS("","txtva")
CONDVHOR=GF_PARAMETROS("","txtvhor")
CONDVMIN=GF_PARAMETROS("","txtvmin")
CONDVSEG=GF_PARAMETROS("","txtvseg")
CONDVAMPM=GF_PARAMETROS("","txtvampm")
if (CONDVAMPM <> "AM") or ((CONDVAMPM <> "PM")) then CONDVAMPM=""
P_chk=GF_PARAMETROS7("__CLK__","",3)
'Tomo los momento actuales
MmtoConsultaActual=GF_VerFechaDato()
MmtoSistemaActual=GF_VerFechaSistema()
if (CONDED = "") then CONDED=day(MmtoConsultaActual)
if (CONDEM = "") then CONDEM=month(MmtoConsultaActual)
if (CONDEA = "") then CONDEA=year(MmtoConsultaActual)
if (CONDEHOR = "") then CONDEHOR=hour(MmtoConsultaActual)
if (CONDEMIN = "") then CONDEMIN=Minute(MmtoConsultaActual)
if (CONDESEG = "") then CONDESEG=Second(MmtoConsultaActual)
if (CONDEAMPM = "") then CONDEAMPM=right(MmtoConsultaActual,2)
if (CONDVD = "") then CONDVD=day(MmtoSistemaActual)
if (CONDVM = "") then CONDVM=month(MmtoSistemaActual)
if (CONDVA = "") then CONDVA=year(MmtoSistemaActual)
if (CONDVHOR = "") then CONDVHOR=Hour(MmtoSistemaActual)
if (CONDVMIN = "") then CONDVMIN=minute(MmtoSistemaActual)
if (CONDVSEG = "") then CONDVSEG=second(MmtoSistemaActual)
if (CONDVAMPM = "") then CONDVAMPM=right(MmtoSistemaActual,2)
'Correccion de hora
if (CONDEAMPM <> "") and (CONDEHOR >= 13) then CONDEHOR= CONDEHOR-12
if (CONDVAMPM <> "") and (CONDVHOR >= 13) then CONDVHOR= CONDVHOR-12
'Si se pidio una operacion se realiza
if (ACCION = "CONTROLAR") then
	mensaje=ControlarFecha(CONDED,CONDEM,CONDEA,CONDEHOR,CONDEMIN,CONDESEG)
	mensaje=ControlarFecha(CONDVD,CONDVM,CONDVA,CONDVHOR,CONDVMIN,CONDVSEG)
end if   
if (ACCION = "GUARDAR") then
   mensaje=ControlarFecha(CONDVD,CONDVM,CONDVA,CONDVHOR,CONDVMIN,CONDVSEG)
   if (mensaje = "") then
      call GF_setMomentoDato(CONDED,CONDEM,CONDEA,CONDEHOR,CONDEMIN,CONDESEG,CONDEAMPM)      
      call GF_setMomentoSistema(CONDVD,CONDVM,CONDVA,CONDVHOR,CONDVMIN,CONDVSEG,CONDVAMPM)                  
      'Guardo datos en la session
      sessino("MG940/Prmtr/__CLK__") = P_chk      
   end if 
end if
%>
<html>
<head>
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<br>
<%=GF_TITULO("Clock.gif","Establecer Momento de Consulta")%>
<br>
<FORM NAME="frmsel" METHOD="POST" ACTION="MG940_MomentoConsulta.ASP">
<table border="1" CELLPACING="2" align="center" bgcolor="#FFFFFF"> 
 <% if (mensaje <> "") then %>
 <tr class="TDERROR">
  <td align="center" colspan="2"><%=mensaje%></td>
 </tr>
 <% end if %>
 <tr>
	<TD align="center"><b><font size=2><%=gf_traducir("Consulta de datos al: ")%></font></b></td> 
    <TD align="center"> 
     <INPUT TYPE="text" NAME="txted" VALUE="<%=CONDED%>" size="2" maxlength="2" > /
	 <INPUT TYPE="text" NAME="txtem" VALUE="<%=CONDEM%>" size="2" maxlength="2" > /
     <INPUT TYPE="text" NAME="txtea" VALUE="<%=CONDEA%>" size="4" maxlength="4" > 
     <INPUT TYPE="text" NAME="txtehor" VALUE="<%=CONDEHOR%>" size="2" maxlength="2" > :
	 <INPUT TYPE="text" NAME="txtemin" VALUE="<%=CONDEMIN%>" size="2" maxlength="2" > :
     <INPUT TYPE="text" NAME="txteseg" VALUE="<%=CONDESEG%>" size="2" maxlength="2" > 
     <INPUT TYPE="text" NAME="txteampm" VALUE="<%=CONDEAMPM%>" size="2" maxlength="2" > 
	</td>
 </tr>
 <tr>	
	 <TD align="center"><b><font size=2><%=GF_TRADUCIR("Ingreso a sistema: ")%></font></b></td>
     <TD align="center">
       <INPUT TYPE="text" NAME="txtvd" VALUE="<%=CONDVD%>" size="2" maxlength="2" > /
	   <INPUT TYPE="text" NAME="txtvm" VALUE="<%=CONDVM%>" size="2" maxlength="2" > /
       <INPUT TYPE="text" NAME="txtva" VALUE="<%=CONDVA%>" size="4" maxlength="4" > 
       <INPUT TYPE="text" NAME="txtvhor" VALUE="<%=CONDVHOR%>" size="2" maxlength="2" > :
	   <INPUT TYPE="text" NAME="txtvmin" VALUE="<%=CONDVMIN%>" size="2" maxlength="2" > :
       <INPUT TYPE="text" NAME="txtvseg" VALUE="<%=CONDVSEG%>" size="2" maxlength="2" > 
       <INPUT TYPE="text" NAME="txtvampm" VALUE="<%=CONDVAMPM%>" size="2" maxlength="2" > 
      </td>
 </tr> <tr>
  <td colspan=2 align="center">
    <INPUT TYPE="Submit" NAME="P_CONTROLAR"  VALUE="<% =GF_TRADUCIR("CONTROLAR") %>" TABINDEX=1 onClick="LF_OPERACION('CONTROLAR')">
  </td>
</tr>
<tr>
  <td colspan=2 align = "center"> 
    <INPUT TYPE="Submit" NAME="P_GUARDAR" VALUE="<% =GF_TRADUCIR("GUARDAR") %>" TABINDEX=2 onClick="LF_OPERACION('GUARDAR')">
  </td>
</tr>
<tr>
  <td colspan=2 align = "center"> 
    <% P_set="OFF"
       if (Ucase(P_chk)="ON") then 
          P_chk="CHECKED" 
          P_set="ON"
       end if   
    %>  
    <INPUT TYPE="checkbox" class="NOBORDER" NAME="chkCLK" TABINDEX=3 onClick="LF_SET()" <% =P_chk %>>&nbsp;<% =GF_TRADUCIR("Automatizar") %>
  </td>
</tr>
</TABLE>
<input type="hidden" name="P_ACCION" value="CONTROLAR">
<input type="hidden" name="__CLK__" value="<% =P_set %>">
</form>
<script language="javascript">
   function LF_OPERACION(P_OPR)
   {
      frmsel.P_ACCION.value=P_OPR;
   }
   function LF_SET()
   {
      if (frmsel.__CLK__.value=="OFF")
         frmsel.__CLK__.value="ON";
      else   
         frmsel.__CLK__.value="OFF";
   }
</script>
</body>
</html>
