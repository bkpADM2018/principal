<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientossql.asp"--> 
<!--#include file="Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientospaginacion.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->
<% ProcedimientoControl "MGBKS" %> 
<%
' Mostrar registros del maestro general
dim my_fecha, intIndice
dim con,sql,rs
dim my_kcanterior
dim i, V,D,n,vword, vbuscar
dim myNombreCompleto
dim P_km,P_KC,P_DS,P_KR,P_ACCESO,MG_km,MG_KC,MG_DS,MG_KR,P_Operacion
dim My_LinkDescri, My_LinkCodigo,My_Texto
dim p_dssel, p_KCsel, MY_SELECT
dim kcpri, kcult
dim lpp,intMostrar,P_BUSQUEDA
dim my_order,my_wherekc, my_likeds, my_likekc 
dim btn, My_BtnCambiar, My_WhereMaestro
dim fecha, hasta, My_BtnEstado, My_BtnTodosSM_Estado
dim  p_smkr, p_smds, p_AccesoMaestro, p_Error, p_Acceso_DS, p_sdds_DS, p_sdkr_DS, p_sdkr, p_sdds
dim p_sdkr_KC, p_sdds_KC, p_Acceso_KC, My_TypeKC, My_TypeDS, p_MgKcOriginal(200), p_MgDsOriginal(200)
dim strNombres, strValores,strCadenaL, strCadenaR, strCadena, p_MgKc(70), p_MgDs(200), p_MgKr(200)
dim strLinkPagina,MAXLPP,P_ORDEN,strLinkFlechas
My_BtnEstado = ""
My_TypeKC = "Hidden"
My_TypeDS = "Hidden"
MAXLPP=50
strNombres = split(session("nombres"),"$$")
strValores = split(session("valores"),"$$")
strCadena=""
if isarray(strNombres) then
  for i = 0 to ubound(strNombres)-1
    strCadena = strCadena & strNombres(i) & "=" & strValores(i) & "&"
  next
  i = instr(1,strCadena,"?")
  if (len(strCadena)>0) then
    strCadenaL = mid(strCadena,1,i-1)
    strCadenaR = mid(strCadena,i+1,(len(strCadena)-1))
  end if 
end if 

My_BtnTodosSM_Estado = ""
P_KM = "SM"
P_KR = GF_ParametrosNumericos("P_KR")
P_KC = GF_ParametrosForm("P_KC")
P_OPERACION = GF_PARAMETROSFORM("P_OPERACION")
P_BUSQUEDA = GF_PARAMETROSFORM("P_BUSQUEDA")
P_ORDEN = GF_PARAMETROS("P_ORDEN","P_ORDEN")
GF_MGC p_km,p_kc,p_kr,p_ds 
My_BtnCambiar = "Si Cambios"
if p_Operacion = "MODIFICAR" then
   'Controlas acceso a modificar
   'El acceso al maestro se controlo antes
   'Controlar acceso a los datos
   P_ACCESO_DS = GF_ControlAcceso("ACCESO","SD","__MGDS",p_KR)
   if p_Acceso_DS > 1 then My_TypeDS = "Text"

   P_ACCESO_KC = GF_ControlAcceso("ACCESO","SD","__MGKC",p_KR)
   if p_Acceso_KC > 1 then My_TypeKC = "Text"

   if (p_Acceso_Kc > 1 or p_Acceso_Ds > 1) then 
        RecuperarValoresCambiados
        My_BtnCambiar = "No Cambios"
        My_BtnTodosSM_Estado = "Disabled"
   else
        My_BtnCambiar = "Sin Acceso"
		My_BtnEstado = "Disabled"
   end if
end if

P_ACCESO = GF_CONTROLACCESOKS("ACCESO",P_km,P_kc,P_kr,p_ds)
IF P_ACCESO  < "1" then p_kc = ""
IF P_KM <> "SM" OR P_KC = "" THEN  
   P_KM = "SM"
   P_KC = "SM"
   GF_MGKS P_KM,P_KC,P_KR,P_DS
END IF   
MY_FECHA = SESSION("MOMENTODATO")

' Condiciones de busqueda
p_KCsel = GF_ParametrosForm("P_KCSEL")
my_likekc = GF_LIKE ("MG_KC", p_kcsel)
p_dssel = GF_ParametrosForm("P_DSSEL")
my_likeds = GF_LIKE ("MG_DS", p_dssel)
MY_SELECT = my_likekc & my_likeds

if P_BUSQUEDA = "TODOS_SM" and My_Select <> "" then
   My_WhereMaestro = " where Mg_Km <> 'Basura'"
   My_BtnEstado = "Disabled"
END IF   

if p_kc <> "SM" then My_BtnTodosSM_Estado = "Disabled"
btn = GF_ParametrosForm("btn")
CargarLinksEnSession

if My_WhereMaestro = "" then My_WhereMaestro = "WHERE mg_km = '" & p_kc & "'" 

my_order = " ORDER BY mg_kc ,mg_mmmv DESC "
if (P_ORDEN = "KCUP") then my_order = " ORDER BY mg_kc ,mg_mmmv DESC "
if (P_ORDEN = "KCDOWN") then my_order = " ORDER BY mg_kc DESC,mg_mmmv DESC "
if (P_ORDEN = "DSUP") then my_order = " ORDER BY mg_ds ,mg_mmmv DESC "
if (P_ORDEN = "DSDOWN") then my_order = " ORDER BY mg_ds DESC,mg_mmmv DESC "
if (P_ORDEN = "MMUP") then my_order = " ORDER BY mg_mmmv"
if (P_ORDEN = "MMDOWN") then my_order = " ORDER BY mg_mmmv DESC "
if (P_ORDEN = "KRUP") then my_order = " ORDER BY mg_kr ,mg_mmmv DESC "
if (P_ORDEN = "KRDOWN") then my_order = " ORDER BY mg_kr DESC,mg_mmmv DESC "
'Response.Write my_order

sql = "SELECT MG_KM,MG_KC,MG_DS,MG_MMMV,MG_MMSY,MG_KR FROM mg " & My_WhereMaestro & MY_SELECT 
sql = sql & my_wherekc & my_order  
'RESPONSE.WRITE SQL
gf_bd_control rs, con, "OPEN", SQL 


'--------------------------------------------------------------------------------------------------------
sub RecuperarValoresCambiados
dim j
for j = 1 to MAXLPP
  'response.Write "KR(" & p_MgKr(j) & ")"
    p_MgKc(j) = GF_ParametrosForm ("p_MgKc" & j)
    p_MgDs(j) = GF_ParametrosForm("p_MgDs" & j)
    p_MgKcOriginal(j) = GF_ParametrosForm("p_MgKcOriginal" & j)
    p_MgDsOriginal(j) = GF_ParametrosForm("p_MgDsOriginal" & j)

    if p_MgKc(j) = "" then p_MgKc(j) = p_MgKcOriginal(j)
    if p_MgDs(j) = "" then p_MgDs(j) = p_MgDsOriginal(j)
    'Response.Write "<br>KC(" & p_MgKc(j) & ")DS(" & p_MgDs(j) & ")KCO(" & p_MgKcOriginal(j) & ")DSO(" & p_MgDsOriginal(j) & ")"
	if p_MgDs(j) <> p_MgDsOriginal(j) or p_MgKc(j) <> p_MgKcOriginal(j) then
	   'Guardar cambios
        p_MgKr(j) = GF_ParametrosNumericos("p_MgKr" & j)    
	    Actualizar_MG P_KC, p_Mgkr(j), p_MgKc(j), p_MgDs(j) 
	end if   
next 
end sub
'--------------------------------------------------------------------------------------------------------
sub Actualizar_MG (p_mgkm, p_mgkr, p_mgkc, p_mgds)
'Esta rutina actualiza el codigo y la descripcion de un determinado kr
dim SqlUpd, RsUpd, CnUpd
p_mgkm = GF_ControlarInputKc(p_mgkm)
p_mgkc = GF_ControlarInputKc(p_mgkc)

SqlUpd = "Select *from MG where MG_KM='" & p_mgkm & "' and MG_KC='" & p_mgkc & "' and MG_KR <> " & p_mgkr
GF_BD_CONTROL RsUpd, CnUpd, "OPEN", SqlUpd
if RsUpd.eof then
   SqlUpd = "Update MG SET MG_KC='" & p_mgkc & "' , MG_DS='" & p_mgds & "' where MG_KR=" & p_mgkr 
   Call executeQueryDb(DBSITE_SQL_INTRA, rsX, "EXEC", SqlUpd)   
end if
GF_BD_CONTROL RsUpd, CnUpd, "CLOSE", SqlUpd
end sub
'--------------------------------------------------------------------------------------------------------
sub NombreSession
    I = I + 1
    N = "MGBKS.LINK" & I
	v = session(n)
end sub
'--------------------------------------------------------------------------------------------------------
Function  PrepararUnRegistro	 
if rs("mg_kc") <> my_kcAnterior then
            PrepararUnRegistro = True
			i= i + 1
            my_kcAnterior = rs("mg_kc")
			IF P_OPERACION = "SELECT" THEN
				   if strCadenaL <> "" and strCadenaR <> "" then 
				     P_KR=0
				     My_LinkDescri = strValores(ubound(strValores)) & "?" & strCadenaL & rs("mg_kc") & strCadenaR
     		       else
				     My_LinkDescri = "MGSELECT.asp?P_KC=" & rs("mg_kc")
				   end if
			   ELSE
			   if rs("MG_KM") = "SM" THEN
			      My_LinkCodigo = "MG210.asp?P_KR=" & rs("mg_kR")
			      My_LinkDescri = "MGBKS.asp?P_KC=" & rs("mg_kC")
			      else
			      My_LinkCodigo = "MG210.ASP?P_KR=" & P_KR
			      My_LinkDescri = "MG210.ASP?P_KR=" & rs("mg_kr")
			   END IF 
			END IF 
else
                  PrepararUnRegistro = false
end if			
end Function			
'--------------------------------------------------------------------------------------------------------
SUB CargarLinksEnSession
n = "MGBKS.P_KC"
if P_KC <> session(N) then
   session(N) = p_KC
   sql = "SELECT * FROM RelacionesConsulta WHERE SRO1KM = 'SR' AND SRO1KC = 'SMSMLINK' AND SRO2KM = 'SM'"
   sql = sql & " AND SRO2KC = '" & P_KC & "'"
   sql = sql & " AND SRO3KM = 'SM' "
   gf_bd_control rs, con, "OPEN", SQL 
   I = 0
   while I < 4 
         NombreSession
		 v = "" 
		 if not RS.EOF then
		    MG_KR  = rs("SRO3KR")
			GF_MGC MG_KM,MG_KC,MG_KR,MG_DS
			V = MG_KC & "(" & MG_DS & ")"
			rs.movenext
	      end if
		  session(N) = v
   wend
   	    gf_bd_control rs ,con ,"CLOSE", sql
end if
end sub
%>

<script language=javascript>
   var VecPOP = new Array();
   
   function LF_VERDEFINICION(P_KR, P_KM)
   {
       var i = 0;
	   
	   for(i in VecPOP)
	   {
	      VecPOP[i]="hidden";
		  eval("POP"+i).style.visibility=VecPOP[i];       
	   }
	   if (P_KM == 1)
	   {
	      VecPOP[P_KR]="visible";
	      eval("POP"+P_KR).style.visibility=VecPOP[P_KR];      
	   }	  
   }
   function LF_OCULTARDEFINICION(P_KR)
   {
       var i = 0;
	   
	   for(i in VecPOP)
	   {
	      VecPOP[i]="hidden";
		  eval("POP"+i).style.visibility=VecPOP[i];       
	   } 
   }
</script>

<html>
<head>
<title>Browse de Maestros</title> <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"> 
<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" onClick="LF_OCULTARDEFINICION()">
<br>
<%=GF_TITULO("TablaMG.gif","Tablas del Maestro General")%>
<%
'Colocar los links vinculados
I = 0
  while I < 4 AND P_OPERACION <> "SELECT" 
     NombreSession
      if V <> "" then
         D = GF_TRADUCIR(RIGHT(V,LEN(V)-2)) 
         V = LEFT(V,2)
         %>
<a href="MGBKS.ASP?P_KC=<%=V%> onMouseOver="window.status='Enlace a la página'"><%=D%></a> 
<% 
      End if
  wend
%>
<form id="frmsel"  NAME="frmsel" METHOD="post" ACTION="MGBKS.asp?P_KR=<%=P_KR %>%20"&P_OPERACION="<%=P_OPERACION%>">
<% lpp = 10
   strLinkFlechas="MGBKS.asp?P_KC=" & P_KC & "&P_OPERACION=" & P_OPERACION & "&P_DSSEL=" & p_dssel & "&P_KCSEL=" & p_kcsel
   strLinkPagina="MGBKS.asp?P_KC=" & P_KC & "&P_OPERACION=" & P_OPERACION & "&P_DSSEL=" & p_dssel & "&P_KCSEL=" & p_kcsel & "&P_ORDEN=" & P_ORDEN   
   GF_PAGINAR "N",strLinkPagina,lpp,MAXLPP,rs %>
<table align="center" width="95%" border="0" cellspacing=0> 
<TBODY>
<TR>
 <TD> 
 	  <input type=image  class="NOBORDER" src ="images/Aceptar.gif" align=absMiddle alt="<%=gf_traducir("Aceptar")%>" >&nbsp;<% =GF_TRADUCIR("Aceptar") %> | 	   	  
      <input type=image  class="NOBORDER" src ="images/Agregar.gif"  align=absMiddle alt="<%=gf_traducir("Agregar")%>" onClick = btnAgregar_onclick() >&nbsp;<% =GF_TRADUCIR("Agregar") %> | 
  <% if (My_BtnEstado = "") then
      if (My_BtnCambiar = "Si Cambios") then %>
         <input type=image  class="NOBORDER" src ="images/edit-16x16.png" align=absMiddle alt="<%=gf_traducir(My_BtnCambiar)%>" onClick =Btn_Cambiar_onclick() >&nbsp;<% =GF_TRADUCIR("Modificar") %> | 
   <% else %>
         <input type=image  class="NOBORDER" src ="images/ModificarNO.gif" align=absMiddle alt="<%=gf_traducir(My_BtnCambiar)%>" onClick =Btn_Cambiar_onclick() >&nbsp;<% =GF_TRADUCIR("Aceptar Cambios") %> | 
   <% end if %>      
      <input TYPE="hidden" value="<% =My_BtnCambiar %>" name="Btn_Cambiar">
  <% end if %>    
  <% if (My_BtnTodosSM_Estado = "") then %>
      <input type=image class="NOBORDER" src ="images/Search-16x16.png" align=absMiddle alt="<%=gf_traducir("Buscar Todos")%>" onclick=BtnTodosSM_onclick() >&nbsp;<% =GF_TRADUCIR("Buscar Todos") %> | 
  <% else 
		if (p_kc <> "SM") then %>          
      <a href="MGBKS.ASP?P_KM=SM"><img src ="images/Subir.gif" align=absMiddle alt="<%=gf_traducir("Volver al Maestro General")%>" style="cursor:hand;"></a>&nbsp;<% =GF_TRADUCIR("Subir un Nivel") %> |        
  <%    end if
	 end if %>    
  
 </td>
</TR> 
</TBODY>
</table>
<input type="hidden" name="P_ORDEN" value="<% =P_ORDEN %>">
<table class=reg_header align="center" width="100%" cellSpacing=1 cellPadding=2>
  <tr> 
      <% if P_BUSQUEDA = "TODOS_SM" then %>
     <td align="center" >
        &nbsp;
     </td>
     <% end if %>
 	 <td align="center">
 	    &nbsp;
	 </td>
	   <% My_Texto = GF_TRADUCIR("Maestro: ") & p_kc & "<br>" & GF_TRADUCIR("Descripción:") & P_ds%> 
	 <td align="center" height="1">
	    <table><tr><td>&nbsp;</td><td align="left"><b><%=my_texto%></b></TD></tr></table>
	 </td>
	 <TD ALIGN="center" >
  	    &nbsp;
     </td> 
     <TD ALIGN="center" >
  	  &nbsp;
     </td>
  </tr> 
  <tr class=reg_header_nav>
      <% if P_BUSQUEDA = "TODOS_SM" then %>
     <td>
        &nbsp;
     </td>
     <% end if %>
	  <td align="center" width=10%>
          <INPUT TYPE="text" size=9 NAME="p_KCsel" VALUE="<%=p_KCsel%>" >
      </td>
	  <td align="center" width=40%> 
	      <INPUT TYPE="text" NAME="p_dssel" size=50 VALUE="<%=p_dssel%>">
	  </td>
	  
	  <td align="center" nowrap>
   	      &nbsp;
      </td>
	  <td align="center" nowrap>
          &nbsp;
	  </td>
  </tr>

  <tr class=reg_header_nav >
     <% if P_BUSQUEDA = "TODOS_SM" then %>
     <td>
        <div align="center"><font size="3"><%=GF_TRADUCIR("Maestro")%></font></div>
     </td>
     <% end if %>
     <td>
        <div align="center">
			<A href="<% =strlinkFlechas %>&P_ORDEN=KCUP"><IMG src="images/arrow_up.gif" align=absMiddle border=0></a>
			<%=GF_TRADUCIR("Codigo")%>
			<A href="<% =strlinkFlechas %>&P_ORDEN=KCDOWN"><IMG src="images/arrow_down.gif" align=absMiddle border=0></a>
		</div>
     </td>
     <td>
        <div align="center">
			<A href="<% =strlinkFlechas %>&P_ORDEN=DSUP"><IMG src="images/arrow_up.gif" align=absMiddle border=0></a>
			<%=GF_TRADUCIR("Descripcion")%>
			<A href="<% =strlinkFlechas %>&P_ORDEN=DSDOWN"><IMG src="images/arrow_down.gif" align=absMiddle border=0></a>
		</div>
     </td>
     <td width="20%"> 
        <div align="center">
        <A href="<% =strlinkFlechas %>&P_ORDEN=MMUP"><IMG src="images/arrow_up.gif" align=absMiddle border=0></a>
        <%=GF_TRADUCIR("Fecha")%>
        <A href="<% =strlinkFlechas %>&P_ORDEN=MMDOWN"><IMG src="images/arrow_down.gif" align=absMiddle border=0></a>
        </div>
     </td>
     <td> 
		<div align="center">
		<A href="<% =strlinkFlechas %>&P_ORDEN=KRUP"><IMG src="images/arrow_up.gif" align=absMiddle border=0></a>
       <%=GF_TRADUCIR("Kr")%>
       <A href="<% =strlinkFlechas %>&P_ORDEN=KRDOWN"><IMG src="images/arrow_down.gif" align=absMiddle border=0></a>
       </div>
     </td>
 </tr> 
 
<%

my_kcAnterior = "???"
i = 0 
intIndice=0
  if (rs.recordcount = 1) and (P_OPERACION = "SELECT") then 
    PrepararUnRegistro
    response.redirect My_LinkDescri
  end if
  
  while not rs.eof and CInt(i) < CInt(lpp)
    if PrepararUnRegistro()  then
     %>
      <tr class=reg_header_navdos>
      
      <% if P_BUSQUEDA = "TODOS_SM" then %>
         <td ALIGN=CENTER > 
		 <%=rs("Mg_Km")%> </td>
      <% end if %>
      <TD align=left > 
	      <input name="p_MgKc<%=i%>" type="<%=My_TypeKC%>" maxlength="15" size=9 value="<%=rs("MG_kc")%>" >
	  	  <input name="p_MgKcOriginal<%=i%>" type="hidden" value="<%=rs("MG_kc")%>">
		   <% 
		   if My_TypeKC = "Hidden" then
		      if P_BUSQUEDA <> "TODOS_SM" then 
		         Response.Write "<A HREF=" & My_LinkDescri & ">" & rs("MG_kc") & "</A>"
	             else
	             Response.write rs("MG_kc")
	          end if 
	       end if   
	       %>
        </td>
           <%
           IF P_OPERACION = "SELECT" THEN 
           %>
	       <TD > 
	         <%=rs("MG_ds")%> 
	       </td>
	       <%
	       else
	         if (p_Kc = "SM") then
		        intMostrar = 1
	         else
		        intMostrar = 0
		     end if   
	       %>  
          <TD > 
	         <input name="p_MgDs<%=i%>" type="<%=My_TypeDS%>" value="<%=rs("MG_Ds")%>" size=50>
  	         <input name="p_MgDsOriginal<%=i%>" type="hidden" value="<%=rs("MG_Ds")%>">
		      <%
		      if My_TypeDS = "Hidden" then
		         if P_BUSQUEDA <> "TODOS_SM" then 
		         %>
    	           <A HREF=<%=My_LinkDescri%> onMouseOver="LF_VERDEFINICION(<% =intIndice %>,<% =intMostrar %>)"> <%=rs("MG_ds") %> </A>
	               <table cellpadding=0 id="POP<% =intIndice%>" style="border-width:1;border-style:outset;visibility:hidden;position:absolute;background-color:white;border-width:1;border-style:outset;" cellspacing=0>
		             <tr>
		               <td>
	                     <A HREF=<%=My_LinkCodigo%>><font color=Navy>VER DEFINICION MAESTRO</font></a>
    	               </td>
    	             </tr>
			       </table>
		         <%
		         else
		            Response.write rs("MG_ds")
		         end if
		      end if   
		      %>
	      </td>
	        <%
	        END IF
	  ' Modificar si tiene la hora
	   fecha = trim(rs("MG_mmmv"))
	   if len(fecha) > 10  then 
	      fecha = mid(fecha,1,instr(fecha," "))
	   end if
	   %>  
      <TD align=center > <%=fecha%> </td> 
        
      <TD width=10% align=center> 
          <input name="p_MgKr<%=i%>" type="hidden" value="<%=rs("MG_KR")%>">	  
	      <%=RS("MG_KR")%> 
	  </td>
    </TR> 
     <%
      end if 
  rs.movenext 
  intIndice= intIndice+1
 wend
%> 
<input type="hidden" name="hdnkm" value="<% =p_km %>"> 
<input type="hidden" name="P_OPERACION" value="<% =p_OPERACION %>"> 
<% gf_bd_control rs ,con ,"CLOSE","" %> 
</table>
</FORM>

<tr> 
  <td WIDTH="108">&nbsp;
    
  </td>
</tr> 
</html>
<script language="javascript" >
function btnAgregar_onclick()
{ 
	document.frmsel.action="MAESTROGENERALALTAS.ASP?P_KM=<%=P_KC%>&P_OPERACION=<%=P_OPERACION%>"
}
function BtnTodosSM_onclick()
{ 
	document.frmsel.action="MGBKS.ASP?P_kr=<%=P_kr %>&P_BUSQUEDA=TODOS_SM"
}
function Btn_Cambiar_onclick()
{
	if (document.frmsel.Btn_Cambiar.value == "Si Cambios") 
	{
		document.frmsel.action="MGBKS.ASP?P_kr=<%=P_kr %>&P_OPERACION=MODIFICAR"
	}
	else
	{
		document.frmsel.action="MGBKS.ASP?P_kr=<%=P_kr %>&P_OPERACION=Empty"
	}
}
</SCRIPT>


