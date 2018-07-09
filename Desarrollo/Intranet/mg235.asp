<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->
<% ProcedimientoControl "MG235" %>
<%
'Administrar relaciones Permite visualizar, Modificar, Dar de alta.
dim con,strBajas,strRelacionados 
dim rs, Sql, mensaje,C,V, KM, KC, VL 
dim MG_km , MG_kc , MG_Ds , MG_Kr
dim p_km , p_kc , p_Ds , p_Kr, P_VL, MY_VL
dim p_km1, p_kc1, p_Ds1, p_Kr1 ,p_smkr1,p_smds1,p_smacceso1
dim p_km2, p_kc2, p_Ds2, p_Kr2,p_smkr2,p_smds2,p_smacceso2
dim p_km3, p_kc3, p_Ds3, p_Kr3, p_smkr3,p_smds3,p_smacceso3
dim P_VL3
dim N_3okr
DIM N_VL
DIM P_3okr
dim my_kc, my_km, my_ds, MY_3OKR
dim p_ACCION
dim My_AccesoRelacion, My_AccesoRelacionante, My_AccesoRelacionado
dim P_RELACIONADOS
dim P_BAJAS
Dim My_Contador, dioAlta
dioAlta = "TDERROR"
'--------------------------------------------------------------------
sub TomarParametros
P_KM1 = GF_PARAMETROS("P_km1","P_KM1")
P_KC1 = GF_PARAMETROS("P_kC1","P_KC1")
P_KM2 = GF_PARAMETROS("P_km2","P_KM2")
P_KC2 = GF_PARAMETROS("P_kC2","P_KC2")
P_KM3 = GF_PARAMETROS("P_km3","P_KM3")
P_KC3 = GF_PARAMETROS("P_kC3","P_KC3")
P_VL3 = GF_PARAMETROS("P_VL3","P_VL3")
P_KR1 = GF_PARAMETROS("P_KR1","")  + 0
P_KR2 = GF_PARAMETROS("P_KR2","")  + 0
P_KR3 = GF_PARAMETROS("P_KR3","")  + 0
IF P_KR1 > 0 THEN GF_MGC P_KM1,P_KC1,P_KR1,P_DS1
IF P_KR2 > 0 THEN GF_MGC P_KM2,P_KC2,P_KR2,P_DS2
IF P_KR3 > 0 THEN GF_MGC P_KM3,P_KC3,P_KR3,P_DS3
P_ACCION = ucase(GF_PARAMETROS("P_ACCION","??"))
if p_accion = "NEXTO1" THEN BuscarRelacionSiguiente
if p_accion = "NEXTO2" THEN BuscarRelacionanteSiguiente
P_bajas = GF_PARAMETROSFORM("P_BAJAS")
IF P_Accion = "BAJAS" or p_bajas = "" THEN DefinirSiNo p_bajas
P_RELACIONADOS = gf_parametrosform("P_RELACIONADOS")
IF P_Accion = "RELACIONADOS" or p_relacionados = ""  THEN DefinirSiNo p_relacionados
end sub
'--------------------------------------------------------------------
SUB DefinirSiNo(byref p_parametro)
   IF P_Parametro = "NO" then
      P_Parametro = "SI"
	  else
	  P_Parametro = "NO"
   end if
end sub  
'--------------------------------------------------------------------
sub ControlarRelacion
	My_AccesoRelacion  = GF_mg_acceso(P_KM1,p_KC1,P_KR1,P_ds1,p_smkr1,p_smds1,p_smacceso1,mensaje)
	if MENSAJE <> ""  then  mensaje = "Relacion: " &  mensaje
end sub
'--------------------------------------------------------------------
sub ControlarRelacionante
	if mensaje = "" then   
	   My_AccesoRelacionante  = GF_mg_acceso(P_KM2,p_KC2,P_KR2,P_ds2,p_smkr2,p_smds2,p_smacceso2,mensaje)
	   if mensaje <> "" then  mensaje = "Relacionante: " & MENSAJE
	end if
end sub
'--------------------------------------------------------------------
sub ControlarRelacionado
	if mensaje = "" then
		if p_km3 <> "" or p_kc <> "" then
		   P_KR3 = 0
		   My_AccesoRelacionado  = GF_mg_acceso(P_KM3,p_KC3,P_KR3,P_ds3,p_smkr3,p_smds3,p_smacceso3,mensaje)
		   if mensaje <> "" then  mensaje = "Relacionado: " & MENSAJE
		end if
	end if
end sub
'--------------------------------------------------------------------
sub DardealtaNuevaRelacion   
	if p_ACCION = "ADD" and mensaje = "" THEN
		if p_km3 = "" then 
		      mensaje = "Debe indicar un maestro"
		else 
			if p_kc3 = "" then
				mensaje = "Debe indicar un codigo"
			else 
				IF P_VL3 = "" THEN
					Mensaje = Mensaje & "Debe tener un valor"
				else 
					dioAlta = "TDSUCCESS"
					IF GF_MGSRADD(P_kr1, P_kr2, P_kr3, P_VL3, P_3okr) THEN
						Mensaje = mensaje & "Se dio de alta"
					else
						Mensaje = Mensaje & "Se modifico un valor previo"
					end if
				end if   	  
			end if
		end if   
		if mensaje <> "" then mensaje = "La nueva relacion " & mensaje  
	end if
end sub
'--------------------------------------------------------------------
sub BuscarRelacionSiguiente
' Buscar la relacion siguiente.
   sql = "SELECT TOP 1 * FROM relacionesconsulta WHERE (sro1Km + sro1kc ) > '" & P_km1 & p_kc1 & "' ORDER BY sro1km, sro1kc "   
   gf_bd_control rs,con,"OPEN",Sql      
   IF NOT RS.EOF THEN
      P_KM1 = RS("SRO1KM")
	  P_KC1 = RS("SRO1KC")
	  p_kr1 = rs("SRO1KR")
	  P_KM2 = ""
	  P_KC2 = ""
	  P_KM3 = ""
	  P_KC3 = ""
  	  BuscarRelacionanteSiguiente
   ElSE		
	  mensaje = GF_TRADUCIR("ERROR") & ":_No hay mas relaciones"	  
   END IF      
   gf_bd_control rs,con,"CLOSE",Sql     
end sub   
'--------------------------------------------------------------------
sub buscarRelacionanteSiguiente  
DIM MY_SQL, MY_RS, my_con 
   mY_SQL = "SELECT TOP 1 * FROM RELACIONESCONSULTA WHERE sro1km= '" & P_km1 & "' and sro1kc = '" & p_kc1 & "'"
   my_sql = my_sql & " AND (SRO2KM + SRO2KC) > '" & P_KM2 & p_KC2 & "' ORDER BY sro2kM , SRO2KC "
   gf_bd_control My_rs,my_con,"OPEN",mY_SQL   
   IF NOT My_rs.EOF THEN
      P_KM2 = My_rs("SRO2KM")
	  P_KC2 = My_rs("SRO2KC")
	  else
	  p_km2 = ""
	  P_KC2 = ""
   END IF   
   gf_bd_control My_rs,my_con,"CLOSE",mY_SQL  
END SUB   
'-----------------------------------------------------------------------------

Mensaje = ""
Tomarparametros
ControlarRelacion
if mensaje = "" then ControlarRelacionante
if mensaje = "" then ControlarRelacionado
if mensaje = "" THEN DardeAltaNuevaRelacion
%>
<html>
<head>
<title>Codigo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
</head>
<body >
<br>
<% =GF_TITULO("Relaciones.gif", "Administracion de Relaciones") %>
<form name="form1" method="post" action="MG235.asp">
<table align="center" width="95%" border="0" cellspacing=1 cellpadding=2> 
	<TR>
		<TD>
			<INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Aceptar.gif" NAME="agregar" >&nbsp;<%=gf_traducir("Aceptar")%> |
			<INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Agregar.gif" NAME="Add" onclick="Add_onclick()">&nbsp;<%=gf_traducir("Agregar")%> |
			<% 
			strBajas= "No" 
			if (p_Bajas = "SI") then strBajas=""
			strRelacionados= "No" 
			if (p_relacionados = "SI") then strRelacionados=""
			%>
			<INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/RelacionBorrada.gif" NAME="Bajas" style="font-size:8pt;" onclick="Bajas_onclick()" TabIndex=12>&nbsp;<% =GF_TRADUCIR(strBajas & " Ver Bajas")%> |
			<INPUT align=absMiddle class="NOBORDER" TYPE="image" src ="images/Relacionados.gif" NAME="RELACIONADOS" style="font-size:8pt;" onclick="RELACIONADOS_onclick()" TabIndex=13>&nbsp;<%=GF_TRADUCIR(strRELACIONADOS & " Ver Relacionados")%> |
		</td>
	</TR> 
</table>
<input type="HIDDEN" name="P_BAJAS" value="<%=P_BAJAS%>" >
<input type="HIDDEN" name="P_RELACIONADOS" value="<%=P_RELACIONADOS%>" >
<table border=1 align="center" width="98%" cellspacing=1 cellpadding=0>
	<tr align="center" class="reg_Header_nav"> 
		<th> 
			<div ><font size="4"><%=GF_TRADUCIR("Tipo")%></font></div>
		</th>
		<th > 
			<div ><font size="4"><%=GF_TRADUCIR("Maestro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Descripcion maestro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Codigo")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Descripcion registro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Valor")%></font></div>
		</th>
	</tr>

    <tr class="reg_Header"> 
		<td> <%=gf_traducir("Relacion")%> </TD>
		<th bgcolor="#FFFFFF"> <input type="text" name="p_km1" value="<%=P_km1%>" size="2" TabIndex=1> </th>
		<th bgcolor="#FFFFFF"><A HREF=<%="mg210.asp?P_KR=" & p_smkr1%>> <%=p_smds1%></A></th>
		<th bgcolor="#FFFFFF"> <input type="text" name="p_kc1" value="<%=P_Kc1%>" size="14" TabIndex=2></th>
		<th bgcolor="#FFFFFF"><A HREF=<%="mg210.asp?P_KR=" & p_kr1%>> <%=p_ds1%></A></th>
		<td align="center">&nbsp;</TD>
    </tr>
    <tr class="reg_Header"> 
		<td> <%=gf_traducir("Relacionante")%> </TD>
		<th bgcolor="#FFFFFF"> <input type="text" name="p_km2" value="<%=P_KM2%>" size="2" TabIndex=3></th>
		<th bgcolor="#FFFFFF"> <A HREF=<%="mg210.asp?P_KR=" & p_smkr2%>> <%=p_smds2%></A></th>
		<th bgcolor="#FFFFFF"> <input type="text" name="p_kc2" value="<%=P_Kc2%>" size="14" TabIndex=4></th>
		<th bgcolor="#FFFFFF"> <A HREF=<%="mg210.asp?P_KR=" & p_kr2%>> <%=p_ds2%></A></th>
		<td align="center">&nbsp;</TD>
	</tr>
	<tr class="reg_Header"> 
     	<td><%=gf_traducir("Relacionado")%>  </TD>
        <th bgcolor="#FFFFFF"> <input type="text" name="p_km3" value="<%=P_KM3%>" size="2" TabIndex=5></th>
		<th bgcolor="#FFFFFF"> <A HREF=<%="mg210.asp?P_KR=" & p_smkr3%>> <%=p_smds3%></A></th>
        <th bgcolor="#FFFFFF"> <input type="text" name="p_kc3" value="<%=P_Kc3%>" size="14" TabIndex=6>      </th>
        <th bgcolor="#FFFFFF"> <A HREF=<%="mg210.asp?P_KR=" & p_kr3%>> <%=p_ds3%></A></th>
		<th bgcolor="#FFFFFF"> <input type="TEXT" name="P_VL3" VALUE="<%=P_VL3%>" size="11" TabIndex=7></th>
	</tr>
</table>
<br>
<table border=1 align="center" width="98%" cellspacing=1 cellpadding=0>

	<tr class="reg_Header_nav" align="center" > 
		<th > 
			<div ><font size="4"><%=GF_TRADUCIR("Tipo")%></font></div>
		</th>
		<th > 
			<div ><font size="4"><%=GF_TRADUCIR("Maestro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Descripcion maestro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Codigo")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Descripcion registro")%></font></div>
		</th>
		<th > 
			<div><font size="4"> <%=GF_TRADUCIR("Valor")%></font></div>
		</th>
	</tr>	    
    <%
	if mensaje <> "" then %>
	<tr class="reg_Header"> 
		<TD colspan=6 align="center">
			<table align="center"  border=0 width="100%" cellspacing=0>
				<tr> 
					<th class="<%=dioAlta%>" align="center" colspan=4 height="21" ><%=mensaje%></th>
				</tr>
			</table>
		</td>		
	</tr>	  
	<% 
	end if     
    
    if p_kr1 <> 0 and p_kr2 <> 0 and  P_RELACIONADOS = "NO" then 
		sql = "Select sro1km,sro1kc,srvalor,sr3okr,sro3km,sro3kc,sro3ds,sro3kr from RELACIONESCONSULTA WHERE sro1KR= " & P_kr1 & " and sro2kr= " & P_kr2   
		IF P_BAJAS = "SI" THEN  SQL = SQL & " AND SRVALOR <> '*' " 'No mostrar bajas
		sql = sql & " order by srvalor,sro3km,sro3kc"
		call GF_BD_CONTROL (rs,con,"OPEN",Sql)
		My_Contador = 3
		while not rs.eof 
			My_Contador = My_Contador + 1
			'generar los nombres de los campos de input
			if rs("SRO3KM") <> MG_KC THEN 
				MG_KC = RS("SRO3KM")
				MG_KR = 0 
				GF_MGC "SM",MG_KC,MG_KR,MG_DS 
			END IF   
			N_3OKR = "P_3OKR" & My_Contador 
			N_VL = "P_VL" & My_Contador 
			MY_3OKR = RS("SR3OKR")
			P_VL = RTRIM(LTRIM(GF_PARAMETROSFORM(N_VL)))
			P_3OKR = GF_PARAMETROSFORM(N_3OKR) 
			IF NOT ISNUMERIC(P_3OKR) THEN P_3OKR = 0
			p_3okr = p_3okr + 0
			MY_VL = RTRIM(LTRIM(RS("SRVALOR") ))
			'RESPONSE.WRITE "VALOR MY_3OKR(" & MY_3OKR & ")"
			'RESPONSE.WRITE "VALOR P_3OKR (" & P_3OKR & ")"
			iF P_3OKR = MY_3OKR THEN 
				'RESPONSE.WRITE "Mismo KR VALORES(" & P_VL & ") Y (" & MY_VL & ")"
				IF P_VL&"X" <> MY_VL&"X" then 
					GF_MGSRADD P_KR1,P_KR2,RS("SRO3KR"),P_VL,MY_3OKR
					MY_VL = P_VL
				END IF 	  
			END IF   
			%>
			<tr class="reg_Header"> 
				<td> <%=gf_traducir("Relacionado")%> </TD>
				<input type="HIDDEN" name="P_3OKR<%=My_Contador%>" value="<%=MY_3OKR%>" > 
				<td bgcolor="#FFFFFF" align="center">  <%=rs("sro3km")%> </TD>
				<td bgcolor="#FFFFFF" align="left"> <A HREF=<%="mg210.asp?P_KR=" & mg_kr%>> <%=mg_ds%></A></td>
				<td bgcolor="#FFFFFF" align="center"> 
					<A HREF="mg235.asp?P_Kc2=<% =rs("sro3kc")%>&P_Km2=<% =rs("sro3km")%>&P_Kc1=<% =rs("sro1kc")%>&P_Km1=<% =rs("sro1km")%>">
					<%=rs("sro3kc")%> </A>
				</td>
				<td bgcolor="#FFFFFF" align="left"> <A HREF=<%="mg210.asp?P_KR=" & rs("sro3kr")%>> <%=rs("sro3ds")%></A></td>
				<td bgcolor="#FFFFFF" align=center> <input type="TEXT" name="P_VL<%=My_Contador%>" VALUE="<%=MY_VL%>" size="11" TabIndex=14>      </td>
			</tr>
			<%
			rs.movenext
		wend
		call GF_BD_CONTROL (rs,con,"CLOSE","")
		if my_contador > 8 and 1=2 then %>
			<tr> 
				<TD colspan=6 align="center" class="reg_header_nav" >
					<INPUT TYPE="Submit" NAME="NextO1" VALUE="Relacion + " onclick="NextO1_onclick()">
					<INPUT TYPE="Submit" NAME="NextO2" VALUE="Relacionante +" onclick="NextO2_onclick()">
				</td>
			</tr>
	    <%
		end if
	end if
	%>   
</table>
</form>
</body>
</html>
<SCRIPT LANGUAGE = "javascript">
	function Add_onclick(){
		document.form1.action="MG235.ASP?P_ACCION=Add"
	}
	function NextO1_onclick(){
		document.form1.action="MG235.ASP?P_ACCION=NextO1"
	}
	function NextO2_onclick(){
		document.form1.action="MG235.ASP?P_ACCION=NextO2"
	}
	function Bajas_onclick(){
		document.form1.action="MG235.ASP?P_ACCION=Bajas"
	}
	function RELACIONADOS_onclick(){
		document.form1.action="MG235.ASP?P_ACCION=RELACIONADOS"
	}
</SCRIPT>