<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 18/01/2008                         -->
'<!--               Pagina     : AUPGetSistemas.ASP                 -->
'<!--               Descripcion: Listado de sistemas                -->
'<!------------------------------------------------------------------->

Const TASK_TOKEN = "///"
Const TASK_KEY_ARR = "TK_AA"
Const TASK_KEY_TRA = "TK_TA"
Const TASK_KEY_PIE = "TK_PA"
Const TASK_KEY_EXP = "TK_EA"
Const TASK_KEY_N = "SCW"
Const TASK_SC_ARR = "TK_AA"
Const TASK_SC_TRA = "TK_TA"
Const TASK_SC_PIE = "TK_PA"
Const TASK_SC_EXP = "TK_EA"
Const TASK_SC_N = "SC_N_"
'----------------------------------------------------------------------------------------------
function PrepareWord2(pWord)
dim auxWord
	auxWord = replace(pWord," ","%20")
	PrepareWord2 = pWord
end function
'----------------------------------------------------------------------------------------------
function PrepareWord(pWord)
dim auxWord
	auxWord = replace(pWord,"á","a")
	auxWord = replace(auxWord,"é","e")
	auxWord = replace(auxWord,"í","i")
	auxWord = replace(auxWord,"ó","o")
	auxWord = replace(auxWord,"ú","u")
	auxWord = replace(auxWord,"ñ","n")
	auxWord = replace(auxWord,"(","{")
	auxWord = replace(auxWord,")","}")
	PrepareWord = auxWord
end function
'----------------------------------------------------------------------------------------------
function getPreparedWord(pWord)
	dim rtrn
	rtrn = replace(pWord,"Arroyo - ","")
	rtrn = replace(rtrn,"Transito - ","")
	rtrn = replace(rtrn,"Exportacion - ","")
	rtrn = replace(rtrn,"Piedrabuena - ","")
	'rtrn = replace(rtrn,"Nivel ","")
	getPreparedWord = rtrn
end function
'----------------------------------------------------------------------------------------------
function isValidIndex(pIndex)
	if pIndex < dicContabilidad.Count then isValidIndex = true
end function
'----------------------------------------------------------------------------------------------
function putInputCredencial(p_idPersona)
	dim sql, oConn, rsCredencial
	sql = "select NroCredencial from Profesionales where idProfesional = " & p_idPersona
	call GF_BD_CONTROL (rsCredencial,oConn,"OPEN",sql)
	if rsCredencial.eof then
		response.write "no esta registrado este profesional en el sistema"
	else
		%>
		<tr class=reg_header_navdos>
		<td colspan="2" align=left>
			<%=GF_Traducir("Numero de credencial")%>
			<input type="text" id="nroCredencial" name="nroCredencial" size=10 maxlength=8 style="border-style:none;" value="<%=rsCredencial("NroCredencial")%>">
		</td></tr>
		<%
	end if
end function
'----------------------------------------------------------------------------------------------
'Funcion que permite saber si un determinado sistema agrupa sus tareas para las diferentes divisiones.
Function isArrangedByDivision(myTask)
	ret = false
	Select case (myTask)
		case  "SCW"
			ret = true
		case "SAW"
			ret = true
	End Select
	isArrangedByDivision = ret
End Function
'----------------------------------------------------------------------------------------------
Function GF_MGSR3(P_o1kr, p_o2kr, P_o3kr, byref p_valor, byref P_3okr)
DIM CON, RS, strSQL
  strSQL = "SELECT * FROM RelacionesConsulta where sro1kr = " & p_o1kr & "  and sro2kr = " & p_o2kr & " and sro3kr = " & p_o3kr & " order by SRMMDT desc"
  GF_BD_Control rs,con, "OPEN", strSQL
  P_3OKR = ""
  p_valor = ""
  GF_MGSR3 = false
  if not RS.EOF then 
	if rs("srValor") <> "*" then
		p_valor = rs("srValor")
		p_3okr = rs("sr3okr")
		GF_MGSR3 = true
	end if	
  end if	   
  GF_BD_Control rs,con,"CLOSE",strSQL 
end function
'----------------------------------------------------------------------------------------------
'	COMIENZO PAGINA
'----------------------------------------------------------------------------------------------
dim myKey, myOutput, myTask, typeOfInput, myNameFunction, myWidth
dim elementAA, elementPA, elementEA, elementTA,elementN,  index, index2, flag
dim mySRO1KR, myO1KRPS, FrmDic
dim strSQL, rsTareas, oConn, myChecked, myAccion
dim myTTKC, myTTDS, MyStatus, myIndex
dim mySelectedX, mySelectedY, mySelectedA, mySelectedU, myValor

'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary")
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next

if trim(FrmDic("pEgreso")) = "V" then MyStatus = "Disabled"
	
'Response.Write PrepareWord2(FrmDic("pSisDS"))
call GF_MGC("SR","PRTT",myO1KRPS,"")
call GF_MGC("SR","TSTT",mySRO1KR,"")

strSQL= "Select * from RelacionesConsulta where sro1kr=" & mySRO1KR & " and sro2kr=" & FrmDic("pSisKR") & " and srvalor<>'*' order by sro3kc asc"
'Response.write strSQL
call GF_BD_CONTROL (rsTareas,oConn,"OPEN",strSQL)
%>
<table width="90%" border=0 align="center" class="reg_header" cellpadding="2" cellspacing="1">
<%
if rsTareas.eof then
%>
	<tr>
		<td>
			<font color=red>&nbsp;&nbsp;&nbsp;&nbsp;<%=gf_traducir("No existen tareas asignadas al sistema!")%></font>
		</td>
	</tr>
<% 
else
	set dicContabilidad = CreateObject ("Scripting.Dictionary")
	while not rsTareas.eof
		myTask = ucase(rsTareas("sro2kc"))
		mySelectedY = ""
		mySelectedA = ""
		mySelectedU = ""
		mySelectedX = ""
		if GF_MGSR3(myO1KRPS, FrmDic("IdPersona"), rsTareas("sr3okr"), myValor, "") then
			myChecked = "Checked"
			myAccion = "D"
			'Response.Write "<br>" & myValor
			select case myValor
				case "Y"
					mySelectedY = "SELECTED"
					myAccion = "A"
				case "A"
					mySelectedA = "SELECTED"
					myAccion = "A"
				case "U"
					mySelectedU = "SELECTED"
					myAccion = "A"
				case "X"
					mySelectedX = "SELECTED"
					myAccion = "A"
			end select		
		else
			myChecked = ""
			myAccion = "A"
			mySelectedX = "SELECTED"
		end if
		myTTKC = rsTareas("sro3kc")
		myTTDS = PrepareWord(rsTareas("sro3ds"))
		
		if myTask = "SAW" or myTask = "SC" then
			myKey = left(rsTareas("sro3kc"),5)			
			typeOfInput = "CheckBox"	
			myNameFunction = "cargarLista"
			myOutput = "<td nowrap align=left><font>"
			myOutput = myOutput & "<input style='border-style:none;cursor:pointer;' type='" & typeOfInput & "' onclick=" & myNameFunction  & "(" & rsTareas("sr3okr") & ",'" & replace(FrmDic("pSisDS")," ","%20") & "','" & myTTKC & "','" & replace(myTTDS," ","%20") & "','1','" & MyAccion & "') " & myChecked & " id='CheckBox1' name='CheckBox1' " & MyStatus & ">"
			myOutput = myOutput & getPreparedWord(rsTareas("sro3ds")) & "</font></td>"
			if dicContabilidad.Exists(myKey) then
				dicContabilidad(myKey) = dicContabilidad(myKey) & TASK_TOKEN & myOutput
			else
				dicContabilidad.Add myKey, myOutput
			end if
		
		elseif myTask = "SCW" then
			if left(rsTareas("sro3kc"),3) = "SCW" then
				myKey = left(rsTareas("sro3kc"),3)
			else
				myKey = left(rsTareas("sro3kc"),5)
			end if	
			typeOfInput = "checkbox"	
			myNameFunction = "cargarLista"
			'myOutput = myOutput & "<input style='border-style:none;cursor:pointer;' type='" & typeOfInput & "' onclick=" & myNameFunction  & "(" & rsTareas("sr3okr") & ",'" & PrepareWord2(FrmDic("pSisDS")) & "','" & myTTKC & "','" & replace(myTTDS," ","%20") & "','" & MyAccion & "') " & myChecked & " id='CheckBox1' name='CheckBox1' " & MyStatus & ">"
			'Response.Write "<hr>(" & rsTareas("sro3kc") & ")(" & mid(rsTareas("sro3kc"),8,1) & ")" & rsTareas("sr3okr") & ",'" & PrepareWord2(FrmDic("pSisDS")) & "','" & myTTKC & "','" & myTTDS & "','1', '" & MyAccion & "')' " & myChecked & " " & MyStatus
			if mid(rsTareas("sro3kc"),8,1) = "0" then
				myOutput = "<td nowrap align=left><font>"
				myOutput = myOutput & getPreparedWord(rsTareas("sro3ds")) & "</font></td>"
				myOutput = myOutput & "<td><select style='border-style:none;cursor:pointer;' onchange=cargarListaPre(this," & rsTareas("sr3okr") & ",'" & replace(FrmDic("pSisDS")," ", "%20") & "','" & myTTKC & "','" & replace(myTTDS," ", "%20") & "','" & MyAccion & "') " & myChecked & " id='select1' name='select1'><option value='X' " & mySelectedX & ">" & GF_Traducir("Denegado") & "</option><option value='U' " & mySelectedU & ">" & GF_Traducir("Usuario") & "</option><option value='Y' " & mySelectedY & ">" & GF_Traducir("Auditor") & "</option><option value='A' " & mySelectedA & ">" & GF_Traducir("Admin") & "</option></select></td>"
			elseif mid(rsTareas("sro3kc"),8,1) = "1" then
				myOutput = "<td colspan='2' nowrap align=left><font>"
				myOutput = myOutput & "<input style='border-style:none;cursor:pointer;' type='" & typeOfInput & "' onclick=" & myNameFunction  & "(" & rsTareas("sr3okr") & ",'" & replace(FrmDic("pSisDS")," ","%20") & "','" & myTTKC & "','" & replace(myTTDS," ", "%20") & "','1','" & MyAccion & "') " & myChecked & " id='CheckBox1' name='CheckBox1' " & MyStatus & ">"
				myOutput = myOutput & getPreparedWord(rsTareas("sro3ds")) & "</font></td>"
			else
				myOutput = "<td colspan='2' nowrap align=left>"
				myOutput = myOutput & "<input style='border-style:none;cursor:pointer;' type='" & typeOfInput & "' onclick=" & myNameFunction  & "(" & rsTareas("sr3okr") & ",'" & replace(FrmDic("pSisDS")," ","%20") & "','" & myTTKC & "','" & replace(myTTDS," ", "%20") & "','1','" & MyAccion & "') " & myChecked & " id='CheckBox1' name='CheckBox1' " & MyStatus & ">"
				myOutput = myOutput & "<font>" & getPreparedWord(rsTareas("sro3ds")) & "</font></td>"				
			end if	
			
			if dicContabilidad.Exists(myKey) then
				dicContabilidad(myKey) = dicContabilidad(myKey) & TASK_TOKEN & myOutput
			else
				dicContabilidad.Add myKey, myOutput
			end if
		else
			if (rsTareas("sro2kc") = "AAB" and myIndex = 0) then
				call putInputCredencial(FrmDic("IdPersona"))
			end if
			myIndex = myIndex + 1
			if myIndex mod 2 <> 0 then Response.Write "<tr class=reg_header_navdos>"
			%>
				<td width="50%" align=left>
					<font>
						<input style="border-style:none;cursor:pointer;" type="CheckBox" onclick="cargarLista(<%=rsTareas("sr3okr")%>, '<%=PrepareWord(FrmDic("pSisDS"))%>' ,'<%=myTTKC%>','<%=myTTDS%>', '1', '<%=MyAccion%>')" <%=myChecked%> id="CheckBox1" name="CheckBox1" <%=MyStatus%>>
						<%=myTTDS%><%=MyStatus%>
					</font>
				</td>
			<%
			if myIndex mod 2 = 0 then Response.Write "</tr>"
		end if
		rsTareas.movenext
	wend
	flag = true
	if myTask = "SC" then	
	'if (isArrangedByDivision(myTask)) then	'Proceso de los datos de contabilidad.				
		elementAA = split(dicContabilidad(TASK_SC_ARR), TASK_TOKEN)
		elementEA = split(dicContabilidad(TASK_SC_EXP), TASK_TOKEN)
		elementTA = split(dicContabilidad(TASK_SC_TRA), TASK_TOKEN)
		elementPA = split(dicContabilidad(TASK_SC_PIE), TASK_TOKEN)
		elementN = split(dicContabilidad(TASK_SC_N), TASK_TOKEN)
		while index < ubound(elementAA)+1
			Response.Write "<tr class=reg_header_navdos>"
			if index = 0 then Response.Write "<td class=reg_header_nav>Arroyo</td><td class=reg_header_nav>Piedrabuena</td><td class=reg_header_nav>Transito</td><td class=reg_header_nav>Exportacion</td></tr><tr class=reg_header_navdos>"
			Response.write elementAA(index)
			Response.write elementPA(index)
			Response.write elementTA(index)
			Response.write elementEA(index)		
			index = index + 1
		wend		
			Response.Write "</tr><tr class=reg_header_nav>"
			Response.Write "<td colspan=4 class=reg_header_nav>Niveles</td>"	
			Response.Write "<tr><td colspan=4><table width=100% class=reg_header_navdos><tr class=reg_header_navdos>"	
			index = 0
			while index < ubound(elementN)+1
				Response.write elementN(index)
				index = index + 1
			wend
			Response.Write "</tr></table></td></tr>"
			Response.Write "</tr>"
			%>
			<td colspan=4 valign="bottom" align=center style='BACKGROUND-COLOR: #FFEECD;'><a style='cursor:pointer;' onclick="createPopUpWindow();">[Definicion de procesos]</td>
			<%
	elseif myTask = "SAW" then	
	'if (isArrangedByDivision(myTask)) then	'Proceso de los datos de contabilidad.				
		elementAA = split(dicContabilidad(TASK_KEY_ARR), TASK_TOKEN)
		elementEA = split(dicContabilidad(TASK_KEY_EXP), TASK_TOKEN)
		elementTA = split(dicContabilidad(TASK_KEY_TRA), TASK_TOKEN)
		elementPA = split(dicContabilidad(TASK_KEY_PIE), TASK_TOKEN)
		elementN = split(dicContabilidad(TASK_KEY_N), TASK_TOKEN)
		while index < ubound(elementAA)+1
			Response.Write "<tr class=reg_header_navdos>"
			if index = 0 then Response.Write "<td class=reg_header_nav>Arroyo</td><td class=reg_header_nav>Piedrabuena</td><td class=reg_header_nav>Transito</td><td class=reg_header_nav>Exportacion</td></tr><tr class=reg_header_navdos>"
			Response.write elementAA(index)
			Response.write elementPA(index)
			Response.write elementTA(index)
			Response.write elementEA(index)		
			index = index + 1
		wend	
	elseif (myTask = "SCW") then
		'Response.Write "ARROYO" & dicContabilidad(TASK_KEY_ARR)
		elementAA = split(dicContabilidad(TASK_KEY_ARR), TASK_TOKEN)
		elementEA = split(dicContabilidad(TASK_KEY_EXP), TASK_TOKEN)
		elementTA = split(dicContabilidad(TASK_KEY_TRA), TASK_TOKEN)
		elementPA = split(dicContabilidad(TASK_KEY_PIE), TASK_TOKEN)
		elementN = split(dicContabilidad(TASK_KEY_N), TASK_TOKEN)
		while index < ubound(elementAA)+1
			if 1=2 then
			else
				Response.Write "<tr class=reg_header_navdos>"
				if index = 0 then Response.Write "<td colspan='2' class=reg_header_nav>Arroyo</td><td colspan='2' class=reg_header_nav>Piedrabuena</td><td colspan='2' class=reg_header_nav>Transito</td><td colspan='2' class=reg_header_nav>Exportacion</td></tr><tr class=reg_header_navdos>"
				Response.write elementAA(index)
				Response.write elementPA(index)
				Response.write elementTA(index)
				Response.write elementEA(index)		
			end if	
			index = index + 1
		wend	
		index = 0
		Response.write "<tr class=reg_header_nav2><td colspan='2'><b>" & GF_Traducir("Firmas") & "</b></td>"
		while index < ubound(elementN)+1
			Response.write elementN(index)
			index = index + 1
		wend	
	end if
end if

call GF_BD_CONTROL (rsTareas,oConn,"CLOSE",strSQL)
%>
</table>
