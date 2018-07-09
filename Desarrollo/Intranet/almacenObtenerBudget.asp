<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
			
<%
dim myIdObra, myIdBudgetArea, myIdBudgetDetalle, rsBudget, antArea, myReadOnly
dim myTexto
myIdObra = GF_PARAMETROS7("idObra",0 ,6)
myIdBudgetArea = GF_PARAMETROS7("idBudgetArea", 0, 6)
myIdBudgetDetalle = GF_PARAMETROS7("idBudgetDetalle", 0, 6)
myReadOnly = GF_PARAMETROS7("readOnly", 0, 6)
accion = GF_PARAMETROS7("accion","" ,6)
idDivision = GF_PARAMETROS7("idDivision",0 ,6)
%>
<%  if accion = ACCION_PROCESAR Then%>
<%	if (myReadOnly=0) then	 
				Set rsBudget = obtenerListaBudgetObra(myIdObra, 0, 0)
%>
				<select id="idBudgetDetalle" name="idBudgetDetalle" onBlur="javascript:readBudgetArea(this)" class="inputs">
					<option value="0">Ninguno...</option>
					<%	
					while not rsBudget.eof
							if (CLng(rsBudget("IDAREA")) <> antArea) then 
								antArea = CLng(rsBudget("IDAREA"))
								%>
								<optgroup label="<%=rsBudget("IDAREA")%>-<%=rsBudget("DSBUDGET")%>"></optgroup>
								<%
							else
								%>
								<option alt="<%=rsBudget("IDAREA")%>" value="<%=rsBudget("IDDETALLE")%>" <%if ((CLng(rsBudget("IDDETALLE")) = myIdBudgetDetalle) and (antArea = myIdBudgetArea)) then response.write "selected='true'"%>>
									&nbsp;&nbsp;<%=rsBudget("IDDETALLE")%>.<%=rsBudget("DSBUDGET")%>
								</option>
							<% 
							end if
							rsBudget.MoveNext()
					wend 	
					%>									
				</select>		
<%	else	        
	    Set rsBudget = obtenerListaBudgetObra(myIdObra, myIdBudgetArea, myIdBudgetDetalle)
	    if (not rsBudget.eof) then myTexto = rsBudget("DSBUDGET")        
		response.write myTexto	%>	
		<input type="hidden" name="idBudgetDetalle" id="idBudgetDetalle" value="<%=myIdBudgetDetalle%>">
<%	end if	%>	
	<input type="hidden" name="idBudgetArea" id="idBudgetArea" value="<%=myIdBudgetArea%>">
<% else %>	
	<select name="masterSelect" id="masterSelect" size="1" onchange="actualizarAreaDetalle(this)" >			
	<%	Set rs = obtenerListaObras("", "", "",idDivision,OBRA_ACTIVA)
		If (not rs.Eof) then %>
			<optgroup label="<%=GF_TRADUCIR("PARTIDA PRESUPUESTARIA")%>" id="0">
		<%	while not rs.eof			
				myClass = ""
				myValue = rs("IDOBRA")
				if(rs("ESINVERSION") = OBRA_MANTENIMIENTO) Then	
					mySelect = "selected='selected'"
				else 
					mySelect = ""
				end if	%>
				<option value="<%=myValue %>" <%=myClass%> <%=mySelect%> >
					&nbsp;&nbsp;<%= rs("CDOBRA") & "-" & rs("DSOBRA")%>
				</option>
		<%		rs.movenext
			wend%>
			</optgroup>
		<%end if
		Set rs = obtenerSectores("")
		If (not rs.Eof) then %>
			<optgroup label="<%=GF_TRADUCIR("SECTORES")%>" id="1">
		<%	while not rs.eof  %>
				<option value="<%=rs("IDSECTOR") %>" >
					&nbsp;&nbsp;<%= rs("IDSECTOR") & "-" & rs("DSSECTOR")%>
				</option>
		<%		rs.movenext
			wend
		%></optgroup><%	
		end if %>
	</select>
<% end if %>
