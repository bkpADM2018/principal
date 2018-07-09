<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<%
dim idObra,idArea,idDeta,modificable
idObra = GF_PARAMETROS7("idObra",0 ,6)
idArea = GF_PARAMETROS7("idArea",0 ,6)
idDeta = GF_PARAMETROS7("idDeta",0 ,6)
modificable = GF_PARAMETROS7("modificable","",6)
%>
<script>
	function getAreaDetalle(me)
	{
		aux = $(me).val().split("-")
		$("#cmbIdArea").val(aux[0]);
		$("#cmbIdDeta").val(aux[1]);
	}
</script>


<%if (UCASE(modificable) = "TRUE") then%>
<select name="AreaDetalle", id="AreaDetalle" size="1" onchange="getAreaDetalle(this)">	
	<option value ="0"><%=GF_TRADUCIR("- Seleccione (Opcional) -")%></option>
	<%
	Set rs = leerBudget(idObra)
	while not rs.eof 
		
		myValue = rs("IDAREA") & "-" & rs("IDDETALLE")
		if (rs("IDDETALLE")=0) then 
			myClass =  "style='font-weight: bold' " 
		else 
			myClass = ""
		end if
		if (cint(idArea)=rs("IDAREA")) and (cint(idDeta) = rs("IDDETALLE")) then 
			mySelect = "selected='selected'"
		else 
			mySelect = ""
		end if
		
		%>
  <option value="<%=myValue %>" <%=myClass%> <%=mySelect%> >
			<%if (rs("IDDETALLE")<>0) then
				response.write rs("IDAREA") & " - " & rs("IDDETALLE") & " - " &  rs("DSBUDGET")
			else
				response.write rs("IDAREA") & " - " & rs("DSBUDGET")
			end if%>
  </option>
	<%rs.movenext
	wend%>
</select>
<input type="hidden" id="cmbIdArea" name="cmbIdArea" value="<%=idArea%>">
<input type="hidden" id="cmbIdDeta" name="cmbIdDeta" value="<%=idDeta%>">
<% else 
	Set rs = leerBudget(idObra)
	while not rs.eof 
		if (cdbl(rs("idarea")) = cdbl(idarea) and cdbl(rs("iddetalle")) = cdbl(idDeta)) then 	%>
			<label><%=rs("idArea") & "-" & rs("idDetalle") & " " & rs("dsbudget")%></label>
			<input type="hidden" id="cmbIdArea" name="cmbIdArea" value="<%=rs("idArea")%>">
			<input type="hidden" id="cmbIdDeta" name="cmbIdDeta" value="<%=rs("iddetalle")%>">
	<% 	end if
		rs.movenext
	wend%>
<% end if %>