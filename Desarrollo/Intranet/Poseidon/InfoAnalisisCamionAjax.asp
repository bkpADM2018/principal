<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientosMG.asp"-->
<!--#include file="../includes/procedimientossql.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<%
'---------------------------------------------------------------------------------------------------------------
Dim g_strPuerto, g_idCamion, g_dtContable, g_ctaPorte,g_sqCalada

g_strPuerto = GF_Parametros7("Pto","",6)
g_dtContable = GF_Parametros7("dtContable","",6)
g_ctaPorte = GF_Parametros7("ctaPorte","",6)
g_idCamion = GF_Parametros7("idCamion","",6)
g_sqCalada = GF_Parametros7("sqCalada",0,6)


Set g_rsCaladas = getCaladaCamionBySqCalada (g_dtContable, g_idCamion, g_sqCalada, g_strPuerto)
%>
<link rel="stylesheet" type="text/css" href="../css/main.css" />
<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
<%if not g_rsCaladas.eof then	 %>
	<table class="datagridlv2" align="center" width="80%" id="tblCalada_<%=g_sqCalada%>">
    	<thead>
        	<tr> 
        		<th>&nbsp;</th>       	    
            	<th colspan="2"><%=GF_Traducir("Datos generales")%></th>
            	<th>&nbsp;</th>
            </tr>
        </thead>
    	<tbody>
            <tr>
                <td width="25%"><b><%=GF_Traducir("Humedad")%>:</b></td>
                <td width="25%"><%=g_rsCaladas("vlhumedad")%></td>
                <td width="25%"><b><%=GF_Traducir("Proteina")%>:</b></td>
                <td width="25%"><%=g_rsCaladas("vlproteina")%></td>                
            </tr>
            <tr>                
                <td><b><%=GF_Traducir("Merma")%>.:</b></td>
                <td><%=g_rsCaladas("pcMerma")%></td>
                <td><b><%=GF_Traducir("Tipo Calada")%>.:</b></td>
                <td><%=g_rsCaladas("ictipocalada") %></td>
            </tr>
            <tr>                
                <td><b><%=GF_Traducir("Camara")%>.:</b></td>
                <td><%=g_rsCaladas("iccamara")%></td>
                <td><b><%=GF_Traducir("Humedimetro")%>.:</b></td>
                <td><%=g_rsCaladas("ichumedimetro") %></td>
            </tr>
            <tr>                
                <td><b><%=GF_Traducir("Sticker")%>.:</b></td>
                <td><%=g_rsCaladas("NUBARRAS") %></td>
                <td><b><%=GF_Traducir("Aceptacion")%>.:</b></td>
                <td><%=g_rsCaladas("dsaceptacion") %></td>
            </tr>            
            <% if (g_rsCaladas("CDMOTIVORECHAZO") > 0) then %>
				<tr>
					<td><b><%=GF_Traducir("Motivo Rechazo")%>.:</b></td>
					<td colspan="3"><%=getDSMotivoRechazo(g_rsCaladas("CDMOTIVORECHAZO"))%></td>
				</tr>	
            <% end if %>
            <tr>			
                <td><b><%=GF_TRADUCIR("Observaciones")%>.:</b></td>
                <% g_dsObservaciones = "-"
                   if (Len(Trim(g_rsCaladas("DSOBSERVACIONES"))) > 0) then g_dsObservaciones = Trim(g_rsCaladas("DSOBSERVACIONES")) %>
                <td colspan="3"><%= g_dsObservaciones%></td>
            </tr>
        </tbody>
    </table>		    
<%else	
	Response.Write "La calada "& g_sqCalada &" no disponible en estos momentos"
end if%>
<BR>
<% Call getDatosHumedimetro(g_sqCalada, g_dtContable, g_idCamion, g_rsHumedimetro)%>
    <table class="datagrid" align="center" width="80%" id="tblHumedimetro_<%=g_sqCalada%>">
    	<thead>
        	<tr>
        		<th rowspan="2"><%=GF_Traducir("Muestra")%></th>
            	<th rowspan="2"><%=GF_Traducir("Humedad")%></th>
            	<th rowspan="2"><%=GF_Traducir("Peso Hectolitrico")%></th>
            	<th rowspan="2"><%=GF_Traducir("Temperatura")%></th>
            </tr>
        </thead>
        <tbody>
        <% if not g_rsHumedimetro.eof then 
				getPromediosYMaximos g_rsHumedimetro, g_fltPromedioHumedad, g_fltPromedioPesoHect, g_fltPromedioTemp, g_fltMaxHumedad, g_fltMinPesoHect, g_fltMaxTemp 
				while not g_rsHumedimetro.eof %>
					<tr>
						<td align="center"><%=g_rsHumedimetro("sqmuestra")%></td>
						<td align="center" <%if CDbl(g_fltMaxHumedad) = CDbl(g_rsHumedimetro("vlhumedad")) then response.write "bgcolor='#98FB98'"%>><%=g_rsHumedimetro("vlhumedad")%></td>
						<td align="center" <%if CDbl(g_fltMinPesoHect) = CDbl(g_rsHumedimetro("vlpeso")) then response.write "bgcolor='#98FB98'"%>><%=g_rsHumedimetro("vlpeso")%></td>
						<td align="center" <%if CDbl(g_fltMaxTemp) = CDbl(g_rsHumedimetro("vltemperatura")) then response.write "bgcolor='#98FB98'"%>><%=g_rsHumedimetro("vltemperatura")%></td>
					</tr>
				<%	g_rsHumedimetro.movenext
				wend  %>
				<tr>				
					<td ><%=GF_Traducir("Promedio")%></td>
					<td align="center" ><%=formatnumber(g_fltPromedioHumedad, 2)%></td>
					<td align="center" ><%=formatnumber(g_fltPromedioPesoHect, 2)%></td>
					<td align="center" ><%=formatnumber(g_fltPromedioTemp, 2)%></td>
				</tr>	
		<%else%>
			<tr><td colspan="4" align="center"><%=GF_Traducir("Sin datos de humedimetro")%></td></tr>
		<%end if%>
		</tbody>
	</table>
	
	<table class="datagrid" align="center" width="80%" >
		<thead>
        	<tr>
        		<th rowspan="2"><%=GF_Traducir("Rubro")%></th>
            	<th rowspan="2"><%=GF_Traducir("Valor")%></th>
            	<th rowspan="2"><%=GF_Traducir("Merma")%></th>
            	<th rowspan="2"><%=GF_Traducir("Autom./Manual")%></th>
            </tr>
        </thead>
        <tbody>
		<% Set g_rsResultados = getResultadosCalada(g_dtContable, g_idCamion, g_sqCalada)
		   if (not g_rsResultados.eof) then 						
				while (not g_rsResultados.eof) %>
		   		<tr>
		   			<td><% =g_rsResultados("DSRUBRO") %></td>
		   			<td align="center"><% =g_rsResultados("VLBONREBAJA") %></td>
		   			<td align="center"><% =g_rsResultados("VLMERMA") %>%</td>
		   			<td align="center">
		   			<%  if (g_rsResultados("ICINGMANUAL") = "N") then
	   						response.Write "Automatico"
		   				else
		   					Response.Write "Manual"
		   				end if			   				
		   			%>
		   			</td>			   		
		   		</tr>
		   		<%
		   			g_rsResultados.MoveNext()
		   		wend			   	
		   else	%>
				<tr><td colspan="5" align="center" class="reg_header_navdos"><%=GF_Traducir("Sin datos de Rubros")%></td></tr>			   
		<%  end if	%>
		</tbody>
	</table>
<BR>	
<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getDSMotivoRechazo(cdMotivo)
	Dim strSQL, rtrn
	strSQL = " SELECT DSMOTIVO FROM MOTIVOS WHERE CDMOTIVO = " & cdMotivo
	Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)
	if (not rs.EoF) Then rtrn = rs("DSMOTIVO")
	getDSMotivoRechazo = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getCaladaCamionBySqCalada(p_dtContable, p_idCamion,p_sqCalada, p_Pto)
	dim strSql, diaHoy, rs

	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)
	auxDTcontable = Left(p_dtContable,4) &"-"& Mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)

	strSql = "Select ictipocalada, vlhumedad, vlproteina, dsaceptacion, pcmerma, iccamara, ichumedimetro, dsobservaciones,CDMOTIVORECHAZO,COALESCE(AC.CDACEPTACION,0) CDACEPTACION,NUBARRAS"&_
			 " from"&_
			 "		(Select '" & diaHoy & "' DTCONTABLE, IDCAMION, SQCALADA, ICTIPOCALADA, VLHUMEDAD, VLPROTEINA, CDACEPTACION, CDRUBROPPAL, CDGRADO, ICCONDICIONFABRICA, CDUSERNAME, DTCALADA, DSOBSERVACIONES, NUBARRAS, CDMOTIVORECHAZO, PCMERMA, ICCAMARA, ICHUMEDIMETRO, DSOBSHUMEDIMETRO, CDSUPERVISOR, NUBALDE, CDFUERASTD, IDCONTRATO from caladadecamiones where idCamion = '" & p_idCamion & "' and sqcalada ="& p_sqCalada &_
			 "     union"&_
			 "      Select * from hcaladadecamiones where DTCONTABLE='" & auxDTcontable & "' and idCamion = '" & p_idCamion & "' and sqcalada ="& p_sqCalada &") CC"&_
			 " inner join aceptacioncalidad AC on CC.CDACEPTACION=AC.CDACEPTACION"&_
			 " where DTCONTABLE='" & auxDTcontable & "'"
	
	Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSql)
	Set getCaladaCamionBySqCalada = rs
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Función que devuelve el resultado de cada uno de los rubros cargados en la calada para cada una de sus secuencias.
Function getResultadosCalada(p_dtContable, p_idCamion, p_sqCalada) 

    Dim strSql, diaHoy, rs,auxDTcontable
	
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2) 
	auxDTcontable = Left(p_dtContable,4) &"-"& Mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)

    strSQL= "Select * from " & _    
            "((Select '" & diaHoy & "' DTCONTABLE, IDCAMION, SQCALADA, A.CDRUBRO, DSRUBRO, VLBONREBAJA, CDSUPERVISOR,	VLMERMA,	ICINGMANUAL " & _
            "from RUBROSVISTEOCAMIONES A " & _
	        "INNER JOIN RUBROS B on A.CDRUBRO=B.CDRUBRO " & _	        
            "where SQCALADA=" & p_sqCalada & " and IDCAMION='" & p_idCamion & "'" & _
            ") UNION (" & _
            "Select DTCONTABLE, IDCAMION, SQCALADA, A.CDRUBRO, DSRUBRO, VLBONREBAJA, CDSUPERVISOR,	VLMERMA,	ICINGMANUAL " & _
            "from HRUBROSVISTEOCAMIONES A " & _
	        "INNER JOIN RUBROS B on A.CDRUBRO=B.CDRUBRO " & _	        
            "where DTCONTABLE='" & auxDTcontable & "' and SQCALADA=" & p_sqCalada & " and IDCAMION='" & p_idCamion & "')) TABLA " & _            
            "Order by SQCALADA DESC, CDRUBRO"            
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSql)
    
    Set getResultadosCalada = rs
    
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'JAS - CORREGIDO EL TEMA DE LOS JOINS GENERALES y AGREGADO LA CONSULTA PARA CAMIONES HISTORICOS
Function getDatosHumedimetro(p_sqCalada, p_dtContable, p_idCamion, p_rsHumedimetro)
	Dim strSql, diaHoy,auxDTcontable
	
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2) 
	auxDTcontable = Left(p_dtContable,4) &"-"& Mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)

	strSql = "Select * from"
	strSql = strSql & " (Select '" & diaHoy & "' DTCONTABLE, IDCAMION, SQCALADA, SQMUESTRA, VLHUMEDAD, VLTEMPERATURA, VLPESO from muestrashumedcamiones where idCamion = '" & p_idCamion & "' and sqcalada = " & p_sqCalada
	strSql = strSql & " union"
	strSql = strSql & " select * from hmuestrashumedcamiones where dtContable='" & auxDTcontable & "' and idCamion = '" & p_idCamion & "' and sqcalada = " & p_sqCalada & ") T"
	strSql = strSql & " where DTCONTABLE='" & auxDTcontable & "'"
	
	Call GF_BD_Puertos(g_strPuerto, p_rsHumedimetro, "OPEN", strSql)
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
sub	getPromediosYMaximos(p_rsHumedimetro, byref p_fltPromedioHumedad, byref p_fltPromedioPesoHect, byref p_fltPromedioTemp, byref p_fltMaxHumedad, byref p_fltMinPesoHect, byref p_fltMaxTemp)
'calcula promedios y maximos/minimos de los parametros de humedimetro
dim l_intContador
dim l_fltSumaHumedad, l_fltSumaPesoHect, l_fltSumaTemp	
	p_fltMaxHumedad = 0
	p_fltMinPesoHect = 10000000 
	p_fltMaxTemp = 0
	while not p_rsHumedimetro.eof
		l_intContador = l_intContador + 1
		l_fltSumaHumedad = l_fltSumaHumedad + CDbl(p_rsHumedimetro("vlhumedad"))
		l_fltSumaPesoHect = l_fltSumaPesoHect + CDbl(p_rsHumedimetro("vlpeso"))
		l_fltSumaTemp = l_fltSumaTemp + CDbl(p_rsHumedimetro("vltemperatura"))
		if CDbl(p_fltMaxHumedad) < Cdbl(p_rsHumedimetro("vlhumedad")) then p_fltMaxHumedad = p_rsHumedimetro("vlhumedad")
		if CDbl( p_fltMinPesoHect) > CDbl(p_rsHumedimetro("vlpeso")) then p_fltMinPesoHect = p_rsHumedimetro("vlpeso")
		if CDbl(p_fltMaxTemp) < CDbl(p_rsHumedimetro("vltemperatura")) then p_fltMaxTemp = p_rsHumedimetro("vltemperatura")
		p_rsHumedimetro.movenext
	wend
	p_fltPromedioHumedad = l_fltSumaHumedad/l_intContador
	p_fltPromedioPesoHect = l_fltSumaPesoHect/l_intContador
	p_fltPromedioTemp = l_fltSumaTemp/l_intContador	
	g_rsHumedimetro.movefirst
end sub
%>