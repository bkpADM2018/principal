<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<%

Dim cdAviso, cdPuerto, currProducto, rsDatos, rsCargas, kilosSinCosecha, kilosDS, fechaDS, cantRegDatos

Function getEmbarquesDatos(pCdPuerto, pCdAviso)
	Dim strSQL
	
	strSQL = " SELECT ED.CDPRODUCTO," &_ 
			 "		  CDCOSECHA,    " &_
			 "		  KILOS,		" &_
			 "		  PERMISO,      " &_
			 "		  P.DSPRODUCTO  " &_			 
			 "  FROM EMBARQUESDATOS ED " &_
			 "	  LEFT JOIN PRODUCTOS P ON ED.CDPRODUCTO = P.CDPRODUCTO " &_
			 "  WHERE CDAVISO=" & pCdAviso &_
			 "  ORDER BY ED.CDPRODUCTO"
	Call GF_BD_Puertos (pCdPuerto, rs, "OPEN",strSQL)
	Set getEmbarquesDatos = rs
	
End Function
'----------------------------------------------------------------------------------------------------
'Funcion que obtiene las cargas registradas por balanza en un aviso de embaque determinado.
'Totaliza las cargas registradas por producto.
Function getCargasEmbarqueBalanza(pCdPuerto, pCdAviso)
    
    Dim strSQL, rs
    
    strSQL= "Select CE.CDPRODUCTO, DSPRODUCTO, SUM(VLKILOS) KILOS, MAX(DTCARGA) FECHABZA, DS.KILOSDS, IDDRAFT,FECHADRAFT" &_
            " from CARGASEMBARQUE CE " &_
			" left join (select Sum(TOTALDRAFT) as KILOSDS ,CDPRODUCTO ,IDDRAFT, FECHADRAFT from TBLEMBARQUESDRAFTSURVEY "&_
            "            where CDAVISO = "&pCdAviso&" and CDESTADO IN("&ESTADO_ACTIVO&","&ESTADO_AUTORIZADO&") group by IDDRAFT,CDPRODUCTO,FECHADRAFT) DS "&_
            "       on DS.CDPRODUCTO = CE.CDPRODUCTO 	"&_
            "	inner join PRODUCTOS P on CE.CDPRODUCTO = P.CDPRODUCTO " &_
            " where CDAVISO=" & pCdAviso &_
            " group by CE.CDAVISO,CE.CDPRODUCTO, DSPRODUCTO ,DS.KILOSDS,IDDRAFT,FECHADRAFT " &_
            " order by CE.CDPRODUCTO"            
    Call GF_BD_Puertos (pCdPuerto, rs, "OPEN",strSQL)
    Set getCargasEmbarqueBalanza = rs
    
End Function
'----------------------------------------------------------------------------------------------------
'Funcion que devuelve true si un determinado Codigo de aviso tiene draft Survey activo, se la utiliza para saber si 
'es necesario mostrar la columna del Draft survey, debido a que un Cdaviso puede tener varios productos y no todos tiene Draft
Function getDraftSurveyByCdAviso(pCdPuerto ,pCdAviso)
	Dim strSQL, rs, rtrn		
	rtrn = false
    strSQL = " Select * from TBLEMBARQUESDRAFTSURVEY where CDAVISO = "&pCdAviso&" and CDESTADO IN("&ESTADO_ACTIVO& ","&ESTADO_AUTORIZADO&")"     
    Call GF_BD_Puertos (pCdPuerto, rs, "OPEN",strSQL)
    if not rs.Eof then  rtrn = true
    getDraftSurveyByCdAviso = rtrn
End Function
'----------------------------------------------------------------------------------------------------
cdAviso = GF_Parametros7("cdAviso",0,6)
cdPuerto = GF_Parametros7("pto","",6)
fechaDS = GF_Parametros7("issuedate","",6)
if (fechaDS = "") then fechaDS = Left(session("MmtoSistema"),8)
Set rsCargas = getCargasEmbarqueBalanza(cdPuerto, cdAviso)
Set rsDatos = getEmbarquesDatos(cdPuerto, cdAviso)
cantRegDatos = rsDatos.recordcount
flagDraft = false
if (getDraftSurveyByCdAviso(cdPuerto,cdAviso)) then flagDraft = true
%>
<table border="0" width="100%" align="left" >							
    <tr>
	    <td colspan=3 align='center'> 
		<% if not rsCargas.Eof then %>
		    <table border=0 width='90%'>
			    <tr class='reg_header_warning'>
			        <td align='center'><font size='2'><% =GF_Traducir("Producto") %></font></td>
				    <td align='center'><font size='2'><% =GF_Traducir("Cosecha") %></font></td>
				    <td align='center'><font size='2'><% =GF_Traducir("Kilos Bza.") %></font></td>
				    <% if flagDraft then%>
				    <td align='center'><font size='2'><% =GF_Traducir("Kilos Draft") %></font></td>
				    <% end if %>
				    <td align='center'><font size='2'><% =GF_Traducir("Permiso") %></font></td>
				    <td align='center'></td>
			    </tr>
			    <% while (not rsCargas.eof) 
			        'Tomo el total de kilos que indicó la balanza y el producto a procesar.
			        kilosSinCosecha = CDbl(rsCargas("KILOS"))
			        currProducto = CInt(rsCargas("CDPRODUCTO"))			        
			        if not isNull(rsCargas("IDDRAFT")) then						
				        kilosTotalDraft = CDbl(rsCargas("KILOSDS"))
				        idDraft = rsCargas("IDDRAFT")
				        kilosParcialDraft = 0
				    end if
				    kilosCosecha = 0
				    flagKilosDraft = true
			        while (not rsDatos.eof)						
			            if (CInt(rsDatos("CDPRODUCTO")) = currProducto) then
			            'Muetro los kilos asignados a una determinada cosecha del producto en proceso.
			            kilosCosecha = Cdbl(rsDatos("KILOS")) + kilosCosecha
					%>
					<tr class='reg_header_2'>
					    <td align='left'><font size='2'><% =rsDatos("CDPRODUCTO") & " - " & rsDatos("DSPRODUCTO") %></font></td>
					    <td align='center'><font size='2'><% =rsDatos("CDCOSECHA") %></td>
					    <td align='right'><font size='2'><% =GF_EDIT_DECIMALS(CDbl(rsDatos("KILOS"))*100,2) %></font></td>					    
						<% if(idDraft > 0)then%>
						<td align='right'><%
							  if(flagKilosDraft)then
								  if(kilosTotalDraft >= kilosCosecha) then 
									  kilosParcialDraft = Cdbl(rsDatos("KILOS"))
								  else	
									  kilosParcialDraft = Cdbl(rsDatos("KILOS")) - (kilosCosecha - kilosTotalDraft)
									  flagKilosDraft = false
								  end if	%>								
								  <font size='2'><% =GF_EDIT_DECIMALS(CDbl(kilosParcialDraft)*100,2) %></font>		
						<%	  end if	%>	
						</td>
						<%	else if flagDraft then %>
								<td align='right'></td>
							<%	end if %>
						<%	end if %>
					    <td align='center'><font size='2'><% =rsDatos("PERMISO") %></font></td>
					    <td align='center'><font><span style="cursor:pointer;" onclick="cargaEdit(<% =cdAviso %>, <% =rsDatos("CDPRODUCTO") %>, '<%=rsDatos("DSPRODUCTO")%>', <%=rsDatos("CDCOSECHA") %>, <%=rsDatos("KILOS") %>, '<% =rsDatos("PERMISO") %>')" color="blue"><u><% =GF_Traducir("Editar") %></u></span></font> |
					                       <font><span style="cursor:pointer;" onclick="cargaDel(<% =cdAviso %>, <% =rsDatos("CDPRODUCTO") %>, <%=rsDatos("CDCOSECHA") %>, <%=rsDatos("KILOS") %>, '<% =rsDatos("PERMISO") %>')" color="blue"><u><% =GF_Traducir("Quitar") %></u></span></font>
					    </td>
					    
					</tr> 		
					<%     kilosSinCosecha = kilosSinCosecha - CDbl(rsDatos("KILOS"))			           
			            end if			            
			            rsDatos.MoveNext() 
			        wend 			        
			        'Muestro el saldo de los kilos sin cosecha. 
			        if (kilosSinCosecha > 0) then    %>			        
					<tr class='reg_header_2'>
				        <td align='left'><font size='2'><% =rsCargas("CDPRODUCTO") & " - " & rsCargas("DSPRODUCTO") %></font></td>
						<td align='center'></td>
						<td align='right'><font size='2'><% =GF_EDIT_DECIMALS(kilosSinCosecha*100,2) %></font></td>
						<% if(idDraft > 0)then%>
							<td align='right'>
								<%if(flagKilosDraft) then%>									
									<font size='2'><% if(CDbl(kilosTotalDraft - kilosCosecha) > 0) then Response.Write GF_EDIT_DECIMALS(CDbl(kilosTotalDraft - kilosCosecha)*100,2)%></font>
									<input type="hidden" id="kilosDraftSinCosecha_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" value="<%= CDbl(kilosTotalDraft - kilosCosecha)%>">
								<%else%>
									<input type="hidden" id="kilosDraftSinCosecha_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" value="0">
								<%end if%>
							</td>	
						<%	else if flagDraft then %>
								<td align='right'></td>
							<%	end if %>
						<%	end if %>
						<td></td>
						<td align='center'><font><span style="cursor:pointer;" onclick="cargaEdit(<% =cdAviso %>, <% =rsCargas("CDPRODUCTO") %>,  '<%=rsCargas("DSPRODUCTO")%>', '', <%=kilosSinCosecha %>, '')" color="blue"><u><% =GF_Traducir("Asignar Cosecha") %></u></span></font></td>						
					</tr>
				<%	end if				
					'Totalizo los kilos de cada Producto(con y sin cosecha), aca se puede aplicar el Draft Survey  %>										
					<tr class='reg_header_total'>						
						<td align='left'><font size='2'><% =rsCargas("CDPRODUCTO") & " - " & rsCargas("DSPRODUCTO") %></font></td>
						<td align='center'><font size='2'><% =GF_TRADUCIR("Total") %></font></td>
						<td align='right'><font size='2'><% =GF_EDIT_DECIMALS(CDbl(rsCargas("KILOS"))*100,2) %></font><input type="hidden" id="kilosBza_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" name="kilosBza_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" value="<%=rsCargas("KILOS")%>"></td>
						<% if(idDraft > 0)then%>
						<td align='right'>
							<font size='2'><% if (not isNull(rsCargas("IDDRAFT"))) then Response.Write GF_EDIT_DECIMALS(Cdbl(rsCargas("KILOSDS"))*100,2) %></font>							
						</td>
						<%	else if flagDraft then %>
								<td align='right'></td>
							<%	end if %>
						<%	end if %>
						<input type="hidden" id="kilosDraft_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" name="kilosDraft_<%=cdAviso%>_<%=rsCargas("CDPRODUCTO")%>" value="<%=kilosTotalDraft%>">
						<td align='right'></td>		
						<td align='center'>				
						<% if (idDraft > 0) then %>
							<font><span style="cursor:pointer;" onclick="cargaEditDraft(<%=cdAviso%>,<%=idDraft%>,<%=kilosTotalDraft%>,<%=rsCargas("CDPRODUCTO")%>,'<%=rsCargas("DSPRODUCTO")%>',<%= rsCargas("FECHADRAFT")%>,'<%= GF_FN2DTE(rsCargas("FECHADRAFT"))%>')" color="blue"><u><% =GF_Traducir("Editar Draft") %></u></span></font> |
	                        <font><span style="cursor:pointer;" onclick="cargaDelDraft(<%=cdAviso%>,<%=idDraft%>)" color="blue"><u><% =GF_Traducir("Quitar Draft") %></u></span></font>
						<% else	  %>
							<a style="cursor:pointer;" onclick="newDraft(<%=cdAviso%>,<%=rsCargas("CDPRODUCTO")%>,'<%=rsCargas("DSPRODUCTO")%>')"><font color="black"><u>Agregar Draft Survey</u></font></a>
						<% end if %>
						</td>
					</tr>
					<input type="hidden" name="kilosCosecha_<%=cdAviso%>_<%=currProducto%>" id="kilosCosecha_<%=cdAviso%>_<%=currProducto%>" value="<%=kilosCosecha %>">
					<input type="hidden" name="kilosSinCosecha_<%=cdAviso%>_<%=currProducto%>" id="kilosSinCosecha_<%=cdAviso%>_<%=currProducto%>" value="<%=kilosSinCosecha%>">
					<input type="hidden" id="fechaBza_<%=cdAviso%>_<%=currProducto%>" name="fechaBza_<%=cdAviso%>_<%=currProducto%>" value="<%=GF_STANDARIZAR_FECHA_RTRN(rsCargas("FECHABZA"))%>">
			    <%  kilosVal = 0  
					rsCargas.MoveNext()
				    ' Si hay mas Cargas en balanza, vuelvo a iniciar los Embarques ya que hay mas de 1 producto	
  				    if not rsCargas.EoF then 
  						 if cantRegDatos > 0 then rsDatos.MoveFirst() 
  					end if	
			       wend
			    %>
		    </table>	
		 <%end if%>   	
	    </td> 
	</tr>
	<tr>
		<td width="5%">&nbsp;</td>
		<td width="20%"><span id="strProducto_<%=cdAviso%>"><%=GF_Traducir("Producto:")%></span></td>
		<td >
			<span id="spanProducto_<%=cdAviso%>"></span>
			<input type="hidden" id="cdProducto_<%=cdAviso%>" name="cdProducto_<%=cdAviso%>">
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><span id="strCosecha_<%=cdAviso%>"><%=GF_Traducir("Cosecha:")%></span></td>
		<td>
			<input type="text" maxlength="8" size="8" name="cosecha_<%=cdAviso%>" id="cosecha_<%=cdAviso%>" value="" onkeypress="return controlIngreso(this, event, 'N')">						
			<div style="visibility:hidden;display:inline-block;float:left;position:absolute;margin-right:5px;" id="issuedateDiv_<%=cdAviso%>" name="issuedateDiv_<%=cdAviso%>" class="labelStyle"><% =GF_FN2DTE(fechaDS) %></div>
			<a style="visibility:hidden;" id="issuedateLink_<%=cdAviso%>" name="issuedateLink_<%=cdAviso%>" href="javascript:MostrarCalendario('imgLimite_<%=cdAviso%>', SeleccionarCalEmision)"><img id="imgLimite_<%=cdAviso%>" src="../images/DATE.gif"></a>			
			<input type="hidden" id="issuedate_<%=cdAviso%>" name="issuedate_<%=cdAviso%>" value="<% =fechaDS %>" />
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><%=GF_Traducir("Kilos:")%></td>
		<td >
			<input type="text" name="kilos_<%=cdAviso%>" id="kilos_<%=cdAviso%>" value="" onkeypress="return controlIngreso(this, event, 'N')">			
		</td>
	</tr>
	<tr id="trkgToepfer_<%=cdAviso%>" style="visibility:hidden; position:absolute;">
		<td>&nbsp;</td>
		<td><%=GF_Traducir("Kg Bza. Cargador Toepfer:")%></td>
		<td >
			<input type="text" name="kgToepfer_<%=cdAviso%>" id="kgToepfer_<%=cdAviso%>" value="" onkeypress="return controlIngreso(this, event, 'N')">			
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><span id="strPermiso_<%=cdAviso%>"><%=GF_Traducir("Permiso:")%></span></td>		
		<td >
			<input type="text" name="permiso_<%=cdAviso%>" id="permiso_<%=cdAviso%>" value="">			
			<div style="visibility:hidden;" id="dsFile_<%=cdAviso%>"></div>
			<input type="hidden" id="dsFilePath_<%=cdAviso%>" name="dsFilePath_<%=cdAviso%>" value="">	
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td align="right" colspan="2">			
			<a style="cursor:pointer;" id="actulizar_<%=cdAviso%>" ><font color="blue"><u>Actualizar</u> |</font></a>
			<a style="cursor:pointer;" id="cancelar_<%=cdAviso%>" onclick="cancelDetails(<%=cdAviso%>)"><font color="blue"><u>Cancelar</u></font></a>			
			<span style="position:absolute;visibility:hidden;" id="SPAN_<%=cdAviso%>">
				<img src="images/loading_small_black.gif">
			</span>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	<!--Imprimir linea de errores -->				
	<tr>
		<td>&nbsp;</td>
		<td colspan="2">
			<div id="MSG_<%=cdAviso%>"></div>
		</td>	
	</tr>	
	<!--Fin imprimir linea de errores -->	
</table>