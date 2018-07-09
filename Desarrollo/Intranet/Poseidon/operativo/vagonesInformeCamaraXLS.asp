<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="includes/procedimientosVIC.asp"-->
<%
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
Call GP_CONFIGURARMOMENTOS()

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
pTipo = GF_PARAMETROS7("pTipo", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)
totalVagones = 0
totalKilosNetos = 0
call getParametros()
strSQL = generarSQL()
call GF_BD_Puertos(pto, rsGeneral, "OPEN", strSQL)
Call GF_createXLS("vagones")
%>
<html>
<head>
<title><%=GF_TRADUCIR("Puertos - Visteo Calada")%></title>
</script>
</head>
<table border="1" cellpadding="0" cellspacing="0" width="60%">
      <tr>
	    <td bgcolor="#517B4A" align="center" colspan="16">
			<b>
				<font color="white" style="font-family:courier;" size="5">
					Informe Calada Vagones
				</font>
			</b>
		</td>
      </tr>
      <tr>
	    <td align="left" colspan="16">
			<%if myTurnoDesde <> "" then %>
				<b><font style="font-family:courier;" size="1">Turno Desde........:</font></b>
				<font style="font-family:courier;" size="1"><%=myTurnoDesde%></font>
				<br>
			<%end if%>	    
			<%if myturnoHasta <> "" then %>
				<b><font style="font-family:courier;" size="1">Turno Hasta........:</font></b>
				<font style="font-family:courier;" size="1"><%=myturnoHasta%></font>
				<br>
			<%end if%>	    
			<%if myIncluir <> "" then %>
				<b><font style="font-family:courier;" size="1">Incluir............:</font></b>
				<font style="font-family:courier;" size="1">
					<%select case myIncluir
						case "T"
							Response.Write "Todos"
						case "C"
							Response.Write "Con Analisis"	
						case "S"	
							Response.Write "Sin Analisis"
					  end select		
					%></font>
				<br>
			<%end if%>	    
			<%if myOperativo <> "" then %>
				<b><font style="font-family:courier;" size="1">Operativo..........:</font></b>
				<font style="font-family:courier;" size="1"><%=myOperativo%></font>
				<br>
			<%end if%>	    
			<%if myCdAceptacion <> "" then %>
				<b><font style="font-family:courier;" size="1">Aceptacion.........:</font></b>
				<font style="font-family:courier;" size="1"><%=myCdAceptacion%></font>
				<br>
			<%end if%>	    
			<%if myStickerDesde <> "" then %>
				<b><font style="font-family:courier;" size="1">Sticker Desde......:</font></b>
				<font style="font-family:courier;" size="1"><%=myStickerDesde%></font>
				<br>
			<%end if%>	    
			<%if myStickerHasta <> "" then %>
				<b><font style="font-family:courier;" size="1">Sticker Hasta......:</font></b>
				<font style="font-family:courier;" size="1"><%=myStickerHasta%></font>
				<br>
			<%end if%>	    
			<%if myFecContableDesde <> "" then %>
				<b><font style="font-family:courier;" size="1">Fecha Desde........:</font></b>
				<font style="font-family:courier;" size="1"><%=GF_FN2DTE(myFecContableDesde)%></font>
				<br>
			<%end if%>	    
			<%if myFecContableHasta <> "" then %>
				<b><font style="font-family:courier;" size="1">Fecha Hasta........:</font></b>
				<font style="font-family:courier;" size="1"><%=GF_FN2DTE(myFecContableHasta)%></font>
				<br>
			<%end if%>	    
			<%if myCdCoordinador <> "" then %>
				<b><font style="font-family:courier;" size="1">Coordinador........:</font></b>
				<font style="font-family:courier;" size="1"><%=myDsCoordinador%></font>
				<br>
			<%end if%>	    
			<%if myCdCoordinado <> "" then %>
				<b><font style="font-family:courier;" size="1">Coordinado.........:</font></b>
				<font style="font-family:courier;" size="1"><%=myDsCoordinado%></font>
				<br>
			<%end if%>	    
			<%if myCdProducto <> "" then %>
				<b><font style="font-family:courier;" size="1">Producto...........:</font></b>
				<font style="font-family:courier;" size="1"><%=myCdProducto%></font>
				<br>
			<%end if%>	    
			<%if myCdCorredor <> "" then %>
				<b><font style="font-family:courier;" size="1">Corredor...........:</font></b>
				<font style="font-family:courier;" size="1"><%=myDsCorredor%></font>
				<br>
			<%end if%>	    
			<%if myCdVendedor <> "" then %>
				<b><font style="font-family:courier;" size="1">Vendedor...........:</font></b>
				<font style="font-family:courier;" size="1"><%=myDsVendedor%></font>
				<br>
			<%end if%>	    
			<%if myCdEntregador <> "" then %>
				<b><font style="font-family:courier;" size="1">Entregador.........:</font></b>
				<font style="font-family:courier;" size="1"><%=myDsEntregador%></font>
				<br>
			<%end if%>	    
		</td>
	</tr>
		<%
		if rsGeneral.eof then 
			Response.Write "<tr><td colspan=15 align=center>No se encontraron camiones</td></tr>"
		else	
		%>		 																								
			<TR class="reg_Header_nav">
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Fecha")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Coordinador")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Coordinado")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Producto")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Corredor")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Entregador")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Vendedor")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Localidad")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Nro Vagon")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Carta Porte")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Merma")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Kilos Netos")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Barras")%> </TD>
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Grado")%> </TD>				
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Aceptacion")%> </TD>								
				<TD bgcolor='#ffeecd' align="center">	<%=GF_Traducir("Hora")%> </TD>								
			</TR>
			<%
			CargarGrados
			while not rsGeneral.eof
				cont = cont + 1
				myKilosNetos = Clng(rsGeneral("Bruto"))-Clng(rsGeneral("Tara"))
				myGradoParticular =  VerGrado (pto, rsGeneral("cdProducto"), rsGeneral("cdAceptacion"), rsGeneral("Barras"), rsGeneral("fecha"),myIncluir)
				If myGradoParticular <> "XXX" Then
					totalVagones = totalVagones + 1	
					totalNetoAcumulado = totalNetoAcumulado + myKilosNetos
					call Sumar_Totales (myKilosNetos, totalVagones)
					call SumarResumen (myGradoParticular,myKilosNetos)
				End If			
				if cont mod 2 then
					color = "#ffffff"
				else
					color =	"#dcdcdc"
				end if
					%>
					<TR>
						<TD bgcolor="<%=color%>" align="center"> <%=GF_FN2DTE(Left(rsGeneral("DTPESADA"),8))%></TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Coordinador")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Coordinado")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Producto")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Corredor")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Entregador")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Vendedor")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Localidad")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("NoVagon")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=cstr(rsGeneral("CartaPorte"))%> &nbsp;</TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Merma")%> </TD>
						<TD bgcolor="<%=color%>" align="right">		<%=cdbl(myKilosNetos)%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Barras")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=GF_Traducir("XXX")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=rsGeneral("Aceptacion")%> </TD>
						<TD bgcolor="<%=color%>" align="center">	<%=Right(GF_FN2DTE(rsGeneral("DTPESADA")),8)%></TD>
					</TR>
					<%
				rsGeneral.movenext
			wend
			%>
					<tr>
						<td colspan=16>&nbsp;</td>
					</tr>	
					<TR>
						<TD align="left" colspan=11>
							<b>
								<%=GF_Traducir("TOTAL VAGONES") & ": " & totalVagones%> 
							</b>
						</td>	
						<TD align="right" colspan="1"><b><%=cdbl(totalNetoAcumulado)%> </b></TD>
						<TD align="left" colspan="4">&nbsp;</TD>
					</TR>
			
			
					<tr>
						<td colspan=16>&nbsp;</td>
					</tr>	
					<tr>
						<td colspan=16><B><%=GF_Traducir("RESUMEN")%></B></td>
					</tr>	
					<tr>
						<td colspan=1>&nbsp;</td>
						<!--<td colspan=2>-->
									<!--<td>&nbsp;</td>-->
									<td bgcolor='#ffeecd' align="center" colspan="2" rowspan="2"><b><%=GF_Traducir("Items")%></b></td>
									<td bgcolor='#ffeecd' align="center" colspan="2"><B><%=GF_Traducir("VAGONES")%></B></td>
									<td bgcolor='#ffeecd' align="center" colspan="2"><B><%=GF_Traducir("KILOGRAMOS")%></B></td>				
								</tr>	
								<tr class="reg_Header_nav">
									<td colspan=1>&nbsp;</td>
									<td bgcolor='#ffeecd' align="center"><b><%=GF_Traducir("Cantidad")%></b></td>
									<td bgcolor='#ffeecd' align="center"><b><%=GF_Traducir("%")%></b></td>
									<td bgcolor='#ffeecd' align="center"><b><%=GF_Traducir("Cantidad")%></b></td>
									<td bgcolor='#ffeecd' align="center"><b><%=GF_Traducir("%")%></b></td>
								</tr>	
								<%
								call SumarPorcentajeResumen(totalVagones, totalNetoAcumulado, totalVagonesRegistrados, totalKilosNetosRegistrados)
								For i = 0 To 13
									Response.Write "<tr>"
										Response.Write "<td colspan=1></td>"
										Response.Write "<td colspan=2>" & myGrado(i,1) & "</td>"
										Response.Write "<td align='right' colspan=1>" & myGrado(i,2) & "</td>"
										Response.Write "<td align='right' colspan=1>" & myGrado(i,3) & "</td>"
										Response.Write "<td align='right' colspan=1>" & cdbl(myGrado(i,4)) & "</td>"
										Response.Write "<td align='right' colspan=1>" & myGrado(i,5) & "</td>"
									Response.Write "</tr>"
								Next
								%>
								<tr>
									<td colspan=1>&nbsp;</td>
									<td bgcolor=LightGrey colspan=2><B><%=GF_Traducir("TOTAL")%></B></td>
									<td bgcolor=LightGrey align="right"><b><%=totalVagonesRegistrados%></b></td>
									<td bgcolor=LightGrey align="right"><b><%=GF_EDIT_DECIMALS(10000,2)%></b></td>
									<td bgcolor=LightGrey align="right"><b><%=totalKilosNetosRegistrados%></b></td>
									<td bgcolor=LightGrey align="right"><b><%=GF_EDIT_DECIMALS(10000,2)%></b></td>		
								</tr>												
							<!--</table>
						</td>
					</tr>				-->
			<%	
	end if
	%>
      <tr>
	    <td bgcolor="#ffeecd" align="center" colspan="16">
			<b>
				<font style="font-family:courier;" size="1">
					Fin Reporte
				</font>
			</b>
		</td>
      </tr>		
   </table>
</body>
</html>
