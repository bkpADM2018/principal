<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="reporteBoletinesAnalisisCommon.asp"-->
<% Response.Buffer = False %>
<%'-----------------------------------------------------------------------------------------------------------
Function imprimirFiltros()
	Dim auxCoordinador, auxCordinado, auxProducto, auxFechaDesdeD,auxFechaDesde,auxCalador  %>
	<table style="font-size:12; font-weight:bold; font-family:courier;">
		<tr>
			<td colspan="2"><%=GF_TRADUCIR("Puerto")%>:</td>
			<td colspan="2" align="left"><% =g_Pto %></td>
		</tr>	
		<tr>
	<%	auxCoordinador = "Todos"
		if (g_Coordinador > 0) then	auxCoordinador = Trim(g_Coordinador)&"-"&Trim(getDsEmpresa(g_Coordinador)) %>	
			<td colspan="2"><%=GF_TRADUCIR("Coordinador")%>:</td>
			<td colspan="2" align="left"><% =auxCoordinador %></td>
	<%	auxCoordinado = "Todos"
		if (g_Coordinado > 0) then	auxCoordinado = Trim(g_Coordinado)&"-"&Trim(getDsCliente(g_Coordinado)) %>
			<td colspan="2"><%=GF_TRADUCIR("Coordinado")%>:</td>
			<td colspan="2" align="left"><% =auxCoordinado %></td>
		</tr>
		<tr>
	<%	auxFechaDesde = "Todas" 
		if (g_FechaDesde <> "") then auxFechaDesde = g_FechaDesde %>
			<td colspan="2"><%=GF_TRADUCIR("Fecha Desde")%>:</td>
			<td colspan="2" align="left"><% =auxFechaDesde %></td>
	<%	auxFechaHasta = "Todas" 
		if (g_FechaHasta <> "") then auxFechaHasta = g_FechaHasta%>
			<td colspan="2"><%=GF_TRADUCIR("Fecha Hasta")%>:</td>
			<td colspan="2" align="left"><% =auxFechaHasta %></td>
		</tr> 
		<tr>
	<%	auxProducto = "Todos"
		if (g_Producto) then auxProducto = Trim(g_Producto)  & "-" & getDsProducto(g_Producto)%>		
			<td colspan="2"><%=GF_TRADUCIR("Producto")%>:</td>
			<td colspan="2" align="left"><% =auxProducto %></td>
	<%	auxCalador = "Todos"
		if(g_Calador <> "")then auxCalador = g_Calador &"-"& getUserDescription(g_Calador) %>
			<td colspan="2"><%=GF_TRADUCIR("Calador")%>:</td>
			<td colspan="2" align="left"><% =auxCalador %></td>
		</tr>
		<tr>	
	<%	auxSticker = "Todos"
		if(g_Sticker <> "")then auxSticker = Trim(g_Sticker)%>
			<td colspan="2"><%=GF_TRADUCIR("Sticker")%>:</td>
			<td colspan="2" align="left"><% =auxSticker %></td>
	<%	auxCertificado = "Todos"
		if(g_Certificado <> "")then auxCertificado = Trim(g_Certificado)%>	
			<td colspan="2"><%=GF_TRADUCIR("Certificado")%>:</td>
			<td colspan="2" align="left"><% =auxCertificado %></td>
		</tr>
		<tr>
	<%	auxGrados = "Todos"
		if(g_Grado > 0)then auxGrados = getDsGrado(g_Grado)%>
			<td colspan="2"><%=GF_TRADUCIR("Grado")%>:</td>
			<td colspan="2" align="left"><% =auxGrados %></td>
		</tr>
	</table>
<%End Function
'------------------------------------------------------------------------------------------------------------------
Function writeFieldCabecera(pField,pValue)	%>	
		<tr>
			<td colspan="3"></td>		
			<td align="left"><%=pField%></td>		
			<td colspan="5" align="left"><% =pValue %></td>	
		</tr><%
End Function
'-----------------------------------------------------------------------------------------------------------------
Function drawTitulosDetalle(pArr) %>	
	<tr><td colspan=8 class="border">
		<table style="font-size:10;">	
			<tr class="titulos">
			<% for x = 0 to UBound(pArr)%>				
				<td><%=pArr(x)%></td>
			<% next %>	
			</tr> 
		</table>
	</td></tr>	<%
End Function
'-----------------------------------------------------------------------------------------------------------------
Function drawDetalle(pArr)
	Dim myField, h, rtrn,myRegistro %>		
	<%	Call drawTitulosDetalle(pArr)
		myRegistro = Split(str, DETAIL_TOKEN) %>
		<tr><td colspan=8 class="border">
			<table style="font-size:10;">
		<%	For h = 0 To UBound(myRegistro)
				myField = Split(myRegistro(h), FIELD_TOKEN) %>
				<tr>				
		<%		myInit = 0
				if h > 0 then myInit = 3		
				For z = 0 To UBound(myField) %>
		<%			if myInit > 0 and z < 3 then  %>
						<td></td>
		<%			else
						rtrn = Split(myField(z), "=") 	%>
						<td align="<%=arrAlignDetalle(z)%>"><%= rtrn(1) %></td>
		<%			end if
				Next %>
				</tr>		
		<%	Next %>
			</table>
		</td></tr> <%		
End Function
'-----------------------------------------------
Function drawCabecera(pArr) 
	Dim myField, h, rtrn %>
	<tr><td colspan=8 class="border">
		<table style="font-size:10;">
	<%	myField = Split(str, FIELD_TOKEN)		
		For h = 0 To UBound(myField)
			rtrn = Split(myField(h), "=")%>
			<tr>
				<td colspan="2"></td>		
				<td  align="left"><%=rtrn(0)%></td>		
				<td colspan="5" align="left"><% =rtrn(1) %></td>	
			</tr>			
	<%	Next %>
		</table>
	</td></tr>	<%
End Function
'---------------------------------------------------------------
Function drawTotales(pArr) 
	Dim myField, h, rtrn, auxTotalNeto, auxGradoAnalisis%>
	<tr><td colspan=8 class="border">
		<table style="font-size:10;">
			<tr class="titulos">	
		<%	myRegistro = Split(str, FIELD_TOKEN)	
			For h = 0 To UBound(myRegistro)
				myField = Split(myRegistro(h), "=")	%>
				<td colspan=3 align="right"><%=myField(0) & ": " & myField(1) %></td>
		<%	Next %>
				<td colspan=2 align="right"></td>
			</tr>
		</table>
	</td></tr><%
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'   ETAPA 2 - GENERACION DEL EXCEL
'*****************************************************************************************
Dim str,index,txtLine,arrData, flagHayResultado
g_Pto		  = GF_PARAMETROS7("pto", "", 6)
g_Coordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
g_Coordinado  = GF_PARAMETROS7("cdCoordinado", "", 6)
g_Producto    = GF_PARAMETROS7("cmbCdProducto", 0, 6)
g_FechaDesdeD = GF_PARAMETROS7("fechaDesdeD", "", 6)
g_FechaDesdeM = GF_PARAMETROS7("fechaDesdeM", "", 6)
g_FechaDesdeA = GF_PARAMETROS7("fechaDesdeA", "", 6)
g_FechaHastaD = GF_PARAMETROS7("fechaHastaD", "", 6)
g_FechaHastaM = GF_PARAMETROS7("fechaHastaM", "", 6)
g_FechaHastaA = GF_PARAMETROS7("fechaHastaA", "", 6)
g_Sticker	  = GF_PARAMETROS7("sticker", "", 6)
g_Calador	  = GF_PARAMETROS7("cdCalador", "", 6)
g_Certificado = GF_PARAMETROS7("certificado", "", 6)
g_Grado		  = GF_PARAMETROS7("grado", 0, 6)

g_FechaDesde = g_FechaDesdeA & "-" & g_FechaDesdeM & "-" & g_FechaDesdeD
g_FechaHasta = g_FechaHastaA & "-" & g_FechaHastaM & "-" & g_FechaHastaD

fname = "BOLETINES_" & g_Pto
Call GF_createXLS(fname)
%>
<html>
<head>
	<style type="text/css">
		.border { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
		}

		.titulos {
			background-color:#D8D8D8;
			font-weight:bold;
		}

		.areas {
			background-color:#CECEF6;
			font-weight:bold;
		}
	</style>
</head>
<body onLoad="bodyOnLoad()">	
	<table class="border" style="background-color:#FFFACD; font-weight:bold">		
		<tr><td colspan=8 align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan=8 align="center" style="font-size:24"><% =GF_TRADUCIR("REPORTE DE BOLETINES DE ANALISIS") %></td></tr>		
	</table>	
	<table>
		<tr><td colspan=8>
	<%		Call imprimirFiltros() %>
		</td></tr>	<%			
			index = 0
			Set fs = Server.CreateObject("Scripting.FileSystemObject")			
			flagHayResultado = false
			while index <= maxSegment	
				pStrPath = Server.MapPath("Temp/BOLETINES_ANALISIS_" & session("Usuario") & "_" & index & ".txt")
				if (fs.FileExists(pStrPath)) then
					Set fadm = fs.OpenTextFile(pStrPath, 1)	
					while (not fadm.AtEndOfStream)					
						txtLine = fadm.ReadLine()
						arrData = Split(txtLine, SECTOR_TOKEN)
						str = arrData(0)
						Call drawCabecera(arrTitulosCabecera)
						str = arrData(1)
						Call drawDetalle(arrTitulosDetalle)
						str = arrData(2)
						Call drawTotales(arrTitulosTotal)
						flagHayResultado = true
					wend
					Set fadm = nothing					
				end if
				index = index + 1			
				fs.DeleteFile(pStrPath)			
			wend %>
		</td></tr>	
	<%	if not flagHayResultado then %>		
		<tr><td colspan=8 class="border">
			<table style="font-size:10;">				
				<tr><td colspan=8 align="center" class="titulos"><%=GF_TRADUCIR("No se encontraron resultados")%></td></tr>					
			</table>
		</td></tr>
	<% end if%>
	</table>
</body>
</html>