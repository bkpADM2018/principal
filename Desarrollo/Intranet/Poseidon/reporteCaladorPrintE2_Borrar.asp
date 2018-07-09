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
<!--#include file="reporteCaladorCommon.asp"-->
<% Response.Buffer = False 
Const INDEX_CDRUBRO = 0
Const INDEX_ABRUBRO = 1 
Const INDEX_VLRUBRO = 2
Const INDEX_DSRUBRO = 3
Const INDEX_NETO_SIN_MERMA = 13
%>
<%'-----------------------------------------------------------------------------------------------------------
Function imprimirFiltros()
	Dim auxCoordinador, auxCordinado, auxProducto, auxFechaDesdeD,auxFechaDesde,auxCalador  %>
	<table style="font-size:12;" width="100%">
		<tr>
	<%	auxFechaDesde = "Todas" 
		if (g_FechaDesde <> "") then auxFechaDesde = g_FechaDesde %>
			<td colspan="2" align="left"><%=GF_TRADUCIR("Fecha Desde")%>:</td>
			<td colspan="4" align="left"><% =auxFechaDesde %></td>
	<%	auxFechaHasta = "Todas" 
		if (g_FechaHasta <> "") then auxFechaHasta = g_FechaHasta%>
			<td colspan="2" align="left"><%=GF_TRADUCIR("Fecha Hasta")%>:</td>
			<td colspan="4" align="left"><% =auxFechaHasta %></td>
		</tr> 
		<tr>
	<%	auxUsuario = "Todos"
		if (g_cdUsuario <> "") then	auxUsuario = Trim(g_cdUsuario)&"-"&Trim(getUserDescription(g_cdUsuario)) %>	
			<td align="left" colspan="2"><%=GF_TRADUCIR("Usuario")%>:</td>
			<td colspan="4" align="left"><% =auxUsuario %></td>
	<%	auxCoordinado = "Todos"
		if (g_cdCoordinado > 0) then	auxCoordinado = Trim(g_cdCoordinado)&"-"&Trim(g_dsCoordinado) %>
			<td colspan="2"><%=GF_TRADUCIR("Coordinado")%>:</td>
			<td colspan="4" align="left"><% =auxCoordinado %></td>
		</tr>	
		<tr>
	<%	auxCorredor = "Todos"
		if (g_cdCorredor > 0) then	auxCorredor = Trim(g_cdCorredor)&"-"&Trim(g_dsCorredor) %>
			<td align="left" colspan="2"><%=GF_TRADUCIR("Corredor")%>:</td>
			<td colspan="4" align="left"><% =auxCorredor %></td>
	<%	auxVendedor = "Todos"
		if (g_cdVendedor > 0) then	auxVendedor = Trim(g_cdVendedor)&"-"&Trim(g_dsVendedor) %>
			<td colspan="2"><%=GF_TRADUCIR("Vendedor")%>:</td>
			<td colspan="4" align="left"><% =auxVendedor %></td>
		</tr>	
		<tr>
	<%	auxAceptacion = "Todos"
		if (g_cdAceptacion) then auxAceptacion = Trim(getDsAceptacion(g_cdAceptacion))%>
			<td colspan="2"><%=GF_TRADUCIR("Aceptacion")%>:</td>
			<td colspan="4" align="left"><% =auxAceptacion %></td>
	<%	auxProducto = "Todos"
		if(g_cdProducto <> 0)then auxProducto = Trim(g_cdProducto) & "-" & Trim(getDsProducto(g_cdProducto))%>
			<td colspan="2"><%=GF_TRADUCIR("Productos")%>:</td>
			<td colspan="4" align="left"><% =auxProducto %></td>
		</tr>
		<tr>	
	<%	auxChkCamiones = "NO"
		if(g_chkCamiones = 1) then auxChkCamiones = "SI" %>
			<td colspan="2"><%=GF_TRADUCIR("Ver Camiones")%>:</td>
			<td colspan="4" align="left"><% =auxChkCamiones %></td>	
	<%	auxChkVagones = "NO"
		if(g_chkVagones = 1) then auxChkVagones = "SI" %>
			<td colspan="2"><%=GF_TRADUCIR("Ver Vagones")%>:</td>
			<td colspan="4" align="left"><% =auxChkVagones %></td>
		</tr>
		<tr>
	<%	auxRubro = "Todos"
		if(g_cdRubro > 0)then auxRubro = g_cdRubro & "-" & getDsRubro(g_cdRubro)%>
			<td colspan="2"><%=GF_TRADUCIR("Rubro")%>:</td>
			<td colspan="4" align="left"><% =auxRubro %></td>
	<%	auxChkPromediar = "NO"
		if(g_chkPromediar = 1) then auxChkPromediar = "SI" %>
			<td colspan="2"><%=GF_TRADUCIR("Promediar")%>:</td>
			<td colspan="4" align="left"><% =auxChkPromediar %></td>
		</tr>
	</table>
<%End Function
'-----------------------------------------------------------------------------------------------------------------
Function drawTitulosDetalle() %>	
	<tr><td colspan=8 class="border">
		<table style="font-size:10;">	
			<tr class="titulos">
				<td colspan="5"></td>
			<% for x = 0 to UBound(arrTitulosRubros) - 1%>				
				<td align="center"><%=arrTitulosRubros(x)%></td>
			<% next %>	
			</tr> 
		</table>
	</td></tr>	<%
End Function
'-----------------------------------------------------------------------------------------------------------------
Function drawDetalle(pStr,pNetoSMerma)
	Dim myField, h, rtrn,myRegistro,columnInit,auxCdRubro,auxVlRubro %>		
	<%	Call drawTitulosDetalle()
		myRegistro = Split(pStr, DETAIL_TOKEN) %>
		<tr><td colspan=8 class="border">
			<table style="font-size:10;">
		<%	For h = 0 To UBound(myRegistro)
				myField = Split(myRegistro(h), FIELD_TOKEN) %>
				<tr>
					<td colspan="5"></td>
		<%			For z = 0 To UBound(myField)
						rtrn = Split(myField(z), "=")
						Call loadPropertyRubro(z, rtrn(1))
						if (z < INDEX_DSRUBRO)then %>
							<td align="center"><%= rtrn(1) %></td>
		<%				end if
					Next %>
				</tr>
		<%		'sumo los valores de los rubros para totalizarlo al final
				auxKey = Trim(g_CdRubro&"|"& Trim(g_AbRubro) &"|"& Trim(g_DsRubro))
				if (not dicTest.Exists(auxKey)) Then
					Call dicTest.Add(auxKey,pNetoSMerma*g_VlRubro)
					Call dicContRubro.Add(auxKey,pNetoSMerma)
				else					
					if (pNetoSMerma > 0) then 
						dicTest.Item(auxKey) = Cdbl(dicTest.Item(auxKey)) + (pNetoSMerma * g_VlRubro)
						dicContRubro.Item(auxKey) = dicContRubro.Item(auxKey) + Cdbl(pNetoSMerma)
					end if	
				end if
			Next %>
			</table>
		</td></tr> <%		
End Function
'-------------------------------------------------------------------------------------------------------------
Function loadPropertyRubro(pIndex,pVl)
	Select Case pIndex
		Case INDEX_CDRUBRO
			g_CdRubro = pVl
		Case INDEX_ABRUBRO
			g_AbRubro = pVl
		Case INDEX_VLRUBRO
			g_VlRubro = pVl
		Case INDEX_DSRUBRO
			g_DsRubro = pVl
	End Select		
End Function
'-------------------------------------------------------------------------------------------------------------
'Dibuja la cebecera y devuelve el valor de Kilos netos sin merma
Function drawCabecera(pArr) 
	Dim myField, h, rtrn,auxNetoSMerma 	
	auxNetoSMerma = 0
	myField = Split(pArr, FIELD_TOKEN) %>
	<tr>	
<%	For h = 0 To UBound(myField)
		rtrn = Split(myField(h), "=")		
		if (h = INDEX_NETO_SIN_MERMA) then auxNetoSMerma = rtrn(1)%>
		<td class="reg_header_navdos" align="left"><%=rtrn(1)%></td>
<%	Next %>
	</tr>
<%  drawCabecera = auxNetoSMerma
End Function
'------------------------------------------------------------------------------------------------------------
'Dibuja los titulos de la Cabecera de Camiones
Function drawTituloCabecera(pArr) %>
	<TR>	
<%	for i = 0 to UBound(pArr) %>
		<TD class="reg_header_nav" align="center">	<%=pArr(i)%> </TD>
<%	next  %>
	</TR>
	<%
End Function
'------------------------------------------------------------------------------------------------------------
Function imprimirTitulosTotales()
	Dim h %>
	<tr>	
		<td colspan="4">
<%	For h = 0 To UBound(arrTitulosTotales) %>
		<td class="reg_header_nav" colspan="2" align="center"><%=arrTitulosTotales(h)%></td>
<%	Next %>
	</tr>
<%End Function
'------------------------------------------------------------------------------------------------------------
Function imprimirTotales()
	Dim h, aux
	Call imprimirTitulosTotales()
	for each strKey in dicTest.Keys 
		aux = Split(strKey,"|")
		%>
	<tr>	
		<td colspan="4">
		<td class="reg_header_navdos" colspan="2" align="left"><%=aux(1)%></td>
		<td class="reg_header_navdos" colspan="2" align="left"><%=aux(2)%></td>
		<td class="reg_header_navdos" colspan="2" align="right"><%=round(Cdbl(dicTest.Item(strKey))/Cdbl(dicContRubro.Item(strKey)),2)%></td>
	</tr>
<%	Next
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'   ETAPA 2 - GENERACION DEL EXCEL
'*****************************************************************************************
Dim str,index,txtLine,arrData, flagHayResultado,fadm,auxCdRubro,auxVlRubro,auxDsRubro,auxAbRubro,dicContRubro,totNetoSinMerma

g_FechaDesde = g_FechaDesdeA & "-" & g_FechaDesdeM & "-" & g_FechaDesdeD
g_FechaHasta = g_FechaHastaA & "-" & g_FechaHastaM & "-" & g_FechaHastaD

fname = "REPORTE_CALADOR_" & g_Pto
Call GF_createXLS(fname)
index = 0
Set dicTest = Server.CreateObject("Scripting.Dictionary")
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set dicCam = Server.CreateObject("Scripting.Dictionary")
Set dicVag = Server.CreateObject("Scripting.Dictionary")
Set dicContRubro = Server.CreateObject("Scripting.Dictionary")
while index <= maxSegment
	pStrPath = Server.MapPath("Temp/REPORTE_CALADOR_" & session("Usuario") & "_" & index & ".txt")
	if (fs.FileExists(pStrPath)) then
		Set fadm = fs.OpenTextFile(pStrPath, 1)
		while (not fadm.AtEndOfStream)
			txtLine = fadm.ReadLine()
			if (Trim(txtLine) = REPORTE_CAMIONES) then
				isCamion = true
				isVagon  = false
			else if (Trim(txtLine) = REPORTE_VAGONES) then
					isCamion = false
					isVagon  = true
				else
					if isCamion then Call dicCam.add(cont,txtLine)
					if isVagon then  Call dicVag.add(cont,txtLine)
				end if
			end if
			cont = cont + 1
		wend
		Set fadm = nothing
		fs.DeleteFile(pStrPath)
	end if
	index = index + 1
wend


%>
<html>
<head>
	<style type="text/css">
        .reg_header
        {
            BORDER-BOTTOM: #f4b800 1px solid;
            BORDER-LEFT: #f4b800 1px solid;
            BACKGROUND-COLOR: #ffeecd;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            FONT-SIZE: 10px;
            BORDER-TOP: #f4b800 1px solid;
            BORDER-RIGHT: #f4b800 1px solid;
            TEXT-DECORATION: none;
            -moz-border-radius: 5px 5px 5px 5px
        }
        .reg_header_error
        {
            BORDER-BOTTOM: #f80800 1px solid;
            BORDER-LEFT: #f40800 1px solid;
            BACKGROUND-COLOR: #ffaa99;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            COLOR: #ffffff;
            FONT-SIZE: 10px;
            BORDER-TOP: #f40800 1px solid;
            FONT-WEIGHT: bold;
            BORDER-RIGHT: #f40800 1px solid;
            TEXT-DECORATION: none
        }
        .reg_header_nav
        {
            BACKGROUND-COLOR: #517b4a;
            COLOR: #ffffff;
            FONT-SIZE: 10px;
            FONT-WEIGHT: bold
        }
        .reg_header_navdos
        {
            BACKGROUND-COLOR: #dcdcdc;
            COLOR: #006400;
            FONT-SIZE: 10px;
            FONT-WEIGHT: bold
        }
        .titu_header
        {
            BORDER-BOTTOM: #006400 1px solid;
            BORDER-LEFT: #006400 1px solid;
            BACKGROUND-COLOR: #517b4a;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            COLOR: white;
            FONT-SIZE: 12px;
            BORDER-TOP: #006400 1px solid;
            FONT-WEIGHT: bold;
            BORDER-RIGHT: #006400 1px solid;
            TEXT-DECORATION: none
        }
    </style>
</head>
<body>
	<table  border="1" cellpadding="0" cellspacing="0" width="60%">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td colspan="19" align="center" class="titu_header"><%=GF_Traducir("Reporte de Calador")%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="border">
				<%Call imprimirFiltros %>
			</td>
		</tr>
		<% if (g_chkCamiones = 1) Then %>
			<tr>
				<td>
					<table class="reg_header_navdos" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">
						<tr>
							<td class="reg_header_nav" align="center" colspan="18"><%=GF_TRADUCIR("CAMIONES")%></td>
						</tr>
				<%		if (dicCam.Count > 0) Then
							flagHayResultado = false
							for each strItem in dicCam.Items
								if not flagHayResultado then Call drawTituloCabecera(arrTitulosCamiones)
								str = Split(strItem, SECTOR_TOKEN)
								strCabecera = str(0)
								strDetalle  = str(1)
								netoSinMerma = drawCabecera(strCabecera)
								Call drawDetalle(strDetalle,netoSinMerma)
								flagHayResultado = true
							Next
							Call imprimirTotales()
							dicTest.RemoveAll
							dicContRubro.RemoveAll
						else	%>
							<tr>
								<td align="center" colspan="18"><%=GF_TRADUCIR("No se encontraron camiones")%></td>
							</tr>
				<%		end if	%>
					</table>
				</td>
			</tr>
		<% end if %>	
		<tr><td colspan="18"></td></tr>
		<% if (g_chkVagones = 1) Then %>		
			<tr>
				<td>
					<table class="reg_header_navdos" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">
						<tr>
							<td class="reg_header_nav" align="center" colspan="19"><%=GF_TRADUCIR("VAGONES")%></td>
						</tr>
				<%		if (dicVag.Count > 0) Then
							flagHayResultado = false
							for each strItem in dicVag.Items
								if not flagHayResultado then Call drawTituloCabecera(arrTitulosVagones)
								str = Split(strItem, SECTOR_TOKEN)
								strCabecera = str(0)
								strDetalle  = str(1)
								netoSinMerma = drawCabecera(strCabecera)
								Call drawDetalle(strDetalle,netoSinMerma)								
								flagHayResultado = true
							Next
							Call imprimirTotales()
							dicTest.RemoveAll
							dicContRubro.RemoveAll
						else	%>
							<tr>
								<td align="center" colspan="18"><%=GF_TRADUCIR("No se encontraron vagones")%></td>
							</tr>
				<%		end if	%>
					</table>
				</td>
			</tr>
		<% end if %>	
	</table>
</body>


	