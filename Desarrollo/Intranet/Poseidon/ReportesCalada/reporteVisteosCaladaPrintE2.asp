<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteVisteosCaladaCommon.asp"-->
<% Response.Buffer = False %>
<%

'-----------------------------------------------------------------------------------------
'Ordenamiento de vector. Se utiliza un metodo precario pero efectivo debido a que no hay muchos elementos en el vector. 
'Como maximo al momento de definir este algoritmo, un vector puede tener 42 elementos. En promedio tienen 5.
function sortArray(arrShort)

    for i = 0 To UBound(arrShort)-1
			for j= i+1 to UBound(arrShort)
			    if (arrShort(i)>arrShort(j)) then
			        temp=arrShort(j)
			        arrShort(j)=arrShort(i)
			        arrShort(i)=temp
			    end if
			next
    next
    sortArray = arrShort

End function
'-----------------------------------------------------------------------------------------
function imprimirFiltros(pto,fechaD,fechaH,pidcamion,pnuCartaPorte,pcdProducto, pDsProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado, pcdTransporte)	
%>
	<table style="font-size:16; font-weight:bold; font-family:courier">
	<tr>
		<td colspan="4">Puerto......: <% =pto %></td>
		<td></td>
<%
	auxCorredor = "Todos"
	if(pcdCorredor > 0)then
		auxCorredor = Trim(pcdCorredor)&" - "&Trim(getDsCorredor(pcdCorredor))	
	end if	
%>	
		<td colspan="4">Corredor...: <% =auxCorredor %></td>
	</tr>
<%
	myFormatFecha = GF_FN2DTE(fechaD)
%>
	<tr>
		<td colspan="4">Fecha Desde.: <% =myFormatFecha %></td>
		<td></td>
<%
	auxVendedor = "Todos"
	if(pcdVendedor > 0)then 
		auxVendedor = Trim(pcdVendedor)&" - "&Trim(getDsVendedor(pcdVendedor))			
	end if
%>
		<td colspan="4">Vendedor...: <% =auxVendedor %></td>
	</tr>
<%
	myFormatFecha = GF_FN2DTE(fechaH)
%>
	<tr>
		<td colspan="4">Fecha Hasta.: <% =myFormatFecha %></td>
		<td></td>
<%
	auxCliente = "Todos"
	if(pcdCliente > 0)then 
		auxCliente = Trim(pcdCliente)&" - "&Trim(getDsCliente(pcdCliente))			
	end if
%>

		<td colspan="4">Cliente....: <% =auxCliente %></td>
	</tr>
<%
	if(pidcamion > 0)then 
		auxCamion = pidcamion		
	else
		if (nuCartaPorte <> "") then
			auxCamion = GF_EDIT_CTAPTE(GF_nChars(nuCartaPorte, 16, "0", CHR_AFT))
		else
			auxCamion = "Todos" 
		end if
	end if
%>
	<tr>
		<td colspan="4">Camion......: <% =auxCamion %></td>
		<td></td>
<%
	auxEntregador = "Todos"
	if(pcdEntregador > 0)then 
		auxEntregador = Trim(pcdEntregador)&" - "&Trim(getDsEntregador(pcdEntregador))	
	end if
%>	
		<td colspan="4">Entregador.: <% =auxEntregador %></td>
	</tr>
<%
	auxProducto = "Todos"
	if(pcdProducto > 0)then auxProducto = Trim(pcdProducto)&" - "&Trim(pDsProducto)	
%>
	<tr>
		<td colspan="4">Producto....: <% =auxProducto %></td>
		<td></td>
		<td colspan="4">Estado......: <% 
		    if (CInt(pcdEstado) = 0) then
		        aux = "Descargados OK"
            else
                aux = getDsEstado(pcdEstado) 
            end if		         
		    response.write pcdEstado & "-" & aux
		    %></td>
	</tr>
	<tr>
		<td colspan="4">Transporte..: <% =getDsTransporte(pcdTransporte) %></td>
		<td></td>
		<td colspan="4"></td>
	</tr>
	</table>
<%
End function
'-----------------------------------------------------------------------------------------
Function imprimirTitulos(pMaxSQCalada, pArrTitulos, pArrAlign, pArrRubros, pDicRubros)
	Dim i, j, arr(), index
					
	Redim arr(UBound(pArrTitulos) + (pMaxSQCalada * UBound(arrTitulosVisteo)) + (pMaxSQCalada*pDicRubros.Count) + pMaxSQCalada, 1)
%>	
		<tr style="background-color:#E3F6CE; font-weight:bold">
<%
		'Titulos del camion.
		For i = 0 to UBound(pArrTitulos)
			arr(i,0) = pArrTitulos(i)
			arr(i,1) = pArrAlign(i)
%>
			<td class="border" align="center"><% =GF_TRADUCIR(arr(i,0)) %></td>
<%
		Next		
		index = UBound(pArrTitulos) + 1
		
		For i = 1 to pMaxSQCalada						
			'Titulos de las cabeceras de visteos
			arr(index,0) = arrTitulosVisteo(0)
			arr(index,1) = pArrAlign(0)
%>
			<td class="border" align="center"><% =GF_TRADUCIR(arr(index,0)) & "_" & i %></td>			
<%			
			index=index+1
			arr(index,0) = arrTitulosVisteo(1)
			arr(index,1) = pArrAlign(1)
%>
			<td class="border" align="center"><% =GF_TRADUCIR(arr(index,0)) & "_" & i %></td>
<%			
			index=index+1
			arr(index,0) = arrTitulosVisteo(2)
			arr(index,1) = pArrAlign(2)
%>			
			<td class="border" align="center"><% =GF_TRADUCIR(arr(index,0)) & "_" & i %></td>
			
<%			index=index+1
			'Titulos de los rubros.
			For j = 0 to UBound(pArrRubros)			
				arr(index,0) = pArrRubros(j)
				arr(index,1) = "right"
%>
			<td class="border" align="center"><% =pDicRubros.item(pArrRubros(j)) & "_" & i %></td>
<%				index=index+1
			Next
		Next	%>				
		</tr>	
<%	
	imprimirTitulos = arr
End function
'-----------------------------------------------------------------------------------------
Function imprimirCampo(data, control, txtAlign)
	Dim i, arr
	
	arr	= Split(data, VALUE_TOKEN)
	if (arr(0) = control) then
		imprimirCampo = true
%>
	<td align="<% =txtAlign %>"	class="border">	<% =arr(1) %></td>
<%	
	else
		imprimirCampo = false
%>
	<td class="border"></td>
<%	
	end if		

End function
'-----------------------------------------------------------------------------------------
Function imprimirDatosXLS(pStrPath, arrAlignCamion, arrAlignVisteos, arrAlignRubros, arrTitulos)
	Dim fs, arch, txtLine, h, w, ret, data
	Dim arr1, arr2
	
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Set arch = fs.OpenTextFile(pStrPath, 1)
	
	while (not arch.AtEndOfStream)
%>
		<tr>
<%	
		txtLine = arch.ReadLine()
		arrData = Split(txtLine, FIELD_TOKEN)
		w = 0
		For h = 0 To UBound(arrTitulos,1)
				data = NO_DATA
				if (w <= UBound(arrData)) then data = arrData(w)
				ret =  imprimirCampo(data, arrTitulos(h,0), arrTitulos(h,1))
				if (ret) then w = w+1
		Next
%>
		</tr>
<%	
	wend
	
	arch.close()
	
	Set arch = Nothing
	Set fs = nothing
	
End Function
'------------------------------------------------------------------------------------------
Function imprimirEncabezado(titulo)
	dim division, conn, rsDivision, strSQL
%>
	<table class="border">
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="center" style="font-size:24"><% =GF_TRADUCIR(titulo) %></td></tr>
	</table>
<%	
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'*****************************************************************************************

'*****************************************************************************************
'* ETAPA 2 - GENERACION DEL EXCEL
'*****************************************************************************************

arrRubros = dicRubros.Keys()
arrRubros = sortArray(arrRubros)


filename = "VISTEOS_CALADA_" & g_Puerto
if (dsProducto <> "") then filename = filename & "_" & dsProducto
filename = filename & "_" & Replace(g_fechaDesde,"-", "_") & "_al_" & Replace(g_fechaHasta,"-", "_")

Call GF_createXLS(filename)
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
		<tr><td><% Call imprimirEncabezado("DATOS VISTEOS CALADA") %></td></tr>		
	</table>	
<%		
	Call imprimirFiltros(g_Puerto,g_fechaDesde,g_fechaHasta,g_idCamion,g_cPorte,g_Producto, dsProducto,g_Vendedor,g_Corredor,g_Cliente,g_Corredor, g_Estado, g_transporte)
%>
	<table class="border">
<%
	arrTitulosCompletos = imprimirTitulos(g_MaxSQCalada, arrTitulosExcel, arrAlignExcel, arrRubros, dicRubros)	
	'Establezco la ruta y el nombre del archivo a crear
	Call imprimirDatosXLS(strPath, arrAlignExcel, arrAlignVisteos, arrAlignRubros, arrTitulosCompletos)
	
	'Borro archivos de trabajo
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	'Se borra el archivo temporal.
	Call fs.deleteFile(strPath, true)
	Call fs.deleteFile(strPathAdm, true)
	Set fs = nothing
%>		
	</table>
</body>
</html>
