<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<% Response.Buffer = False 

Const INDEX_KG_DESCARGADOS = 0
Const INDEX_KG_MERMA = 1
Const INDEX_KG_TOTAL = 2
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function imprimirFiltros(p_Pto,p_FechaDesde,p_FechaHasta,p_CdProducto) %>
	<table style="border-style:solid;border-width:thin;" >
	    <tr>
    		<td colspan="2" class="titulos" ><%=GF_TRADUCIR("Puerto: ") %></td>
            <td colspan="4" class="titulos" align="left" ><%=pto %></td>
        </tr>    
    	<tr>
		    <td colspan="2" class="titulos" ><%=GF_TRADUCIR("Fecha Desde: ") %></td>
            <td colspan="4" class="titulos" align="left" ><% =GF_FN2DTE(p_FechaDesde) %></td>
        </tr>
        <tr>
            <td colspan="2" class="titulos" ><%=GF_TRADUCIR("Fecha Hasta: ") %></td>
            <td colspan="4" class="titulos" align="left" ><% =GF_FN2DTE(p_FechaHasta) %></td>
        </tr>
        <tr>
            <td colspan="2" class="titulos" ><%=GF_TRADUCIR("Producto: ") %></td>
            <td colspan="4" class="titulos" align="left" >
               <% if (Cdbl(p_CdProducto) <> 0) then                    
                    Response.Write getDsProducto(p_CdProducto) 
                  else
                    Response.Write "Todos"
                   end if%>
            </td>
        </tr>
    </table>
<%
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function corteControlFechaMermaVolatil(p_Rs,p_Fecha)
    corteControlFechaMermaVolatil = false
    if (not p_Rs.Eof) then
        if (Cdbl(p_Fecha) = Cdbl(p_Rs("DTCONTABLE"))) then corteControlFechaMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function corteControlProdcutoMermaVolatil(p_Rs,p_Producto,p_Fecha)
    corteControlProdcutoMermaVolatil = false
    if (not p_Rs.Eof) then
        if ((Cdbl(p_Producto) = Cdbl(p_Rs("CDPRODUCTO")))and(Cdbl(p_Fecha) = Cdbl(p_Rs("DTCONTABLE")))) then corteControlProdcutoMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function corteControlClienteMermaVolatil(p_Rs,p_Cliente,p_Fecha,p_Producto)
    corteControlClienteMermaVolatil = false
    if (not p_Rs.Eof) then
        if ((Cdbl(p_Cliente) = Cdbl(p_Rs("CDCLIENTE")))and(Cdbl(p_Fecha) = Cdbl(p_Rs("DTCONTABLE")))and(Cdbl(p_Producto) = Cdbl(p_Rs("CDPRODUCTO")))) then corteControlClienteMermaVolatil = true
    end if
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function iniciarArrayTotales(ByRef p_arrTotal)
    p_arrTotal(INDEX_KG_DESCARGADOS) = 0
    p_arrTotal(INDEX_KG_MERMA) = 0
    p_arrTotal(INDEX_KG_TOTAL) = 0
End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function armarCuerpoReporte(p_Pto,p_FechaDesde,p_FechaHasta,p_Transporte,p_CdProducto)
    Dim rs, fecha,arrTotalCliente(3),arrTotalProducto(3),arrTotalFecha(3),arrTotalPeriodo(3),cdProducto, dsProducto,cdCliente,dsCliente,auxRatio, vecClientes(50, 2), vecProductos(50, 2)
    Dim strSQL
    
    
    if (CInt(p_Transporte) = TIPO_TRANSPORTE_CAMION) then            
        Call executeSP_Puertos(rs, p_Pto, "HCAMIONESDESCARGA_GET_MERMAVOLATIL_BY_PARAMETERS", GF_FN2DTCONTABLE(p_FechaDesde) &"||"& GF_FN2DTCONTABLE(p_FechaHasta) &"||"& p_CdProducto &"||||")
    else
        Call executeSP_Puertos(rs, p_Pto, "HVAGONES_GET_MERMAVOLATIL_BY_PARAMETERS", GF_FN2DTCONTABLE(p_FechaDesde) &"||"& GF_FN2DTCONTABLE(p_FechaHasta) &"||"& p_CdProducto&"||||")
    end if

    if not rs.Eof then %>
        <tr>
            <td colspan="3"></td>
            <td align="center" style="font-size:10px;font-weight:bold;"><%=GF_TRADUCIR("Descargado") %></td>
            <td align="center" style="font-size:10px;font-weight:bold;"><%=GF_TRADUCIR("Merma") %></td>
            <td align="center" style="font-size:10px;font-weight:bold;"><%=GF_TRADUCIR("Neto") %></td>
        </tr>
    <%  while (not rs.Eof)
            Call iniciarArrayTotales(arrTotalFecha)
            fecha = rs("DTCONTABLE") %>
            <tr>
                <td ><%=GF_TRADUCIR("Fecha: ") %></td>
                <td colspan="5" align="left"><%=GF_FN2DTE(fecha) %></td>
            </tr>
            <tr><td colspan="6"></td></tr>
         <% while (corteControlFechaMermaVolatil(rs, fecha))
                Call iniciarArrayTotales(arrTotalProducto)
                cdProducto = rs("CDPRODUCTO") 
                dsProducto = Ucase(Trim(rs("DSPRODUCTO")))
                %> 
                <tr>
                    <td></td>
                    <td><%=GF_TRADUCIR("Producto: ") %></td>
                    <td colspan="4"><%=dsProducto %></td>
                </tr>
             <% while (corteControlProdcutoMermaVolatil(rs, cdProducto,fecha))
                    Call iniciarArrayTotales(arrTotalCliente)
                    dsCliente = Ucase(Trim(rs("DSCLIENTE")))
                    cdCliente = rs("CDCLIENTE") %>
                <%  while (corteControlClienteMermaVolatil(rs, cdCliente,fecha,cdProducto))
                        arrTotalCliente(INDEX_KG_DESCARGADOS) = Cdbl(arrTotalCliente(INDEX_KG_DESCARGADOS)) + Cdbl(rs("PESO"))
                        auxRatio = 0
                        arrTotalCliente(INDEX_KG_MERMA) = Cdbl(arrTotalCliente(INDEX_KG_MERMA)) + CDbl(rs("MERMAVOLATIL"))
                        rs.MoveNext()
                    wend 
                    'Finalizo la cuenta del cliente, redondeo y totalizo el neto.
                    arrTotalCliente(INDEX_KG_DESCARGADOS) = round(Cdbl(arrTotalCliente(INDEX_KG_DESCARGADOS)), 0)
                    arrTotalCliente(INDEX_KG_MERMA) = round(Cdbl(arrTotalCliente(INDEX_KG_MERMA)), 0)
                    arrTotalCliente(INDEX_KG_TOTAL) = Cdbl(arrTotalCliente(INDEX_KG_DESCARGADOS)) - Cdbl(arrTotalCliente(INDEX_KG_MERMA))
                %>
                    <tr>
                        <td colspan="2"></td>
                        <td><%=dsCliente %></td>
                        <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalCliente(INDEX_KG_DESCARGADOS))*100,2) & " KG" %></td>
                        <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalCliente(INDEX_KG_MERMA))*100,2) & " KG" %></td>
                        <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalCliente(INDEX_KG_TOTAL))*100,2) & " KG" %></td>
                   </tr>
           <%       arrTotalProducto(INDEX_KG_DESCARGADOS) =  Cdbl(arrTotalProducto(INDEX_KG_DESCARGADOS)) + Cdbl(arrTotalCliente(INDEX_KG_DESCARGADOS))
                    arrTotalProducto(INDEX_KG_MERMA) =  Cdbl(arrTotalProducto(INDEX_KG_MERMA)) + CDbl(arrTotalCliente(INDEX_KG_MERMA))                    
               wend 
               'Finalizo, totalizo el neto de producto.
               arrTotalProducto(INDEX_KG_TOTAL) =  Cdbl(arrTotalProducto(INDEX_KG_DESCARGADOS)) - CDbl(arrTotalProducto(INDEX_KG_MERMA))               
           %>
                <tr>
                    <td ></td>
                    <td colspan="2" align="left" class="titulos"><%=GF_TRADUCIR("Total Producto ")  %></td>
                    <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalProducto(INDEX_KG_DESCARGADOS))*100,2) & " KG" %></td>
                    <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalProducto(INDEX_KG_MERMA))*100,2) & " KG" %></td>
                    <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalProducto(INDEX_KG_TOTAL))*100,2) & " KG" %></td>
                </tr>
                <tr><td colspan="6"></td></tr>
        <%      Call totalzarKilos(vecProductos, dsProducto, Cdbl(arrTotalProducto(INDEX_KG_DESCARGADOS))*100, Cdbl(arrTotalProducto(INDEX_KG_MERMA))*100)  
                arrTotalFecha(INDEX_KG_DESCARGADOS) =  Cdbl(arrTotalFecha(INDEX_KG_DESCARGADOS)) + Cdbl(arrTotalProducto(INDEX_KG_DESCARGADOS))
                arrTotalFecha(INDEX_KG_MERMA) =  Cdbl(arrTotalFecha(INDEX_KG_MERMA)) + Cdbl(arrTotalProducto(INDEX_KG_MERMA))
            wend  
            'Finalizo fecha, totalizo neto de fecha.
            arrTotalFecha(INDEX_KG_TOTAL) =  Cdbl(arrTotalFecha(INDEX_KG_DESCARGADOS)) - Cdbl(arrTotalFecha(INDEX_KG_MERMA))                
        %>
            <tr>
                <td colspan="3" class="titulos"><%=GF_TRADUCIR("Total Fecha ")  %></td>
                <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalFecha(INDEX_KG_DESCARGADOS))*100,2) & " KG" %></td>
                <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalFecha(INDEX_KG_MERMA))*100,2) & " KG" %></td>
                <td align="right" class="titulos"><%= GF_EDIT_DECIMALS(Cdbl(arrTotalFecha(INDEX_KG_TOTAL))*100,2) & " KG" %></td>
            </tr>
    <%      arrTotalPeriodo(INDEX_KG_DESCARGADOS) =  Cdbl(arrTotalPeriodo(INDEX_KG_DESCARGADOS)) + CDbl(arrTotalFecha(INDEX_KG_DESCARGADOS))
            arrTotalPeriodo(INDEX_KG_MERMA) =  Cdbl(arrTotalPeriodo(INDEX_KG_MERMA)) + CDbl(arrTotalFecha(INDEX_KG_MERMA))            
        wend 
        'Finalizo periodo, totalizo neto.
        arrTotalPeriodo(INDEX_KG_TOTAL) =  Cdbl(arrTotalPeriodo(INDEX_KG_DESCARGADOS)) - CDbl(arrTotalPeriodo(INDEX_KG_MERMA))    
    %>

<%
        salir = false
        idx = LBound(vecProductos)%>
        <tr><td colspan="6"></td></tr>
        <% if p_Transporte = TIPO_TRANSPORTE_CAMION then %>
            <tr><td colspan="6" class="titulos"><%=GF_TRADUCIR("Totales por Producto de CAMIONES")%></td></tr>
            <%else%>
            <tr><td colspan="6" class="titulos"><%=GF_TRADUCIR("Totales por Producto de VAGONES")%></td></tr>
        <%end if
        while ((idx <= UBound(vecProductos)) and (not salir))    
            if (vecProductos(idx, 0) <> "") then %>
	            <tr><td colspan = "3" class="titulos"><%=vecProductos(idx, 0) %></td>
		            <td align="right" ><%=GF_EDIT_DECIMALS(Cdbl(vecProductos(idx, 1)),2) &" KG." %></td>
                    <td align="right" ><%=GF_EDIT_DECIMALS(Cdbl(vecProductos(idx, 2)),2) &" KG." %></td>
                    <td align="right" ><%=GF_EDIT_DECIMALS((Cdbl(vecProductos(idx, 1)) - Cdbl(vecProductos(idx, 2))),2) &" KG." %></td>
	            </tr>	            
        <%
                idx = idx + 1
            else
                salir = true
            end if	            
        wend
%>        
        <tr><td colspan="3" class="titulos"><%=GF_TRADUCIR("Total periodo ") & GF_FN2DTE(p_FechaDesde) &" al "& GF_FN2DTE(p_FechaHasta)  %></td>
            <td align="right" class="titulos"><%=GF_EDIT_DECIMALS(Cdbl(arrTotalPeriodo(INDEX_KG_DESCARGADOS))*100,2) &" KG." %></td>
            <td align="right" class="titulos"><%=GF_EDIT_DECIMALS(Cdbl(arrTotalPeriodo(INDEX_KG_MERMA))*100,2) &" KG." %></td>
            <td align="right" class="titulos"><%=GF_EDIT_DECIMALS(Cdbl(arrTotalPeriodo(INDEX_KG_TOTAL))*100,2) &" KG." %></td>
        </tr>
<% else %>
        <tr><td colspan="6"><%=GF_TRADUCIR("No se encontraron resultados") %></td></tr>
 <% end if

End function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function totalzarKilos(ByRef pVec, pClave, pKgDescargados, pKgMerma)
    Dim x, encontrado
    'Busco la clave en el vector.
    encontrado = false
    x = LBound(pVec)
    while ((x <= UBound(pVec)) and (not encontrado))
        if ((pVec(x, 0) = "") or (pVec(x, 0) = pClave)) then
            'Llegue al final de las claves cargadas o encontre la calve buscada, sumo los kilos.
            pVec(x, 0) = pClave
            if (pVec(x, 0) = "") then 
                pVec(x, 1) = 0
                pVec(x, 2) = 0
            end if                
            pVec(x, 1) = CDbl(pVec(x, 1)) + CDbl(pKgDescargados)
            pVec(x, 2) = CDbl(pVec(x, 2)) + CDbl(pKgMerma)
            encontrado = true            
        end if            
        x = x + 1
    wend
End Function
'******************************************************************************************************
'**************************************** COMIENZO DE PAGINA ******************************************
'******************************************************************************************************
Dim filename, pto, fechaDesde, fechaHasta, cdProducto

filename = "MERMAVOLATIL_" & g_Puerto & "_" & Left(Session("MmtoDato"),8) & ".xls"

fechaDesde = GF_Parametros7("fechaDesde","",6)
fechaHasta = GF_Parametros7("fechaHasta","",6)
pto = GF_Parametros7("pto","",6)
g_strPuerto = pto
cdProducto = GF_Parametros7("cdProducto",0,6)

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
			font-size:14;
            font-weight:bold; 
		}
	</style>	
</head>
<body>
	<table class="border" style="background-color:#FFFACD; font-weight:bold">
		<tr><td colspan=6 align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan=6 align="center" style="font-size:24"><% =GF_TRADUCIR("REPORTE DE MERMA VOLATIL") %></td></tr>
	</table>

<%		
	Call imprimirFiltros(pto,fechaDesde,fechaHasta,cdProducto)
%>
    <table>
        <tr><td colspan="6"><H3><%=GF_TRADUCIR("CAMIONES") %></H3></td></tr>
<%	
	Call armarCuerpoReporte(pto,fechaDesde,fechaHasta,TIPO_TRANSPORTE_CAMION,cdProducto)
%>
        <tr><td colspan="6"></td></tr>

        <tr><td colspan="6"><H3><%=GF_TRADUCIR("VAGONES") %></H3></td></tr>
<%
    Call armarCuerpoReporte(pto,fechaDesde,fechaHasta,TIPO_TRANSPORTE_VAGON,cdProducto)
%>		
	</table>
</body>
</html>
