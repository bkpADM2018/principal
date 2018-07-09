<!--#include file="consultaCamionesCommon.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<%
Dim xlsCircuito, xlsEstado
'Llamar a la pagina que arma el EXCEL.
filename = "Camiones_" & pto & "_" & session("Usuario") & "_" & session("MmtoDato")
Call GF_createXLS(filename)
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />   
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
<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
	    <td colspan="12" align="center" class="titu_header"><%=GF_Traducir("Consulta de Camiones")%></td>
    </tr>
</table>
<table width="95%" align="center" border="0">
    <tr>
		<td align="right"><b><% = GF_TRADUCIR("Id Camion") %>:</b></td>
		<td ><% =idCamion %></td>
		<td></td>
		<td align="right"><b><% = GF_TRADUCIR("Fecha Contable") %>:</b></td>
		<td colspan="2"><% =GF_FN2DTE(fecContable) & "  " & GF_TRADUCIR(" al ") & "  " & GF_FN2DTE(fecContableH) %></td>		
		<td align="right"><b><% = GF_TRADUCIR("Cod. Cupo") %>:</b></td>
		<td ><% =codCupo %></td>		
    </tr>                     
    <tr>
		<td align="right"><b><% = GF_TRADUCIR("Nro C. Porte") %>:</b></td>
		<td ><% =nuCartaPorte1 %>-<% =nuCartaPorte2 %></td>
		<td></td>
		<td align="right"><b><% = GF_TRADUCIR("Pat. Chasis") %>:</b></td>
		<td ><% =patChasis1 & "-" & patChasis2 %></td>
		<td></td>
		<td align="right"><b><% = GF_TRADUCIR("Pat. Acoplado") %>:</b></td>
		<td ><% =patAcoplado1 & "-" & patAcoplado2 %></td>		
    </tr>                     
    <tr>
		<td align="right"><b><% = GF_TRADUCIR("Cliente") %>:</b></td>
		<td colspan="2"><%=cdCliente & "-" & dsCliente%></td>		
		<td align="right"><b><% = GF_TRADUCIR("Corredor") %>:</b></td>
		<td colspan="2"><%=cdCorredor & "-" & dsCorredor%></td>
		<td align="right"><b><% = GF_TRADUCIR("Vendedor") %>:</b></td>
		<td colspan="2"><%=cdVendedor & "-" & dsVendedor%></td>									
    </tr>  
    <tr>
		<td align="right"><b><% = GF_TRADUCIR("Chofer") %>:</b></td>
		<td colspan="2"><%=cdChofer & "-" & dsChofer%></td>
		<td align="right"><b><% = GF_TRADUCIR("Transportista") %>:</b></td>
		<td colspan="2"><%=cdTransportista & "-" & dsTransportista%></td>
		<td align="right"><b><% = GF_TRADUCIR("Tipo Camión") %>:</b></td>
		<td >
		    <%  xlsCircuito = GF_Traducir("TODOS")
		        if (cdCircuito = CIRCUITO_CAMION_CARGA) then xlsCircuito = GF_Traducir("CARGA")
		        if (cdCircuito = CIRCUITO_CAMION_DESCARGA) then xlsCircuito = GF_Traducir("DESCARGA")							            
		        response.Write xlsCircuito
		     %>								
		</td>
    </tr>  
    <tr>
		<td align="right"><b><% = GF_TRADUCIR("Producto") %>:</b></td>
		<td colspan="2"><% =getDsProducto(cdProducto) %></td>
		<td align="right"><b><% = GF_TRADUCIR("Estado") %>:</b></td>
		<td colspan="2">
		    <%  Call executeQueryDb(pto, rsEstados, "OPEN", "SELECT * FROM ESTADOS where CDESTADO=" & cdEstado)				
		        xlsEstado="TODOS"
		        if (not rsEstados.eof) then xlsEstado = rsEstados("DSESTADO")
			    response.Write xlsEstado	
			%>																
		</td>
		<td align="right"><b><% = GF_TRADUCIR("Solo c/Muestras Extra") %>:</b></td>
		<td>
		    <%
		        if (chkMuestrasAud = MUESTRAS_AUDITORIA_ONLY) then
		            response.Write "SI"
		        else
		            response.Write "NO"
		        end if
		     %>							    
		</td>
    </tr>  
    <tr><td>&nbsp;</td></tr>
</table>

<% Call crearTabla(mostrar, accion) %>

</body>
</html>