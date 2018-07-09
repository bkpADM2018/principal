<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosFacturacionCalidad.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<%
'-----------------------------------------------------------------------------------------------------------------------------------
Function getMatrizDescarga(pStrPuerto)
    Dim rsClientes, rsProductos, intCantCamionesAux, intTotalesCliente, intTotales
	
	Dim myPrecioPto1, myPrecioPto2, myPrecioPto3, myMonedaSecado, rs
       
    '***************
    '*** SECADO ***
    '***************
    strSQL="SELECT PTODESDE, PRECIO,PRECIOADICIONAL"&_
            " FROM PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_SECADO & " ORDER BY VIGENCIADESDE DESC, PTODESDE ASC" 			
    Call executeQueryDb(pStrPuerto, rs, "OPEN", strSQL)    
    myPrecioPto1 = 0
    myPrecioPto2 = 0
    myPrecioPto3 = 0
    myMonedaSecado = MONEDA_DOLAR
    if (not rs.eof) then 
        myPrecioPto1 = rs("PRECIO")
        rs.MoveNext()
    end if        
    if (not rs.eof) then         
        myPrecioPto2 = rs("PRECIO")
        rs.MoveNext()
    end if        
    if (not rs.eof) then         
        myPrecioPto3 = rs("PRECIO")        
        myPrecioAdic = rs("PRECIOADICIONAL")
    end if
    '***************
    '*** ZARANDA ***
    '***************
    strSQL="SELECT CDMONEDA,PRECIO FROM PRECIOSERVICIOS where CDCONCEPTO= " & SERVICIO_ACOND_ZARANDA & " ORDER BY VIGENCIADESDE DESC"	
    Call executeQueryDb(pStrPuerto, rs, "OPEN", strSQL)   
    myPrecioZaranda = 0
    myMonedaZaranda = MONEDA_PESO
    if (not rs.eof) then 
        myPrecioZaranda = rs("PRECIO")
        myMonedaZaranda = rs("CDMONEDA")
    end if        
%>
	<table width="100%">	        
	        <tr style="height: 30px">
	          <td style="border:2px solid #396E8F; text-align: center;" width="70%"><% =GF_TRADUCIR("SECADO (" & getSimboloMoneda(myMonedaSecado) & " x Tn):") %> Pto 1: <% =GF_EDIT_DECIMALS(CDbl(myPrecioPto1)*100, 2) %> / Pto 2: <% =GF_EDIT_DECIMALS(CDbl(myPrecioPto2)*100, 2) %> / Pto 3: <% =GF_EDIT_DECIMALS(CDbl(myPrecioPto3)*100, 2) %> / Pto. Adic: <% =GF_EDIT_DECIMALS(CDbl(myPrecioAdic)*100, 2) %></td>	          
	          <td style="border:2px solid #396E8F; text-align: center;"><% =GF_TRADUCIR("ZARANDEO (" & getSimboloMoneda(myMonedaZaranda) & " x Tn): ") & GF_EDIT_DECIMALS(CDbl(myPrecioZaranda)*100, 2) %> </td>
            </tr>            
    </table>    
    <br />
<%	
	'****************
	'*** CAMIONES ***
	'****************
	strSQL="Select CL.CDCLIENTE, CL.DSCLIENTE, P.CDPRODUCTO, DSPRODUCTO, count(*) CANT " &_
			" from CAMIONES C " &_
			"	inner join CAMIONESDESCARGA CD on C.IDCAMION=CD.IDCAMION " &_
			"	inner join PRODUCTOS P on P.CDPRODUCTO=C.CDPRODUCTO " &_
			"	inner join CLIENTES CL on CL.CDCLIENTE=CD.CDCLIENTE " &_		
			" where C.cdEstado <> " & CAMIONES_ESTADO_BAJA		
	if (not IsToepfer(session("KCOrganizacion"))) then				
		strSQL = strSQL & " and	( CD.cdcliente in (Select CDCLIENTE from clientes where NUCUIT = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & " OR CD.cdvendedor in (Select CDVENDEDOR from VENDEDORES where NUDOCUMENTO = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & " OR CD.cdcorredor in (Select CDCORREDOR from CORREDORES where NUCUIT = '" & session("CuitOrganizacion") & "')) "
	end if
	strSQL = strSQL & " group by CL.CDCLIENTE, CL.DSCLIENTE, P.CDPRODUCTO, DSPRODUCTO" &_
			" order by CL.DSCLIENTE, DSPRODUCTO "
	Call drawMovTable(pStrPuerto, strSQL, "Descargas", 1)	
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
Function getMatrizCargas(pStrPuerto)
    
	Dim strSQL
	
	'****************
	'*** CAMIONES ***
	'****************
	strSQL="Select CL.CDCLIENTE, CL.DSCLIENTE, P.CDPRODUCTO, DSPRODUCTO, count(*) CANT " &_
			" from CAMIONES C " &_
			"	inner join CAMIONESCARGA CD on C.IDCAMION=CD.IDCAMION " &_
			"	inner join PRODUCTOS P on P.CDPRODUCTO=C.CDPRODUCTO " &_
			"	inner join CLIENTES CL on CL.CDCLIENTE=CD.CDCLIENTE " &_		
			" where C.cdEstado <> " & CAMIONES_ESTADO_BAJA
	if (not IsToepfer(session("KCOrganizacion"))) then				
		strSQL = strSQL & "	and (CD.cdcliente in (Select CDCLIENTE from clientes where NUCUIT = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & " OR CD.cdvendedor in (Select CDVENDEDOR from VENDEDORES where NUDOCUMENTO = '" & session("CuitOrganizacion") & "') "
		strSQL = strSQL & " OR CD.cdcorredor in (Select CDCORREDOR from CORREDORES where NUCUIT = '" & session("CuitOrganizacion") & "')) "
	end if
	strSQL = strSQL & " group by CL.CDCLIENTE, CL.DSCLIENTE, P.CDPRODUCTO, DSPRODUCTO" &_
			" order by CL.DSCLIENTE, DSPRODUCTO "
	Call drawMovTable(pStrPuerto, strSQL, "Cargas", 2)	
	
End Function
'-------------------------------------------------------------------------------------------------------------
Function drawMovTable(pStrPuerto, query, pTitle, pIdType)	
	Dim rs, arrMovimientos(100, 100), rowCount, colCount
	Dim oldClient, oldProduct, colIdx

%>
	<table align="center" cellspacing="1" cellpadding="1" width="100%">	
<%	
	Call executeQueryDb(pStrPuerto, rs, "OPEN", query)

	if (not rs.eof) then
		'Paso los movimientos a una tabla.
		rowCount = 1 'Nro maximo de filas
		colCount = 1 'Nro maximo de columnas
		colIdx = 0	 'Columna del producto buscado.
		oldClient = 0
		oldProduct = 0
		arrMovimientos(1, 1) = pTitle
		arrMovimientos(1, 100) = "Total"
		arrMovimientos(100, 1) = "Total"	
		while (not rs.eof)
			if (CLng(oldClient) <> CLng(rs("CDCLIENTE"))) then 
				rowCount = rowCount + 1
				oldClient = rs("CDCLIENTE")
				arrMovimientos(rowCount, 0) = oldClient
				arrMovimientos(rowCount, 1) = Trim(rs("DSCLIENTE"))
			end if
			if (CLng(oldProduct) <> CLng(rs("CDPRODUCTO"))) then 				
				colIdx = 0
				For i = 2 to colCount
					if (CLng(arrMovimientos(0, i)) = CLng(rs("CDPRODUCTO"))) then colIdx = i					
				Next
				if (colIdx = 0) then
					'No se encontro el producto en la tabla, se agrega.							
					colCount = colCount + 1
					colIdx = colCount					
					arrMovimientos(0, colIdx) = Trim(rs("CDPRODUCTO"))
					arrMovimientos(1, colIdx) = Trim(rs("DSPRODUCTO"))
				end if
				oldProduct = rs("CDPRODUCTO")
			end if
			'Grabo el valor del Cliente-Producto
			arrMovimientos(rowCount, colIdx) = rs("CANT")
			'Grabo los totales
			arrMovimientos(100, colIdx) = CLng(arrMovimientos(100, colIdx)) + CLng(rs("CANT"))
			arrMovimientos(rowCount, 100) = CLng(arrMovimientos(rowCount, 100)) + CLng(rs("CANT"))
			arrMovimientos(100, 100) = CLng(arrMovimientos(100, 100)) + CLng(rs("CANT"))
			rs.MoveNext()
		wend		
		'Mustra la tabla.
%>    
        <thead>
            <tr>				
			<% For i = 1 to colCount %>
	            <th align="center"><% =arrMovimientos(1, i) %></th>
			<% Next %>
				<th align="center"><% =arrMovimientos(1, 100) %></th>
            </tr>
        </thead>
        <tbody>			
			<%  For y = 2 to rowCount	%>
			<tr>				
					<td align="center" style="border: 1px solid #ccc;background-color:#396E8F;color:#fff;font-size:11px;font-weight:bold;" width="20%"><% =arrMovimientos(y, 1) %></td>
			<%		For x = 2 to colCount %>
						<td align="center" style="border: 1px solid #ccc;">
							<a href='javascript:irConsultaCamiones(<%=arrMovimientos(y, 0) %>,<%=arrMovimientos(0, x)%>, <% =pIdType %>)' style="color:#2e6b4d;"><% =arrMovimientos(y, x) %></a>
						</td>
			<% 		Next	%>
					<td class="rtotal" style="border: 1px solid #ccc;background-color:#999; color: #FFFFFF; font-weight:bold;" align="center">
						<a style="font-weight:bold;" href='javascript:irConsultaCamiones(<%=arrMovimientos(y, 0) %>, 0, <% =pIdType %>)' style="color:#2e6b4d;"><% =arrMovimientos(y, 100) %></a>
					</td>
			</tr>
			<%	Next %>	
			<!-- Linea totales x producto -->
			<tr class="rtotal">			
			<% For i = 1 to colCount %>
				<td align="center">
					<a style="font-weight:bold;" href='javascript:irConsultaCamiones("", <%=arrMovimientos(0, i) %>, <% =pIdType %>)' >
						<% =arrMovimientos(100, i) %>
					</a>
				</td>
			<% Next %>
				<td align="center"><a title='TODOS' href='javascript:irConsultaCamiones("",0, <% =pIdType %>)' style="font-weight:bold;"><% =arrMovimientos(100, 100) %></a></td>
			</tr>			
		</tbody>
	<% else %>	
		<thead>
			<th align="center"><% =GF_Traducir(pTitle) %></th>
		<thead>
		<tbody>
			<tr><TD align="center"><%=GF_Traducir("No se regsitra descarga de camiones")%></TD></tr>	
		</tbody>
	<% end if %>		
    </table>
    <%
End Function
'********************************************************************************************************************************************
'******************************************************** COMIENZO DE PAGINA ****************************************************************
'********************************************************************************************************************************************
Dim g_strPuerto
g_strPuerto = GF_Parametros7("pto","",6)

%>
<html>
<head>
   <link rel="stylesheet" type="text/css" href="../css/main.css" />
   <script type="text/javascript">
   function irConsultaCamiones(pCliente, pProducto, pCircuito) {
       window.open('consultaCamiones.asp?pto=<%=g_strPuerto%>&cdCliente=' + pCliente + '&cdProducto=' + pProducto + '&cdCircuito=' + pCircuito, '<%=g_strPuerto%>', 'width=940,height=640,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
   }
   </script> 
</head>
<body style="background: url('../images/descargas_BG.png') 0px 30px no-repeat;">
    <% Call getMatrizDescarga(g_strPuerto) %>
    <br />
    <% Call getMatrizCargas(g_strPuerto) %>
</body>
</html>