<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<%
'------------------------------------------------------------------------------------------------------------------------
Dim idFactura,strSQL,rs,comprobante
idFactura = GF_PARAMETROS7("idFactura","",6)
comprobante = GF_PARAMETROS7("cdComprobante","",6)

strSQL = "SELECT * FROM TFFL.TF105F1 WHERE FCRGSG ="& idFactura &" ORDER BY FCSQSG"
Call executeQuery(rs, "OPEN", strSQL)

%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE FACTURACION - FIRMAS REGISTRADAS</title>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="../css/main.css" />	
    
	<script type="text/javascript" src="../scripts/channel.js"></script>
	<script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript">			
		
	</script>
</head>
<BODY>
	<div class="col66"></div>
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Nro.Registro")%></div>
        <div class="col26"> <%= idFactura %>  </div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Comprobante")%></div>
        <div class="col26"> <%= comprobante %>  </div>
    </div>
        

<div class="col66"></div>

	<%if not rs.eof then%>
	<table class="datagrid datagridlv1" width="100%" >
        <thead>
            <tr>
                <th width="30%" align="center"> <%=GF_Traducir("Usuario")%> </th>
                <th width="25%" align="center"> <%=GF_Traducir("Fecha firma ")%> </th>
                <th width="25%" align="center"> <%=GF_Traducir("Hora firma")%> </th>
            </tr>
        </thead>
        <tbody>
		<%  while not rs.eof %>
			<tr>				
				<td align="left"><%=getUserDescription(rs("CUSRSG"))%></td>
				<td align="center"><%=left(GF_FN2DTE(rs("MMTOSG")),10)%></td>
				<td align="center"><%=Right(GF_FN2DTE(rs("MMTOSG")),8)%></td>
			<tr>
		<%	rs.MoveNext()
		wend%>
        </tbody>	
	</table>	    
    <%else%>
    <table class="datagrid datagridlv1" width="100%" >
		<tr><td align="center">No se encontraron firmas registradas</td></tr>
	</table>	    
	<%end if%>

</BODY>
</html>