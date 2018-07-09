<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
Dim nroLote,fechaLote
nroLote   = GF_PARAMETROS7("nroLote",0,6)
fechaLote = GF_PARAMETROS7("fechaLote","",6)

Call executeSP(rsFir, "EJIFL.TBLPROVISIONESFIRMAS_GET_BY_PARAMETERS", nroLote &"||"& fechaLote)

%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE PROVISIONES - FIRMAS REGISTRADAS</title>
	<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="css/main.css" />	
    
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript">
		
	</script>
</head>
<BODY>
	<div class="col66"></div>
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Nro.Lote")%></div>
        <div class="col26"> <%= nroLote %>  </div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Lote")%></div>
        <div class="col26"> <%= GF_FN2DTE(fechaLote) %>  </div>
    </div>

    <div class="col66"></div>

	<%if not rsFir.eof then%>
	<table class="datagrid datagridlv1" width="100%" >
        <thead>
            <tr>
                <th width="30%" align="center"> <%=GF_Traducir("Usuario")%> </th>
                <th width="25%" align="center"> <%=GF_Traducir("Fecha firma ")%> </th>
                <th width="25%" align="center"> <%=GF_Traducir("Hora firma")%> </th>
            </tr>
        </thead>
        <tbody>
		<%  while not rsFir.eof %>
			<tr>				
				<td align="left"><%=getUserDescription(rsFir("CDUSUARIO"))%></td>
				<td align="center"><%=left(GF_FN2DTE(rsFir("FECHAFIRMA")),10)%></td>
				<td align="center"><%=Right(GF_FN2DTE(rsFir("FECHAFIRMA")),8)%></td>
			<tr>
		<%	rsFir.MoveNext()
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