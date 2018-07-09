<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<% 
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim minuta, fecha, tipoCambioDia,rsTC,importePesosDia

minuta = GF_PARAMETROS7("minuta","",6)
fecha  = GF_PARAMETROS7("fecha","",6)

strSQL = "SELECT CASE WHEN B.CBIOX9 IS NULL THEN 0 ELSE B.CBIOX9 END AS TIPOCAMBIOCBTE, "&_
	     "       A.IMPORTEPESOSCBTE, "&_
	     "       A.FECHACBTE, "&_
	     "       A.MINUTA  "&_
         "FROM ( SELECT DSQJNB AS IMPORTEPESOSCBTE,  "&_
	     "	            DSQFNB AS MINUTA, "&_
         "              '20' || SUBSTR(DSCRDT,2,6) FECHACBTE "&_
         "   FROM PROVFL.ACDSREP     "&_
         "   WHERE DSQFNB = "& minuta &") A  "&_
         "LEFT JOIN PROVFL.ACD9REP B ON A.MINUTA = B.NINGX9 "
Call executeQuery(rsTC, "OPEN", strSQL)

%>
<html>
<body>
<div>
	<hr /><h3>INFORMACION TIPO DE CAMBIO</h3><hr />
</div>
    <div class="tableaside size100">
        <div class="alertmsj" style="height:15px;width:80%;margin:0 auto;">
            <%=GF_TRADUCIR("La factura posee un tipo de cambio que no se corresponde con el de su dia de emisión") %>
        </div>
	    <table class="datagrid" align="center" width="80%">
            <thead>
	            <tr>
                    <th width="70%" align="center" class="thicon"><%=GF_Traducir("Observaciones")%></th>   
                    <th width="30%" align="center" class="thicon"><%=GF_Traducir("Importe Total")%></th>
			    </tr>
	        </thead>
		    <tbody>
		    <%	if (not rsTC.eof) then 
                    tipoCambioDia= getTipoCambioCV(MONEDA_DOLAR, rsTC("FECHACBTE"), T_CAMBIO_VENDEDOR) 
                    importePesosDia = (Cdbl(rsTC("IMPORTEPESOSCBTE"))/Cdbl(rsTC("TIPOCAMBIOCBTE")))*Cdbl(tipoCambioDia)
                    diferenciaPesos =  Cdbl(rsTC("IMPORTEPESOSCBTE")) - Cdbl(importePesosDia)
                %>
		            <tr>
                        <td align="left"><%=GF_TRADUCIR("Tipo cambio Factura: ") & TIPO_MONEDA_PESO & " " & GF_EDIT_DECIMALS(Cdbl(rsTC("TIPOCAMBIOCBTE"))*100,2) %></td>
                        <td align="right"><%=TIPO_MONEDA_PESO & " " & GF_EDIT_DECIMALS(Cdbl(rsTC("IMPORTEPESOSCBTE"))*100,2)%></td>
                    </tr>
                    <tr>
                        <td align="left"><%=GF_TRADUCIR("Tipo cambio Día: ") & TIPO_MONEDA_PESO & " " & GF_EDIT_DECIMALS(Cdbl(tipoCambioDia)*100,2) %></td>
                        <td align="right"><%=TIPO_MONEDA_PESO & " " & GF_EDIT_DECIMALS(CDbl(importePesosDia)*100,2)%></td>
                    </tr>
                    <tr>
                        <td align="left"><%=GF_TRADUCIR("Diferencia")  %></td>
                        <td align="right"><%=TIPO_MONEDA_PESO & " " & GF_EDIT_DECIMALS(CDbl(diferenciaPesos)*100,2)%></td>
                    </tr>
		    <%  else  %>
			        <tr><td colspan="2" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td></tr>
		    <%  end if  %>
		    </tbody>
	    </table>
    </div>
</body>
</html>