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
Dim idCotizacion, rsPIC

idCotizacion = GF_PARAMETROS7("idCotizacion","",6)

strSQL = " SELECT A.FECHSA AS FECHAEMISIONFAC, "&_
	     "        B.MOMENTO AS FECHAEMISIONPIC "&_
         " FROM MERFL.MER301F1 A "&_
         " LEFT JOIN TOEPFERDB.TBLCTZCABECERA B ON A.MINUSA = B.IDCOTIZACION "&_
         " WHERE A.MINUSA = "&idCotizacion& " AND A.EVENSA = '"& AUTH_TYPE_PICF &"'"
Call executeQuery(rsPIC, "OPEN", strSQL)

%>
<html>
<body>
<div>
	<hr /><h3>INFORMACION FECHA FACTURA</h3><hr />
</div>
    <div class="tableaside size100">
	    <table class="datagrid" align="center" width="80%">
            <thead>
	            <tr>
                    <th width="70%" align="center" class="thicon"><%=GF_Traducir("Fecha emisión Factura")%></th>   
                    <th width="30%" align="center" class="thicon"><%=GF_Traducir("Fecha emisión PIC")%></th>
			    </tr>
	        </thead>
		    <tbody>
		    <%	if (not rsPIC.eof) then 
                     while (not rsPIC.eof) %>
		                <tr>
                            <td align="center"><%=GF_FN2DTE(rsPIC("FECHAEMISIONFAC")) %></td>
                            <td align="center"><%=GF_FN2DTE(Left(rsPIC("FECHAEMISIONPIC"),8))%></td>
                        </tr>
                  <%    rsPIC.movenext()
                     wend %>
		    <%  else  %>
			        <tr><td colspan="2" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td></tr>
		    <%  end if  %>
		    </tbody>
	    </table>
    </div>
</body>
</html>