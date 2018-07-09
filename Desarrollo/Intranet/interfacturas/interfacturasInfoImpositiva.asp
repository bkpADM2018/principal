<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<%
Dim reg, rs , g_Observaciones ,rs_percepcion
Dim strSQL, simboloMoneda, myIVA , myPercepcionIIBB
Dim myGravado, myNoGravado , myPersepcionIVA , myTasaIVA

    reg = GF_PARAMETROS7("id","",6)

	strSQL= "Select FCCMTF, " &_
	        " CASE WHEN FAIMPN IS NULL THEN 0 ELSE FAIMPN*100 END AS GRAVADO, " &_
	        " CASE WHEN FAIMPN IS NULL THEN 0 ELSE FAIMPN*100 END AS NO_GRAVADO, " &_
	        " CASE WHEN FAIVAI IS NULL THEN 0 ELSE FAIVAI*100 END AS IVA, " &_
	        " CASE WHEN FAIMP IS NULL THEN 0 ELSE FAIMP*100 END AS PERCEPCION_IVA, " &_
	        " CASE WHEN FAIMP2 IS NULL THEN 0 ELSE FAIMP2*100 END AS PERCEPCION_IIBB, " &_
	        " FCMNCD, " &_
	        " CASE WHEN D.FASTOT IS NULL THEN 0 ELSE D.FASTOT*100 END AS SUB_TOTAL_GRAVADO, " &_
	        " F2IIPR TASA_IVA " &_
		    " from TFFL.TF100F1 C " &_
		    " left join TFFL.TF111 D on C.FCRGNR=D.FANRRG " &_
		    " left join TFFL.TF102F1 IVA on C.FCRGNR=IVA.F2RGNR " &_
		    " where FCRGNR=" & reg
    Call executeQuery(rs, "OPEN", strSQL)
    
    if (not rs.Eof) Then
        simboloMoneda = getSimboloMoneda(rs("FCMNCD"))
        myPercepcionIIBB    = rs("PERCEPCION_IIBB") 
        myTasaIVA           = rs("TASA_IVA")
        myIVA               = rs("IVA")
        myPersepcionIVA     = rs("PERCEPCION_IVA")    
        'Si es una factura B, no se debe mostrar el importe de IVA discriminado.
        myGravado = Cdbl(rs("SUB_TOTAL_GRAVADO")) - Cdbl(rs("GRAVADO"))
        myNoGravado = rs("NO_GRAVADO")
        if (rs("FCCMTF") = "B") then             
            myGravado = 0
            myNoGravado = Cdbl(rs("SUB_TOTAL_GRAVADO")) - Cdbl(rs("GRAVADO"))  + CDbl(rs("NO_GRAVADO"))
        end if
        if (CDbl(myPercepcionIIBB) > 0) then
            'Obtengo el detalle de las percepciones
            strSQL="Select * from TFFL.TF114 D inner join MERFL.MER1K2F1 P on P.CODIPO=D.FAPRO4 where FANRR4=" & reg
            Call executeQuery(rs_percepcion, "OPEN", strSQL)
            if (not rs_percepcion.eof) then
                g_Observaciones = g_Observaciones & "Det IIBB:"
                while (not rs_percepcion.eof)
                    g_Observaciones = g_Observaciones & " " & Trim(rs_percepcion("DESCPO")) & "=" & simboloMoneda & " " & GF_EDIT_DECIMALS(CDbl(rs_percepcion("FAIMP4"))*100, 2)
                    rs_percepcion.MoveNext()
                wend        
            end if
        end if
    end if

%>
<html>
<head>
    <link rel="stylesheet" href="../css/main.css" type="text/css">
</head>
<body>
<div>
	<hr /><h3>INFORMACION APERTURA IMPOSITIVA</h3><hr />
</div>

<table class="datagrid" align="center" width="80%">
<tbody>
        <thead>
            <tr>
                <th align="center">CONCEPTO</th>
                <th align="center">IMPORTE</th>
            </tr>
        </thead>
<%	if (not rs.eof) then %>
		<tr>
		    <td align="left"><%=GF_TRADUCIR("Gravado")%></td>
		    <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(myGravado, 2) %></td>
		</tr>
        <tr>
		    <td align="left"><%=GF_TRADUCIR("No Gravado")%></td>
		    <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(myNoGravado, 2) %></td>
		</tr>
        <tr>
		    <td align="left"><%=GF_TRADUCIR("IVA Inscr.") & " (" & myTasaIVA & "%)" %></td>
		    <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(myIVA, 2) %></td>
		</tr>
        <tr>
		    <td align="left"><%=GF_TRADUCIR("Percepcion IVA")%></td>
		    <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(myPersepcionIVA, 2) %></td>
		</tr>
        <tr>
		    <td align="left"><%=GF_TRADUCIR("Percepcion IIBB")%></td>
		    <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(myPercepcionIIBB, 2) %></td>
		</tr>
        <tr>
            <td colspan="2" align="left"><% =g_Observaciones %></td>
        </tr>
<%  else  %>
		<tr><td colspan="2" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td></tr>
<%  end if  %>
</tbody>
</table>
</body>
</html>
