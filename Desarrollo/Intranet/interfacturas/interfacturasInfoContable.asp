<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<%
    Dim reg, rs
    reg = GF_PARAMETROS7("id","",6)
    Call executeSP(rs, "TFFL.TF101F1_GET_TOTALIZAR_X_CUENTA_CCOSTO", reg)
%>
<html>
<head>
    <link rel="stylesheet" href="../css/main.css" type="text/css">
</head>
<body>
<div><hr /><h3>INFORMACION CONTABLE</h3><hr /></div>
    <table class="datagrid" align="center" width="80%">
        <thead>
            <tr>
                <th>TIPO</th>
                <th>CONCEPTO</th>
                <th>CUENTA</th>
                <th>C. COSTOS</th>
                <th>IMPORTE</th>
            </tr>
        </thead>
        <tbody>
<%    
    if (not rs.eof) then
        while (not rs.eof)
			myDecome = rs("DECOME")
			if isnull(myDecome) then myDecome = 0
            if ((CInt(myDecome) <> 0) or _
                (Trim(rs("FDCTCD")) <> "") or _
                (CDbl(rs("IMPORTE")) <> 0)) then
%>    
            <tr>
                <td align="center"><% =GF_TRADUCIR("Items") %></td>                
                <td align="center"><% =rs("DECOME") %></td>
                <td align="center"><% =rs("FDCTCD") %></td>
                <td align="center"><% =rs("FDCTCS") %></td>
                <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(CDbl(rs("IMPORTE"))*100, 2) %></td>
            </tr>    
<%    
            end if
            rs.MoveNext()
        wend
        rs.MoveFirst()
%>    
            <tr>
                <td align="center" colspan="2"><% =GF_TRADUCIR("Contrapartida") %></td>                
                <td align="center"><% =rs("FCCTCD") %></td>
                <td align="center"><% =rs("FCCTCS") %></td>
                <td align="right"><% =getSimboloMoneda(rs("FCMNCD")) & " " & GF_EDIT_DECIMALS(CDbl(rs("FCTTGR"))*100, 2) %></td>
            </tr>
<%
    end if
%>         
        </tbody>
    </table>
</body>
</html>