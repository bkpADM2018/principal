<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%

'-----------------------------------------------------------------------------------------------------------

Function CalcularMerma(pCdProducto, pCdRubro, pValor, ByRef pMerma, ByRef pEsZaranda)
    Dim strSQL, rs
    
    pMerma = 0
    pEsZaranda = false
    if (InStr(1, g_listaRubrosHumedad, "," & pCdRubro & ",") > 0) then
        strSQL= "Select (VLMERMAXTABLA + MERMAXMANIPULEO) PORCMERMA from " &_                
                " MERMAXSECADO MXS " &_
                " INNER JOIN GASTOSXSECADO GXS ON GXS.CDPRODUCTO=MXS.CDPRODUCTO" &_
                " where MXS.CDPRODUCTO=" & pCdProducto & " and MXS.VLHUMEDAD=" & pValor                
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
        if (not rs.eof) then pMerma = CDbl(rs("PORCMERMA"))
    end if    
    if (InStr(1, g_listaRubrosZaranda, "," & pCdRubro & ",") > 0) then
        pEsZaranda = true
        strSQL="Select * from MERMASAUTOMATICASPENALIZACION where CDPRODUCTO=" & pCdProducto & " and CDRUBRO=" & pCdRubro & " and VALORMINIMO<=" & pValor & " and VALORMAXIMO>=" & pValor                
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
        if (not rs.eof) then pMerma = CDbl(rs("MERMAVARIABLE"))
    end if     
End Function
'-----------------------------------------------------------------------------------------------------------
Function corteCamion(pRs, pDtContable, pIdCamion)
    Dim ret 
    
    ret = true
    if (not pRs.eof) then
        if ((pRs("DTCONTABLE") = pDtContable) and (pRs("IDCAMION") = pIdCamion)) then ret = false
    end if    
    corteCamion = ret
End Function
'-----------------------------------------------------------------------------------------------------------
Function corteVagon(pRs, pDtContable, pCdOperativo, pCdVagon)
    Dim ret 
    
    ret = true
    if (not pRs.eof) then
        if ((pRs("DTCONTABLE") = pDtContable) and (pRs("CDOPERATIVO") = pCdOperativo) and (pRs("CDVAGON") = pCdVagon)) then ret = false
    end if    
    corteVagon = ret
End Function
'-----------------------------------------------------------------------------------------------------------
'                                               CAMIONES
'-----------------------------------------------------------------------------------------------------------
Function procesarCamiones(pPto, pDtContable)

    Dim strSQL, rs1, myTotalMermaDetalleCalc, myMerma, myTotalMermaDetalleTabla, strLog, vecRubros(100, 5), cantOK, cantError
    Dim ret, esZaranda, printIt

    '--Ppia Produccion
    strSQL="Select * from (" &_
           "Select  HCD.NUCARTAPORTE,   " &_
           "        HCD.CTG,     " &_ 
           "        HCD.CDCORREDOR,     " &_ 
           "        HCD.CDVENDEDOR,     " &_
           "        'V' PpiaProd,    " &_
           "        DO.NCESTABLEPROCE,    " &_
           "        DO.CRAZONDESTINATARIO,           " &_
           "        HC.NUCUITREM, " &_         
           "        '' NUDOCUMENTO, " &_
           "        CDPRODUCTO " &_
           "        from HCAMIONESDESCARGA HCD  " &_
           "        inner join HCAMIONES  HC on HC.DTCONTABLE=HCD.DTCONTABLE and HC.IDCAMION=HCD.IDCAMION   " &_
           "        inner join DATOSONCCA DO on HCD.NUCARTAPORTE=DO.NCCARTAPORTE    " &_
           "        inner join VENDEDORES V on V.CDVENDEDOR=HCD.CDVENDEDOR  " &_
           "        where HCD.DTCONTABLE='" & pDtContable & "'   " &_
           "        and CAST(DO.NCESTABLEPROCE as BIGINT) = 1     " &_
           "        and HCD.CDCLIENTE=1 " &_
           "        and HC.NUCUITREM = V.NUDOCUMENTO    " &_
           "        and HC.CDESTADO in (6, 8)" &_
           " Union  "
    '--No Ppia Produccion
    strSQL= strSQL & "Select  HCD.NUCARTAPORTE,   " &_
           "        HCD.CTG,     " &_ 
           "        HCD.CDCORREDOR,     " &_           
           "        HCD.CDVENDEDOR,     " &_
           "        'F' PpiaProd,    " &_
           "        DO.NCESTABLEPROCE,   " &_
           "        DO.CRAZONDESTINATARIO,   " &_
           "        HC.NUCUITREM, " &_
           "        case when HC.NUCUITREM = V.NUDOCUMENTO then '' else V.NUDOCUMENTO end NUDOCUMENTO,   " &_
           "        CDPRODUCTO " &_
           "        from HCAMIONESDESCARGA HCD  " &_
           "        inner join HCAMIONES  HC on HC.DTCONTABLE=HCD.DTCONTABLE and HC.IDCAMION=HCD.IDCAMION   " &_
           "        inner join DATOSONCCA DO on HCD.NUCARTAPORTE=DO.NCCARTAPORTE    " &_
           "        inner join VENDEDORES V on V.CDVENDEDOR=HCD.CDVENDEDOR  " &_
           "        where HCD.DTCONTABLE='" & pDtContable & "'   " &_
           "        and (cast(DO.NCESTABLEPROCE as BIGINT) <> 1   " &_
	       "            or HCD.CDCLIENTE<>1 " &_
	       "            or HC.NUCUITREM <> V.NUDOCUMENTO)" &_	               
	       "        and HC.CDESTADO in (6, 8)" &_
	       ") T order by NUCARTAPORTE"	       
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    'Obtengo el contrato donde esta aplicada la descarga
    strSQL="Select DISTINCT CPROR6, CSUCR6, COPER6, NCTOR6, ACOSR6, CPORR6, MCPDRJ from MERFL.MER311F6 left join MERFL.MER311FJ on CPRORJ=CPROR6 and CSUCRJ=CSUCR6 and COPERJ=COPER6 and NCTORJ=NCTOR6 and ACOSRJ=ACOSR6 where KGNER6>0 and FECDR6=" & Replace(pDtContable, "-", "") & " order by CPORR6"
    'response.Write strSQL
    Call executeQuery(rs2, "OPEN", strSQL)        
    while (not rs.eof)        
        printIt = false
        strCtos = ""
        if (rs2.eof) then                    
            strCtos = "SIN APLICACION"
            printIt = true
        else
            salir=false
            while ((not rs2.eof) and (not salir))
                if (CDbl(rs2("CPORR6")) < CDbl(rs("NUCARTAPORTE"))) then
                    rs2.MoveNext()
                else
                    salir=true                        
                end if
            wend 
            if (rs2.eof) then                    
                strCtos = "SIN APLICACION"
                printIt = true
            else           
                if (CDbl(rs2("CPORR6")) > CDbl(rs("NUCARTAPORTE"))) then                         
                    strCtos = "SIN APLICACION"
                    printIt = true
                else                        
                    salir=false
                    while ((not rs2.eof) and (not salir))
                        if (CDbl(rs2("CPORR6")) = CDbl(rs("NUCARTAPORTE"))) then
                        'Si es la misma carta de porte - hay aplicacion.                       
                            if ((rs2("MCPDRJ") = rs("PpiaProd")) or (CInt(rs2("COPER6")) = 4)) then
                                aplicacion = "OK"
                            else
                                aplicacion = "ERROR!"                                    
                                printIt = true
                            end if                                
                            if (printIt) then strCtos = strCtos & rs2("CPROR6") & "-" & rs2("CSUCR6") & "-" & rs2("COPER6") & "-" & rs2("NCTOR6") & "/" & rs2("ACOSR6") & "  (" & aplicacion & ")<br />"
                            rs2.MoveNext()
                        else
                            salir=true                                                                 
                        end if                
                     wend                                                                
                end if                         
            end if
        end if       
        if (printIt) then            
%>        
    <tr>
        <td><% =rs("NUCARTAPORTE") %></td>
        <td><% =rs("CTG") %></td>
        <td><% =rs("CDCORREDOR") %></td>
        <td><% =rs("CDVENDEDOR") %></td>
        <td align="center"><% =rs("CDPRODUCTO") %></td>
        <td align="center"><% =rs("PpiaProd") %></td>
        <td align="center"><% =strCtos %></td>
        <td align="center"><% if (CLNG(rs("NCESTABLEPROCE")) < 100) then response.Write "S/D" else response.Write rs("NCESTABLEPROCE") end if %></td>
        <td><% =rs("CRAZONDESTINATARIO") %></td>
        <td align="center"><% =rs("NUCUITREM") %></td>
        <td align="center"><% =rs("NUDOCUMENTO") %></td>
    </tr>    
<%       
        end if               
        rs.MoveNext()
    wend
    
End Function
'-----------------------------------------------------------------------------------------------------------
'                                               VAGONES
'-----------------------------------------------------------------------------------------------------------
Function procesarVagones(pDtContable, pAction)
   
             
End Function
'***********************************************
'*****      COMIENZO DE LA PAGINA          *****
'***********************************************
Dim strSQL, rs, DtContable, myHoy, g_listaRubrosHumedad, g_listaRubrosZaranda, listaRubros
Dim myHasta, transporte, myNext, logMig, rsX, resultC, resultV, action

'----------------------------------
'PARAMETROS DEL APLICATIVO
'
' pto: Puerto ARROYO, TRANSITO, PIEDRABUENA
' f: Fecha a controlar [Opcional]
' t: Transporte C: Camiones, V:Vagones [Opcional]
'----------------------------------
g_strPuerto = GF_PARAMETROS7("pto", "", 6)
'Asumo que se va a migrar los datos del dia de ayer.
myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))
myHoy = GF_DTEADD(myHoy,-1,"D")
'Si pidio una fecha particular tomo esa fecha.
if (GF_PARAMETROS7("f", 0, 6) <> 0) then myHoy = GF_PARAMETROS7("f", 0, 6)
DtContable = Left(myHoy, 4) & "-" & mid(myHoy, 5, 2) & "-" & Right(myHoy, 2)	 

transporte = "T"
if (GF_PARAMETROS7("t", "", 6) <> "") then transporte = GF_PARAMETROS7("t", "", 6)

%>
<html>
<head>
</head>
<body onload="bodyOnLoad()">
PUERTO:<% =g_strPuerto %><br />
FECHA DESCARGA: <% =DtContable %><br />
Mmto Generacion: <% =day(Now()) & "/" & month(Now()) & "/" & year(Now()) & " " & hour(Now()) & ":" & minute(Now()) & ":" & second(Now()) %><br /> 
<table border="1">
    <tr>
        <td rowspan="2" align="center">Carta de Porte</td>
        <td rowspan="2" align="center">CTG</td>
        <td rowspan="2" align="center">Corredor</td>
        <td rowspan="2" align="center">Vendedor</td>
        <td rowspan="2" align="center">Producto</td>
        <td rowspan="2" align="center">Ppia Prod</td>
        <td rowspan="2" align="center">Contratos</td>
        <td colspan="4" align="center">CONTROL</td>
        
    </tr>
    <tr>        
        <td>Estab. Proce.</td>
        <td>Destinatario</td>
        <td>Titular</td>
        <td>Rte. Comercial</td>
    </tr>
<%
if ((transporte = "C") or (transporte = "T")) then Call procesarCamiones(g_strPuerto, DtContable)
'if ((transporte = "V") or (transporte = "T")) then resultV = procesarVagones(g_strPuerto, DtContable)
%>
</table>
</body>
</html>