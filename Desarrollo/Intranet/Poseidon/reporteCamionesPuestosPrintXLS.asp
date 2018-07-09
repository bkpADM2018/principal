<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<%

Const TRANSACCION_INGRESO = 1 'Esta transaccion incluye Ingresado y Sin cupo
Const TRANSACCION_VISTEO = 2 'Esta transaccion incluye Calados y Demorados
Const TRANSACCION_BRUTO = 7 'Esta transaccion incluye Pesado Bruto
Const TRANSACCION_TARA = 8 'Esta transaccion incluye Pesado Tara

Const ESTADO_INGRESADO = 1
Const ESTADO_SIN_CUPO = 10
Const ESTADO_CALADO = 2
Const ESTADO_DEMORADO = 11
Const ESTADO_PESADO_BRUTO = 5
Const ESTADO_PESADO_TARA = 8
'--------------------------------------------------------------------------------------------------------------------
Function imprimirFiltros(p_fechaInicio, p_fechaFin)
    
    call writeXLS("<tr><td colspan=7 align='right' style='font-weight:normal; font-size:10'>"&GF_FN2DTE(session("MmtoSistema")) &"</td></tr>")
    call writeXLS("<tr><td colspan=7 align='center' style='font-size:24'>" & GF_TRADUCIR("REPORTE DE CAMIONES POR PUESTO") &"</td></tr>")
    call writeXLS("<tr><td align='left' style='font-size:12'>" & GF_TRADUCIR("Fecha Inicio: ")  &"</td>")
    call writeXLS("<td align='left' style='font-size:12' colspan='6'>" & GF_FN2DTE(p_fechaInicio) & "</td></tr>")
    call writeXLS("<tr><td align='left' style='font-size:12'>" & GF_TRADUCIR("Fecha Fin: ")& "</td>")
    call writeXLS("<td align='left' style='font-size:12' colspan='6'>" & GF_FN2DTE(p_fechaFin) & "</td></tr>")
	
End function
'--------------------------------------------------------------------------------------------------------------------
Function imprimirTitulos()
    call writeXLS("<tr><td align='center' style='font-size:12;width:120px;' rowspan='2' class='titulos' >" & GF_TRADUCIR("Camion") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;' colspan='2' class='titulos' >" & GF_TRADUCIR("Ingreso") & "</td>")
    call writeXLS("<td align='center' style='font-size:12' colspan='2' class='titulos'>"  & GF_TRADUCIR("Calado") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;width:120px;' rowspan='2' class='titulos'>" & GF_TRADUCIR("Pesado bruto") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;width:120px;' rowspan='2' class='titulos'>" & GF_TRADUCIR("Pesado tara") & "</td></tr>")
    call writeXLS("<tr><td align='center' style='font-size:12;width:120px;' class='titulos'>" & GF_TRADUCIR("Ingreso") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;width:120px;' class='titulos'>" & GF_TRADUCIR("Sin cupo") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;width:120px;' class='titulos'>" & GF_TRADUCIR("Calado") & "</td>")
    call writeXLS("<td align='center' style='font-size:12;width:120px;' class='titulos'>" & GF_TRADUCIR("Demorado") & "</td></tr>")
End Function
'--------------------------------------------------------------------------------------------------------------------
Function generarCorteControlCamion(p_Rs, p_IdCamion, p_dtContable, p_Seguir)
    generarCorteControlCamion = false
    if (p_Seguir) then
        if (not p_Rs.Eof) then
            if ((p_Rs("IDCAMION") = p_IdCamion)and(p_Rs("DTCONTABLE") = p_dtContable)) then generarCorteControlCamion = true
        end if
    end if
End function
'--------------------------------------------------------------------------------------------------------------------
Function generarCorteControlPuesto(p_Rs, p_IdCamion, p_dtContable, p_Estado, p_Seguir)
    generarCorteControlPuesto = false
    if (p_Seguir) then
        if (not p_Rs.Eof) then
            if ((p_Rs("IDCAMION") = p_IdCamion)and(p_Rs("DTCONTABLE") = p_dtContable)and(CInt(p_Rs("CDESTADOPOSTERIOR")) = Cint(p_Estado))) then generarCorteControlPuesto = true
        end if
    end if
End function
'--------------------------------------------------------------------------------------------------------------------
Function imprimirTotalesPuesto(p_DiccTotales, p_VectCol)
    Dim i,key
    writeXLS("<TFOOT>")
    writeXLS("<TR><TD align='center' class='titulos'>"& GF_TRADUCIR("Total") &"</TD>")
    For i = 0 to UBound(p_VectCol)
        key = p_VectCol(i)    
        if (p_DiccTotales.Exists(key)) then
            writeXLS("<TD align='center' class='titulos'>"& p_DiccTotales.Item(key) &"</TD>")
        else
            writeXLS("<TD align='center' class='titulos'>0</TD>")
        end if
    Next
    writeXLS("</TR></TFOOT>")
End function 
'--------------------------------------------------------------------------------------------------------------------
Function armarCuerpoReporte(p_fechaInicio, p_fechaFin, pPto)
    Dim vectCol,dicTotales, IdCamion, dtContable, seguir, fechaPuesto
    Set dicTotales = Server.CreateObject("Scripting.Dictionary")	
    call writeXLS("<table>")
    Call imprimirFiltros(p_fechaInicio, p_fechaFin)
    Call imprimirTitulos()
    Set rs = armarSQLCamionesPuestos(p_fechaInicio, p_fechaFin, pPto)
    if not rs.Eof then
        vectCol = Array (ESTADO_INGRESADO, ESTADO_SIN_CUPO, ESTADO_CALADO, ESTADO_DEMORADO, ESTADO_PESADO_BRUTO, ESTADO_PESADO_TARA)
        while not rs.Eof
            seguir= true	
			i = 0
            IdCamion = rs("IDCAMION")
            dtContable = rs("DTCONTABLE")
            writeXLS("<TR><td align='center' valign='top'>" &  GF_nDigits(IdCamion,10) & "</td>")
			while(generarCorteControlCamion(rs, IdCamion, dtContable, seguir)) 
                estadoPuesto = rs("CDESTADOPOSTERIOR")
                fechaPuesto = ""
                'Se genera el corte de control para los puesto por que hay casos en que un camion pase por un mismo puesto mas de una ves para un mismo circuito
                while(generarCorteControlPuesto(rs, IdCamion, dtContable, estadoPuesto, seguir))
                    if (CInt(vectCol(i)) = CInt(rs("CDESTADOPOSTERIOR"))) then
                        'Se aplica el salto de liena por si hay duplicados de estado por puesto para un camion
                        fechaPuesto = fechaPuesto & GF_FN2DTE(rs("INGRESOFECHA") & GF_nDigits(rs("INGRESOHORA"),6)) & "<BR>"
                        'Cuento las veces que los camiones pasan por un puesto (si un camion duplica su paso tambien se contempla)
                        if (not dicTotales.Exists(estadoPuesto)) then
                            Call dicTotales.Add(estadoPuesto,1)
                        else                
                            dicTotales.item(estadoPuesto) = Cdbl(dicTotales.item(estadoPuesto)) + 1
                        end if
                        rs.MoveNext()
                    else
                        writeXLS("<TD align='center' ></TD>")
                        i = i + 1
					    if (i > UBound(vectCol)) then seguir = false
                    end if
                wend
                'quito el ultimo salto de linea que tiene la fecha/hora
                fechaPuesto = left(fechaPuesto,len(fechaPuesto)-4)
                writeXLS("<TD align='center' valign='top'>"& fechaPuesto &"</TD>")
                i = i + 1
            wend
            writeXLS("</TR>")
        wend
        Call imprimirTotalesPuesto(dicTotales,vectCol)
    else
        writeXLS("<TD align='center' colspan='7'>"& GF_TRADUCIR("No se encontraron resultados") &"</TD>")
    end if
    call writeXLS("</table>")
End Function
'--------------------------------------------------------------------------------------------------------------------
Function armarSQLCamionesPuestos(p_fechaInicio, p_fechaFin,pPto)
    Dim strSQL, auxFechaDesde, auxFechaHasta
    
    auxFechaDesde = left(p_fechaInicio,4) & "-" & Mid(p_fechaInicio, 5, 2) & "-" & Mid(p_fechaInicio, 7, 2) &" "& Mid(p_fechaInicio, 9, 2) &":"& Mid(p_fechaInicio, 11, 2) &":"& Mid(p_fechaInicio, 13, 2) 
    auxFechaHasta = left(p_fechaFin,4) & "-" & Mid(p_fechaFin, 5, 2) & "-" & Mid(p_fechaFin, 7, 2) &" "& Mid(p_fechaFin, 9, 2) &":"& Mid(p_fechaFin, 11, 2) &":"& Mid(p_fechaFin, 13, 2) 
    diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()), 2) & "-" & GF_nDigits(Day(Now()), 2)

    strSQL= "select A.*,B.DSESTADO as ESTADOPOSTERIOR, C.dstransaccion "&_
            "from ( "&_
	        "    (select ((Year(dtauditoria) * 10000) + (Month(dtauditoria) * 100) + Day(dtauditoria))AS INGRESOFECHA, "&_
		    "           ((DATEPART(HOUR, dtauditoria) * 10000) + (DATEPART(MINUTE, dtauditoria) * 100) + DATEPART(SECOND, dtauditoria)) AS INGRESOHORA, "&_
		    "           dtauditoria, "&_
		    "           '"& diaHoy &"' AS DTCONTABLE, "&_
		    "           IDCAMION,  "&_
		    "           cdtransaccion, "&_
		    "           cdestadoposterior "&_
	        "     FROM AUDITORIACAMIONES "&_
	        "where cdtransaccion in ("& TRANSACCION_INGRESO &","& TRANSACCION_VISTEO &","& TRANSACCION_BRUTO &","& TRANSACCION_TARA &") "&_
	        "   and dtauditoria >= '"& auxFechaDesde &"' and dtauditoria <= '"& auxFechaHasta &"'"&_
	        "     )UNION "&_
	        "    (select ((Year(dtauditoria) * 10000) + (Month(dtauditoria) * 100) + Day(dtauditoria))AS INGRESOFECHA, "&_
		    "           ((DATEPART(HOUR, dtauditoria) * 10000) + (DATEPART(MINUTE, dtauditoria) * 100) + DATEPART(SECOND, dtauditoria)) AS INGRESOHORA, "&_
		    "           dtauditoria, "&_
		    "           DTCONTABLE, "&_
		    "           IDCAMION, "&_
		    "           cdtransaccion, "&_
		    "           cdestadoposterior "&_
	        "     FROM HAUDITORIACAMIONES "&_
	        "where cdtransaccion in ("& TRANSACCION_INGRESO &","& TRANSACCION_VISTEO &","& TRANSACCION_BRUTO &","& TRANSACCION_TARA &") "&_
	        "   and dtauditoria >= '"& auxFechaDesde &"' and dtauditoria <= '"& auxFechaHasta &"'"&_
            "    )) A "&_
            "left join estados B on A.CDESTADOPOSTERIOR = B.CDESTADO "&_
            "left join transacciones C on C.cdtransaccion = A.cdtransaccion "&_
            "ORDER BY DTCONTABLE,IDCAMION,cdtransaccion,B.CDESTADO "
            
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
    Set armarSQLCamionesPuestos = rs
End Function

'******************************************************************************************************
'**************************************** COMIENZO DE PAGINA ******************************************
'******************************************************************************************************
Dim pto, fechaInicio, fechaFin, filename,arrColumnaEstado

pto = GF_PARAMETROS7("pto", "", 6)
fechaInicio = GF_PARAMETROS7("fechaInicio", "", 6)
fechaFin = GF_PARAMETROS7("fechaFin", "", 6)


filename = "CamionesPorPuesto_"& pto &"_"& left(Session("MmtoDato"),8) &".xls"
Call GF_createXLS(filename)

call writeXLS("<html>")
call writeXLS("<head>")
call writeXLS("<style type='text/css'>")
call writeXLS(" .border {border-color:#666666;border-style:solid;border-width:thin;}")
call writeXLS(" .titulos {background-color:#D8D8D8;font-weight:bold;border-style:solid;border-width:thin;}")
call writeXLS(" .areas {background-color:#CECEF6;font-weight:bold;}")
call writeXLS("</style>")
call writeXLS("</head>")
call writeXLS("<body>")
Call armarCuerpoReporte(fechaInicio,fechaFin,pto)
call writeXLS("</body")
%>  