<%
'-------------------------------------------------------------------------------------------------------------------
' Función:	
'           loadReportPagoDuplicado
' Autor: 	
'           JCZ - Joanthan G. Costilla
' Fecha: 	
'           22/04/2016
' Objetivo:	
'			Generar reporte donde se puedan detectar Pagos Duplicados A Proveedores
' Parametros:
'			El pModo (XLS_STREAM_MODE/XLS_FILE_MODE)
' Devuelve:
'			fname [string ] --> Es el nombre del archivo generado 
'-------------------------------------------------------------------------------------------------------------------
FUNCTION loadReportPagoDuplicado(pModo,p_fechaDesde,p_fechaHasta)

fname = "REPORTE_PAGO_DUPLICADO_" & RIGHT(session("MmtoSistema"),8)

call GF_setXLSMode(pModo)

Call GF_createXLS(fname)

call writeXLS("<html><head><style type='text/css'>")
    call writeXLS(".border {border-color:#666666;border-style:solid; border-width:thin;}")
    call writeXLS(".titulos {background-color:#D8D8D8;font-weight:bold;}")
call writeXLS("</style></head><body>")

call writeXLS("<table class='border' style='background-color:#FFFFFF; font-weight:bold'>")
    Call writeFilter(p_fechaDesde,p_fechaHasta) 
    call writeXLS("<tr style='font-size:10;'>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Minuta")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Proveedor") & "</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Nro. de factura")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Fecha del pago")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Monto")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("PIC")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Minuta duplicada")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("Fecha del duplicado")&"</td>")
        call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("PIC duplicado")&"</td>")
    call writeXLS("</tr>")		
    Call drawDetalle(p_fechaDesde, p_fechaHasta) 
call writeXLS ("</table></body></html>")

    Call closeXLS()
    loadReportPagoDuplicado = fname

END FUNCTION 
'-------------------------------------------------------------------------------------------------------------------
Function writeFilter(p_fechaDesde, p_fechaHasta) 
    call writeXLS("<tr>")
        call writeXLS("<td colspan='9' align='right' style='font-weight:normal; font-size:10; background:#FFFADA;'>"&GF_FN2DTE(session("MmtoSistema"))&"<br>"&session("usuario")&"</td>")
    call writeXLS("</tr>")
    call writeXLS("<tr>")
        call writeXLS("<td colspan='9' align='center' style='font-size:16;background:#FFFADA;'>"&GF_TRADUCIR("REPORTE DE PAGO DUPLICADO A PROVEEDORES")&"</td>")
    call writeXLS("</tr>")
    call writeXLS("<tr>")
        call writeXLS("<td colspan='2' style='font-size:12;background:#FFFADA;'>Fecha Desde: </td><td style='font-size:12;background:#FFFADA;' colspan='7' align='left'>"&GF_FN2DTE(20&RIGHT(p_fechaDesde,6))&"</td>")
    call writeXLS("</tr>")
    call writeXLS("<tr>")
        call writeXLS("<td colspan='2' style='font-size:12;background:#FFFADA;'>Fecha Hasta: </td><td style='font-size:12;background:#FFFADA;' colspan='7' align='left'>"&GF_FN2DTE(20&RIGHT(p_fechaHasta,6))&"</td>")
    call writeXLS("</tr>")
End Function
'-------------------------------------------------------------------------------------------------------------
Function drawDetalle(p_fechaDesde,p_fechaHasta)
    sp_parameter =  p_fechaDesde &"||"& p_fechaHasta
    Set sp_ret = executeSP(rs, "TESFL.TES134F1_GET_PAGO_DUPLICADO_A_PROV_BY_PARAMETERS", sp_parameter)
    if not rs.Eof then 
	    while not rs.Eof 
		    call writeXLS("<tr style='font-size:10;'>")
                call writeXLS("<td class='border' align='center'>"&rs("MINUTA")&"</td>")
                call writeXLS("<td class='border' align='left'>"&rs("PROVEEDOR") &"-"& rs("RAZONSOCIAL")&"</td>")
                call writeXLS("<td class='border' align='left'>"&rs("TIPOCBTE") & " " & GF_EDIT_CBTE(rs("FACTURA"))&"</td>")
			    call writeXLS("<td class='border' align='left'>"&GF_FN2DTE(rs("FECHAPAGO"))&"</td>")
                call writeXLS("<td class='border' align='left'>"&getSimboloMoneda(rs("MONEDA"))&" "&GF_EDIT_DECIMALS(Cdbl(rs("IMPORTE"))*100,2)&"</td>")
                call writeXLS("<td class='border' align='center'>"&rs("IDCOTIZACION")&"</td>")
                call writeXLS("<td class='border' align='center'>"&rs("MINUTADUP")&"</td>")
                call writeXLS("<td class='border' align='center'>"&GF_FN2DTE(rs("FECPAGODUP"))&"</td>")
                call writeXLS("<td class='border' align='center'>"&rs("IDCOTDUPLICADO")&"</td>") 
		    call writeXLS("</tr>")
	    rs.MoveNext()
	    wend		 
    else
        call writeXLS("<tr>")
            call writeXLS("<td colspan='9' class='border' align='center'>"&GF_TRADUCIR("No se encontraron resultados")&"</td>")
        call writeXLS("</tr>")
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------
%>