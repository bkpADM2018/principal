<%'************************************************************************************************************
'-------------------------------------------------------------------------------------------------------------------
' Función:	
'           loadReportFacturasDuplicadas
' Autor: 	
'           JCZ - Joanthan G. Costilla
' Fecha: 	
'           25/04/2016
' Objetivo:	
'			Generar reporte donde se puedan detectar Facturas Duplicados De Proveedores
' Parametros:
'			El pModo (XLS_STREAM_MODE/XLS_FILE_MODE)
' Devuelve:
'			fname [string ] --> Es el nombre del archivo generado 
'-------------------------------------------------------------------------------------------------------------------
FUNCTION loadReportFacturasDuplicadas(pModo,p_fechaDesde, p_fechaHasta)

fname = "REPORTE_CBTES_DUPLICADOS_" & RIGHT(session("MmtoSistema"),8)

if pModo <> XLS_STREAM_MODE then
    call GF_setXLSMode(pModo)
else
    call GF_setXLSMode(pModo)
end if
Call GF_createXLS(fname)

call writeXLS("<html><head><style type='text/css'>")
    call writeXLS(".border {border-color:#666666;border-style:solid; border-width:thin;}")
    call writeXLS(".titulos {background-color:#D8D8D8;font-weight:bold;}")
    call writeXLS(".detalle {border-style:solid; border-width:thin;}")
call writeXLS("</style></head><body>")

call writeXLS("<table class='border' style='background-color:#FFFFFF; font-weight:bold'>")

Call writeFilter(p_fechaDesde, p_fechaHasta) 
	call writeXLS("<tr style='font-size:10;'>")		
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("PROVEEDOR")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("CUIT")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("FECHA MINUTA")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("MINUTA")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("IMPORTE")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("COMPROBANTE")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("FECHA MINUTA DUPLICADO")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("MINUTA DUPLICADO")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("IMPORTE DUPLICADO")&"</td>")
		call writeXLS("<td class='titulos' align='center'>"&GF_TRADUCIR("CBTE. DUPLICADO")&"</td>")
	call writeXLS("</tr>")
Call drawDetalle(p_fechaDesde, p_fechaHasta)

call writeXLS("</table></body></html>")
    Call closeXLS()
    loadReportFacturasDuplicadas = fname
    
END FUNCTION
'-------------------------------------------------------------------------------------------------------------
Function writeFilter(p_fechaDesde, p_fechaHasta)
    call writeXLS("<tr>")
        call writeXLS("<td colspan='8' align='right' style='font-weight:normal; font-size:10; background:#FFFADA;'>"&GF_FN2DTE(session("MmtoSistema"))&"<br>"&session("usuario")&"</td>")
    call writeXLS("</tr>")
	call writeXLs("<tr>")
        call writeXLs("<td colspan='8' align='center' style='font-weight:normal; font-size:18; background:#FFFADA;'>"&GF_TRADUCIR("REPORTE DE COMPROBANTES DUPLICADOS DE PROVEEDORES")&"</td>")
	call writeXLs("<tr>")
        call writeXLs("<td colspan='2' style='font-size:12;background:#FFFADA;'>Fecha Desde: </td><td colspan='6' align='left' style='font-size:12;background:#FFFADA;'>"&GF_FN2DTE(20&p_fechaDesde)&"</td>")
    call writeXLs("</tr>")
	call writeXLs("<tr>")
        call writeXLs("<td colspan='2' style='font-size:12;background:#FFFADA;'>Fecha Hasta: </td><td colspan='6' align='left' style='font-size:12;background:#FFFADA;'>"&GF_FN2DTE(20&p_fechaHasta)&"</td>")
    call writeXLs("</tr>")
End Function
'-------------------------------------------------------------------------------------------------------------
Function drawDetalle(p_fechaDesde, p_fechaHasta)
    dim color, minutaOld
    minutaOld= 0
    sp_parameter = 1 & p_fechaDesde &" || "& 1 & p_fechaHasta    
    Set sp_ret = executeSP(rs, "PROVFL.ACDSREP_GET_FAC_DUPLICADAS_DE_PROVEEDOR_BY_PARAMETERS", sp_parameter)    
    if not rs.Eof then 
        while not rs.Eof 
            if minutaOld <> Cdbl(rs("MINUTA")) then
                if color <> "#d7f6b3" then
                    color = "#d7f6b3"
                ELSE
                    color = "#FFFFFF"
                end if 
            end if 
            call writeXLS("<tr>")                
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='left'>"&rs("IDPROVEEDOR") & " " & rs("RAZONSOCIAL")&"</td>")
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='left'>"&GF_STR2CUIT(rs("CUIT"))&"</td>")
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center' >"&GF_FN2DTE("20" & Right(rs("FECHAMINUTA"), 6))&"</td>")
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center' >"&rs("MINUTA")&"</td>")
                call writeXLS("<td class='detalle' style='background-color:"&color&";' align='right'>"&getSimboloMoneda(rs("MONEDA")) &" "& GF_EDIT_DECIMALS(Cdbl(rs("IMPORTE"))*100, 2) &"</td>")
                call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center'>"&rs("TIPOCBTE") & " " &  GF_EDIT_CBTE(rs("FACTURA"))&"</td>")			    
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center' >"&GF_FN2DTE("20" & Right(rs("FECHAMINUTADUPLICADA"), 6))&"</td>")
			    call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center' >"&rs("MINUTADUPLICADA")&"</td>")
                call writeXLS("<td class='detalle' style='background-color:"&color&";' align='right'>"&getSimboloMoneda(rs("MONEDADUPLICADA")) &" "& GF_EDIT_DECIMALS(Cdbl(rs("IMPORTEDUPLICADA"))*100, 2) &"</td>")
                call writeXLS("<td class='detalle' style='background-color:"&color&";' align='center'>"&rs("TIPOCBTEDUPLICADA") & " " & GF_EDIT_CBTE(rs("FACTURADUPLICADA"))&"</td>")			    
		    call writeXLS("</tr>")			
	        minutaOld = Cdbl(rs("MINUTA"))
                rs.MoveNext()
       
	    wend
	else
		call writeXLS("<tr>")
			call writeXLS("<td colspan='5' class='border' align='center'>"&GF_TRADUCIR("No se encontraron resultados")&"</td>")
		call writeXLS("</tr>")		
	end if
End Function
%>
	