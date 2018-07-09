<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="interfacturas.asp"-->
<%
Function armarExelReporte(p_Desde,p_Hasta,p_Importe) 
    Dim rs
writeXLS("<TABLE>")
    writeXLS("<THEAD style='width:100%;border-bottom:1px solid #ddd;'>")
        writeXLS("<tr>")
        if p_Importe = 0 then
            writeXLS("<TH align='center' colspan='6' style='font-size:20px;' class='xls_detail_report'>Reporte Facturas - Total Cero</TH>")
        else
            writeXLS("<TH align='center' colspan='7' style='font-size:20px;' class='xls_detail_report'>Reporte Facturas - Items Negativos</TH>")            
        end if
        writeXLS("</tr>")
        writeXLS("<tr>")
            writeXLS("<TH align='center' class='xls_detail_report'>Fecha Desde:</TH>")
            writeXLS("<TH align='center' class='xls_detail_report'>" & GF_FN2DTE(p_Desde) & "</TH>")
            writeXLS("<TH></TH>")
            writeXLS("<TH align='center' class='xls_detail_report'>Fecha Hasta:</TH>")
            writeXLS("<TH align='center' class='xls_detail_report'>" & GF_FN2DTE(p_Hasta) & "</TH>")
        writeXLS("</tr>")
        writeXLS("<tr>")                
        writeXLS("<TH align='center' class='xls_title_report'>Fecha</TH>")            
        writeXLS("<TH align='center' class='xls_title_report'>Tipo Comprobante</TH>")            
        writeXLS("<TH align='center' class='xls_title_report'>Nº Comprobante</TH>")            
        If cint(p_Importe) <> 0 then
            writeXLS("<TH align='center' class='xls_title_report'>Detalle</TH>")            
        end if            
        writeXLS("<TH align='center' class='xls_title_report'>Importe</TH>")            
        writeXLS("<TH align='center' class='xls_title_report'>Sector</TH>")            
        writeXLS("<TH align='center' class='xls_title_report'>Cliente</TH>")            
        writeXLS("</tr>")
    writeXLS("</THEAD>")
    writeXLS("<TBODY style='width:100%;'>")
        If cint(p_Importe) = 0 then
            set sp_ret = executeSP(rs, "TFFL.TF100F1_GET_IMPORTE_CERO_BY_FCCMFC", p_Desde & "||" & p_Hasta & "||1||0$$res1||res2||res3")
        else
            Set sp_ret = executeSP(rs, "TFFL.TF100F1_GET_ITEM_NEGATIVO_BY_FCCMFC", p_Desde & "||" & p_Hasta & "||1||0$$res1||res2||res3")
        end if
        if (not rs.Eof) then
            while (not rs.Eof)
                writeXLS("<TR>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& GF_FN2DTE(Trim(rs("FCCMFC"))) & "</TD>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& getTipoFactura(rs("FCCMTP")) &" " & Trim(rs("FCCMTF")) & "</TD>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>Nro. "& GF_EDIT_CBTE(GF_nDigits(rs("FCCMDV"),4) & GF_nDigits(rs("FCCMNR"),8)) &"</TD>")
                    If cint(p_Importe) <> 0 then
                        writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& rs("DETALLE") &"</TD>")
                    end if
                    writeXLS("<TD align='right' valign='top' class='xls_detail_report'>"& getSimboloMoneda(rs("FCMNCD"))&" "& GF_EDIT_DECIMALS(CDbl(rs("IMPORTE"))*100,2)&"</TD>")                    
                    If Trim(rs("NOSESC")) <> "" then
                        writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& Trim(rs("FCSCNR")) &"-"&Trim(rs("NOSESC")) &"</TD>")
                    Else
                        writeXLS("<TD align='left' valign='top' class='xls_detail_report'></TD>")
                    end If
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>" & rs("FCCLNR") & "-" & rs("DSCLIENTE") & "</TD>")
                writeXLS("</TR>")
                rs.MoveNext() 
            wend
        end if
    writeXLS("</TBODY>")
writeXLS("</TABLE>")
End Function
'---------------------------------------------------------------------------------------------------------------------------------
' Función:	
'           loadReportFactura
' Autor: 	
'           CNA - Joanthan G. Costilla
' Fecha: 	
'           08/03/2016
' Objetivo:	
'			Carga los valores para generar el reporte de Facturas con importes negativos o iguales a cero.
' Parametros:
'			p_Desde [BIGINT] --> Carga la fecha DESDE que momento desea ver las facturas
'           p_Hasta [BIGINT] --> Carga la fecha HASTA que momento desea ver las facturas
'           p_modoImporte [INT] --> Setea si es IMPORTE igual a CERO o NEGATIVO.
'           p_modo [INT] --> modo de vista del Excel
' Devuelve:
'			fileName [string ] --> Es el nombre del archivo generado 
'---------------------------------------------------------------------------------------------------------------------------------
Function loadReportFactura(p_Desde,p_Hasta,p_modoImporte,p_modo)
    Dim fileName
  
    'Array que contiene la descripcion del detalle de las primeras columnas (Entre medio se debe mostrar los datos del Location Number)
    If cint(p_modoImporte) = 0 then
        fileName = "Reportes_con_importe_0_" & session("MmtoDato")
    else
        fileName = "Reportes_con_importe_Negativo_"& session("MmtoDato")
    end if
    Call GF_setXLSMode (p_modo)
    Call GF_createXLS(fileName)
    writeXLS("<html>")
    writeXLS("<head>")
    writeXLS("<style type='text/css'>")
    writeXLS(".xls_title_report {font-size:11;font-weight:bold;background-color:#CEE3F6;border:1px;}")
	writeXLS(".xls_detail_report {font-size:11;border:1px;}")
	writeXLS("</style>")
	writeXLS("</head>")
    writeXLS("<body>")
    Call armarExelReporte(p_Desde,p_Hasta,p_modoImporte)
    writeXLS("</body>")
    writeXLS("</html>")
    Call closeXLS()
    loadReportFactura = fileName
End Function
'*****************************************************************************************************************
'************************************ COMIENZO DE LA PAGINA ******************************************************
'*****************************************************************************************************************
Dim  fechaDesde,fechaHasta,modoImporte,modo
    fechaDesde = GF_PARAMETROS7 ("fechaDesde","",6)
    fechaHasta = GF_PARAMETROS7 ("fechaHasta","",6)
    modoImporte = GF_PARAMETROS7 ("modoImporte",0,6)
    if (LCase(request.servervariables("script_name")) = "/actisaintra/interfacturas/interfacturasreporteprintxls.asp") then
        Call loadReportFactura(fechaDesde,fechaHasta,modoImporte,XLS_STREAM_MODE)   
    end if
    
%>