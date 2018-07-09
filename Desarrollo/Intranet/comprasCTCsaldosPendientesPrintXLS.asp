<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<%
Function armarExelContratoSaldoPendiente() 
    Dim rs
writeXLS("<TABLE>")
    writeXLS("<THEAD style='width:100%;border-bottom:1px solid #ddd;'>")
        writeXLS("<tr>")
            writeXLS("<TD align='right' colspan='7' style='font-size:11;' class='xls_detail_report'>"&GF_FN2DTE(session("mmtosistema"))&"</TD>")
        writeXLS("</tr>")
        writeXLS("<tr>")
            writeXLS("<TD align='right' colspan='7' style='font-size:11;' class='xls_detail_report'>"&session("usuario")&"</TD>")
        writeXLS("</tr>")
        writeXLS("<tr>")
            writeXLS("<TH align='center' colspan='7' style='font-size:20px;' class='xls_detail_report'>Contratos Con Saldo Pendiente</TH>")
        writeXLS("</tr>")
        writeXLS("<tr>")
            For i = 0 To UBound(arrTituloReporte)
                writeXLS("<TH align='center' class='xls_title_report'>"&GF_TRADUCIR(arrTituloReporte(i))&"</TH>")
            Next
        writeXLS("</tr>")
    writeXLS("</THEAD>")
    writeXLS("<TBODY style='width:100%;'>")
        Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACONTRATOS_GET_SALDO_PENDIENTE", "")

        if (not rs.Eof) then
            while (not rs.Eof)
                writeXLS("<TR>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& trim(rs("CDCONTRATO")) & "</TD>")
                    writeXLS("<TD align='center' valign='top' class='xls_detail_report'>"& trim(rs("CDOBRA")) & "</TD>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& trim(rs("TITULO")) & "</TD>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& getUserDescription(rs("CDRESPONSABLE")) & "</TD>")
                    writeXLS("<TD align='left' valign='top' class='xls_detail_report'>"& rs("IDPROVEEDOR")&"-"& trim(rs("DSEMPRESA")) &"</TD>")
                    writeXLS("<TD align='rigth' valign='top' class='xls_detail_report'>"& GF_FN2DTE(Trim(rs("FECHAVTO"))) &"</TD>")
                  
                    writeXLS("<TD align='right' valign='top' class='xls_detail_report'>" &getSimboloMoneda(rs("CDMONEDA"))&GF_EDIT_DECIMALS(cDbl(rs("SALDO")),2)&"</TD>")                    
                
                writeXLS("</TR>")
                rs.MoveNext() 
            wend
        end if
    writeXLS("</TBODY>")
writeXLS("</TABLE>")
End Function
'---------------------------------------------------------------------------------------------------------------------------------
' Función:	
'           loadReporteContrato
' Autor: 	
'           CNA - Joanthan G. Costilla
' Fecha: 	
'           00/00/0000
' Objetivo:	
'			Carga los valores para generar el reporte de Contrato que tienen saldo pendiente de pago.
' Parametros:
'			No Recibe parametros
' Devuelve:
'			fileName [string ] --> Es el nombre del archivo generado 
'---------------------------------------------------------------------------------------------------------------------------------
Function loadReporteContrato()
    Dim fileName
  
    'Array que contiene la descripcion de los titulos del reporte
    arrTituloReporte = Array("Cod. Contrato", "Partida", "Titulo","Responsable","Proveedor","Fecha de vto.","Saldo pendiente")
    'Incluir el código de contrato, el titulo, el responsable, el proveedor, la fecha de vencimiento y el saldo pendiente.
    'Array que contiene la descripcion del detalle de las primeras columnas (Entre medio se debe mostrar los datos del Location Number)
    fileName = "Contrato_Saldo_Pendiente" & session("MmtoDato")
    Call GF_createXLS(fileName)
    writeXLS("<html>")
    writeXLS("<head>")
    writeXLS("<style type='text/css'>")
    writeXLS(".xls_title_report {font-size:13;font-weight:bold;background-color:#517B4A;border:1px;color:#FFFFFF;}")
	writeXLS(".xls_detail_report {font-size:11;border:1px;}")
	writeXLS("</style>")
	writeXLS("</head>")
    writeXLS("<body>")
    Call armarExelContratoSaldoPendiente()
    writeXLS("</body>")
    writeXLS("</html>")
    Call closeXLS()
    loadReporteContrato = fileName
End Function
'*****************************************************************************************************************
'************************************ COMIENZO DE LA PAGINA ******************************************************
'*****************************************************************************************************************
Dim arrTituloReporte, arrPrimerDetalleFijoReporte
    if (LCase(request.servervariables("script_name")) = SITE_ROOT & "comprasctcsaldospendientesprintxls.asp") then
        Call loadReporteContrato()   
    end if
    
%>