<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%

Call GP_CONFIGURARMOMENTOS()

'Se toman todos los proveedores que se publican en el archivo NVA y se copian al log para establecer un nuevo punto de control.
Call executeSP(rsNVA, "QRY.PAGOSULT_GET_BY_PARAMETER", "D||" & Left(session("MmtoDato"), 8))

session("Usuario") = "AGS"
while (not rsNVA.eof)
    Call grabarProveedorLog(CLng(rsNVA("NROPROV")))
    rsNVA.MoveNext()
wend

%>
