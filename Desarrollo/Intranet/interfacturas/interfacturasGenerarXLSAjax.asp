<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientosObras.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="interfacturas.asp"-->
<%
dim rs, conn, strSQL
dim idProveedor, dtDesde, dtHasta, flagDatos, fileName
flagDatos = false
Const TOTAL_COLUMNAS    = 10

'-----------------------------------------------------------------------------------------
Function writeDatosMovimientos()
	Dim KilosLiquidados, precioGasto
	
	While (not rsFAC.eof)		
	    if (rsFAC("ConceptoCAB") = FAC_CODIGO_CONCEPTO_GP) then
		    KilosLiquidados = rsFAC("KilosLiquidados")
		    if isNull(KilosLiquidados) then KilosLiquidados = 0
		    KilosLiquidados = cdbl(KilosLiquidados)/100
		    precioGasto = rsFAC("PrecioGasto")
		    if isNull(precioGasto) then precioGasto = 0
		    precioGasto = GF_EDIT_DECIMALS(cdbl(precioGasto)*100,2)
		    total = GF_EDIT_DECIMALS(kilosMerma*precioGasto,2)
		    call writeXLS("<tr>")
		    call writeXLS("<td class='xls_align_center'>" & GF_FN2DTE(rsFAC("FECHACBTE")) & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & GF_EDIT_CBTE(GF_nDigits(rsFAC("PtoVta"),4) & GF_nDigits(rsFAC("NroCbte"),8)) & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & rsFAC("CONTRATO") & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & rsFAC("CTOCORREDOR") & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & GF_FN2DTE(rsFAC("FechaDESCARGA")) & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & rsFAC("CtaPte") & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & rsFAC("ConceptoDET") & "</td>")		
		    call writeXLS("<td class='xls_align_right'>" & KilosLiquidados & "</td>")
		    call writeXLS("<td class='xls_align_center'>" & getNombreMoneda(rsFAC("CDMONEDA")) & "</td>")
		    call writeXLS("<td class='xls_align_right'>" & precioGasto & "</td>")
		    call writeXLS("<td class='xls_align_right'>" & total & "</td>")
		    call writeXLS("</tr>")
        end if
		rsFAC.movenext
	Wend
	call writeXLS("<tr><td></td></tr>")
	call writeXLS("<tr><td></td></tr>")
	call writeXLS("<tr><td></td></tr>")
	'call closeXLS()
End Function
'-----------------------------------------------------------------------------------------
Function dibujarEncabezado(titulo, pIdProveedor, pDtDesde, pDtHasta)
	dim division, conn, rsDivision, strSQL
	call writeXLS("<table class='xls_border_left'>")
	call writeXLS("<tr><td colspan=" & TOTAL_COLUMNAS & " align='right' style='font-weight:normal; font-size:10'>" & GF_FN2DTE(session("MmtoSistema")) & "<br>" & session("usuario") & "</td></tr>")
	call writeXLS("<tr><td colspan=" & TOTAL_COLUMNAS & " align='center' style='font-size:24'>" & GF_TRADUCIR(titulo) & "</td></tr>")
	call writeXLS("</table>")
	call writeXLS("<table style='font-size:16; font-weight:bold; font-family:courier'>")
	call writeXLS("<tr><td></td></tr>")
	call writeXLS("<tr><td>Proveedor.....:	</td><td align='left'>" & getDescripcionProveedor(pIdProveedor) & "(" & pIdProveedor & ")" & "</td></tr>")
	call writeXLS("<tr><td>Fecha Desde...:	</td><td align='left'>" & pDtDesde & "</td></tr>")
	call writeXLS("<tr><td>Fecha Hasta...:	</td><td align='left'>" & pDtHasta & "</td></tr>")
	call writeXLS("<tr><td></td></tr>")
	call writeXLS("</table>")
End Function
'-----------------------------------------------------------------------------------------
'**************************************************************************
'**************************** INICIO PAGINA *******************************
'**************************************************************************
idProveedor = GF_Parametros7("idProveedor", 0, 6)
dtDesde = GF_Parametros7("dtDesde", "", 6)
dtHasta = GF_Parametros7("dtHasta", "", 6) 
accion = GF_Parametros7("accion", "", 6)  
fileName = "RPT_FACTURAS_" & idProveedor
myUsuario = session("Usuario")
if not isToepfer(session("KCOrganizacion")) then myUsuario = FAC_USER_WEB

call executeSP(rsFAC, "TFFL.TF100F1_GET_FACTURAS_BY_PARAMETERS", FAC_AUTORIZADA & "||" & GF_DTE2FN(dtDesde) & "||" & GF_DTE2FN(dtHasta) & "||" & idProveedor & "||" & myUsuario & "||" & SEC_SYS_FACTURACION)
if not rsFAc.eof then 
	if accion = ACCION_BACH then call GF_setXLSMode(XLS_FILE_MODE)
	Call GF_createXLS(fileName)
	flagDatos = true
	call ArmadoXLS()
	call closeXLS()
end if	


sub ArmadoXLS()
	call writeXLS("<html><head><style type='text/css'>")
	call writeXLS(".xls_border_left {border-color:#666666;border-style:solid; border-width:thin;}")
	call writeXLS(".xls_align_center {border-color:#666666;border-style:solid;border-width:thin;text-align: center;}")
	call writeXLS(".xls_align_right {border-color:#666666;border-style:solid;border-width:thin;text-align: right;}")
	call writeXLS("</style></head><script>parent.habilitarLoading('hidden','absolute');</script><body>")
	call writeXLS("<table class='xls_border_left' style='background-color:#FFFACD; font-weight:bold';>")
	call writeXLS("<tr><td>" & dibujarEncabezado("REPORTE DE FACTURAS", idProveedor, dtDesde, dtHasta) & "</td></tr>")
	call writeXLS("</table> <table  class='xls_border_left' style='background-color:#E0E0F8; font-weight:bold'>")
	call writeXLS("<tr>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("FECHA") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("NUMERO") & "</td>")		
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("CTO TOEPFER") & "</td>")	
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("CTO CORREDOR") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("FECHA DESCARGA") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("CARTA PORTE") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("CONCEPTO") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("TN LIQUIDADAS") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("MONEDA") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("PRECIO x TN") & "</td>")
	call writeXLS("<td class='xls_align_center'>" & GF_TRADUCIR("TOTAL") & "</td>")
	call writeXLS("</tr>")
	call writeXLS("</table>")
	call writeXLS("<table class='border'>")
	call writeXLS(writeDatosMovimientos())
	call writeXLS("</table>")
end sub

	 if not flagDatos then %>
			 <script>
				parent.sinDatos();
			</script>
	<% 
	else
		if (accion = ACCION_BACH) then
			dirMail = getMailFacturacionProveedores(idProveedor, FACTURACION_LISTA_MAIL_ARCHIVO)
			if (dirMail <> "") then		
				auxAsunto  = GF_TRADUCIR(getDescripcionProveedor(CD_TOEPFER) & " - Facturación periodo " & dtDesde & " al " & dtHasta)
				auxMensaje = "Se adjunta el archivo con las facturas emitidas en el período " & dtDesde & " al " & dtHasta & "." &vbcrlf&vbcrlf &_
		                 "Atentamente."&vbcrlf&vbcrlf&_
					     "Departamento de Tesoreria"&vbcrlf& getDescripcionProveedor(CD_TOEPFER)&vbcrlf&"Tel (011) 4317-0000"
				Call GP_ENVIAR_MAIL_ATTACHMENT(auxAsunto, auxMensaje, SENDER_FACTURACION, dirMail, server.MapPath("..") & ("\temp\" & filename & ".xls"))			
			end if
		end if	
	end if %>
	
	
	
