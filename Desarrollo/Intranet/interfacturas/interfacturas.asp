<%
	Dim oXML,root
	
	Dim myRegNbr
	
	Dim acumDesc,GV_TipoComprobante,codigoMoneda,tipoCambio,idioma,importeTotal,codigoDestino	
	Dim totalMonedaOrigen,TotalMonedaDolar,XML_Nro_Lote,cantRegistros
	 
	Const CODIGO_CBTE_FAC_A = "1A"
	Const CODIGO_CBTE_FAC_B = "1B"
	Const CODIGO_CBTE_FAC_E = "1E"
	
	Const CODIGO_CBTE_NCR_A = "1A"
	Const CODIGO_CBTE_NCR_B = "3B"
	Const CODIGO_CBTE_NCR_E = "3E"
	
	Const CODIGO_CBTE_NDB_A = "2A"
	Const CODIGO_CBTE_NDB_B = "2A"
	Const CODIGO_CBTE_NDB_E = "2E"
	
	Const FAC_PROCESAMIENTO_ELECTRONICO = "E"	
	
	Const FAC_FORMULARIO_IMPRESION_ST	 = "FAMI" ' Facturas Locales
	Const FAC_FORMULARIO_IMPRESION_EX	 = "FAEX" ' Facturas de Exportacion
	
	Const FAC_CODIGO_CONCEPTO_GP	 = "GP"		
	
	Const FACTURACION_LISTA_MAIL_ERROR = "E"
	Const FACTURACION_LISTA_MAIL_DAFAULT = "F"
	Const FACTURACION_LISTA_MAIL_RECALCULO = "C"
	Const FACTURACION_LISTA_MAIL_ARCHIVO = "A"
	
	Const FAC_SIS_ORG_FACT = "F"
	Const FAC_COD_DET_A = "A"
	
	Const FAC_USER_WEB = "WEB"
	
	
'-------------------------------------------------------------------------
'Función responsable por determinar el tipo de comprobante para AFIP a 
'partir del tipo y letra del comrpobante (En los archivo de facturacion se necesitan ambos codigos para la determinacion.
'Autor   : Javier A. Scalisi - Fecha: 09/02/2012
'Modifico: Javier A. Scalisi - Fecha: 29/04/2014
Function getCodigoCbteAFIP(pTipoCbteToepfer, pLetraCbteToepfer)
	
	    Dim rs, strSQL
	    
		strSQL="Select * from MET068A where tipcpte=" & pTipoCbteToepfer & " and letra= '" & pLetraCbteToepfer & "'"
		Call executeQueryDb(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)	        
	    getCodigoCbteAFIP = "XX"
	    if (not rs.eof) then getCodigoCbteAFIP = CInt(rs("cod_afip_compras"))
		
End Function
'-------------------------------------------------------------------------
Function obtenerPermisoExportacion(pNroReg)
	Dim strSQL, rs,conn,rtrn
	
	rtrn = "0"
	
	strSQL = "select * from TFFL.TF118 where FENRRG =" & pNroReg
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	
	if (not rs.EoF) then
		rtrn = rs("FENPDE")
	end if
	
	obtenerPermisoExportacion = rtrn
End Function	
'*************************************************************************
'***********************   INICIO DE PAGINA   ****************************
'*************************************************************************
'-------------------------------------------------------------------------------------------------------------------
Function getMailFacturacionProveedores(pIdProveedor, pLista)
	Dim auxDestino,rsMail
	auxDestino = ""
	if (pLista <> "") then
	    Set sp_ret = executeSP(rsMail, "TFFL.TF220F4_GET_BY_IDPROVEEDOR_TIPO", pIdProveedor & "||"& pLista &"||1||0" & "$$totalRegistros")
	    while (not rsMail.Eof)
		    auxDestino = auxDestino & Trim(rsMail("MAIL")) & "; "
		    rsMail.MoveNext()
	    wend
	    if (Len(auxDestino) > 0) then auxDestino = Left(auxDestino, Len(auxDestino)-2)
    end if	   
	getMailFacturacionProveedores = auxDestino
End Function
'********************************************************************
Function getTipoFactura(p_tipo)
	dim rtrn
	get_tipoFactura = "-"
	Select case (cInt(p_tipo))
		case 1	:	rtrn = "Factura"
		case 2	:	rtrn = "Nota de Debito"
		case 3	:	rtrn = "Nota de Credito"
	end select
	getTipoFactura = rtrn
end Function
'********************************************************************
%>