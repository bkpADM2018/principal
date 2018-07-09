<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<%
'--------------------------------------------------------------------------------------------
' Función:	loadImporteCTCSinAjuste
' Autor: 	CNA - Ajaya Nahuel
' Fecha: 	06/11/2013
' Objetivo:	
'			Carga el importe (pesos y dolares) asignado al inicio del contrato, esto significa que excluye todos los ajustes 
'			que tiene el contrato. Es aplicable para anticipos
'--------------------------------------------------------------------------------------------
Function loadImporteCTCSinAjuste()
	Dim strSQL
	strSQL = "SELECT SUM(IMPORTEPESOS) IMPORTEPESOS, SUM(IMPORTEDOLARES) IMPORTEDOLARES "&_
			 "FROM TBLOBRACTCAJUSTES " &_
			 "WHERE IDCONTRATO = " & CTC_idContrato &_
			 "	  AND APLICADO = '" & TIPO_AFIRMACION &"'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.EoF then
		if not IsNull(rs("IMPORTEPESOS")) then CTC_ContratoPesos = CTC_ContratoPesos - Cdbl(rs("IMPORTEPESOS"))
		if not IsNull(rs("IMPORTEDOLARES")) then CTC_ContratoDolares = CTC_ContratoDolares - Cdbl(rs("IMPORTEDOLARES"))	
	end if
End Function
'--------------------------------------------------------------------------------------------
' Función:	getPjeAnticipos
' Autor: 	?  
' Fecha: 	?
' Modificacion: CNA - Ajaya Nahuel 
' Fecha modificacion: 10/04/2013
' Objetivo:	
'			Devuleve el porcentaje pagado en concepto de anticipos que contiene el contrato
' Parametros:
'			pIdContrato 	[int] 	ID CONTRATO
' Devuelve:
'			porcentaje		[decimal] 
'--------------------------------------------------------------------------------------------
Function getPjeAnticipos(idContrato, moneda)
	Dim rtrn, rs, conn, strSQL, auxPorcentaje
	auxPorcentaje = 100
	if (moneda = MONEDA_PESO) then
	    strSQL = "			 SELECT Sum(DET.IMPORTEPESOS) as IMPORTE 		"
	else
	    strSQL = "			 SELECT Sum(DET.IMPORTEDOLARES) as IMPORTE 		"
	end if
	strSQL = strSQL & "	 FROM												"
	strSQL = strSQL & "		 ( SELECT *										"
	strSQL = strSQL & "		   FROM TBLCTZCABECERA				"
	strSQL = strSQL & "		   WHERE IDCONTRATO = " & idContrato
	strSQL = strSQL & "			  AND ESTADO <> '" & CTZ_ANULADA & "' ) CTZ "
	strSQL = strSQL & "		INNER JOIN TBLCTZDETALLE DET			"
	strSQL = strSQL & "			ON CTZ.IDCOTIZACION = DET.IDCOTIZACION		"
    strSQL = strSQL & "	 WHERE DET.IDARTICULO = " & ITEM_ANTICIPO_OBRAS_EN_CURSO
    strSQL = strSQL & "	       AND DET.IMPORTEPESOS > 0"    
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not isNull(rs("IMPORTE")) then
		auxImporte = cdbl(rs("IMPORTE"))/100
		if (moneda = MONEDA_PESO) then
		    auxTotal   = CTC_ContratoPesos 
		else		
		    auxTotal   = CTC_ContratoDolares 
		end if
		auxTotal   = auxTotal/100
		auxPorcentaje = (auxImporte *  100) / auxTotal
	else
	    auxPorcentaje = 0
	end if
	getPjeAnticipos = auxPorcentaje	
End Function
'*******************************************************************
'***********************  COMIENZO DE PAGINA  **********************
'*******************************************************************
'ESTA PAGINA CALCULA LOS IMPORTES Y PORSENTAJES CORRESPONDIENTES ***
'AL PAGO QUE SE ESTA GENERANDO *************************************
'*******************************************************************

Dim rsCTC, flagPje, CTC_cdMonedaPago

CTC_idContrato = GF_PARAMETROS7("idContrato",0,6)
Set rsCTC = readCTC(CTC_idContrato)

if (not rsCTC.eof) then	
	CTC_cdMonedaPago = GF_PARAMETROS7("CTC_cdMoneda","",6)
	CTC_Importe = GF_PARAMETROS7("CTC_Importe",0,6)
	CTC_idPIC = GF_PARAMETROS7("CTC_idPIC",0,6)
	CTC_tipoPago = GF_PARAMETROS7("CTC_tipoPago",0,6)
	CTC_PjeFReparo = cInt(rsCTC("FONDOREPARO"))
	CTC_ContratoPesos = cDbl(rsCTC("IMPORTEPESOS"))
	CTC_ContratoDolares = cDbl(rsCTC("IMPORTEDOLARES"))	
	CTC_aplicaAnticipo = GF_PARAMETROS7("CTC_aplicaAnticipo",0,6)
	CTC_aplicaFReparo = GF_PARAMETROS7("CTC_aplicaFReparo",0,6)
	CTC_ContratoMoneda = rsCTC("CDMONEDA")
	CTC_ImporteObra = 0	
	CTC_Anticipo = 0
	CTC_FReparo = 0
	CTC_APagar = 0
	CTC_PjePago = 0
	
	'Flag para saber si debe o no calcularse el porcentaje del pago. Se asume que si.
	flagPje = true	
	
	Select case CTC_tipoPago
		case PAGO_OBRA:
			' se da la opcion de apliacar o no el anticipo y el fondo de reparo, en esta opcion se manejan todos los importes
			CTC_ImporteObra = CTC_Importe			
			if (CTC_aplicaAnticipo = 1) then
				'Cargo el importe original del contrato, se produce cuando se aplica anticipo 
				Call loadImporteCTCSinAjuste()
				CTC_PjeAnticipo = getPjeAnticipos(CTC_idContrato, CTC_ContratoMoneda)								
				CTC_Anticipo = (((CTC_Importe * CTC_PjeAnticipo) / 100) * -1)
			end if
			if (CTC_aplicaFReparo = 1) then
				CTC_FReparo = (((CTC_Importe * CTC_PjeFReparo) / 100) * -1)
			end if
			CTC_APagar = (CTC_Importe + CTC_Anticipo + CTC_FReparo)
		case PAGO_ANTICIPO:
			'en los anticipos el f. reparo no se calcula
			CTC_Anticipo = CTC_Importe
			CTC_APagar = CTC_Anticipo
		case PAGO_RECUPERO_FR:
			'en el recupero el anticipo no se calcula
			CTC_FReparo = CTC_Importe			
			CTC_APagar = CTC_FReparo
			flagPje = false
	End Select
	'Se calcula el porcentaje según moneda del contrato.
	if (flagPje) then
		'Si el importe del contrato es 0 no se calcula el porcentaje de Pago
		if (CTC_ContratoPesos <> 0) then
			if (CTC_MonedaPago = MONEDA_PESO) then
				CTC_PjePago = (CTC_APagar * 100) / CTC_ContratoPesos
			else	
				'Es Dolares		
				CTC_PjePago = (CTC_APagar * 100) / CTC_ContratoDolares
			end if	
		end if	
	end if
%>
	<input type="hidden" id="CTC_ImporteObra" name="CTC_ImporteObra" value="<% =CTC_ImporteObra %>">
	<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td>
				<table class="reg_header" width="100%" border="0" cellspacing="0">
					<tr>
						<td>
							<% =GF_TRADUCIR("Porcentaje de Pago") %>:&nbsp;<b><% =round(CTC_PjePago, 3) %>&nbsp;%</b>
						</td>
						<td >
							<input type="checkbox" id="CTC_aplicaAnticipo" name="CTC_aplicaAnticipo" value="1" onClick="aplicarAnticipo();" <% if (CTC_aplicaAnticipo = 1) then Response.Write "Checked" %>>
							<% =GF_TRADUCIR("Aplicar Anticipo") %>
						</td>
						<td>
							<input type="checkbox" id="CTC_aplicaFReparo" name="CTC_aplicaFReparo" value="1" onClick="aplicarFReparo();" <% if (CTC_aplicaFReparo = 1) then Response.Write "Checked" %>>
							<% =GF_TRADUCIR("Aplicar F. de Reparo") %>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<br>
		<tr>
			<td>
				<table width="100%" align="center" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<table width="50%" class="reg_header" align="center" border="0">
								<tr>
									<td class="reg_header_nav" align="center" colspan="2"><% =GF_TRADUCIR("Importes") %></td>
								</tr>
								<tr>
									<td class="reg_header_navdos" width="40%"><% =GF_TRADUCIR("Anticipo")%></td>
									<td align="right"><b><% =getSimboloMoneda(CTC_cdMoneda)%>&nbsp;<% =GF_EDIT_DECIMALS(CTC_Anticipo, 2) %></b></td>
									<input type="hidden" id="CTC_Anticipo" name="CTC_Anticipo" value="<% =CTC_Anticipo %>">
								</tr>
								<tr>
									<td class="reg_header_navdos"><% =GF_TRADUCIR("F. Reparo")%></td>
									<td class="subr" align="right"><b><% =getSimboloMoneda(CTC_cdMoneda)%>&nbsp;<% =GF_EDIT_DECIMALS(CTC_FReparo, 2) %></b></td>
									<input type="hidden" id="CTC_FReparo" name="CTC_FReparo" value="<% =CTC_FReparo %>">
								</tr>
								<tr>
									<td class="reg_header_navdos"><% =GF_TRADUCIR("A Pagar")%></td>
									<td align="right"><b><% =getSimboloMoneda(CTC_cdMoneda)%>&nbsp;<% =GF_EDIT_DECIMALS(CTC_APagar, 2) %></b></td>
									<input type="hidden" id="CTC_APagar" name="CTC_APagar" value="<% =CTC_APagar %>">
								</tr>
							</table>
						</td>						
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%
else
	Response.Redirect "comprasAccesoDenegado.asp"
end if
%>