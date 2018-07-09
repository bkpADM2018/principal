<%
Const PagosDelDiaHora = 12
dim PagosDelDiaMsg
PagosDelDiaMsg = "<b>IMPORTANTE:</b> Los pagos correspondientes al dia de la fecha estarán disponibles a partir de las " & PagosDelDiaHora & " Hs."
'---------------------------------------------------------------------------------------------
Function GF_ULTIMA_FECHAPAGO(P_TIPO) 
'Esta funcion obtiene la ultima fecha de pago registrada para el usuario.
'P_TIPO: Indica la forma en que se devuelve la fecha, a saber:
'		"DTE": Formato dd/mm/aaaa 
'		"INT": Fecha en Segundos
	Dim strTipo,rtrn,strSQL,oConn,rs,strKC
	Dim intYear,intMonth, intDay, hoy

	strTipo=UCase(P_TIPO)
	intYear= year(date)
    intMonth= month(date)
    intDay= day(date)
	Call GF_STANDARIZAR_FECHA(intDay,intMonth,intYear)
	hoy = intYear & intMonth & intDay
	rtrn = hoy
	'Se obtiene el KC de la organizacion.
	strKC=session("KCOrganizacion")
	strSQL="Select distinct WCFPAG as FechaPago from TESFL.TES960F1 where WCTCBT <> 'RCL' and (WCNPRO= " & strKC & " or WCPRET= " & strKC & ") and WCFPAG <"
	if hour(now) >= PagosDelDiaHora then strSQL = strSQL & "="		
	strSQL= strSQL & hoy & " order by WCFPAG desc"
	'Response.Write "FECHA(" & strSQL  & ")"
	Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
	if not(rs.EOF) then rtrn=Left(rs("FechaPago"), 8)
		
	if (strTipo = "DTE") then rtrn=GF_FN2DTE(rtrn)	
	GF_ULTIMA_FECHAPAGO = rtrn
End Function 
'------------------------------------------------------------------------------------
Function GF_CAB_TOTALES(ByRef rs,P_FechaPago,p_nroProv)
'Esta Funcion obtiene los totales por Corredor-Vendedor
'Autor: Javier A. Scalisi
'Fecha: 21/08/2003
    Dim strSQL,oConn
	Dim strKC, strDS
	
	'Se obtiene el KR del usuario.
	strSQL="Select FechaPago,KCCOR,KCVEN, Sum(Importe) Importe, Sum(ImporteRet) Retencion,Sum(NetoEmpresaUsuario) NetoEmpresaUsuario, Sum(NetoCliente) NetoCliente from ( "
	'Se totalizan las ordenes de pago
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN, -Sum(WCIMBT) Importe, -Sum(WCIMRE) ImporteRet,0 NetoEmpresaUsuario,0 NetoCliente "
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOME=" & p_nroProv & " and WCPOME=WCPOIV and WCPGCB='P'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union"
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,Sum(WCIMBT) Importe, Sum(WCIMRE) ImporteRet,Sum(WCIMME) NetoEmpresaUsuario,Sum(WCIMIV) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOME =" & p_nroProv & " and WCPGCB='P'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union"
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,Sum(WCIMBT) Importe, Sum(WCIMRE) ImporteRet,Sum(WCIMIV) NetoEmpresaUsuario,Sum(WCIMME) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOIV=" & p_nroProv & " and WCPGCB='P'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET "
	strSQL=strSQL & " union "
	'Se totalizan las ordenes de cobro y se restan
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,Sum(WCIMBT) Importe, Sum(WCIMRE) ImporteRet,0 NetoEmpresaUsuario,0 NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOME=" & p_nroProv & " and WCPOME=WCPOIV and WCPGCB='C'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union "
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,-Sum(WCIMBT) Importe, -Sum(WCIMRE) ImporteRet,-Sum(WCIMME) NetoEmpresaUsuario,-Sum(WCIMIV) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOME =" & p_nroProv & " and WCPGCB='C'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union "
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,-Sum(WCIMBT) Importe, -Sum(WCIMRE) ImporteRet,-Sum(WCIMIV) NetoEmpresaUsuario,-Sum(WCIMME) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOIV=" & p_nroProv & " and WCPGCB='C'"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union"
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,Sum(WCIMBT) Importe, Sum(WCIMRE) ImporteRet,0 NetoEmpresaUsuario,Sum(WCIMME+WCIMIV) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOIV!=" & p_nroProv & " and WCPOME !=" & p_nroProv & " and WCPGCB='P' and (WCNPRO=" & p_nroProv & " or WCPRET=" & p_nroProv & ")"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " union "
	strSQL=strSQL & " Select WCFPAG as FechaPago,WCNPRO as KCCOR,WCPRET as KCVEN,-Sum(WCIMBT) Importe, -Sum(WCIMRE) ImporteRet,0 NetoEmpresaUsuario,-Sum(WCIMME+WCIMIV) NetoCliente"
	strSQL=strSQL & " from TESFL.TES960F1 "
	strSQL=strSQL & " where WCTCBT <> 'RCL' and WCFPAG=" & P_FechaPago & " and WCPOIV!=" & p_nroProv & " and WCPOME !=" & p_nroProv & " and WCPGCB='C' and (WCNPRO=" & p_nroProv & " or WCPRET=" & p_nroProv & ")"
	strSQL=strSQL & " group by WCFPAG,WCNPRO,WCPRET"
	strSQL=strSQL & " ) Tabla "
	strSQL=strSQL & " where FechaPago=" & P_FechaPago
	strSQL=strSQL & " Group by FechaPago,KCCOR,KCVEN"
	strSQL=strSQL & " Order by KCCOR,KCVEN"
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL
End FUnction
'------------------------------------------------------------------------------------
Function GF_CAB_LEER_Periodo(Byref rs, p_fechaPagoDesde, p_fechaPagoHasta, p_tipo, p_minuta, strOrderBy, p_nroProv)
    Dim strKC, strDS
	Dim strSQL,oConn

	if (P_Tipo = "" or P_Minuta = "") then
		'Se leen todas las cabeceras de la organizacion.
		'Se obtienen los datos de la base.
		strSQL="Select FechaPago,TipoCbte,Minuta,CbteProveedor,FechaCbte,PC,Orden,KCCOR,KCVEN,Importe,ImporteRet,Sum(NetoEmpresaUsuario) NetoEmpresaUsuario from ("
		strSQL=strSQL & " select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, WCIMME as NetoEmpresaUsuario"
		strSQL=strSQL & " from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG>=" & p_fechaPagoDesde & " and WCFPAG<=" & p_fechaPagoHasta & " and WCPOME=" & p_nroProv & " and WCPOIV !=" & p_nroProv
		strSQL=strSQL & " union"
		strSQL=strSQL & " select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, (WCIMME+WCIMIV) NetoEmpresaUsuario"
		strSQL=strSQL & " from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG>=" & p_fechaPagoDesde & " and WCFPAG<=" & p_fechaPagoHasta & " and WCPOME=" & p_nroProv & " and WCPOIV =" & p_nroProv
		strSQL=strSQL & " union"
		strSQL=strSQL & " select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, WCIMIV NetoEmpresaUsuario"
		strSQL=strSQL & " from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG>=" & p_fechaPagoDesde & " and WCFPAG<=" & p_fechaPagoHasta & " and WCPOME !=" & p_nroProv & " and WCPOIV =" & p_nroProv
		strSQL=strSQL & " union"
		strSQL=strSQL & " select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, 0 NetoEmpresaUsuario"
		strSQL=strSQL & " from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG>=" & p_fechaPagoDesde & " and WCFPAG<=" & p_fechaPagoHasta & " and WCPOME !=" & p_nroProv & " and WCPOIV !=" & p_nroProv & " and (WCNPRO=" & p_nroProv & " or WCPRET=" & p_nroProv & ")"
		strSQL=strSQL & " ) Tabla"
		strSQL=strSQL & " where FechaPago>=" & p_fechaPagoDesde & " and FechaPago<=" & p_fechaPagoHasta
		strSQL=strSQL & " Group by FechaPago,TipoCbte,Minuta,CbteProveedor,FechaCbte,PC,Orden,KCCOR,KCVEN,Importe,ImporteRet"
		strSQL=strSQL & " Order by " & strOrderBy
	else
		strSQL="Select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, WCIMME as importeMerc, WCIMIV as ImporteIVA, WCPOME as KCMERC, WCPOIV as KCIVA, WCCONT as Contrato "
		strSQL=strSQL & "from TESFL.TES960F1 where WCFPAG >= " & p_fechaPagoDesde & " and WCFPAG <= " & p_fechaPagoHasta & " and WCTCBT='" & P_Tipo & "' and WCNING=" & P_Minuta		
	end if	
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL	
end function
'------------------------------------------------------------------------------------
Function GF_CAB_LEER_Ordenado(ByRef rs,P_FechaPago,P_Tipo,P_Minuta, strOrderBy, p_nroProv)
'Esta funcion lee la cabecera de las ordenes de pago y
'la devuelve en un recordset.
    call GF_CAB_LEER_Periodo(rs, p_fechaPago, p_fechaPago, p_tipo, p_minuta, strOrderBy, p_nroProv)
end function
'------------------------------------------------------------------------------------
Function GF_CAB_LEER(ByRef p_rs,P_FechaPago,P_Tipo,P_Minuta, p_nroProv)
'Esta funcion lee la cabecera de las ordenes de pago y
'la devuelve en un recordset.
'Autor: Javier A. Scalisi
'Fecha: 21/08/2003
    call GF_CAB_LEER_Ordenado(p_rs, p_fechaPago, p_tipo, p_minuta, "KCCOR,KCVEN", p_nroProv)
end function
'------------------------------------------------------------------------------------
function GF_CAB_CALCULAR(P_strKCPRoveedor,P_intFechaPago,ByRef P_Importe,ByRef P_Retenciones,ByRef P_Mercaderias, ByRef P_IVA)
'Esta funcion obtiene los totales generales a mostrar en la pagina de cabecera.
'Autor: Javier A. Scalisi
'Fecha: 21/08/2003
	Dim strSQL,oConn,rs,rs2
	Dim strKC, strDS
	strKC=P_strKCPRoveedor
	P_Importe=0
	P_Retenciones=0
	P_Mercaderias=0
	P_IVA=0
	'Calculo los totales de  retenciones, IVA, y Mercaderias
	strSQL="Select Sum(WCIMRE) Retenciones, Sum(WCIMME) Mercaderia, Sum(WCIMIV) IVA From TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and (WCNPRO=" & strKC & " or WCPRET=" & strKC & ") and WCPGCB='P'"
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL	
	if not(rs.EOF) then
		if (rs("Retenciones") <> "") then P_Retenciones=cDBl(rs("Retenciones"))
		if (rs("Mercaderia") <> "") then P_Mercaderias=cDBl(rs("Mercaderia"))
		if (rs("IVA") <> "") then P_IVA=cDBl(rs("IVA"))
	end if
	GF_BD_AS400_2 rs,oConn, "CLOSE",strSQL
	strSQL="Select Sum(WCIMRE) Retenciones, Sum(WCIMME) Mercaderia, Sum(WCIMIV) IVA From TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and (WCNPRO=" & strKC & " or WCPRET=" & strKC & ") and WCPGCB='C'"
	GF_BD_AS400_2 rs2,oConn,"OPEN",strSQL	
	if not(rs2.EOF) then
		if (rs2("Retenciones") <> "") then P_Retenciones=P_Retenciones - cDBl(rs2("Retenciones"))
		if (rs2("Mercaderia") <> "") then P_Mercaderias= P_Mercaderias - cDBl(rs2("Mercaderia"))
		if (rs2("IVA") <> "") then P_IVA=P_IVA - cDBl(rs2("IVA"))
	end if
	GF_BD_AS400_2 rs2,oConn, "CLOSE",strSQL
	'Calculo el Impote Neto total.
	strSQL="Select Sum(WCIMME) ImporteMercNeto from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and WCPOME=" & strKC & " and WCPGCB='P'"
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL	
	if not(rs.EOF) and (rs("ImporteMercNeto") > "0") then P_Importe=P_Importe + cDBl(rs("ImporteMercNeto"))
	GF_BD_AS400_2 rs,oConn, "CLOSE",strSQL
	strSQL="Select Sum(WCIMME) ImporteMercNeto from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and WCPOME=" & strKC & " and WCPGCB='C'"
	GF_BD_AS400_2 rs2,oConn,"OPEN",strSQL		
	if not(rs2.EOF) and (rs2("ImporteMercNeto") > "0") then P_Importe=P_Importe - cDBl(rs2("ImporteMercNeto"))
	GF_BD_AS400_2 rs2,oConn, "CLOSE",strSQL
	strSQL="Select Sum(WCIMIV) ImporteIVANeto from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and WCPOIV=" & strKC & " and WCPGCB='P'"
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL	
	if not(rs.EOF) and (rs("ImporteIVANeto") > "0")then P_Importe=P_Importe + cDBl(rs("ImporteIVANeto"))
	GF_BD_AS400_2 rs,oConn, "CLOSE",strSQL
	strSQL="Select Sum(WCIMIV) ImporteIVANeto from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and WCPOIV=" & strKC & " and WCPGCB='C'"	
	GF_BD_AS400_2 rs2,oConn,"OPEN",strSQL	
	if not(rs2.EOF) and (rs2("ImporteIVANeto") > "0")then P_Importe=P_Importe - cDBl(rs2("ImporteIVANeto"))
	GF_BD_AS400_2 rs2,oConn, "CLOSE",strSQL
end function
'------------------------------------------------------------------------------------
Function GF_DET_LEER(ByRef rs,P_FechaPago,P_Tipo,P_Minuta)
'Esta funcion lee el detalle de una ordenes de pago y
'lo devuelve en un recordset.
'Autor: Javier A. Scalisi
'Fecha: 21/08/2003
	'EAB 09-01-2013 WDTCBT <> RCL and 
	Dim strSQL,oConn
	
	'Se obtienen los datos de la base.
	strSQL="Select WDFPAG as FechaPago, WDTCBT as TipoCbte, WDNING as Minuta, WDCODE as KCDetalle, WDIMPP as ImportePesos, WDIMPD as ImporteDolar, WDDBCR as DBCR, WDFOPG as KCPago, WDNRLT as Lote, WDPIVA as MRCPago, WDPOPG as KCPRO, WDCBCO as KCBC, WDSBCO as KCBCSC, WDNCHE as Cheque, WDNRET as RetNro, WDMBAJ as MarcaAnulacion "
	strSQL= strSQL & "from TESFL.TES960F2 where WDFPAG=" & P_FechaPago & " and WDTCBT='" & P_Tipo & "' and WDNING=" & P_Minuta & " order by WDCODE desc, WDDBCR desc"
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL
	
	
end function 
'------------------------------------------------------------------------------------
Function GF_CAB_TFPAG(ByRef P_rs, P_intFechaPago, P_KC,P_ORUsr)
'Esta funcion totaliza por forma de pago para una fecha y proveedor
'determinado.
'Autor: Javier A. Scalisi
'Fecha: 18/09/2003

	Dim strSQL,oConn
	Dim rs	

	strSQL="Select Distinct WCNING as Minuta from TESFL.TES960F1 where WCTCBT <> 'RCL' and WCFPAG=" & P_intFechaPago & " and WCNPRO=" & P_ORUsr & " or WCPRET=" & P_ORUsr
	GF_BD_AS400_2 rs,oConn,"OPEN",strSQL
	if not(rs.eof) then
	   'Se arma la lista de minutas validas
	   strLista="("
	   while (not rs.EOF)
		strLista= strLista & rs("Minuta") & ","
		rs.MoveNext
	   wend
	   strLista= strLista & "0)"
	   'Se totalizan los creditos
		strSQL= "Select KCPAGO, Sum(Importe) Importe From ( "
		strSQL= strSQL & "Select WDFOPG as KCPAGO, Sum(WDIMPP) Importe "
		strSQL= strSQL & "From TESFL.TES960F2 "
		strSQL= strSQL & "Where WDTCBT <> 'RCL' and  WDFPAG=" & P_intFechaPago & " and WDFOPG != ' ' and WDPOPG=" & P_KC & " and WDDBCR=2 and WDNING in " & strLista
		strSQL= strSQL & "Group by WDFOPG "
		strSQL= strSQL & "union "
		strSQL= strSQL & "Select WDFOPG as KCPAGO, -Sum(WDIMPP) Importe "
		strSQL= strSQL & "From TESFL.TES960F2 "
		strSQL= strSQL & "Where WDTCBT <> 'RCL' and  WDFPAG=" & P_intFechaPago & " and WDFOPG != ' ' and WDPOPG=" & P_KC & " and WDDBCR=1 and WDNING in " & strLista
		strSQL= strSQL & "Group by WDFOPG) Tabla "
		strSQL= strSQL & "Group by KCPAGO "
   	   GF_BD_AS400_2 P_rs,oConn,"OPEN",strSQL
	end if
    
End Function
'------------------------------------------------------------------------------------
Function GF_CAB_LEERFP(ByRef P_rs,P_strKCProveedor, P_intFechaPago,P_KCFormaPago)

	Dim strSQL,oConn
	
	strSQL= "Select WDFPAG as FechaPago, WDTCBT as TipoCbte, WDNING as Minuta "
	strSQL= strSQL & "from TESFL.TES960F2 "
	strSQL= strSQL & "where WDTCBT <> 'RCL' and WDPOPG=" & P_strKCProveedor & " and WDFOPG='" & P_KCFormaPago & "' and WDFPAG=" & P_intFechaPago
	strSQL= strSQL & "Group by WDFPAG,WDTCBT,WDNING	"
	GF_BD_AS400_2 P_rs,oConn,"OPEN",strSQL
	
End Function
'------------------------------------------------------------------------------------
Function GF_ES_CORREDOR(P_ORKC)
'Este fincion averigua si es un corredor o un vendedor.
'Devuelve true : Es Corredor.
'         false: Es Vendedor.   
	Dim oConn,rs,strSQL
	
	GF_ES_CORREDOR=false
	'JAS - Vieja manera - strSQL="Select WCNPRO as KCCOR from TESFL.TES960F1 where WCNPRO= " & P_ORKC & " and WCPRET <> " & P_ORKC
	strSQL="Select * from MERFL.MER311F1 where CCORR1=" & P_ORKC & " FETCH FIRST 1 ROW ONLY"
    Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)	
	if not(rs.EOF) then GF_ES_CORREDOR=true
			
End Function
'*******************************************************************************
Function CorteCbte(pRs, pCbte)
    CorteCbte = false
    if (not pRs.eof) then
        if (CLng(pRs("Minuta")) = CLng(pCbte)) then
            CorteCbte = true
        end if                    
    end if
End Function
'*******************************************************************************
'/* Funcion: GF_generarPagosImpresion
' * Descripcion: Esta funcion genera un PDF con el resumen de pagos a un periodo.
' * Parametros: p_oPDF objeto PDF.
' *				p_fechaDesde.
' *				p_fechaHasta.
' * Autor: N/N
' *
' * Modificado: Santi Juan Pablo
' * Fecha 15/07/2010
' */
Function GF_generarPagosImpresion(p_oPDF, p_Proveedor, p_fechaDesde, p_fechaHasta)
	dim coord_y, coord_x, aux_Importe, aux_Neto
	Server.ScriptTimeout = 1200
	'Leo las cabeceras de los comprobantes correspondientes
	strSQL="Select MTTCBT TIPO, substr(DSQGNB, 1, 12) as Comprobante, DIVIS3 PTOVENTA, COMPR3 NROFACTURA, mtning as Minuta, mtfpag as FechaPago, mtnpro as CodProveedor, mtNOPC as OrdenPago, MTPGCB as PagoCobro " &_
           " , case when MTPGCB='P' then mtimpp else mtimpp*-1 end as Importe " &_
           " , MTCODE as CodigoDetalle " &_
           "        from tesfl.tes134f1 " &_
           "           LEFT JOIN PROVFL.ACDSREL0 ON MTNING = DSQFNB" &_
           "           LEFT JOIN TESFL.TES111F3 ON TCBTA3=MTTCBT  AND NRCBA3=MTNING AND FCBTA3=MTFPAG " &_
           " where MTFPAG >= " & p_fechaDesde & " and MTFPAG <= " & p_fechaHasta & " and mtnpro= " & p_Proveedor &_
           " and MTCODE in ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'P', 'Q', 'R', 'T', 'U') " &_
           " ORDER BY  mtnpro, mtning, mtcode"
    Call GF_BD_AS400_2(rs, con, "OPEN", strSQL)          
	'Inicializao acumuladores
	acumTotal  = 0
	acum_Ret_1 = 0
	acum_Ret_2 = 0
	acum_Ret_3 = 0
	acum_Ret_4 = 0
	acum_Ret_5 = 0
	acum_Ret_6 = 0
	acum_Ret_7 = 0
	acum_Ret_8 = 0
	acum_Ret_9 = 0
	acum_Ret_P = 0
	acumNeto  = 0
	'Recorro los comprobantes e imprimo	
	while not rs.eof
		'Imprimo titulos, encabezado y cuadro
		Call dibujarEncabezadoPagos(p_oPDF, 50, p_Proveedor, p_fechaDesde, p_FechaHasta)		
		call dibujarTablaPagos(p_oPDF, 100)
		Call dibujarTitulosPagos(p_oPDF)		
		cont = 0
		coord_x = 135
		call GF_setFont(p_oPDF, "COURIER", 7, 0)		
		while (not rs.eof) and (cont < 42)
		    oldCbte = rs("Minuta")
		    myFechaPago = rs("FechaPago")
		    if ((rs("TIPO") = "FCA") or (rs("TIPO") = "NCA") or (rs("TIPO") = "NDA") or (rs("TIPO") = "FCB") or (rs("TIPO") = "NCB")) then		        
		        myCbte = GF_nDigits(cDBl(rs("PTOVENTA")),4) & GF_nDigits(cDBl(rs("NROFACTURA")),8)
		    else
		        myCbte = rs("Comprobante")
            end if		        
			aux_Importe = 0
			retencion1 = 0
			retencion2 = 0
			retencion3 = 0
			retencion4 = 0
			retencion5 = 0
			retencion6 = 0
			retencion7 = 0
			retencion8 = 0
			retencion9 = 0
			retencionP = 0			
			aux_Neto = 0
			while (CorteCbte(rs, oldCbte))				
				select case rs("CodigoDetalle")
				    case "A":
				        'Es el importe del cbte.
				        aux_Importe = CDbl(rs("Importe"))
				        acumTotal = acumTotal + aux_Importe		
				        aux_Neto = aux_Neto + aux_Importe
					case "C", "E", "T":
						retencion1 = CDbl(rs("Importe"))
						acum_Ret_1 = acum_Ret_1 + retencion1
						aux_Neto = aux_Neto - retencion1
					case "B", "R":
						retencion2 = CDbl(rs("Importe"))
						acum_Ret_2 = acum_Ret_2 + retencion2
						aux_Neto = aux_Neto - retencion2
					case "H":
						retencion3 = CDbl(rs("Importe"))
						acum_Ret_3 = acum_Ret_3 + retencion3
						aux_Neto = aux_Neto - retencion3
					case "D":
						retencion4 = CDbl(rs("Importe"))
						acum_Ret_4 = acum_Ret_4 + retencion4
						aux_Neto = aux_Neto - retencion4
					case "I":
						retencion5 = CDbl(rs("Importe"))
						acum_Ret_5 = acum_Ret_5 + retencion5
						aux_Neto = aux_Neto - retencion5
					case "Q":
						retencion6 = CDbl(rs("Importe"))
						acum_Ret_6 = acum_Ret_6 + retencion6
						aux_Neto = aux_Neto - retencion6
					case "K", "L", "M", "P":
						retencion7 = CDbl(rs("Importe"))
						acum_Ret_7 = acum_Ret_7 + retencion7
						aux_Neto = aux_Neto - retencion7
					case "G", "J":
						retencion8 = CDbl(rs("Importe"))
						acum_Ret_8 = acum_Ret_8 + retencion8
						aux_Neto = aux_Neto - retencion8
					case "U":
						retencion9 = CDbl(rs("Importe"))
						acum_Ret_9 = acum_Ret_9 + retencion9
						aux_Neto = aux_Neto - retencion9
					case "F":
						retencionP = CDbl(rs("Importe"))
						acum_Ret_P = acum_Ret_P + retencionP
						aux_Neto = aux_Neto - retencionP
				end select				
				rs.movenext
			wend	
			call GF_writeVerticalTExt(p_oPDF, coord_x, 840, GF_FN2DTE(myFechaPago), 50, PDF_ALIGN_CENTER)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 790, GF_EDIT_Cbte(GF_nDigits(cDBl(myCbte),12)), 60, PDF_ALIGN_CENTER)									
			call GF_writeVerticalTExt(p_oPDF, coord_x, 730, GF_EDIT_DECIMALS(aux_Importe*100 ,2 ), 75, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 650, GF_EDIT_DECIMALS(retencion1*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 595, GF_EDIT_DECIMALS(retencion2*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 540, GF_EDIT_DECIMALS(retencion3*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 485, GF_EDIT_DECIMALS(retencion4*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 430, GF_EDIT_DECIMALS(retencion5*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 375, GF_EDIT_DECIMALS(retencion6*100, 2),	50, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 320, GF_EDIT_DECIMALS(retencion7*100, 2),	55, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 260, GF_EDIT_DECIMALS(retencion8*100, 2),	55, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 200, GF_EDIT_DECIMALS(retencion*100, 2),	55, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 140, GF_EDIT_DECIMALS(retencionP*100, 2),	55, PDF_ALIGN_RIGHT)
			call GF_writeVerticalTExt(p_oPDF, coord_x, 80,  GF_EDIT_DECIMALS(aux_Neto*100, 2), 65, PDF_ALIGN_RIGHT)
			acumNeto = acumNeto + aux_Neto
			coord_x = coord_x + 10
			cont = cont + 1
        wend			
	    if cont = 42 then
		    call GF_newPage(p_oPDF)
		    Call PDFGirarHoja(90)
	    end if
	wend
	'Imprimo totales
	coord_x = coord_x + 10
	call GF_setFont(p_oPDF,"COURIER",7,8)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 790, "Totales",60, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 730, GF_EDIT_DECIMALS(acumTotal*100, 2),  78, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 650, GF_EDIT_DECIMALS(acum_Ret_1*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 595, GF_EDIT_DECIMALS(acum_Ret_2*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 540, GF_EDIT_DECIMALS(acum_Ret_3*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 485, GF_EDIT_DECIMALS(acum_Ret_4*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 430, GF_EDIT_DECIMALS(acum_Ret_5*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 375, GF_EDIT_DECIMALS(acum_Ret_6*100, 2), 53, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 320, GF_EDIT_DECIMALS(acum_Ret_7*100, 2), 58, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 260, GF_EDIT_DECIMALS(acum_Ret_*100, 2), 58, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 200, GF_EDIT_DECIMALS(acum_Ret_9*100, 2), 58, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 140, GF_EDIT_DECIMALS(acum_Ret_P*100, 2), 58, PDF_ALIGN_RIGHT)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 80,  GF_EDIT_DECIMALS(acumNeto*100, 2), 68, PDF_ALIGN_RIGHT)

end function
'------------------------------------------------------------------------------------
sub dibujarTablaPagos(p_oPDF, p_x)
    'Se dibuja el recuadro
	Call GF_squareBox(p_oPDF, p_x,		 10,  479, 830, 0, "#FFFFFF", "#006400", 2, PDF_SQUARE_ROUND)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 80,  478,  60, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 200, 478,  60, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 320, 478,  55, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 430, 478,  55, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 540, 478,  55, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 650, 478,  80, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 730, 478,  60, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 0.5, 80,   15, 570, 0, "#FFFFFF", "#006400", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(p_oPDF, p_x + 30, 10, 1, 830, 0, "#006400", "", 0, PDF_SQUARE_NORMAL)
end sub
'------------------------------------------------------------------------------------
Function dibujarEncabezadoPagos(p_oPDF, coord_x, p_Proveedor, p_fechaDesde, p_FechaHasta)
	'Determino si mi usuario es corredor o vendedor
	'Call GF_writeImage(p_oPDF, Server.MapPath("..\Images\kogge256.jpg"), 10, 810, 75, 57, 90)    
    Call GF_verticalLine(p_oPDF, 40, 40, 690)
    call GF_setFont(p_oPDF, "Arial", 20, 8)
    call GF_writeVerticalText(p_oPDF, 10, 350, GF_Traducir("Resumen de Pagos"),  300, 1)    
    call GF_setFont(p_oPDF, "Arial", 10, 0)
    Call GF_writeVerticalText(p_oPDF, coord_x, 730, GF_Traducir("Empresa") & " : ",  700, 0)
	call GF_setFont(p_oPDF, "", "", 12)
	call GF_writeVerticalText(p_oPDF, coord_x, 675, GetDSEnterprise2(p_Proveedor),  700, 0)	
    call GF_setFont(p_oPDF, "Arial", 10, 0)
    Call GF_writeVerticalText(p_oPDF, coord_x, 260, GF_Traducir("Momento Impresión") & " : ",  200, 0)
	call GF_setFont(p_oPDF, "", "", 12)
	call GF_writeVerticalText(p_oPDF, coord_x, 160, now(),  700, 0)
	call GF_setFont(p_oPDF, "", "", 0)	
    coord_x = coord_x + 20
	Call GF_writeVerticalText(p_oPDF, coord_x, 730, GF_Traducir("Período Pagos") & " : ",  700, 0)
	call GF_setFont(p_oPDF, "", "", 12)
	call GF_writeVerticalTExt(p_oPDF, coord_x, 650, GF_FN2DTE(p_fechaDesde) & " - " & GF_FN2DTE(p_FechaHasta),700, 0)
	call GF_setFont(p_oPDF, "", 9, 4)
    Call GF_writeVerticalText(p_oPDF, coord_x, 260, GF_Traducir("Los importes estan expresados en Pesos Argentinos"),  300, 0)
    
end Function
'------------------------------------------------------------------------------------
Function dibujarTitulosPagos(p_oPDF)
	call GF_setFont(p_oPDF, "", 9, 0)
    call GF_writeVerticalText(p_oPDF, 103, 650, "Retenciones"	,570, PDF_ALIGN_CENTER)
    call GF_writeVerticalText(p_oPDF, 110, 840, "Fecha"			,50, PDF_ALIGN_CENTER)
    call GF_writeVerticalText(p_oPDF, 110, 790, "Comprobante"	,60, PDF_ALIGN_CENTER)
	call GF_writeVerticalText(p_oPDF, 110, 730, "Total Cbte."	,80, PDF_ALIGN_CENTER)
	call GF_setFont(p_oPDF, "", 8, 0)
    call GF_writeVerticalText(p_oPDF, 118, 650, "IVA"	        ,55, PDF_ALIGN_CENTER)
	call GF_writeVerticalText(p_oPDF, 118, 595, "Ganancias"	    ,55, PDF_ALIGN_CENTER)
    call GF_writeVerticalTExt(p_oPDF, 118, 540, "IIBB Bs As"   	,55, PDF_ALIGN_CENTER)
    call GF_writeVerticalTExt(p_oPDF, 118, 485, "IIBB Sta Fe" 	,55, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 430, "IIBB CABA" 	,55, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 375, "IIBB L Pampa"	,60, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 320, "S.U.S.S."	    ,55, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 260, "Der.Reg. e Insp",60, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 200, "Biotecnologia"	,60, PDF_ALIGN_CENTER)
	call GF_writeVerticalTExt(p_oPDF, 118, 140, "Percepciones"	,60, PDF_ALIGN_CENTER)
    call GF_setFont(p_oPDF, "", 9, 0)
    call GF_writeVerticalTExt(p_oPDF, 110, 80, "Neto " & strTipoUsuario	,70, PDF_ALIGN_CENTER)
end Function
'------------------------------------------------------------------------------------
Function DeterminarNroFactura(pFecha, pNroCbte)
	Dim ptoVenta, rtrn
	
	ptoVenta = Left(pNroCbte, 4)
	rtrn = Right(pNroCbte, 8)
	if (CLng(pFecha) >= 20140801) then
		'Se asume que la factura es electronica y ya viene en formato XXXXYYYYYYYY.		
		Select case ptoVenta
			Case "0021"
				ptoVenta = "0031"
			Case "0026"
				ptoVenta = "0036"
			Case "0025"
				ptoVenta = "0035"
			Case "0018"
				ptoVenta = "0028"
			Case "0006"
				ptoVenta = "0006"
			Case "0004"
				ptoVenta = "0014"
		End Select	
	end if
	DeterminarNroFactura = ptoVenta & rtrn
End Function
%>
