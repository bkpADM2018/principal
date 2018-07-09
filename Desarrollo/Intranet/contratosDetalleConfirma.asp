<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosAS400.asp"-->
<!-- #include file="Includes/procedimientostraducir.asp"-->
<!-- #include file="Includes/procedimientosfechas.asp"-->
<!-- #include file="Includes/cor-IncludeCTO.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosSQL.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp"-->
<!-- #Include File="Includes/ExternalFunctions.ASP" -->
<!-- #include file="Includes/procedimientosMail.asp"-->
<!-- #include file="Includes/procedimientosValidacion.asp"-->

<%
Const ADDRESSEE_CONFIRM = "bacariniE@toepfer.com"
Const COND_IVA_C = "C"
Const COND_IVA_X = "X"
Call ProcedimientoControl("CONFWEB")
Dim unitDest, intMostrar, cont, auxDs, strTexto
DIM unitDestino, retValue, strLink, param, strCampoOrden, intPrecioPAnterior,precioTonelada
dim hayReconfirmacion, hayTablaFJ, hayTablaFK, hayTablaF2, estado, hayTablaFM
Dim rsPuertos, rsTransportes, flagGrabado
dim parametrosContratoOrig, dicErr
Dim errores, myIndexDic, ppIni, ppFin, dsVendedor,KilosDelNegocio
dim myValueCUITVendedor, myValuePrecio, auxCondicion, auxImp1, auxImp2, auxImp3, auxImpTotal, auxVTon
dim myKrToepfer, myStyleDiv, myPactadosOriginal, myPactadosModificado, myFechaPago, cdMoneda
myStyleDiv = "style='visibility:hidden;position:absolute;'"
hayReconfirmacion = false
hayTablaFJ = false
hayTablaFK = false
hayTablaF2 = false
hayTablaFM = false
Set dicErr = Server.CreateObject("Scripting.Dictionary")

estado = 2
myIndexDic = 0
'Se reciben los parametros.
accion= UCase(GF_Parametros7("accion","",6))
producto= GF_Parametros7("cmbProducto",0,6)
sucursal= GF_Parametros7("txtSucursal",0,6)
operacion= GF_Parametros7("txtOperacion",0,6)
numero= GF_Parametros7("txtNumero",0,6)
cosecha = GF_Parametros7("txtCosecha",0,6)
cdMoneda = GF_Parametros7("cdMoneda","",6)
Set rsOriginal = readDatosOriginales(producto, sucursal, operacion, numero, cosecha)
if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then 
	if cdMoneda = "" then cdMoneda = getMonedaOperacion(rsOriginal("Operacion"))
end if	
'Response.Write rsOriginal.eof
if (rsOriginal.eof) then
	setError(CONTRATO_NO_EXISTE)
else
	ppIni = "bodyOnLoad();"
	flagGrabado = false	
	Set rsModificado = readDatosModificados(producto, sucursal, operacion, numero, cosecha)	
	'Response.write cDbl(datoModificado("Kilos")) 
	'Response.write cDbl(rsOriginal("AmpAnu"))
	myValueCUITVendedor = GF_Parametros7("CUITVendedor","",6)
	myValuePrecio = GF_Parametros7("Precio","",6)
	myValuePrecio = replace(myValuePrecio, ",", ".")
	myFechaPago = GF_Parametros7("FechaPago","",6)
	if myFechaPago = "" then myFechaPago = rsOriginal("FechaPago")
	if (isFormSubmit())then
		'Se controla la info.
		if controlarParametrosConfirmacion then
			if (ucase(accion) = ucase(ACCION_GRABAR)) then
				'Response.End 
				if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then
					call modificarContratoToepfer()
				else
					if hayReconfirmacion then 
						if grabarReconfirmacion() then
							estado = 1
						end if
					else
						call modificarContrato()	
						estado = 3
					end if
				end if
				ppFin = "info();"
				envioMails
			end if
		end if
	end if
end if
'---------------------------------------------------------------------------------------
Function getMonedaOperacion(p_operacion)
	getMonedaOperacion = MONEDA_PESO
	select case CInt(p_operacion)
		case 6,9,10,11,12
			getMonedaOperacion = MONEDA_DOLAR
	end select	
End Function
'---------------------------------------------------------------------------------------
'/**
' * Funcion    : readDatosOriginales
' * Descripcion: Trae el registro de contrato Original
' * Autor: Javier A. Scalisi
' * Fecha: 29/09/2010
' */
Function readDatosOriginales(producto, sucursal, operacion, numero, cosecha)
    Dim strWhere, strORKC, oConn, strSQL, strOrder
    Call mkWhere(strWhere, "CTO.CPROR1", producto	, "=", 1)
    Call mkWhere(strWhere, "CTO.CSUCR1", sucursal	, "=", 1)
    Call mkWhere(strWhere, "CTO.COPER1", operacion	, "=", 1)
    Call mkWhere(strWhere, "CTO.NCTOR1", numero		, "=", 1)
    Call mkWhere(strWhere, "CTO.ACOSR1", cosecha	, "=", 1)
    
    Call mkWhere(strWhere, "CTO.CONFR1", "F", "=", 3)
	
    'Los contratos de cebada solo deben mostrarse cuando sea para la consulta de la caratula
    'pero no para imprimir boleto o confirmar el contrato.
    'Caso especial solicitado por Ronchel
    'Si la Operacion es 9 se deben poder confirmar los contratos de Cebada(17) y Colza(09)
	'strWhere = strWhere & " AND ((CTO.CPROR1 not in (17, 9, 25)) or (CTO.CPROR1 in (17, 9) and CTO.COPER1 = 9)) "         
	'strWhere = strWhere & " AND CTO.COPER1 <> 04 "         
    
    'Se toman contratos de la cosecha 09 en adelante
    Call mkWhere(strWhere, "CTO.ACOSR1", "11", ">=", 1)

    'Se arma la SQL. FechaEntDesde
    strSQL = "Select CTO.CPROR1 as Producto, CTO.CSUCR1 as Sucursal, CTO.COPER1 as Operacion, CTO.NCTOR1 as Numero, CTO.ACOSR1 as Cosecha, CTO.FCCTR1 as FechaConc"
	strSQL = strSQL & ", CTO.CCORR1 as KCCOR, CTO.CTRAR1 as IdTransporte, CTO.CVENR1 as KCVEN, CTO.CDESR1 as PtoRecepcion, CTO.DESTR1 as PtoDevolucion, CTO.FDPER1 as FecEntDesde, CTO.FHPER1 as FecEntHasta, VEN.NRODOC CUITVendedor, CTO.CONCR1 CtoCorredor"
	strSQL = strSQL & ", P.DESCPC as Procedencia, CTO.FDFIR1 as FecFijaDesde, CTO.FHFIR1 as FecFijaHasta, CTO.KGNFR1 as CantKilosMin, CTO.KGMFR1 as CantKilosMax, P.CODIPC CPProcedencia, P.AUXIPC CAProcedencia, case(J.CDIVRJ) when '" & COND_IVA_C & "' then J.CDIVRJ else '" & COND_IVA_X & "' end as CondicionIVA"
	strSQL = strSQL & ", J.MCPDRJ as MercPropia, J.MCONRJ as MercConsigna, J.HUMERJ as MercHumedad, CASE WHEN SIO.SIOGRANOS is Null THEN '' else SIO.SIOGRANOS END CODIGOSIO "	
	strSQL = strSQL & ", case(M.MDOLOM) when 'F' then DOUBLE(CTO.PRECR1) else DOUBLE(CTO.PRECR1*CTO.TIPCR1) end as PrecioP, case(M.MDOLOM) when 'F' then case(CTO.TIPCR1) when 0 then 0 else DOUBLE(CTO.PRECR1/CTO.TIPCR1) end else DOUBLE(CTO.PRECR1) end as PrecioD, DOUBLE(CTO.PORPR1) as PjeParcial, CTO.DIAPR1 as DiasPago, CTO.CFPAR1 as CodigoPago, CASE WHEN (K.CAPARK) IS NULL THEN 0 ELSE K.CAPARK END CAMIONESPACTADOS, CTO.FPACR1 as FechaPago, '' as Observaciones "
	strSQL = strSQL & ", CASE WHEN AAC.TOTAL IS NULL THEN 0 ELSE AAC.TOTAL END as AmpAnu "
	strSQL = strSQL & ", CASE WHEN AAC.TOTAL IS NULL THEN CTO.KGCOR1 ELSE CTO.KGCOR1 + AAC.TOTAL END as Kilos from MERFL.MER311F1 CTO "
	strSQL = strSQL & "       LEFT JOIN"
	strSQL = strSQL & "               ("
	strSQL = strSQL & "               SELECT CPRORB,CSUCRB,COPERB,NCTORB,ACOSRB, SUM(KGCORB) AS TOTAL FROM MERFL.MER311FB GROUP BY CPRORB,CSUCRB,COPERB,NCTORB,ACOSRB"
	strSQL = strSQL & "               )AAC"
	strSQL = strSQL & "           on CTO.CPROR1=AAC.CPRORB and CTO.CSUCR1=AAC.CSUCRB"
	strSQL = strSQL & "           and CTO.COPER1=AAC.COPERB and CTO.NCTOR1=AAC.NCTORB and CTO.ACOSR1=AAC.ACOSRB"
	strSQL = strSQL & "		inner join MERFL.TCB6A1F1 VEN on CTO.CVENR1=VEN.NROPRO "	
	strSQL = strSQL & "		left  join MERFL.MER142F1 P on CTO.CPRDR1=P.CODIPC and CTO.AUXIR1=P.AUXIPC "	
	strSQL = strSQL & "		left  join MERFL.MER311FJ J on CTO.CPROR1=J.CPRORJ and  CTO.CSUCR1=J.CSUCRJ and CTO.COPER1=J.COPERJ and CTO.NCTOR1=J.NCTORJ and CTO.ACOSR1=J.ACOSRJ "
	strSQL = strSQL & "		left  join MERFL.MER341F2 BOL on CTO.CPROR1=BOL.PLCPRO and CTO.CSUCR1=BOL.PLCSUC and CTO.COPER1=BOL.PLCOPE and CTO.NCTOR1=BOL.PLNCTO and CTO.ACOSR1=BOL.PLACOS "
	strSQL = strSQL & "		left  join MERFL.MER311FK K on CTO.CPROR1=K.CPRORK and CTO.CSUCR1=K.CSUCRK and CTO.COPER1=K.COPERK and CTO.NCTOR1=K.NCTORK and CTO.ACOSR1=K.ACOSRK "
	strSQL = strSQL & "		left  join MERFL.MER132F1 M on CTO.COPER1=M.CODIOM "
	strSQL = strSQL & "		left  join MERFL.MER311FM SIO on CTO.CPROR1=SIO.PRODUCTO and CTO.CSUCR1=SIO.SUCURSAL and CTO.COPER1=SIO.OPERACION and CTO.NCTOR1=SIO.NUMERO and CTO.ACOSR1=SIO.COSECHA "
	strSQL = strSQL & strWhere
    'response.write "<hr>" &  strSQL & "<hr>"
    Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    Set readDatosOriginales = rs    
    
  End Function  
'---------------------------------------------------------------------------------------'/**
' * Funcion: readDatosModificados
' * Descripcion: Esta funcion lee los datos mosdificados del contrato.
' * Valor Devuelto: recordset con los datos leidos.
' * Autor: Javier A. Scalisi
' * Fecha 29/09/2010
' */
Function readDatosModificados(producto, sucursal, operacion, numero, cosecha)
	Dim ret
	if (isFormSubmit()) then
		Set ret = readDatosModificadosParam()
	else
		Set ret = readDatosModificadosDB(producto, sucursal, operacion, numero, cosecha)
	end if
	
	Set readDatosModificados = ret
End Function
'---------------------------------------------------------------------------------------
Function readDatosModificadosDB(producto, sucursal, operacion, numero, cosecha)
    Dim strSQL, oConn, rs, strWhere
    Call mkWhere(strWhere, "PRODUCTO"	, producto	, "=", 1)
    Call mkWhere(strWhere, "SUCURSAL"	, sucursal	, "=", 1)
    Call mkWhere(strWhere, "OPERACION"	, operacion	, "=", 1)
    Call mkWhere(strWhere, "NUMERO"		, numero	, "=", 1)
    Call mkWhere(strWhere, "COSECHA"	, cosecha	, "=", 1)
    strSQL = "Select PRODUCTO,	SUCURSAL,	OPERACION,	NUMERO,	COSECHA,	CTOCORREDOR,	CUITVENDEDOR,	CONDICIONIVA,	"
    strSQL = strSQL & " KILOS,	PROCEDENCIA,	CPPROCEDENCIA,	CAPROCEDENCIA,	MERCPROPIA,	MERCCONSIGNA,	DOUBLE(PJEPARCIAL) PJEPARCIAL,	"
    strSQL = strSQL & " CODIGOPAGO,	FECENTDESDE,	FECENTHASTA,	PTORECEPCION,	IDTRANSPORTE,	FECFIJADESDE,	FECFIJAHASTA,"	
    strSQL = strSQL & " CANTKILOSMIN,	CANTKILOSMAX,	OBSERVACIONES,	DOUBLE(PRECIOP) PRECIOP,	DOUBLE(PRECIOD) PRECIOD,	USUARIO,	REGISTRO,	CODIGOSIO "
    strSQL = strSQL & ", CASE WHEN (K.CAPARK) IS NULL THEN 0 ELSE K.CAPARK END CAMIONESPACTADOS from TOEPFERDB.TBLCONTRATOSCONF CC " 
    strSQL = strSQL & "	left join MERFL.MER311FK K on CC.PRODUCTO=K.CPRORK and CC.SUCURSAL=K.CSUCRK and CC.OPERACION=K.COPERK and CC.NUMERO=K.NCTORK and CC.COSECHA=K.ACOSRK "
    strSQL = strSQL & strWhere
    Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)    
    Set readDatosModificadosDB = rs
    
End Function
'---------------------------------------------------------------------------------------
Function readDatosModificadosParam()
	Dim dic
	', temp
	
	Set dic = Server.CreateObject("Scripting.Dictionary")
	
	'Se reciben los parametros.
	Call readDatosParam("ctoCorredor"	, dic)
	Call readDatosParam("KCVEN"			, dic)
	Call readDatosParam("CUITVendedor"	, dic)
	Call readDatosParam("CondicionIVA"	, dic)
	Call readDatosParam("Kilos"			, dic)
	Call readDatosParam("CamionesPactados", dic)
	Call readDatosParam("Procedencia"	, dic)
	Call readDatosParam("CAProcedencia"	, dic)
	Call readDatosParam("CPProcedencia"	, dic)
	temp = GF_Parametros7("TipoMercaderia", "", 6)
	dic("MercPropia")=	"F"
	dic("MercConsigna")= "F"      
	if (temp = TIPO_PROPIA_PRODUCCION)	then dic("MercPropia")	= "V"	
	if (temp = TIPO_CONSIGNACION)		then  dic("MercConsigna") = "V"      	
	'Para el contrato el precio es uno solo, la moneda y el tipo de operacion definen a que corresponde.
	Call readDatosParam("Precio", dic)
	dic("PrecioP")= dic("Precio")
	dic("PrecioD")= dic("Precio")
	Call readDatosParam("PjeParcial"	, dic)
	Call readDatosParam("CodigoPago"	, dic)
	Call readDatosParam("FecEntDesde"	, dic)
	Call readDatosParam("FecEntHasta"	, dic)
	Call readDatosParam("PtoRecepcion"	, dic)
	Call readDatosParam("PtoDevolucion"	, dic)
	Call readDatosParam("IdTransporte"	, dic)
	Call readDatosParam("FecFijaDesde"	, dic)
	Call readDatosParam("FecFijaHasta"	, dic)
	Call readDatosParam("CantKilosMax"	, dic)
	Call readDatosParam("CantKilosMin"	, dic)
	Call readDatosParam("Observaciones"	, dic)
	Call readDatosParam("CODIGOSIO"	, dic)
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Fields.Append "ctoCorredor", 12
	rs.Fields.Append "KCVEN", 12
	rs.Fields.Append "CUITVendedor", 12
	rs.Fields.Append "CondicionIVA", 12
	rs.Fields.Append "Kilos", 12
	rs.Fields.Append "CamionesPactados", 12
	rs.Fields.Append "Procedencia", 12
	rs.Fields.Append "CAProcedencia", 12
	rs.Fields.Append "CPProcedencia", 12
	rs.Fields.Append "MercPropia", 12
	rs.Fields.Append "MercConsigna", 12
	rs.Fields.Append "Precio", 12
	rs.Fields.Append "PrecioP", 12
	rs.Fields.Append "PrecioD", 12
	rs.Fields.Append "PjeParcial", 12
	rs.Fields.Append "CodigoPago", 12	
	rs.Fields.Append "FecEntDesde", 12			
	rs.Fields.Append "FecEntHasta", 12	
	rs.Fields.Append "PtoRecepcion", 12				
	rs.Fields.Append "PtoDevolucion", 12				
	rs.Fields.Append "IdTransporte", 12	
	rs.Fields.Append "FecFijaDesde", 12	
	rs.Fields.Append "FecFijaHasta", 12	
	rs.Fields.Append "CantKilosMax", 12	
	rs.Fields.Append "CantKilosMin", 12				
	rs.Fields.Append "Observaciones", 12	
	rs.Fields.Append "CODIGOSIO", 12	
	rs.Open
    Call rs.AddNew(dic.Keys, dic.Items)
	Set readDatosModificadosParam = rs
End Function
'---------------------------------------------------------------------------------------
Function readDatosParam(paramName, ByRef dic)
	select case ucase(paramName)
		case "KILOS", "CANTKILOSMIN", "CANTKILOSMAX", "CAMIONESPACTADOS"
			dic(paramName)= GF_Parametros7(paramName, 0, 6)
		case "PRECIO", "PJEPARCIAL"
			dic(paramName)= GF_Parametros7(paramName, "", 6)
			dic(paramName) = replace(dic(paramName),",",".")
		case else
			dic(paramName)= GF_Parametros7(paramName, "", 6)
	end select
End Function
'---------------------------------------------------------------------------------------
Function datoOriginal(p_dato)
    'Si es usuario interno, mostrarle el valor de campo original
    if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then
        if (p_dato <> "") then
            datoOriginal = "(" & p_dato & ")"
        else
            datoOriginal = "(...)"
        end if
    else
       datoOriginal = ""
    end if

End Function
'---------------------------------------------------------------------------------------
'Permite chequear si el recordset con datos modificados esta lleno o no de manera de 
'mostrar la info correctamente.
Function datoModificado(p_dato)
	if not rsModificado.eof then	
		datoModificado = rsModificado(p_dato)
	else
		datoModificado = rsOriginal(p_dato)
	end if
	if ucase(p_dato) = "PROCEDENCIA" or UCASE(p_dato)="OBSERVACIONES" then
		if not isNull(datoModificado) then datoModificado = trim(replace(datoModificado,"'","''"))
	end if	
End Function
'---------------------------------------------------------------------------------------
Function getValor(pValue)
	if trim(pValue) = "" then
		getValor = "null"
	else
		getValor = pValue	
	end if
End Function
'---------------------------------------------------------------------------------------
Function getNumero(pValue)
	'Response.Write "(" & pValue & ")"
	if trim(pValue) = "" then
		getNumero = "0"
	else
		getNumero = pValue	
	end if
End Function
'---------------------------------------------------------------------------------------
Function getTrueFalse(pValue)
	if trim(pValue) = "" then
		getTrueFalse = "F"
	else
		getTrueFalse = pValue	
	end if
End Function
'---------------------------------------------------------------------------------------
Function resaltar(p_datoNuevo, p_datoOriginal)
resaltar = ""
    'resalta en rojo el input, si dato NUevo es distinto de dato original. Es solo para usr. Interno    
    if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then
		if (p_datoNuevo <> p_datoOriginal) then
		   resaltar = "INPUTRESALTE"
		end if
    end if
end Function
'---------------------------------------------------------------------------------------
Function esPagoDirectoVendedor(pKcVendedor)
    Dim strWhere, strORKC, oConn, strSQL, strOrder, rsPagoDirecto
    strWhere = ""
    if (pKcVendedor <> "") then
        Call mkWhere(strWhere, "GAOPR1", pKcVendedor,"=",1)
        strSQL= "Select * from PROVFL.AAD4CPP  " & strWhere
        Call GF_BD_AS400_2(rsPagoDirecto,oConn,"OPEN",strSQL)
        if rsPagoDirecto.eof then
            esPagoDirectoVendedor = false
        else
            esPagoDirectoVendedor = true
        end if
        Call GF_BD_AS400_2(rsPagoDirecto,oConn,"CLOSE",strSQL)
    else
        esPagoDirectoVendedor = false
    end if
end function
'-------------------------------------------------------------------------------------------------
function modificarContrato()
dim mySql, myRs, myCn, confimado, myCondIVA
modificarContrato = false
	myCondIVA = datoModificado("CondicionIVA")
	if (datoModificado("CondicionIVA") <> COND_IVA_C) then myCondIVA = ""	
	mySql = "SELECT * FROM MERFL.MER311F1 WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
	call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
	if not myRs.eof then
		confimado = "F"
		if not hayReconfirmacion then confimado = "V"
		mySql = "INSERT INTO MERFL.LOG311F1 SELECT '" & session("Usuario") & "', 'CONFCONT', '', " & replace(session("MomentoDato"),"'","") & ", CPROR1,CSUCR1,COPER1,NCTOR1,ACOSR1,CACOR1,CCORR1,CDESR1,CFFCR1,CFPAR1,CMONR1,COPCR1,CPRDR1,AUXIR1,CTRAR1,CVENR1,CONCR1,CONFR1,REENR1,DESTR1,DIAPR1,HACOR1,DTFLR1,GSDLR1,VLDLR1,FENVR1,FCCTR1,FDPER1,FHPER1,FPACR1,FVTOR1,FDFIR1,FHFIR1,KGCOR1,KGMFR1,KGNFR1,KGFIR1,KGFOR1,KGPAR1,KGRER1,LALDR1,LALHR1,MARBR1,MARDR1,MARLR1,MARTR1,NREGR1,OPEVR1,DFL$R1,GAS$R1,VLM$R1,PORCR1,PORPR1,PRECR1,PTOFR1,REFGR1,REGBR1,PORFR1,TARFR1,TARGR1,TIPCR1,TRADR1,TIPOR1,KGCUR1,MERCR1 FROM MERFL.MER311F1 WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
		'Response.Write "<HR>INSERT F1 LOG <BR>" & mySql 
		call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
		mySql = "UPDATE MERFL.MER311F1 SET CONFR1='" & confimado & "',CONCR1='" & datoModificado("CtoCorredor") & "', CPRDR1=" & getNumero(datoModificado("CPProcedencia")) & ", AUXIR1=" & getNumero(datoModificado("CAProcedencia")) & ", FDPER1=" & getNumero(datoModificado("FecEntDesde")) & ", FHPER1=" & getNumero(datoModificado("FecEntHasta")) & ", FDFIR1=" & getNumero(datoModificado("FecFijaDesde")) & ", FHFIR1=" & getNumero(datoModificado("FecFijaHasta")) & ", CFPAR1='" & datoModificado("CodigoPago") & "', KGNFR1=" & getNumero(datoModificado("CANTKILOSMIN")) & ", KGMFR1=" & getNumero(datoModificado("CANTKILOSMAX")) & " WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
		'Response.Write "<HR>UPDATE F1 <BR>" & mySql 		
		call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
	end if
	'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	if hayTablaFJ then
		mySql = "SELECT * FROM MERFL.MER311FJ WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
		'Response.Write "<HR>SELECT FJ <BR>" & mySql 
		call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
		if not myRs.eof then
			mySql = "INSERT INTO MERFL.LOG311FJ SELECT CPRORJ,CSUCRJ,COPERJ,NCTORJ,ACOSRJ,HUMERJ,CDIVRJ,MCPDRJ,MCONRJ, '" & session("Usuario") & "', " & replace(session("MomentoDato"),"'","") & " FROM MERFL.MER311FJ WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'Response.Write "<HR>INSERT FJ LOG <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
			mySql = "UPDATE MERFL.MER311FJ SET CDIVRJ='" & myCondIVA & "', MCPDRJ='" & datoModificado("MercPropia") & "', MCONRJ='" & datoModificado("MercConsigna") & "' WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'mySql = "UPDATE MERFL.MER311FJ SET CDIVRJ='" & datoModificado("CondicionIVA") & "' WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'Response.Write "<HR>UPDATE FJ <BR>" & mySql2 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		else
			'mySql = "INSERT INTO MERFL.MER311FJ VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & ",''," & datoModificado("CondicionIVA") & "','" & datoModificado("MercPropia") & "','" & datoModificado("MercConsigna") & "')" 
			mySql = "INSERT INTO MERFL.MER311FJ VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & ",''," & myCondIVA & "','" & datoModificado("MercPropia") & "','" & datoModificado("MercConsigna") & "')" 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		end if
		'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	end if	
	if hayTablaFM then
		Call executeSP(myRs, "MERFL.MER311FM_GET_BY_FILTERS", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||||" & session("Usuario") & "||1||0")
		if not myRs.eof then
		    Call executeSP(rs, "MERFL.MER311FM_UPD", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||" & datoModificado("CODIGOSIO"))
		else
			Call executeSP(rs, "MERFL.MER311FM_INS", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||" & datoModificado("CODIGOSIO"))
		end if		
	end if	
	modificarContrato = true
end function
'-------------------------------------------------------------------------------------------------
function modificarContratoToepfer()
dim mySql, myRs, myCn, myCondIVA
modificarContratoToepfer = false
	myCondIVA = datoModificado("CondicionIVA")
	if (datoModificado("CondicionIVA") <> COND_IVA_C) then myCondIVA = ""	
	mySql = "SELECT * FROM MERFL.MER311F1 WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
	call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
	if not myRs.eof then
		mySql = "INSERT INTO MERFL.LOG311F1 SELECT '" & session("Usuario") & "', 'CONFCONT', '', " & replace(session("MomentoDato"),"'","") & ", CPROR1,CSUCR1,COPER1,NCTOR1,ACOSR1,CACOR1,CCORR1,CDESR1,CFFCR1,CFPAR1,CMONR1,COPCR1,CPRDR1,AUXIR1,CTRAR1,CVENR1,CONCR1,CONFR1,REENR1,DESTR1,DIAPR1,HACOR1,DTFLR1,GSDLR1,VLDLR1,FENVR1,FCCTR1,FDPER1,FHPER1,FPACR1,FVTOR1,FDFIR1,FHFIR1,KGCOR1,KGMFR1,KGNFR1,KGFIR1,KGFOR1,KGPAR1,KGRER1,LALDR1,LALHR1,MARBR1,MARDR1,MARLR1,MARTR1,NREGR1,OPEVR1,DFL$R1,GAS$R1,VLM$R1,PORCR1,PORPR1,PRECR1,PTOFR1,REFGR1,REGBR1,PORFR1,TARFR1,TARGR1,TIPCR1,TRADR1,TIPOR1,KGCUR1,MERCR1 FROM MERFL.MER311F1 WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
		'Response.Write "<HR>INSERT F1 LOG <BR>" & mySql 
		call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
		mySql = "UPDATE MERFL.MER311F1 SET CVENR1=" & datoModificado("KCVEN") & ", CONFR1='V', CONCR1='" & datoModificado("CtoCorredor") & "', KGCOR1=" & KilosDelNegocio & ", CPRDR1=" & getNumero(datoModificado("CPProcedencia")) & ", AUXIR1=" & getNumero(datoModificado("CAProcedencia")) & ", FDPER1=" & getNumero(datoModificado("FecEntDesde")) & ", FHPER1=" & getNumero(datoModificado("FecEntHasta")) & ", FPACR1=" & getNumero(myFechaPago) & ", CDESR1=" & getNumero(datoModificado("PtoRecepcion")) & ", CTRAR1=" & getNumero(datoModificado("IdTransporte")) & ", FDFIR1=" & getNumero(datoModificado("FecFijaDesde")) & ", FHFIR1=" & getNumero(datoModificado("FecFijaHasta")) & ", PRECR1=" & datoModificado("Precio") & ",PORPR1=" & datoModificado("PjeParcial") & ", CFPAR1='" & datoModificado("CodigoPago") & "', KGNFR1=" & getNumero(datoModificado("CANTKILOSMIN")) & ", KGMFR1=" & getNumero(datoModificado("CANTKILOSMAX")) & " WHERE CPROR1=" & producto & " AND CSUCR1=" & sucursal & " AND COPER1=" & operacion & " AND NCTOR1=" & numero & " AND ACOSR1=" & cosecha
		'Response.Write "<HR>UPDATE F1 <BR>" & mySql 		
		call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
	end if
	'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	if hayTablaFJ then
		mySql = "SELECT * FROM MERFL.MER311FJ WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
		call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
		if not myRs.eof then
			mySql = "INSERT INTO MERFL.LOG311FJ SELECT CPRORJ,CSUCRJ,COPERJ,NCTORJ,ACOSRJ,HUMERJ,CDIVRJ,MCPDRJ,MCONRJ, '" & session("Usuario") & "', " & replace(session("MomentoDato"),"'","") & " FROM MERFL.MER311FJ WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'Response.Write "<HR>INSERT FJ LOG <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
			mySql = "UPDATE MERFL.MER311FJ SET CDIVRJ='" & myCondIVA & "', MCPDRJ='" & datoModificado("MercPropia") & "', MCONRJ='" & datoModificado("MercConsigna") & "' WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'mySql = "UPDATE MERFL.MER311FJ SET CDIVRJ='" & datoModificado("CondicionIVA") & "' WHERE CPRORJ=" & producto & " AND CSUCRJ=" & sucursal & " AND COPERJ=" & operacion & " AND NCTORJ=" & numero & " AND ACOSRJ=" & cosecha
			'Response.Write "<HR>UPDATE FJ <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		end if
		'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	end if	
	if hayTablaFK then
		mySql = "SELECT * FROM MERFL.MER311FK WHERE CPRORK=" & producto & " AND CSUCRK=" & sucursal & " AND COPERK=" & operacion & " AND NCTORK=" & numero & " AND ACOSRK=" & cosecha
		'Response.Write "<HR>SELECT FK <BR>" & mySql 
		call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
		if not myRs.eof then
			mySql = "UPDATE MERFL.MER311FK SET CAPARK=" & datoModificado("CamionesPactados") & ",CARERK=" & datoModificado("CamionesPactados") & ",MARCRK='V' WHERE CPRORK=" & producto & " AND CSUCRK=" & sucursal & " AND COPERK=" & operacion & " AND NCTORK=" & numero & " AND ACOSRK=" & cosecha
			'Response.Write "<HR>UPDATE FK <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		else
			if cint(datoModificado("CamionesPactados")) then
				mySql = "INSERT INTO MERFL.MER311FK VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & "," & datoModificado("CamionesPactados") & "," & datoModificado("CamionesPactados") & ",'V')" 
				'Response.Write "<HR>Insert FK <BR>" & mySql 
				call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
			end if
		end if
		'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	end if	
	if hayTablaF2 then
		auxImp1 = round((datoModificado("Kilos") / 1000) * datoModificado("Precio"),2) 'Toneladas * Precio
		'replace(datoModificado("Precio"),",",".")
		auxImp2 = round(auxImp1 * 0.105,2) 'IVA 10,5
		auxImp3	= round((auxImp1 + auxImp2) * 0.2,2)
		auxImpTotal = auxImp1 + auxImp2 + auxImp3 
		'Calculo de Valor Tonelada
		auxVTon = datoModificado("Precio")
		'Asignacion de condicion s/ codigo de pago
		select case ucase(datoModificado("CodigoPago"))
			case "J" 
				auxCondicion = 16
			case "K" 
				auxCondicion = 17
			case "R" 
				auxCondicion = 5				
			case "T" 
				auxCondicion = 14				
			case "Z" 
				auxCondicion = 18				
			case else
				auxCondicion = 0
		end select		
		'Response.Write "<br>Valores F2<br>auxImp1=" & auxImp1 & "<br>auxImp2=" & auxImp2 & "<br>auxImp3=" & auxImp3 & "<br>auxImpTotal=" & auxImpTotal & "<br>auxVTon=" & auxVTon & "<br>auxCondicion=" & auxCondicion 

		mySql = "SELECT * FROM MERFL.MER311F2 WHERE CPROR2=" & producto & " AND CSUCR2=" & sucursal & " AND COPER2=" & operacion & " AND NCTOR2=" & numero & " AND ACOSR2=" & cosecha
		'Response.Write "<HR>SELECT F2 <BR>" & mySql 
		call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
		if not myRs.eof then
			mySql = "UPDATE MERFL.MER311F2 SET CONDR2=" & auxCondicion & ",TIPOR2='GC' WHERE CPROR2=" & producto & " AND CSUCR2=" & sucursal & " AND COPER2=" & operacion & " AND NCTOR2=" & numero & " AND ACOSR2=" & cosecha
			'Response.Write "<HR>UPDATE F2 <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		else
			mySql = "INSERT INTO MERFL.MER311F2 VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & "," & auxCondicion & ",'GC'," & datoModificado("Kilos") & "," & auxImpTotal & "," & auxVTon & "," & getNumero(datoModificado("FecEntHasta")) & ",'F',0,1)" 
			'Response.Write "<HR>Insert F2 <BR>" & mySql 
			call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)			
		end if
		'call GF_BD_AS400_2(myRs,myCn,"CLOSE",mySql)
	end if		
	if hayTablaFM then
		Call executeSP(myRs, "MERFL.MER311FM_GET_BY_FILTERS", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||||" & session("Usuario") & "||1||0")
		if not myRs.eof then
		    Call executeSP(rs, "MERFL.MER311FM_UPD", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||" & datoModificado("CODIGOSIO"))
		else
			Call executeSP(rs, "MERFL.MER311FM_INS", producto & "||" & sucursal & "||" & operacion & "||" & numero & "||" & cosecha & "||" & datoModificado("CODIGOSIO"))
		end if		
	end if	
	modificarContratoToepfer = true
end function
'-------------------------------------------------------------------------------------------------
function grabarReconfirmacion()
dim mySql, myRs, myCn
grabarReconfirmacion = false
	mySql = "SELECT * FROM TOEPFERDB.TBLCONTRATOSCONF WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha
	call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
	if myRs.eof then
		mySql = "INSERT INTO TOEPFERDB.TBLCONTRATOSCONF(PRODUCTO,SUCURSAL,OPERACION,NUMERO,COSECHA,CTOCORREDOR,CUITVENDEDOR,CONDICIONIVA,KILOS,CPPROCEDENCIA,CAPROCEDENCIA,MERCPROPIA,MERCCONSIGNA,PRECIOP,PRECIOD,PJEPARCIAL,CODIGOPAGO, PROCEDENCIA, FECENTDESDE, FECENTHASTA, PTORECEPCION, IDTRANSPORTE, FECFIJADESDE, FECFIJAHASTA, CANTKILOSMIN, CANTKILOSMAX, OBSERVACIONES, CODIGOSIO) VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & ",'" & datoModificado("ctoCorredor") & "'," & getValor(datoModificado("CUITVendedor")) & ",'" & datoModificado("CondicionIVA") & "'," & datoModificado("Kilos") & ",'" & datoModificado("CPPROCEDENCIA") & "'," & getValor(datoModificado("CAPROCEDENCIA")) & ",'" & datoModificado("MercPropia") & "','" & datoModificado("MercConsigna") & "'," & datoModificado("PrecioP") & "," & datoModificado("PrecioD") & "," & datoModificado("PjeParcial") & ",'" & datoModificado("CodigoPago") & "','" & datoModificado("PROCEDENCIA") & "'," & getValor(datoModificado("FECENTDESDE")) & "," & getValor(datoModificado("FECENTHASTA")) & "," & getValor(datoModificado("PtoRecepcion")) & "," & getValor(datoModificado("IDTRANSPORTE")) & "," & getValor(datoModificado("FECFIJADESDE")) & "," & getValor(datoModificado("FECFIJAHASTA")) & "," & getValor(datoModificado("CANTKILOSMIN")) & "," & getValor(datoModificado("CANTKILOSMAX")) & ",'" & datoModificado("OBSERVACIONES") & "', '" & datoModificado("CODIGOSIO") & "')"
	else
		mySql = "UPDATE TOEPFERDB.TBLCONTRATOSCONF SET MERCPROPIA='" & datoModificado("MercPropia") & "', MERCCONSIGNA='" & datoModificado("MercConsigna") & "', CTOCORREDOR='" & datoModificado("ctoCorredor") & "', CUITVENDEDOR=" & datoModificado("CUITVendedor") & ",CONDICIONIVA='" & datoModificado("CondicionIVA") & "',KILOS=" & datoModificado("Kilos") & ",PROCEDENCIA='" & datoModificado("PROCEDENCIA") & "', CPPROCEDENCIA='" & datoModificado("CPPROCEDENCIA") & "',CAPROCEDENCIA=" & datoModificado("CAPROCEDENCIA") & ",PRECIOP=" & datoModificado("PrecioP") & ",PRECIOD=" & datoModificado("PrecioD") & ",PJEPARCIAL=" & datoModificado("PjeParcial") & ",CODIGOPAGO='" & datoModificado("CodigoPago") & "',FECENTDESDE=" & datoModificado("FECENTDESDE") & ", FECENTHASTA=" & datoModificado("FECENTHASTA") & ", PTORECEPCION=" & datoModificado("PtoRecepcion") & ", IDTRANSPORTE=" & datoModificado("IDTRANSPORTE") & ", FECFIJADESDE=" & getValor(datoModificado("FECFIJADESDE")) & ", FECFIJAHASTA=" & getValor(datoModificado("FECFIJAHASTA")) & ", CANTKILOSMIN=" & getValor(datoModificado("CANTKILOSMIN")) & ", CANTKILOSMAX=" & getValor(datoModificado("CANTKILOSMAX")) & ", OBSERVACIONES='" & datoModificado("OBSERVACIONES") & "', CODIGOSIO='" & datoModificado("CODIGOSIO") & "' WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha
	end if
	Call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
	grabarReconfirmacion = true
end function
'-------------------------------------------------------------------------------------------------
function grabarRegistroMail(strSubject,strDescription,strFrom,strTo)
dim mySql, myRs, myCn, strMensaje, auxRegistro
grabarRegistroMail = false
strMensaje = "De:" & strFrom
strMensaje = strMensaje & "<br>" & "Para:" & strTo 
strMensaje = strMensaje & "<br>" & "Asunto:" & strSubject 
strMensaje = strMensaje & "<br>" & "Descripcion:" & strDescription & "<hr>"
strMensaje = trim(replace(strMensaje,"'","''"))
	mySql = "SELECT * FROM TOEPFERDB.TBLCONTRATOSCONF WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha
	call GF_BD_AS400_2(myRs,myCn,"OPEN",mySql)
	if myRs.eof then
		mySql = "INSERT INTO TOEPFERDB.TBLCONTRATOSCONF(PRODUCTO,SUCURSAL,OPERACION,NUMERO,COSECHA, USUARIO, REGISTRO) VALUES(" & producto & "," & sucursal & "," & operacion & "," & numero & "," & cosecha & ",'" & session("Usuario") & "','" & strMensaje & "')"
	else
		if not isNull(myRs("REGISTRO")) then
			auxRegistro = " REGISTRO=CONCAT(REGISTRO,'" & strMensaje & "') "
		else
			auxRegistro = "REGISTRO='" & strMensaje & "'"	
		end if
		mySql = "UPDATE TOEPFERDB.TBLCONTRATOSCONF SET USUARIO='" & session("Usuario") & "', " & auxRegistro & " WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha
	end if
	'response.Write "<HR>CONFIRMACION <BR>" & mySql 
	call GF_BD_AS400_2(myRs,myCn,"EXEC",mySql)
	grabarRegistroMail = true
end function
'-------------------------------------------------------------------------------------------------
Function requiereCodigoSIO(pProducto)
    Dim listaProductos
    Dim myProducto
	
    listaProductos = "[4],[5],[8],[15],[19],[22],[23],[24],[26],[33],[50],[60]"
    
	myProducto = "[" & pProducto & "]"
	
    requiereCodigoSIO = false
    if (InStr(1, listaProductos, myProducto) > 0) then requiereCodigoSIO = true
End Function
'-------------------------------------------------------------------------------------------------
Function controlarParametrosConfirmacion ()
dim rsVendedor
controlarParametrosConfirmacion = false

    if (requiereCodigoSIO(producto)) then
        if (datoModificado("CODIGOSIO") = "") then addError "Codigo SIO Granos esta vacio"    
    end if        
    if (datoModificado("ctoCorredor") = "") then addError "Cto Corredor esta vacio"
    'Vendedor
    if datoModificado("CUITVendedor") = "" then addError "Es necesario indicar el CUIT"
    if Session("KCOrganizacion") <> KC_TOEPFER then
		if GF_CONTROL_CUIT(datoModificado("CUITVendedor")) then
			if not getProveedor(datoModificado("KCVEN"), rsVendedor) is nothing then
				'response.write rsVendedor("CUIT") & "-" & datoModificado("CUITVendedor")
				if datoModificado("CUITVendedor") <> cSTR(rsVendedor("CUIT")) then 
					hayReconfirmacion = true
				end if	
			else
				addError "No se encontro al vendedor"
			end if
		else
			addError "Formato de CUIT inválido"
		end if
        'Marca merc propia-Consignacion
	    'if (datoModificado("MercPropia") = "" and datoModificado("MercConsigna") = "") then addError "Es necesario indicar tipo de mercaderia (Propia, En Consig., No propia)"
    else
		if cint(datoModificado("CamionesPactados")) <> cint(rsOriginal("CamionesPactados")) then hayTablaFK = true	
    end if
    
    if (cdMoneda="") then addError "Debe seleccionar la moneda del contrato"
    if (cdMoneda <> getMonedaOperacion(operacion)) then 
		hayReconfirmacion = true
	end if	
    
    if (datoModificado("CODIGOSIO") <> rsOriginal("CODIGOSIO")) then hayTablaFM = true
	'Si la condicion de pago es J,K,R,T,Z hay que grabar tabla MER311F2
	if (instr("J,K,R,T,Z",datoModificado("CodigoPago"))) then hayTablaF2 = true

	'Marcar si al momento de guardar se debe modificar la tabla MER311FJ
	'Response.Write "MP(" & datoModificado("MercPropia") & "-" & rsOriginal("MercPropia") & "), MC(" & datoModificado("MercConsigna") & "-" & rsOriginal("MercConsigna") & "), CI(" & datoModificado("CondicionIVA") & "-" & rsOriginal("CondicionIVA") & ")"
    if (datoModificado("MercPropia")   <> getTrueFalse(rsOriginal("MercPropia")))	then hayTablaFJ = true
	if (datoModificado("MercConsigna") <> getTrueFalse(rsOriginal("MercConsigna"))) then hayTablaFJ = true
	if (datoModificado("CondicionIVA") <> rsOriginal("CondicionIVA"))				then hayTablaFJ = true	
	if (len(datoModificado("Observaciones")) > 1)									then hayReconfirmacion = true	

    if Int(datoModificado("KCVEN")) = 1 then addError "Es necesario confirmar el CUIT con el personal de la empresa"
    'Kilos-Cantidad contratada
    KilosDelNegocio = cdbl(datoModificado("Kilos")) - cdbl(rsOriginal("AmpAnu"))
	'Response.Write "Kilos(" & datoModificado("Kilos") & ")"
	'Response.Write "Ampl(" & cdbl(rsOriginal("AmpAnu")) & ")"
    'Response.Write "Netos(" & KilosDelNegocio & ")"
    if (KilosDelNegocio < 0) then addError "Cantidad de mercaderia contratada incorrecta. Consulte Ampliaciones/Anulaciones"
    if (datoModificado("Kilos") = 0) then addError "Cantidad de mercaderia contratada no puede ser 0"
    'Procedencia
    'Response.Write "(" & datoModificado("CPProcedencia") & ")"
    if (trim(datoModificado("CPProcedencia")) = "" or clng(datoModificado("CPProcedencia")) = 99 or clng(datoModificado("CPProcedencia")) = 0) then addError "Es necesario confirmar la procedencia"
    if (trim(datoModificado("Procedencia")) = "A CONFIRMAR") then addError "Es necesario confirmar la procedencia"
    'Porcentaje parcial
    if (CDbl(datoModificado("PjeParcial")) = 0) then addError "El Porcentaje Parcial no puede ser 0"
	if (CDbl(datoModificado("PjeParcial")) > 10000) then addError "El Porcentaje Parcial no puede ser superior al 100%"
	
    'Control Fechas Entrega
    if (cLng(datoModificado("FecEntDesde")) > cLng(datoModificado("FecEntHasta"))) then addError "Fecha de Entrega Desde mayor a Fecha de Entrega Hasta"
    'Dar dos dias de flexibilidad para la fecha de entrega
    if (abs(cLng(datoModificado("FecEntHasta"))) - CLng(rsOriginal("FecEntHasta")) > 2 )																		then hayReconfirmacion = true
    if (abs(cLng(datoModificado("FecEntDesde"))) - CLng(rsOriginal("FecEntDesde")) > 2 )																		then hayReconfirmacion = true
    if((CLng(datoModificado("PrecioP"))			<> CLng(rsOriginal("PrecioP"))) and (CLng(datoModificado("PrecioD")) <> CLng(rsOriginal("PrecioD"))))	then hayReconfirmacion = true
    if (CLng(datoModificado("Kilos"))			<> CLng(rsOriginal("Kilos")))																			then hayReconfirmacion = true
    if (CDbl(datoModificado("PjeParcial"))		<> CDbl(rsOriginal("PjeParcial")))																		then hayReconfirmacion = true
    if (CLng(datoModificado("PtoRecepcion"))	<> CLng(rsOriginal("PtoRecepcion")))																	then hayReconfirmacion = true
    if (CLng(datoModificado("IdTransporte"))	<> CLng(rsOriginal("IdTransporte")))																	then hayReconfirmacion = true
        
    'Controles en caso de existir las fijaciones
    if(len(datoModificado("FecFijaDesde")) = 8) then
	    if (cLng(datoModificado("CantKilosMin")) = 0) then addError "Kilos Mínimos debe ser mayor a cero"
        if (cLng(datoModificado("CantKilosMax")) = 0) then addError "Kilos Máximos debe ser mayor a cero"
		if (cLng(datoModificado("CantKilosMin")) > cLng(datoModificado("CantKilosMax"))) then addError "Kilos Mínimos no puede ser mayor a Kilos Máximos"
	    if (CLng(datoModificado("CantKilosMin")) <> CLng(rsOriginal("CantKilosMin")))	then hayReconfirmacion = true
	    if (CLng(datoModificado("CantKilosMax")) <> CLng(rsOriginal("CantKilosMax")))	then hayReconfirmacion = true

        if (cLng(datoModificado("FecFijaDesde")) > cLng(datoModificado("FecFijaHasta"))) then addError "Fecha de Fijación Desde mayor a Fecha de Fijación Hasta"
        'Dar dos dias de flexibilidad para la fecha de fijacion
        if (abs(cLng(datoModificado("FecFijaHasta")) - CLng(rsOriginal("FecFijaHasta"))) > 2 )then hayReconfirmacion = true
        if (abs(cLng(datoModificado("FecFijaDesde")) - CLng(rsOriginal("FecFijaDesde"))) > 2 )then hayReconfirmacion = true
        if(CLng(datoModificado("PrecioP")) <> 0) then addError "Contrato es 'A Fijar' - no lleva el Precio"
    else
        'El usuario dejo el precio en blanco en un contrato que no es a Fijar
        if cdbl(datoModificado("Precio")) = 0 then addError "El precio debe ser mayor a 0"
    end if
    if not (hayError) then controlarParametrosConfirmacion = true
end function
'-------------------------------------------------------------------------------------------------
Function getProveedor(p_KcProveedor, ByRef p_rsProveedor)
    Dim strWhere, strORKC, oConn, strSQL, strOrder
    strWhere = ""
    if (p_KcProveedor <> "") then Call mkWhere(strWhere, "IdEmpresa", cSTR(p_KcProveedor),"=",0)
    strSQL= "SELECT * FROM TOEPFERDB.VWEMPRESAS " & strWhere
	Call GF_BD_AS400_2(p_rsProveedor,oConn,"OPEN",strSQL)
	if (p_rsProveedor.eof) then
       getProveedor = nothing
    else
       getProveedor = p_rsProveedor
    end if
End Function
'-------------------------------------------------------------------------------------------------
Function mostrarErrores()
    response.write "<table border=0 align='center' width='95%'><tr><td>"
    For i = 0 to dicErr.count - 1
        Response.Write "<li><font color=#ff0000>" & dicErr(i) & "</font></li>"
    next
    response.write "</td></tr></table>"
end function
'-------------------------------------------------------------------------------------------------
sub addError(pErrorMsg)
	call dicErr.Add (myIndexDic,pErrorMsg)
	myIndexDic = myIndexDic + 1
end sub
'-------------------------------------------------------------------------------------------------
function hayError()
	hayError = true
	if dicErr.Count = 0 then hayError = false
end function
'-------------------------------------------------------------------------------------------------
function envioMails()
'Esto es para que permita seguir y se muestre el mensaje de info.
on error resume next
    if estado = 1 then
        enviarMailConfirmacion()
    elseif estado = 2 then
        enviarMailReconfirmacion()
    elseif estado = 3 then
        enviarMailConfirmacionExterna()
    end if
end function
'-------------------------------------------------------------------------------------------------
Function enviarMailConfirmacion()
'Funcion, envia el mail de pendiente de confirmacion a usuario interno
Dim strFrom,strTo,strSubject,strDescription
Dim strNombre,strEmpresa,strMail,strTexto,strTelefono
Dim errMsg
Dim vectorMails(10), cantMails, iMail
dim auxCantMails
strTo=""
    
	'Mail para Toepfer
	cantMails = 0
    iMail = 0
    cantMails = obtenerMailConfirmaciones(KC_TOEPFER, vectorMails)
    while iMail < cantMails
        strTo = strTo & vectorMails(iMail) & "; "
        iMail = iMail + 1
    wend
    
	'strTo = ADDRESSEE_CONFIRM
    strFrom = SENDER_MERCADERIAS
 	strSubject="Confirmacion via Web. Contrato "
    strSubject=strSubject & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
	strDescription="Se requiere la reconfirmacion del contrato numero "
	
	
	strDescription = strDescription & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
	strDescription = strDescription & chr(13) & chr(10) & chr(13) & chr(10) & "DATOS MODIFICADOS " & chr(13) & chr(10) & chr(13) & chr(10)
	strDescription = strDescription & getDatosModificados()
	strDescription = strDescription & getResponsableMailStandard()
	
    call grabarRegistroMail(strSubject,strDescription,strFrom,strTo)
    call GP_ENVIAR_MAIL(strSubject,strDescription,strFrom,strTo)
end function 
'-----------------------------------------------------------------------------------------
function getResponsableMailStandard()
dim str, strAux
	str = str & chr(13) & chr(10) & chr(13) & chr(10) & chr(13) & chr(10) & "Informacion de control"
	str = str & chr(13) & chr(10) & "-------------------------------------------------------------------------"
	str = str & chr(13) & chr(10) & "Momento de Operacion.:" & day(now()) & "/" & Month(now()) & "/" & year(now()) & " " & hour(now()) & ":" & minute(now()) & ":" & Second(now()) & " "
	str = str & chr(13) & chr(10) & "Corredor.............:" & trim(GetDsEnterprise2(rsOriginal("KCCOR"))) & " "
	str = str & chr(13) & chr(10) & "Vendedor.............:" & trim(GetDsEnterprise2(rsOriginal("KCVEN"))) & " "
   	str = str & chr(13) & chr(10) & "Usuario..............:" & session("Usuario") & " (" & GetDsEnterprise2(Session("KCOrganizacion")) & ")"
    getResponsableMailStandard = str
end function
'-------------------------------------------------------------------------------------------------
Function enviarMailConfirmacionExterna()
'Funcion, envia el mail de pendiente de confirmacion a usuario interno
Dim strFrom,strTo,strSubject,strDescription
Dim strNombre,strEmpresa,strMail,strTexto,strTelefono
Dim errMsg
Dim vectorMails(10), cantMails, iMail
	strTo=""
	'Mail para Corredor/Vendedor
    if len(rsOriginal("KCCOR")) > 0 then
        auxCantMails = obtenerMailConfirmaciones(rsOriginal("KCCOR"), vectorMails)
    else
        auxCantMails = obtenerMailConfirmaciones(rsOriginal("KCVEN"), vectorMails)
    end if
    iMail = 0
    cantMails = auxCantMails
    while iMail < auxCantMails
        strTo = strTo & vectorMails(iMail) & "; "
        iMail = iMail + 1
    wend
    
	'Mail para Toepfer
	auxCantMails = 0
    iMail = 0
    auxCantMails = obtenerMailConfirmaciones(KC_TOEPFER, vectorMails)
    while iMail < auxCantMails
        strTo = strTo & vectorMails(iMail) & "; "
        iMail = iMail + 1
    wend 
    cantMails = auxCantMails + cantMails	
	
	'strTo = ADDRESSEE_CONFIRM
    strFrom = SENDER_MERCADERIAS
 	strSubject = "Confirmacion via Web. Contrato "
    strSubject = strSubject & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
	strSubject = strSubject & " Confirmado OK."
	strDescription="El contrato "
	strDescription = strDescription & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
	strDescription = strDescription & " ha sido confirmado." & chr(13) & chr(10) & chr(13) & chr(10) & "DATOS MODIFICADOS " & chr(13) & chr(10) & chr(13) & chr(10)
	strDescription = strDescription & getDatosModificados()
	strDescription = strDescription & getResponsableMailStandard()
	call grabarRegistroMail(strSubject,strDescription,strFrom,strTo)
    call GP_ENVIAR_MAIL(strSubject,strDescription,strFrom,strTo)
end function
'-------------------------------------------------------------------------------------------------
Function enviarMailReconfirmacion()
'Envia el mail a usuario externo, si es que no necesita reconfirmacion por parte de usuraio interno de toepfer
Dim strFrom,strTo,strSubject,strDescription
Dim strNombre,strEmpresa,strMail,strTexto,strTelefono
Dim errMsg, auxCantMails
Dim vectorMails(10), cantMails, iMail

	strTo=""
	'Mail para Corredor/Vendedor
    if len(rsOriginal("KCCOR")) > 0 then
        auxCantMails = obtenerMailConfirmaciones(rsOriginal("KCCOR"), vectorMails)
    else
        auxCantMails = obtenerMailConfirmaciones(rsOriginal("KCVEN"), vectorMails)
    end if
    iMail = 0
    cantMails = auxCantMails
    while iMail < auxCantMails
        strTo = strTo & vectorMails(iMail) & "; "
        iMail = iMail + 1
    wend
    
	'Mail para Toepfer
	auxCantMails = 0
    iMail = 0
    auxCantMails = obtenerMailConfirmaciones(KC_TOEPFER, vectorMails)
    while iMail < auxCantMails
        strTo = strTo & vectorMails(iMail) & "; "
        iMail = iMail + 1
    wend 
    cantMails = auxCantMails + cantMails
    
    if cantMails <> 0 then
        'strTo = ADDRESSEE_CONFIRM
        strFrom = SENDER_MERCADERIAS
        strSubject="Confirmacion via Web. Contrato "
        strSubject=strSubject & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
        strSubject=strSubject & ". Confirmado OK."
        strDescription="Se han confirmado los datos del contrato numero "
        strDescription = strDescription & GF_EDIT_CONTRATO(producto, sucursal, operacion, numero, cosecha)
        if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then strDescription = strDescription & chr(13) & chr(10) & chr(13) & chr(10) & "DATOS CONFIRMADOS " & chr(13) & chr(10) & chr(13) & chr(10) & getDatosModificados()
        strDescription=strDescription & chr(13) & chr(10) & vbcrlf & "El boleto de compra-venta ya se encuentra disponible en nuestro sitio web. Si desea recibir este y otros boleto por mail, ingrese a nuestro sitio y configure sus direcciones de correo para que podamos enviarselo. " & vbcrlf & "Muchas Gracias"
        strDescription = strDescription & vbcrlf & getResponsableMailStandard()
        call grabarRegistroMail(strSubject,strDescription,strFrom,strTo)
        call GP_ENVIAR_MAIL(strSubject,strDescription,strFrom,strTo)
    end if
end function
'-------------------------------------------------------------------------------------------------
function getDatosModificados()
'arma el string con los datos modificados por usuario interno para armar el mail para usuario externo en el caso de la confirmacion definitiva
dim modif, auxDs1, auxDs2, cdMonedaOriginal
modif = ""
    if (trim(datoModificado("CtoCorredor")) <> trim(rsOriginal("CtoCorredor")))     then        modif = modif & "Cto.Corredor...............:"  & chr(9) & trim(datoModificado("CtoCorredor")) & " (" & trim(rsOriginal("CtoCorredor")) & ")"  & chr(13) & chr(10)
    if (datoModificado("CODIGOSIO")   <> rsOriginal("CODIGOSIO")) 		            then	    modif = modif & "Codigo SIO Granos..........:"	& chr(9)  & datoModificado("CODIGOSIO") & " (Valor Anterior: " & rsOriginal("CODIGOSIO") & ")" & chr(13) & chr(10)    
    'if (CSTR(datoModificado("KCVEN")) <> CSTR(rsOriginal("KCVEN")))				then		modif = modif & "Cto.Corredor...............:"	& chr(9) & trim(datoModificado("CtoCorredor")) & " (Valor Anterior: " & trim(rsOriginal("CtoCorredor")) & ")"  & chr(13) & chr(10)
    if (datoModificado("CondicionIVA")		<> rsOriginal("CondicionIVA"))			then		modif = modif & "Condicion IVA..............:"	& chr(9) & datoModificado("CondicionIVA") & " (Valor Anterior: " & rsOriginal("CondicionIVA") & ")" & chr(13) & chr(10)
    if (datoModificado("CodigoPago")		<> rsOriginal("CodigoPago"))			then		modif = modif & "Condicion Pago.............:"	& chr(9) & getDsPago(datoModificado("CodigoPago")) & " (Valor Anterior: " & getDsPago(rsOriginal("CodigoPago")) & ")" & chr(13) & chr(10)
    
    if Session("KCOrganizacion") <> KC_TOEPFER then
		if not getProveedor(datoModificado("KCVEN"), rsVendedor) is nothing then
			if datoModificado("CUITVendedor") <> cSTR(rsVendedor("CUIT"))			then		modif = modif & "CUIT Vendedor..............:"	& chr(9) & GF_STR2CUIT(datoModificado("CUITVendedor")) & " (Valor Anterior: " & GF_STR2CUIT(cSTR(rsVendedor("CUIT"))) & ")" & chr(13) & chr(10) 			
		end if
    end if    
    
	cdMonedaOriginal = getMonedaOperacion(rsOriginal("Operacion"))
    if (cdMoneda <> cdMonedaOriginal) then		
			simboloMonedaOriginal = getSimboloMoneda(cdMonedaOriginal)
    		modif = modif & "Moneda del Contrato........:"	& chr(9) & getSimboloMoneda(cdMoneda) & " (Valor Anterior: " & simboloMonedaOriginal & ") " & chr(13) & chr(10)
	end if
    
    if (Clng(datoModificado("Kilos"))		<> Clng(rsOriginal("Kilos")))			then		modif = modif & "Cantidad Contratada........:"	& chr(9) & datoModificado("Kilos") & " (Valor Anterior: " & rsOriginal("Kilos") & ") " & chr(13) & chr(10)
    if ((clng(datoModificado("CPProcedencia"))	<> clng(rsOriginal("CPProcedencia"))) or (clng(datoModificado("CAProcedencia"))	<> clng(rsOriginal("CAProcedencia"))))			then		modif = modif & "Procedencia:...............:"	& chr(9) & trim(datoModificado("Procedencia")) & " (Valor Anterior: " & replace(trim(rsOriginal("Procedencia")),"'","''") & ")" & chr(13) & chr(10)
    'Condicion    
    if ((datoModificado("MercPropia") <> rsOriginal("MercPropia")) or (datoModificado("MercConsigna") <> rsOriginal("MercConsigna"))) then    
        auxDs1 = getTextoCondicion(datoModificado("MercPropia"), datoModificado("MercConsigna"))
        auxDs2 = getTextoCondicion(rsOriginal("MercPropia"), rsOriginal("MercConsigna"))
        modif = modif & "Condicion..................:"	& chr(9) & auxDs1 & " (Valor Anterior: " & auxDs2 & ")" & chr(13) & chr(10)
    end if        
    
    select case CInt(operacion)
        case 0,1,2,3,5: 'es en pesos
			if (CLng(datoModificado("PrecioP")) <> CLng(rsOriginal("PrecioP")))		then		modif = modif & "Precio.....................:"	& chr(9) & datoModificado("PrecioP") & " (Valor Anterior: " & rsOriginal("PrecioP") & ")" & chr(13) & chr(10)
        case 6,9,10,11,12: 'es en dolares
			if (CLng(datoModificado("PrecioD")) <> CLng(rsOriginal("PrecioD")))		then		modif = modif & "Precio.....................:"	& chr(9) & datoModificado("PrecioD") & " (Valor Anterior: " & rsOriginal("PrecioD") & ")" & chr(13) & chr(10)
	end select
    if (CDbl(datoModificado("PjeParcial")) <> CDbl(rsOriginal("PjeParcial")))		then		modif = modif & "Parcial (%)................:"	& chr(9) & GF_EDIT_DECIMALS(cdbl(datoModificado("PjeParcial"))*100,2) & " (Valor Anterior: " & GF_EDIT_DECIMALS(cdbl(rsOriginal("PjeParcial"))*100,2) & ")" & chr(13) & chr(10)
    if (CLng(datoModificado("FecEntDesde")) <> CLng(rsOriginal("FecEntDesde")))		then		modif = modif & "Fecha Entrega Desde........:"	& chr(9) & GF_FN2DTE(datoModificado("FecEntDesde")) & " (Valor Anterior: " & GF_FN2DTE(rsOriginal("FecEntDesde")) & ")" & chr(13) & chr(10)
    if (CLng(datoModificado("FecEntHasta")) <> CLng(rsOriginal("FecEntHasta")))		then		modif = modif & "Fecha Entrega Hasta........:"	& chr(9) & GF_FN2DTE(datoModificado("FecEntHasta")) & " (Valor Anterior: " & GF_FN2DTE(rsOriginal("FecEntHasta")) & ")" & chr(13) & chr(10)
    if (Cint(datoModificado("PtoRecepcion")) <> Cint(rsOriginal("PtoRecepcion")))	then
        auxDs1 = getDescripcionDestino(datoModificado("PtoRecepcion"))
        auxDs2 = getDescripcionDestino(rsOriginal("PtoRecepcion"))
		modif = modif & "Puerto Recepcion...........:" & chr(9) & trim(auxDs1) & " (Valor Anterior: " & trim(auxDs2) & ") " & chr(13) & chr(10)
    end if
    if (Cint(datoModificado("IdTransporte")) <> Cint(rsOriginal("IdTransporte"))) then
		auxDs1 = getDescripciontransporte(datoModificado("IdTransporte"))
		auxDs2 = getDescripciontransporte(rsOriginal("IdTransporte"))
		modif = modif & "Transporte.................:" & chr(9) & trim(auxDs1) & " (Valor Anterior: " & trim(auxDs2) & ") " & chr(13) & chr(10)
    end if
    if CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER) and datoModificado("CodigoPago") = "X" then
		if cint(datoModificado("CamionesPactados")) <> cint(rsOriginal("CamionesPactados")) then	modif = modif & "Camiones Pactados..........:"	& chr(9)  & datoModificado("CamionesPactados") & " (Valor Anterior: " & rsOriginal("CamionesPactados") & ")" & chr(13) & chr(10)
    end if    
    if(len(datoModificado("FecFijaDesde")) = 8) then
        'en caso de que es contrato a Fijar																	
        if (Clng(datoModificado("CantKilosMin")) <> Clng(rsOriginal("CantKilosMin"))) then	modif = modif & "Kilos Min a Fijar..........:"	& chr(9)  & datoModificado("CantKilosMin") & " (Valor Anterior: " & rsOriginal("CantKilosMin") & ")" & chr(13) & chr(10)
        if (Clng(datoModificado("CantKilosMax")) <> Clng(rsOriginal("CantKilosMax"))) then	modif = modif & "Kilos Max a Fijar..........:"	& chr(9)  & datoModificado("CantKilosMax") & " (Valor Anterior: " & rsOriginal("CantKilosMax") & ")" & chr(13) & chr(10)
        if (CLng(datoModificado("FecFijaDesde")) <> CLng(rsOriginal("FecFijaDesde"))) then	modif = modif & "Fecha Fijacion Desde.......:"	& chr(9)  & GF_FN2DTE(datoModificado("FecFijaDesde")) & " (Valor Anterior: " & GF_FN2DTE(rsOriginal("FecFijaDesde")) & ")" & chr(13) & chr(10)
        if (CLng(datoModificado("FecFijaHasta")) <> CLng(rsOriginal("FecFijaHasta"))) then	modif = modif & "Fecha Fijacion Hasta.......:"	& chr(9)  & GF_FN2DTE(datoModificado("FecFijaHasta")) & " (Valor Anterior: " & GF_FN2DTE(rsOriginal("FecFijaHasta")) & ")" & chr(13) & chr(10)
    end if
	if (len(datoModificado("Observaciones"))      > 0)								  then 	modif = modif & chr(13) & "Observaciones"	& chr(13) & datoModificado("Observaciones") & chr(13) & chr(10)    
    getDatosModificados = modif
end function
'--------------------------------------------------------------------------------------------
Function getObservacionesBoleto()
'Trae la observacion de boleto para contrato a Fijar, lo devuelve como un string
Dim strWhere, strORKC, oConn, strSQL, strOrder, rsObservaciones
dim observacion
    strWhere = ""
    observacion = ""
    if (producto <> "")		then Call mkWhere(strWhere, "CPRORP", producto,"=",1)
    if (sucursal <> "")		then Call mkWhere(strWhere, "CSUCRP", sucursal,"=",1)
    if (operacion <> "")	then Call mkWhere(strWhere, "COPERP", operacion,"=",1)
    if (numero <> "")		then Call mkWhere(strWhere, "NCTORP", numero,"=",1)
    if (cosecha <> "")		then Call mkWhere(strWhere, "ACOSRP", cosecha,"=",1)
    strSQL="Select * from MERFL.MER311FP " & strWhere
    strSql = strSql & " order by NUMRRP"
    Call GF_BD_AS400_2(rsObservaciones,oConn,"OPEN",strSQL)
    while not(rsObservaciones.eof)
		if not isNull(rsObservaciones("OBSERP")) then
			observacion = observacion & " " & rsObservaciones("OBSERP")	
		end if	
        rsObservaciones.movenext
    wend
    getObservacionesBoleto = observacion
end function
'--------------------------------------------------------------------------------------
function getDsPago(pCodigo)
dim strSQL, rsFormasPago, oConn, rtrn 
rtrn = ""
strSQL="Select * from MERFL.MER2I1F1 where CODIFP='" & pCodigo & "'"
Call GF_BD_AS400_2(rsFormasPago,oConn,"OPEN",strSQL)
if not rsFormasPago.eof then rtrn = GF_Traducir(TRIM(rsFormasPago("DESCFP")))
getDsPago = rtrn
end function
'--------------------------------------------------------------------------------------
Function determineCondicion(dataMP, dataVC, ByRef pp, ByRef vc, ByRef np)
    pp=false
    vc=false    
    np=true
    if (dataMP = "V") then    
        np=false
        pp=true
    else if (dataVC = "V") then  
            np=false
            vc=true
         end if
    end if 
End function
'--------------------------------------------------------------------------------------
Function getTextoCondicion(dataMP, dataVC)
    dim ret, pp, vc, np
    
    Call determineCondicion(dataMP, dataVC, pp, vc, np)
    if (pp) then ret = "ES MERCADERIA DE PROPIA PRODUCCION"
    if (vc) then ret = "ES VENTA EN CONSIGNACION POR CTA y ORDEN DE V/COMITENTES"
    if (np) then ret = "NO ES DE SU PROPIA PRODUCCION"
    
    getTextoCondicion = ret
    
End function
%>
<HTML>
<HEAD>
	<meta charset="utf-8">
    <TITLE>TOEPFER INTERNATIONAL - Contratos</TITLE>
    <link rel="stylesheet" type="text/css" href="css/ActisaIntra-1.css">            
    <link rel="stylesheet" type="text/css" media="all" href="CSS/calendar-win2k-2.css" title="win2k-2" />
    <link rel="stylesheet" type="text/css" href="css/Toolbar.css">
    <link rel="stylesheet" type="text/css" href="css/iwin.css">
    
	

    <script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>      
    <script type="text/javascript" src="scripts/controles.js"></script>
	<script type="text/javascript" src="scripts/formato.js"></script>
	<script type="text/javascript" src="scripts/iwin.js"></script>	
    <script type="text/javascript" src="scripts/channel.js"></script>
    
    <!-- Archivos necesarios para Autocomplete -->
    <link rel="stylesheet" type="text/css" href="css/autocomplete.css">
    <script type="text/javascript" src="scripts/jqueryObject.js"></script>
	<script type="text/javascript" src="scripts/ui/jquery.ui.core.js"></script>
	<script type="text/javascript" src="scripts/ui/jquery.ui.widget.js"></script>
	<script type="text/javascript" src="scripts/ui/jquery.ui.position.js"></script>
	<script type="text/javascript" src="scripts/jQueryAutocomplete.js"></script>
	<!-- Fin archivos necesarios para Autocomplete -->
	
    <script type="text/javascript">
	$(function() {
		$( "#procedencia" ).autocomplete({
			minLength: 2,
			source: "contratosGetLocalidadesAJAX.asp",
			focus: function( event, ui ) {
				$( "#procedencia" ).val(ui.item.label);
				return false;
			},
			select: function( event, ui ) {
				$( "#procedencia" ).val(ui.item.label);
				$( "#CAProcedencia" ).val( ui.item.value );
				$( "#CPProcedencia" ).val( ui.item.id );
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item){
					$( "#procedencia" ).val("");
					$( "#CAProcedencia" ).val(0);
					$( "#CPProcedencia" ).val(0);
				}
				return true;
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.label + "<br><font style='font-size:10;'>" + item.desc + "</font></a>" )
				.appendTo( ul );
		};
	});

  	      
    var IE5 = ((navigator.userAgent.toLowerCase().indexOf('msie')!= -1) && (!window.opera));
    var GKO = (navigator.userAgent.toLowerCase().indexOf('gecko')!= -1);
    var ch = new channel();
	function submitInfo(acc) {		
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}

	function bodyOnLoad() {	
		tb = new Toolbar('toolbar', 8, 'images/contratos/');
		idBtnControl = tb.addButton("control-16x16.png","Confirmar", "submitInfo('<% =ACCION_GRABAR %>',1)");								
		tb.draw();
	}

	function CerrarCal(cal) {
		cal.hide();
	}
	
	function MostrarCalendario(p_ImgId, funcSel) {
		var dte= new Date();		    	    
		var elem= document.getElementById(p_ImgId);
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}		
	function SeleccionarFechaPago(cal, date) {
		var str= new String(date);		
		document.getElementById("FechaPagoF").value = str;
	    document.getElementById("FechaPago").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();	
	}	
	function SeleccionarFechaEntregaDesde(cal, date) {
		var str= new String(date);		
		document.getElementById("FecEntDesdeF").value = str;
	    document.getElementById("FecEntDesde").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();	
	}		
	function SeleccionarFechaEntregaHasta(cal, date) {
		var str= new String(date);		
	    document.getElementById("FecEntHastaF").value = str;
		document.getElementById("FecEntHasta").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);	    
		if (cal) cal.hide();	
	}				
	function SeleccionarFechaFijaDesde(cal, date) {
		var str= new String(date);		
		document.getElementById("FecFijaDesdeF").value = str;
	    document.getElementById("FecFijaDesde").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();	
	}
	function SeleccionarFechaFijaHasta(cal, date) {
		var str= new String(date);		
		document.getElementById("FecFijaHastaF").value = str;
	    document.getElementById("FecFijaHasta").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
		if (cal) cal.hide();	
	}	
	function info() {	
		opener.document.form1.submit();
		popUp = new PopUpWindow('popUp', 'contratosInfoConfirmacion.asp?producto=<%=producto%>&numero=<%=numero%>&estado=<%=estado%>&fechaConc=<%=0%>&sucursal=<%=sucursal%>&operacion=<%=operacion%>&cosecha=<%=cosecha%>&unitDest=<%=unitDest%>&corredor=<%=rsOriginal("KCCOR")%>&FechaConcertacion=<%=rsOriginal("FECHACONC")%>', '550', '190', 'Modificacion Realizada');
	}   
	function closePopUp(){
		self.close();
	}
	function configurarMail(){
		//self.close();
		document.location.href="datosPersonales.asp";
	}
	function HabilitarCamionesPactadosDiv(obj){
		if (obj.value == "X"){
			document.getElementById("CamionesPactadosTR").style.visibility = "visible";
			document.getElementById("CamionesPactadosTR").style.position = "relative";
		}
		else{
			document.getElementById("CamionesPactadosTR").style.visibility = "hidden";
			document.getElementById("CamionesPactadosTR").style.position = "absolute";
			document.getElementById("CamionesPactados").value = 0;
		}
		
	}
</script>
</HEAD>
<body onLoad="<%=ppIni%>;<%=ppFin%>">	
	
<div id="capaPrincipal" style="z-index:1; height: 100%; width: 100%;">
<FORM NAME="frmSel" ID="frmSel" METHOD="POST" ACTION="contratosDetalleConfirma.asp">
      
      <div id="toolbar"></div>
      <!-- Errores -->
		<%
		if dicErr.Count => 0 then 
			call mostrarErrores()
			Response.Write "<br>"
		end if
		%>
		
      <!-- INICIO CONTRATO -->      
      <% if (not rsOriginal.eof) then %>
      <TABLE WIDTH="95%" ALIGN="CENTER" CELLSPACING="0" CELLPADDING="0" BORDER="0">            
             <TR>
                 <TD WIDTH="8"><img src="images/marco_r1_c1.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_r1_c3.gif"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD COLSPAN="3">
                     <TABLE WIDTH="100%">
                            <%
                            if (flagGrabado) then
                                response.write "<tr><td colspan=2 class='TDNOTICE'>" & GF_TRADUCIR("LOS DATOS DEL CONTRATO HAN SIDO GUARDADOS") & "</td></tr>"
                            else
								'Si submitio datos, el estado del contrato puede ser distinto al original, no se muestra mensaje.
								if (not isFormSubmit()) then
									if (rsModificado.eof) then %>
								       <TR><TD COLSPAN="2" class="TDERROR"><% =GF_TRADUCIR("CONTRATO SIN CONFIRMAR")%></TD><TR>
								<%	else	%>
								       <TR><TD COLSPAN="2" class="TDNOHAY"><% =GF_TRADUCIR("CONTRATO PENDIENTE DE CONFIRMACION")%></TD></TR>
							<%		end if
								end if
							end if%>
                            <TR>
                                <TD colspan ="2"><B><% =GF_TRADUCIR("Fecha de Concertacion") %>:</B> <% =GF_FN2DTE(rsOriginal("FechaConc")) %></TD>
                            </TR>
                            <TR>
                                <TD WIDTH="50%"><B><% =GF_TRADUCIR("Cto Corredor") %>:</B> <input size="20" type="text" class="<%=resaltar(trim(datoModificado("CtoCorredor")), trim(rsOriginal("CtoCorredor")))%>" name="ctoCorredor" id="ctoCorredor" value="<% =trim(datoModificado("ctoCorredor")) %>">&nbsp;<% response.write datoOriginal(trim(rsOriginal("CtoCorredor")))%></TD>
                                <TD><B><% =GF_TRADUCIR("Operacion") %>: </B><% =GF_Traducir(getDescripcionOperacion(rsOriginal("Operacion"))) %></TD>
                            </TR>
                            <% 'COLZA y CEBADA CERVECERA QUEDARON FUERA DE LA NORMATIVA.
                                if ((producto <> 9) and (producto <> 17)) then %>
                            <TR>
                                <TD WIDTH="50%"><B><% =GF_TRADUCIR("Codigo SIO Granos") %>:</B> <input size="20" maxlength="20" type="text" class="<%=resaltar(trim(datoModificado("CODIGOSIO")), trim(rsOriginal("CODIGOSIO")))%>" name="CODIGOSIO" id="CODIGOSIO" value="<% =trim(datoModificado("CODIGOSIO")) %>"></TD>
                            </TR>
                            <% end if %>
                     </TABLE>
                 </TD>
                 <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
             </TR>
             <TR>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD COLSPAN="3">
                     <TABLE WIDTH="100%">
                           <TR>
                                <TD><B><% =GF_TRADUCIR("Partes Involucradas") %></B></TD>
                                <td colspan="3">&nbsp;</td>
                            </TR>
                            <TR>
                                <TD ALIGN="RIGHT" width="15%"><% =GF_TRADUCIR("Corredor") %>:</TD>
                                <TD colspan="2" width="85%"><% =GetDSEnterprise2(rsOriginal("KCCOR")) %></TD>
                                <input type="hidden" name="KCCOR" id="KCCOR" value="<%=rsOriginal("KCCOR")%>">
                            </TR>
                             <TR>
                                <TD ALIGN="RIGHT" width="15%"><% =GF_TRADUCIR("CUIT Vendedor") %>:</TD>                                
                                <%
                                'Response.Write datoModificado("CUITVendedor")
                                if (session("KCOrganizacion") <> KC_TOEPFER) then
									if len(myValueCUITVendedor)>0 then
										dsVendedor = GetDsEnterprise3(myValueCUITVendedor)
									else
										if datoModificado("CUITVendedor") <> "" then 
											dsVendedor = GetDsEnterprise3(cDbl(datoModificado("CUITVendedor")))	
										end if										
									end if	                                
									%>
									<td width="10%">
	                                    <input type="text" name="CUITVendedor" id="CUITVendedor" value="<%=myValueCUITVendedor%>" size="12" maxlength="11">
                                    </TD>
                                    <TD><%=dsVendedor%>
										<input class="<%=resaltar(trim(datoModificado("CUITVendedor")),trim(rsOriginal("CUITVendedor")))%>" type="hidden" name="KCVEN" id="KCVEN" value="<%=rsOriginal("KCVEN")%>">
			                        </TD>
									<%
                                else
									if myValueCUITVendedor = "" then myValueCUITVendedor = datoModificado("CUITVendedor")
                                	%>
                                	<td width="10%">
										<input onchange="submitInfo('NOT');" class="<%=resaltar(Cdbl(datoModificado("CUITVendedor")),Cdbl(rsOriginal("CUITVendedor")))%>" type="text" name="CUITVendedor" id="CUITVendedor" value="<%=myValueCUITVendedor%>" size="12" maxlength="11">
									</TD>
									<TD>
									<%
									'dsVendedor = GetDsEnterprise3(datoModificado("CUITVendedor"))
									if isnumeric(TRIM(myValueCUITVendedor)) then
										strSQL = ""
										strSQL = strSQL & " SELECT * "
										strSQL = strSQL & " FROM TOEPFERDB.VWEMPRESAS "
										strSQL = strSQL & " WHERE CUIT = '" & myValueCUITVendedor & "'"
										CALL GF_BD_AS400_2 (rs,con, "OPEN",strSQL)
										%>
										<select name="KCVEN" id="KCVEN">
											<% 
											while (not rs.eof)
												mySelected = ""
												if CDbl(rsOriginal("KCVEN")) = CDbl(rs("IDEMPRESA")) then mySelected = "SELECTED"
												%>
												<option <%=mySelected%> value="<% =rs("IDEMPRESA") %>"><% =GF_TRADUCIR(rs("DSEMPRESA")) %>
												<%
												rs.MoveNext
											wend
											%>
										</select>										
										<%										
										while not rs.eof
												rtrn = trim(rs("DSEmpresa"))
											rs.MoveNext 
										wend	
									else
										rtrn = "-"
									end if
									%>
									</td>
									<%
                                end if
                                %>
                                

                            </TR>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_t.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
             </TR>
           
             <TR>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD WIDTH="50%" VALIGN="TOP">
                     <!-- MARCO DE MERCADERIAS -->
                     <TABLE WIDTH="100%">
                            <TR>
                                <TD COLSPAN="4"><B><% =GF_TRADUCIR("Mercaderia")%></B></TD>
                            </TR>
                            <tr>
                                <TD WIDTH="5%"></TD>                                
                                <TD ALIGN="RIGHT" WIDTH="35%"><% =GF_TRADUCIR("Condicion")%>:</TD>
                                <TD COLSPAN="2">                                                                        
                                    <%  
                                        Call determineCondicion(datoModificado("MercPropia"), datoModificado("MercConsigna"), flagPP, flagVC, flagNP)
                                        if (flagPP) then chkPP="checked"
                                        if (flagVC) then chkVC="checked"
                                        if (flagNP) then chkNP="checked"
                                        'Se analizan los datos originales para mostrarle al usuario interno cual era el estado original del dato.
                                        Call determineCondicion(rsOriginal("MercPropia"), rsOriginal("MercConsigna"), flagPP, flagVC, flagNP)
                                    %>                                
                                    <input type="radio" name="TipoMercaderia" value="<% =TIPO_PROPIA_PRODUCCION %>" <% =chkPP %>/>      <% =datoOriginal(flagPP) %> &nbsp;  <% =GF_TRADUCIR("ES MERCADERIA DE PROPIA PRODUCCION") %> <br />
                                    <input type="radio" name="TipoMercaderia" value="<% =TIPO_CONSIGNACION %>" <% =chkVC %> />          <% =datoOriginal(flagVC) %> &nbsp;  <% =GF_TRADUCIR("ES VENTA EN CONSIGNACI&Oacute;N POR CTA y ORDEN DE V/COMITENTES") %> <br />                                    
                                    <input type="radio" name="TipoMercaderia" value="<% =TIPO_NO_PROPIA_PRODUCCION %>" <% =chkNP %>/>   <% =datoOriginal(flagNP) %> &nbsp;  <% =GF_TRADUCIR("NO ES DE SU PROPIA PRODUCCION") %>                                                               
                                </TD>
                            </tr>
                            <TR>
                                <TD WIDTH="5%"></TD>
                                <TD ALIGN="RIGHT" WIDTH="35%"><% =GF_TRADUCIR("Producto")%>:</TD>
                                <TD colspan="2"><% = getDescripcionProducto(rsOriginal("Producto")) %></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Humedad(V/F)")%>:</TD>
                                <td colspan="2"><%=rsOriginal("MercHumedad")%></td>
                            </TR>
                            <TR>
                                <td></td>
                                <td ALIGN="RIGHT"><% =GF_TRADUCIR("Cond.IVA")%>:</td>
                                <td colspan="2">
                                    <input style="cursor:pointer" type="radio" name="CondicionIVA" id="CondicionIVA" value="<% =COND_IVA_C %>" <% if(datoModificado("CondicionIVA") = COND_IVA_C) then Response.write "CHECKED" %>/>&nbsp;<% =GF_TRADUCIR("P.Canje")%>&nbsp;
                                    <input style="cursor:pointer" type="radio" name="CondicionIVA" id="CondicionIVA" value="<% =COND_IVA_X %>" <% if(datoModificado("CondicionIVA") = COND_IVA_X) then Response.write "CHECKED" %>/>&nbsp;<% =GF_TRADUCIR("Normal")%>
                                </td>
                            </TR>
                            <TR>
                                <td></td>
                                <td ALIGN="RIGHT"><% =GF_TRADUCIR("Cosecha")%>:</td>
                                <td>                                
									<%=GF_nDigits((CInt(rsOriginal("Cosecha")) - 1),2) & "/" & GF_nDigits(CInt(rsOriginal("Cosecha")),2)%>
                                </td>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Contratado")%>:</TD>
                                <TD>
									<input type="text" class="<%=resaltar(CLng(datoModificado("Kilos")),Clng(rsOriginal("Kilos")))%>" size=10 maxlength=15 name="kilos" id="kilos" value="<%response.write datoModificado("Kilos") %>"  onKeyPress=" return controlIngreso(this, event, 'E');">&nbsp;Kg&nbsp;<font color="red"><b>*</b></font>&nbsp;<%= datoOriginal(rsOriginal("Kilos"))%>
                                </TD>
                                <TD></TD>
                            </TR>
                            <% 
                            if datoModificado("CodigoPago") = "X" then myStyleDiv = "style='visibility:visible;position:relative;'" 
							myPactadosOriginal = rsOriginal("CamionesPactados")
							if isNull(myPactadosOriginal) then myPactadosOriginal = 0
							myPactadosModificado = datoModificado("CamionesPactados")
							if isNull(myPactadosModificado) then myPactadosModificado = 0
							%>
                            <TR id="CamionesPactadosTR" <%=myStyleDiv%>>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Camiones Pactados")%>:</TD>
                                <TD>
									
									<%if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then %>
										<input type="text" class="<%=resaltar(CLng(myPactadosModificado),Clng(myPactadosOriginal))%>" size=10 maxlength=15 name="CamionesPactados" id="CamionesPactados" value="<%response.write myPactadosModificado %>"/>&nbsp;<font color="red"><b>*</b></font>&nbsp;<%=datoOriginal(myPactadosOriginal)%>
									<% else %>	
										<%=myPactadosOriginal%>
									<% end if %>	
                                </TD>
                                <TD></TD>
                            </TR>
                            <TR>
                                <td></td>
                                <td ALIGN="RIGHT"><% =GF_TRADUCIR("Procedencia")%>:&nbsp;
                                <br><font style="font-size:8;">(Búsqueda por CP o Nombre)</font>
                                </td>
                                <%	dim auxProcedencia
									'Response.Write "(" & datoModificado("CPProcedencia") & ")"
									if trim(datoModificado("CPProcedencia")) <> "" and clng(datoModificado("CPProcedencia")) <> 0 and clng(datoModificado("CPProcedencia")) <> 99 then 
										auxProcedencia = datoModificado("Procedencia")
									else
										auxProcedencia = ""
									end if	
                                %>
								<td colspan="2">
									<span class="demo">
										<span class="ui-widget">
											
											<input id="procedencia" name="procedencia" value="<%=auxProcedencia%>">
										</span>
									</span>                                
                                </td>
									<input type="hidden" name="CAProcedencia" id="CAProcedencia" value="<%=datoModificado("CAProcedencia")%>">
									<input type="hidden" name="CPProcedencia" id="CPProcedencia" value="<%=datoModificado("CPProcedencia")%>">
                                </td>
                            </TR>
                            <tr>
                                <td></td>
                                <td></td>
                                <td><% response.write datoOriginal(trim(rsOriginal("Procedencia")))%></td>
                            </tr>
                    </TABLE>
                    <!-- FIN MARCO DE MERCADERIAS -->
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_c_v.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD VALIGN="TOP">     
					<!-- MARCO DE FIJACION -->
					<TABLE WIDTH="100%">
						<tr>
							<TD COLSPAN="4"><B><% =GF_TRADUCIR("Fijación") %></B></TD>
                        </tr>
                        <%if (datoModificado("FecFijaDesde") <> "" and len(datoModificado("FecFijaDesde"))=8) then%>
							<tr>
								<td WIDTH="5%"></td>
								<td ALIGN="RIGHT" WIDTH="35%"><% =GF_TRADUCIR("Desde")%>:</TD>
								<td COLSPAN=2>
									<input type="hidden" name="FecFijaDesde" id="FecFijaDesde" value="<%=datoModificado("FecFijaDesde")%>">
									<input type="text" class="<%=resaltar(Clng(datoModificado("FecFijaDesde")),Clng(rsOriginal("FecFijaDesde")))%>" readonly="readonly" name="FecFijaDesdeF" id="FecFijaDesdeF" value="<% =GF_FN2DTE(datoModificado("FecFijaDesde")) %>">
									&nbsp;<font color="red"><b>*</b></font>&nbsp;									
									<img id="imgFecFijaDesde" align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario('imgFecFijaDesde',SeleccionarFechaFijaDesde)">&nbsp;<%=datoOriginal(GF_FN2DTE(rsOriginal("FecFijaDesde")))%>
								</td>
							</tr>
							<tr>
								<td></td>
								<td ALIGN="RIGHT"><% =GF_TRADUCIR("Hasta")%>:</td>
								<td COLSPAN=2>
									<input type="hidden" name="FecFijaHasta" id="FecFijaHasta" value="<%=datoModificado("FecFijaHasta")%>">
									<input type="text" class="<%=resaltar(Clng(datoModificado("FecFijaHasta")),Clng(rsOriginal("FecFijaHasta")))%>" readonly="readonly" name="FecFijaHastaF" id="FecFijaHastaF" value="<% =GF_FN2DTE(datoModificado("FecFijaHasta")) %>">
									&nbsp;<font color="red"><b>*</b></font>&nbsp;
									<img id="imgFecFijaHasta" align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario('imgFecFijaHasta',SeleccionarFechaFijaHasta)">&nbsp;<%=datoOriginal(GF_FN2DTE(rsOriginal("FecFijaHasta")))%>
								</td>
							</tr>
							<tr>
								<td></td>
								<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Cant (Min)")%>:</TD>
								<TD>
									<input type="text" class="<%=resaltar(CLng(datoModificado("CantKilosMin")),Clng(rsOriginal("CantKilosMin")))%>"	size=10 maxlength=15 name="CantKilosMin"	id="CantKilosMin"	value="<%response.write datoModificado("CantKilosMin") %>">&nbsp;Kg&nbsp;<font color="red"><b>*</b></font>&nbsp;<%= datoOriginal(rsOriginal("CantKilosMin"))%>
								</TD>
								<TD></TD>
							</tr>
							<tr>
								<TD></TD>
								<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Cant (Max)")%>:</TD>
								<TD>
									<input type="text" class="<%=resaltar(CLng(datoModificado("CantKilosMax")),Clng(rsOriginal("CantKilosMax")))%>"	size=10 maxlength=15 name="CantKilosMax"	id="CantKilosMax"	value="<%response.write datoModificado("CantKilosMax") %>">&nbsp;Kg&nbsp;<font color="red"><b>*</b></font>&nbsp;<%= datoOriginal(rsOriginal("CantKilosMax"))%>
								</TD>
								<TD></TD>
							</tr>
                        <%else%>
                            <tr valign="center" height="90%">
								<td align="center">
                                   <% =GF_Traducir("A este contrato no se le aplica fijación")%>
                                </td>
                            </tr>
                        <%end if%>
                    </TABLE>
                    <!-- FIN MARCO DE FIJACION -->
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
             </TR>
            
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_plus.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
             </TR>
            
             <TR>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD WIDTH="50%" VALIGN="TOP">
					<!-- MARCO DE ENTREGA -->
					<TABLE WIDTH="100%">
						<TR>
							<TD COLSPAN="4"><B><% =GF_TRADUCIR("Entrega")%></B></TD>
                        </TR>
                        <TR>
                            <TD WIDTH="5%"></TD>
                            <TD ALIGN="RIGHT" WIDTH="35%"><% =GF_TRADUCIR("Desde")%>:</TD>
							<TD COLSPAN=2>
								<input type="hidden" name="FecEntDesde" id="FecEntDesde" value="<%=datoModificado("FecEntDesde")%>">
								<input type="text" class="<%=resaltar(Clng(datoModificado("FecEntDesde")),Clng(rsOriginal("FecEntDesde")))%>" readonly="readonly" name="FecEntDesdeF" id="FecEntDesdeF" value="<% =GF_FN2DTE(datoModificado("FecEntDesde")) %>">
								&nbsp;
								<img id="imgFecEntregaDesde" align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario('imgFecEntregaDesde',SeleccionarFechaEntregaDesde)">&nbsp;<%=datoOriginal(GF_FN2DTE(rsOriginal("FecEntDesde")))%>
								<font color="red"><b>*</b></font>&nbsp;
							</TD>
                        </TR>
                        <TR>
							<TD WIDTH="5%"></TD>
                            <TD ALIGN="RIGHT" WIDTH="30%"><% =GF_TRADUCIR("Hasta")%>:</TD>
                            <TD COLSPAN=2>
								<input type="hidden" name="FecEntHasta" id="FecEntHasta" value="<%=datoModificado("FecEntHasta")%>">
								<input type="text" class="<%=resaltar(Clng(datoModificado("FecEntHasta")),Clng(rsOriginal("FecEntHasta")))%>" readonly="readonly" name="FecEntHastaF" id="FecEntHastaF" value="<% =GF_FN2DTE(datoModificado("FecEntHasta")) %>">
								&nbsp;
								<img id="imgFecEntregaHasta" align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario('imgFecEntregaHasta',SeleccionarFechaEntregaHasta)">&nbsp;<%=datoOriginal(GF_FN2DTE(rsOriginal("FecEntHasta")))%>
								<font color="red"><b>*</b></font>&nbsp;
							</TD>
                        </TR>
                        <%
                        strSQL="select * from MERFL.MER192F1 order by DESCDE asc"
                        call GF_BD_AS400_2(rsPuertos,oConn,"OPEN",strSQL)
                        if cint(datoModificado("PtoRecepcion"))>0 then%>
							<TR>
								<TD></TD>
								<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Puerto Entrega") %>:</TD>
								<TD>
									<%
									auxDs = getDescripcionDestino(datoModificado("PtoRecepcion"))
									%>									
                                    <select class="<%=resaltar(Clng(datoModificado("PtoRecepcion")),Clng(rsOriginal("PtoRecepcion")))%>" name="PtoRecepcion" id="PtoRecepcion">
										<option value="<%=datoModificado("PtoRecepcion")%>"><% =GF_TRADUCIR(auxDs) %>
										<% 
										while (not rsPuertos.eof)
											mySelected = ""
											if clng(datoModificado("PtoRecepcion")) = clng(rsPuertos("CODIDE")) then mySelected = "SELECTED"
											%>
											<option <%=mySelected%> value="<% =rsPuertos("CODIDE") %>"><% =GF_TRADUCIR(rsPuertos("DESCDE")) %>
											<%
											rsPuertos.MoveNext
										wend
										%>
									</select>
									&nbsp; <font color="red"><b>*</b></font>
									&nbsp;
									<% response.write datoOriginal(auxDs)%>
								</td>
							</TR>
                        <%
                        end if
                        if cint(rsOriginal("PtoDevolucion"))>0 then
							%>
							<TR>
								<TD></TD>
								<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Puerto Devol.") %>:</TD>
								<TD>
									<%
									auxDs = getDescripcionDestino(rsOriginal("PtoDevolucion"))
									rsPuertos.MoveFirst
                                    %>
                                    <select class="<%=resaltar(Clng(datoModificado("PtoDevolucion")),Clng(rsOriginal("PtoDevolucion")))%>" name="PtoDevolucion" id="PtoDevolucion">
                                        <option SELECTED value="<%=datoModificado("PtoDevolucion")%>"><% =GF_TRADUCIR(auxDs) %>
										<% 
										while (not rsPuertos.eof)%>
											<option value="<% =rsPuertos("CODIDE") %>"><% =GF_TRADUCIR(rsPuertos("DESCDE")) %>
											<%
											rsPuertos.MoveNext
										wend
										%>
									</select>
									&nbsp; <font color="red"><b>*</b></font>
									&nbsp;
									<% response.write datoOriginal(auxDs)%>
								</td>
							</TR>
						<%
						end if
						%>
						<tr>
							<td></td>
							<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Transporte") %>:</TD>
							<TD>
								<%
							    if(isNull(datoModificado("IdTransporte")) or CLng(datoModificado("IdTransporte")) = 0) then
									auxDs = "Sin definir"
							    else
							        auxDsTransporte = getDescripcionTransporte(datoModificado("IdTransporte"))
							    end if
	                            strSQL="Select * from MERFL.MER182F1 order by DESCTR asc"
			                    call GF_BD_AS400_2(rsTransportes,oConn,"OPEN",strSQL)
							    %>
								<select style="zoom: 1" class="<%=resaltar(clng(datoModificado("IdTransporte")),clng(rsOriginal("IdTransporte")))%>" name="IdTransporte" id="IdTransporte">
									<option SELECTED value="<%=datoModificado("IdTransporte")%>"><% =GF_TRADUCIR(auxDsTransporte) %>
									<% 
									while (not rsTransportes.eof)%>
										<option value="<% =rsTransportes("CODITR") %>"><% =GF_TRADUCIR(rsTransportes("DESCTR")) %>
										<%
										rsTransportes.MoveNext
									wend
									%>
								</select>
								&nbsp; <font color="red"><b>*</b></font>
								<%
							    if(isNull(rsOriginal("IdTransporte"))) then
									auxDsTransporte = "Sin definir"
							    else
									auxDsTransporte = getDescripcionTransporte(datoModificado("IdTransporte"))
							    end if
							    %>
								&nbsp;
								<% response.write datoOriginal(auxDsTransporte)%>
							</td>
	                    </tr>
                    </TABLE>
                 <!-- FIN MARCO DE ENTREGA -->
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_c_v.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD VALIGN="TOP">  
					<!-- MARCO DE PAGO -->            
					<TABLE WIDTH="100%">
                       <TR>
                            <TD COLSPAN="4"><B><% =GF_TRADUCIR("Pago") %></B></TD>
                        </TR>
                        
                        <TR>
							<TD WIDTH="5%"></TD>
								<TD WIDTH="35%" ALIGN="RIGHT"><% =GF_TRADUCIR("Moneda") %>:</TD>
									<TD align="left" colspan=2>
										<input type="radio" style="cursor:pointer" name="cdMoneda" value='<% =MONEDA_PESO %>' <% if cstr(cdMoneda) = MONEDA_PESO then Response.Write "CHECKED"%>> Pesos
										<input type="radio" style="cursor:pointer" name="cdMoneda" value='<% =MONEDA_DOLAR %>' <% if cstr(cdMoneda) = MONEDA_DOLAR then Response.Write "CHECKED"%>> Dolares
									</TD>
						</TR>
                        
                        <TR>
							<TD WIDTH="5%"></TD>
								<%  
								precioTonelada = 0 
								select case CInt(rsOriginal("Operacion"))
									case 0,1,2,3,5
									    simboloMoneda = getSimboloMoneda(MONEDA_PESO)
									    precioTonelada = CDbl(datoModificado("PrecioP"))
									    precioToneladaOriginal = CDbl(rsOriginal("PrecioP"))
									case 6,9,10,11,12
									    simboloMoneda = getSimboloMoneda(MONEDA_DOLAR)
									    precioTonelada = CDbl(datoModificado("PrecioD"))
									    precioToneladaOriginal = CDbl(rsOriginal("PrecioD"))
							    end select
							    %>
								<TD WIDTH="35%" ALIGN="RIGHT"><% =GF_TRADUCIR("Precio") %>:</TD>
								<%
								if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then
									%>
									<TD align="left" colspan=2>
										<%' =simboloMoneda %>
										<input style="text-align:right;" class="<%=resaltar(precioTonelada,precioToneladaOriginal)%>" size="10" type="text" name="precio" id="precio" value="<%=precioTonelada%>" onKeyPress=" return controlIngreso(this, event, 'E');">&nbsp; <font color="red"><b>*</b></font>&nbsp;<% =datoOriginal(GF_EDIT_DECIMALS(cdbl(precioToneladaOriginal)*100,2))%>
									</TD>
									<%
								else
									%>
									<TD align="left" colspan=2>
										
										<input style="text-align:right;" align="right" size="10" type="text" name="precio" id="precio" value="<%=myValuePrecio%>" onKeyPress=" return controlIngreso(this, event, 'E');" >&nbsp; <font color="red"><b>*</b></font>
									</TD>
									<%
								end if
								%>
						</TR>
						<TR>
							<TD></TD>
							<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Parcial") %>:</TD>
							<TD align="left">
								<input align="right" class="<%=resaltar(CDbl(datoModificado("PjeParcial")),CDbl(rsOriginal("PjeParcial")))%>" size="6" maxlength="5" type="text" name="PjeParcial" id="PjeParcial" value="<% =datoModificado("PjeParcial") %>" onKeyPress=" return controlIngreso(this, event, 'I');">&nbsp;%&nbsp;&nbsp;<font color="red"><b>*</b></font>
								<% 
								myVal = datoOriginal(GF_EDIT_DECIMALS(CDbl(rsOriginal("PjeParcial"))*100, 2))
								Response.Write "&nbsp;" & myVal
								%>
							</TD>	
						</TR>                            
						<%
						if (not (isNull(rsOriginal("DiasPago"))) and (rsOriginal("DiasPago") <> "0")) then 
							'si hay dias de pago indicados
							%>
							<TR>
								<TD></TD>
								<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Dias Pago") %>:</TD>
								<TD colspan="2"><%=rsOriginal("DiasPago")%></TD>
							</TR>
							<% 
						end if 
						%>
						<TR>
							<TD></TD>
							<TD ALIGN="RIGHT"><% =GF_TRADUCIR("Forma de Pago") %>:</TD>
							<TD align="left">
							<% 
							if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then 
								strSQL="Select * from MERFL.MER2I1F1 order by DESCFP asc"
								Call GF_BD_AS400_2(rsFormasPago,oConn,"OPEN",strSQL)                               
								%>
					            <select name="CodigoPago" id="CodigoPago" onchange="javascript:HabilitarCamionesPactadosDiv(this)">
									<% 
									while (not rsFormasPago.eof)%>
										<option <% if (datoModificado("CodigoPago") = rsFormasPago("CODIFP")) then Response.write "selected" %> value="<% =rsFormasPago("CODIFP") %>">
											<%
											if CINT(rsOriginal("CamionesPactados")) = 0 then
												myReplacement = ""
											else	
												myReplacement = rsOriginal("CamionesPactados")
											end if	
											response.write replace(GF_Traducir(rsFormasPago("DESCFP")),"X", myReplacement)
										rsFormasPago.MoveNext
									wend 
									%>
								</select>
								<%
							else									
								if ((isNull(rsOriginal("CodigoPago"))) or (rsOriginal("CodigoPago") = "")) then
									response.write GF_Traducir("Sin Forma de Pago definido")                           
								else
									if rsOriginal("CodigoPago") = "X" then 
										myCodigoPago = "A"
									else
										myCodigoPago = rsOriginal("CodigoPago")
									end if	
									response.write getDsPago(myCodigoPago)
								end if
								%>
								<input type="hidden" name="CodigoPago" id="CodigoPago" value="<%=rsOriginal("CodigoPago")%>">
								<%
							end if
							%>
							
							</td>
						</TR>
						<% if (CLng(session("KCOrganizacion")) = CLng(KC_TOEPFER)) then %>
						<tr>
                            <TD></TD>
                            <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Fecha de Pago")%>:</TD>
							<TD COLSPAN=2>
								<input type="hidden" name="FechaPago" id="FechaPago" value="<%=myFechaPago%>">
								<input type="text" class="<%=resaltar(Clng(myFechaPago),Clng(rsOriginal("FechaPago")))%>" readonly="readonly" name="FechaPagoF" id="FechaPagoF" value="<% =GF_FN2DTE(myFechaPago) %>">
								&nbsp;
								<img id="imgFechaPago" align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario('imgFechaPago',SeleccionarFechaPago)">&nbsp;<%=datoOriginal(GF_FN2DTE(rsOriginal("FechaPago")))%>
								<font color="red"><b>*</b></font>&nbsp;
							</TD>						
						</tr>
						<% else 
								if (Len(rsOriginal("FechaPago")) = 8) then
								%>
							
								<TR>
									<TD></TD>
									<TD align="right"><% =GF_TRADUCIR("Fecha de Pago")%>:</TD>
									<TD COLSPAN=2><%
										if not isnull(rsOriginal("FechaPago")) then
											Response.write GF_FN2DTE(rsOriginal("FechaPago")) 
										end if
										%></TD>
								</TR>
								<%
								end if
							end if	
						if (esPagoDirectoVendedor(rsOriginal("KCVEN"))) then
							%>
							<TR>
								<TD COLSPAN=4><% =GF_TRADUCIR("PAGO DIRECTO AL VENDEDOR") %></TD>
							</TR>
							<%
						end if
						%>
					</TABLE>
					<!-- FIN MARCO DE PAGO --> 
				</TD>
                <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
			</TR>
			<TR>
			    <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
			    <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
			    <TD WIDTH="8"><img src="images/marco_T_I.gif"></TD>
			    <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
			    <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
			</TR>             
		<tr>
			<td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
			<td colspan=3><b><% =GF_TRADUCIR("Observaciones de Boleto") %></b></td>
			<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
		</tr>
		<tr id="trLegal">
			<td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
			<td colspan=3>
				<div Align=justify>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<i><%=getObservacionesBoleto()%></i>
				</div>
			</td>
			<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
		</tr>
		<TR>
		    <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
		    <TD colspan="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
		    <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
		</TR>  

		<tr>
			<TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
			<td colspan="3" align='left'>
				<B><% =GF_TRADUCIR("Observaciones") %></B>&nbsp;&nbsp;<font color="red"><b>*</b></font>
			</td>
			<TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
		</tr>
		<tr>
			<TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
			<td colspan="3" align='center'>
				<textarea class="<%=resaltar(datoModificado("Observaciones"),"")%>" id='Observaciones' name='Observaciones' rows='3' cols='114'><%=datoModificado("Observaciones")%></textarea>
			</td>
			<TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
		</tr>
		<tr>
			<TD WIDTH="8"><img src="images/marco_r3_c1.gif"></TD>
			<TD colspan="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
			<TD WIDTH="8"><img src="images/marco_r3_c3.gif"></TD>
		</TR>
		<tr>
			<td colspan="5"><font color="red"><b>*</b></font>&nbsp;<font color="blue"><i><%=GF_Traducir("Cambios en estos campos requeriran confirmacion por parte de personal de Toepfer")%></i></font></td>
		</tr>  
		      
	</table>			
      <!-- FIN CONTRATO -->      
      
	  <% end if %>	
      <INPUT type="HIDDEN" name="accion" id="accion" value="">
      <INPUT TYPE="HIDDEN" NAME="cmbProducto"	ID="cmbProducto"	VALUE="<% =producto%>">
      <INPUT TYPE="HIDDEN" NAME="txtSucursal"	ID="txtSucursal"	VALUE="<% =sucursal%>">
      <INPUT TYPE="HIDDEN" NAME="txtOperacion"	ID="txtOperacion"	VALUE="<% =operacion%>">
      <INPUT TYPE="HIDDEN" NAME="txtNumero"		ID="txtNumero"		VALUE="<% =numero%>">
      <INPUT TYPE="HIDDEN" NAME="txtCosecha"	ID="txtCosecha"		VALUE="<% =cosecha%>">      
</FORM>
</div>
<br>
</BODY>
</HTML>