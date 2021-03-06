<%

Const UNIDAD_KILOS = 1
Const UNIDAD_TONELADAS = 2

Const TIPO_PROPIA_PRODUCCION = "P"
Const TIPO_CONSIGNACION = "C"
Const TIPO_NO_PROPIA_PRODUCCION = "N"   

Const PROVEEDOR_A_CONFIRMAR = 1
Const SIN_CORREDOR = 0
Const KC_TOEPFER = 99999997
'Definicion de variables globales a utilizar
Dim g_intProducto,g_intSucursal,g_intOperacion,g_intFechaFijaDesde
Dim g_intNumero,g_intCosecha,g_intFechaConc,g_intPrecioP, g_intKilosNetos
Dim g_intKCCOR,g_intKCVEN,g_intKilos,g_intAnulaciones,g_intPjeParcial
Dim g_intFechaFijaHasta,g_intFechaEntDesde,g_intFechaEntHasta,g_intKilosMin
Dim g_intKilosMax,g_intPuertoRecepcion,g_intPuertoDevolucion,g_chrMercPropia, g_chrMercConsigna
Dim g_strDiasPago, g_strRecibido, g_strCondicionIVA, g_intFechaPago, g_chrMercConHumedad, g_intTransporte
Dim g_strCtoCorredor,g_chrTipoCto,g_chrMrcConfirma,g_strProcedencia, g_strCAProcedencia, g_strCPProcedencia
Dim g_strKCPago,g_intPrecioD,g_intTipoCambio,g_intKgEntregados, g_strCodigoPago, g_intCamionesPactados, g_strObservacion
Dim g_intFechaDesde, g_intFechaHasta
'Recordsets usados
Dim g_rsContratos, g_rsContratosConf, g_rsDescargas, g_rsAnalisis, g_rsDetAnalisis
Dim g_strCampoOrden, g_tipo
'Var usadas para las descargas
Dim g_intFechaDescarga,g_chrMrcConforme,g_intPuerto, g_intPlanillaCos, g_intPlanillaNro,g_intPlanillaSec,g_intCartaPorte,g_intReciboNro,g_intSolicitudNro,g_intCdeEs, g_intAnalisisGdo
Dim g_flgAgrupar, g_intCantidad, g_intKilosDescarga
'Var usadas para los analisis
Dim g_intFechaAnalisis, g_intBolsa, g_intGrado, g_intCosto, g_intNumeroAnalisis
'Var usadas para los detalles de los analisis
Dim g_strConceptoDs, g_intConcepto, g_intBonif, g_intRebaja, g_intValor

'/**
' * Funcion    : GF_EDIT_CONTRATO
' * Descripcion: Esta funcion Arma el numero de contrato dandole
' *              el formato alfanumerico correcto
' * Parametros : p_cto1 [in] 2 digitos (producto)
' *              p_cto2 [in] 1 digito  (sucursal)
' *              p_cto3 [in] 2 digitos (operacion)
' *              p_cto4 [in] 5 digitos (secuencia)
' *              p_cto5 [in] 2 digitos (cosecha)
' * Valor Devuelto:
' * Si todos los parametros son correctos devuelve el numero
' * de contrato con el formato XX-X-XX-XXXXX/XX.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 01/07/2004
' */
Function GF_EDIT_CONTRATO(p_cto1,p_cto2,p_cto3,p_cto4,p_cto5)

   Dim aux
   'Controlo que los parametros sean numericos
   aux = GF_nDigits(p_cto1,2) & "-" & GF_nDigits(p_cto2,1)
   aux = aux & "-" & GF_nDigits(p_cto3,2) & "-" & GF_nDigits(p_cto4,5) & "/" & GF_nDigits(p_cto5,2)
   GF_EDIT_CONTRATO = aux
   
End Function
'/**
' * Funcion    : GF_INT2CTO
' * Descripcion: Esta funcion transforma un string numerico en un
' *              contrato dandole el formato alfanumerico correcto
' * Parametros : p_cto [in] 12 digitos (formato : XXXXXXXXXXXX)
' * Valor Devuelto:
' * Si todos los parametros son correctos devuelve el numero
' * de contrato con el formato XX-X-XX-XXXXX/XX.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 01/07/2004
' */
Function GF_INT2CTO(p_cto)

   Dim aux
   'Controlo que el string parado tenga la cant de digitos correcta.
   if (Len(p_cto) <> 12) then
      aux = GF_EDIT_CONTRATO("00","0","00","00000","00")
   else
      aux= GF_EDIT_CONTRATO(left(p_cto,2),mid(p_cto,3,1),mid(p_cto,4,3),mid(p_cto,6,5),right(p_cto,2))
   end if
   GF_INT2CTO = aux

End Function
'/**
' * Funcion: initHeader
' * Descripcion: Esta funcion inicializa la lectura de la
' *              cabecera de los contratos.
' * Parametros:
' * Valor Devuelto:
' * Si hay al menos un registro valido devuelve 1 sino 0.
' * Modificacion> se agrego un control if, chequeando el valor de variable g_consultarContratosConf
' * que indica, si se va a consultar la tabla de Contratos o ContratosConf
' * Autor modificacion> Henzel Pavlo
' * Fecha 15/11/2007
' */
Function initHeader()
    Dim strWhere, strORKC, oConn, strSQL, strOrder, strgroup, tt, fechaControl
    fechaControl = GF_DTEADD(mid(session("MmtoSistema"),1,8),-1,"a")
    fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,5,2)-1)*-1,"M")
    fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,7,2)-1)*-1,"D")    
    strWhere = " where (CTO.FHPER1 > '" & fechaControl & "') "
    if g_intKCCOR = "" and g_intKCVEN = "" then
        strORKC = session("KCOrganizacion")
        if cint(strORKC) <> KC_TOEPFER then
            'Si no es el usuario de toepfer, q el usuario logueado sea el corredor o el vendedor
            strWhere = strWhere & " and (CTO.CVENR1=" & strORKC & " or CTO.CCORR1=" & strORKC
            'Si no se estan viendo boletos, mostrar los negocios donde el corredor sea el MAT 
			'para los boletos NO.
            if (g_tipo <> "BOLETO") then strWhere = strWhere & " or MAT.CCORRH= " & strORKC 
            strWhere = strWhere & ")"
        else			
			Call mkWhere(strWhere, "CTO.CCORR1","5454","<>",3) 			
        end if
    else
        if g_intKCCOR <> "" then call mkWhere(strWhere, "CTO.CCORR1",g_intKCCOR,"=",1)
        if g_intKCVEN <> "" then call mkWhere(strWhere, "CTO.CVENR1",g_intKCVEN,"=",1)
    end if
    if (g_intProducto <> "") then Call mkWhere(strWhere, "CTO.CPROR1", g_intProducto,"=",1)
    if (g_intSucursal <> "") then Call mkWhere(strWhere, "CTO.CSUCR1", g_intSucursal,"=",1)
    if (g_intOperacion <> "") then Call mkWhere(strWhere, "CTO.COPER1", g_intOperacion,"=",1)
    if (g_intNumero <> "") then Call mkWhere(strWhere, "CTO.NCTOR1", g_intNumero,"=",1)
    if (g_intCosecha <> "") then Call mkWhere(strWhere, "CTO.ACOSR1", g_intCosecha,"=",1)
	
	if (g_chrMrcConfirma <> "") then
		'Si me determina una marca de confirmaci�n, solo muestro prestamos si quiere ver contratos confirmados, si 
		'quiere ver los no confirmados, puede que sea para BOLETO, con lo cual ah� no se muetran los prestamos.
		strWhere = strWhere & " and "
		if (g_chrMrcConfirma = "V") then strWhere = strWhere & "("		
		strWhere = strWhere & "((CTO.TIPOR1='A' or CTO.TIPOR1='B') and (CTO.CONFR1='" & g_chrMrcConfirma & "' or CTO.CCORR1 = 5454))"
		if (g_chrMrcConfirma = "V") then strWhere = strWhere & " or CTO.TIPOR1='C')"
	else
		'Si se filtra por una operacion especifica, no se limita a contratos de compra y venta dado que se pueden elegir prestamos.		
		if (g_intOperacion = "") then strWhere = strWhere & " and (CTO.TIPOR1='A' or CTO.TIPOR1='B')"
	end if
	
	if (g_chrMrcRecibido <> "") then
        call mkWhere(strWhere, "R.PLRECI", g_chrMrcRecibido, "=", 3)
    end if 

    'Muestra o envia solo los boletos con productos que NO son 9,17,25 o que son 9 o 17, pero con codigo de operacion 9.                
    'strWhere = strWhere & " and (CTO.CPROR1 not in (9,17,25) or ((CTO.CPROR1 = 9 or CTO.CPROR1 = 17) and CTO.COPER1 = 9)) " 
    'strWhere = strWhere & " and (CTO.CPROR1 not in (9,17,25) or (CTO.CPROR1 = 9 and CTO.COPER1 = 9) or (CTO.CPROR1 = 17 and CTO.COPER1 = 9 and CTO.ACOSR1 <> 12)) " 
    if (cint(strORKC) <> 99999997) then
		strWhere = strWhere & " AND ((CTO.CPROR1 not in (9, 17, 25, 28))"
		if (g_tipo <> "BOLETO") then 'Si quiere ver los contratos confirmados o uno en particular, mostrar todo.
			strWhere = strWhere & " or (CTO.CPROR1 = 17) or (CTO.CPROR1 = 9) "
		else
			strWhere = strWhere & " or (CTO.CPROR1 = 17 and CTO.COPER1 = 9 and CTO.ACOSR1 <= 12) " & _
								"	or (CTO.CPROR1 = 9 and CTO.COPER1 in (0, 9) and CTO.ACOSR1=14)" & _
								"	or (CTO.CPROR1 = 17 and CTO.COPER1 in (0, 9) and CTO.ACOSR1 = 14)"
		end if
		strWhere = strWhere & ") " 
	end if	

    if (g_intFechaConc <> "") then strWhere = strWhere & GF_LIKE("CTO.FCCTR1", g_intFechaConc)
    strgroup = " group by CTO.CPROR1, CTO.CSUCR1, CTO.COPER1, CTO.NCTOR1, CTO.ACOSR1, CTO.FCCTR1, CTO.CCORR1, CTO.CVENR1, CTO.KGCOR1, CTO.PORPR1, CTO.FDFIR1, CTO.FHFIR1, CTO.FDPER1, CTO.FHPER1, CTO.KGNFR1, CTO.KGMFR1, CTO.CDESR1, CTO.DESTR1, J.MCPDRJ,"
	strgroup = strgroup & " CTO.CONCR1, CTO.TIPOR1, CTO.CONFR1, P.DESCPC, CTO.CPRDR1, CTO.AUXIR1, F.DESCFP, M.MDOLOM, CTO.TIPCR1, CTO.KGRER1, CTO.CTRAR1, CTO.CFPAR1, J.HUMERJ, K.CAPARK, CTO.FPACR1, J.CDIVRJ, R.PLRECI, CTO.DIAPR1, J.MCONRJ, CTO.PRECR1 "
    strSQL = "Select CTO.CPROR1 as Producto, CTO.CSUCR1 as Sucursal, CTO.COPER1 as Operacion, CTO.NCTOR1 as Numero, CTO.ACOSR1 as Cosecha, CTO.FCCTR1 as FechaConc, CTO.CCORR1 as KCCOR, CTO.CVENR1 as KCVEN, CTO.KGCOR1 as Kilos, "
    strSQL = strSQL & "sum(B.KGCORB) as Anulaciones, "
	strSQL = strSQL & " case(M.MDOLOM) when 'F' then CTO.PRECR1*1 else CTO.PRECR1*CTO.TIPCR1 end as PrecioP, CTO.PORPR1 as PjeParcial, CTO.FDFIR1 as FechaFijaDesde, CTO.FHFIR1 as FechaFijaHasta, CTO.FDPER1 as FechaEntDesde,"
	strSQL = strSQL & " CTO.FHPER1 as FechaEntHasta, CTO.KGNFR1 as KilosMin, CTO.KGMFR1 as KilosMax, CTO.CDESR1 as PuertoRecepcion, CTO.DESTR1 as PuertoDevolucion, J.MCPDRJ as MercPropia, CTO.CONCR1 as CtoCorredor, CTO.TIPOR1 as TipoCto, CTO.CONFR1 as MrcConfirma,"
	strSQL = strSQL & " P.DESCPC as Procedencia, CTO.CPRDR1 as CPProcedencia, CTO.AUXIR1 as CAProcedencia, F.DESCFP as KCPago, case(M.MDOLOM) when 'F' then case(CTO.TIPCR1) when 0 then 0 else CTO.PRECR1/CTO.TIPCR1 end else CTO.PRECR1*1 end as PrecioD, CTO.TIPCR1 as TipoCambio,"
	strSQL = strSQL & " CTO.KGRER1 as SaldoEnt, CTO.CTRAR1 as Transporte, CTO.CFPAR1 as CodigoPago, J.HUMERJ as MercConHumedad, K.CAPARK as CamionesPactados, CTO.FPACR1 as FechaPago, case(J.CDIVRJ) when 'C' then J.CDIVRJ else 'X' end as CondicionIVA,"
	strSQL = strSQL & " R.PLRECI as Recibido, CTO.DIAPR1 as DiasPago, J.MCONRJ as MercConsignacion"	
	strSQL = strSQL & " from MERFL.MER311F1 CTO left join MERFL.MER311FB B on CTO.CPROR1=B.CPRORB and CTO.CSUCR1=B.CSUCRB and CTO.COPER1=B.COPERB and CTO.NCTOR1=B.NCTORB and CTO.ACOSR1=B.ACOSRB left join MERFL.MER311FJ J on CTO.CPROR1=J.CPRORJ and"
    strSQL = strSQL & " CTO.CSUCR1=J.CSUCRJ and CTO.COPER1=J.COPERJ and CTO.NCTOR1=J.NCTORJ and CTO.ACOSR1=J.ACOSRJ left join MERFL.MER311FK K on CTO.CPROR1=K.CPRORK and CTO.CSUCR1=K.CSUCRK and CTO.COPER1=K.COPERK and CTO.NCTOR1=K.NCTORK and CTO.ACOSR1=K.ACOSRK"
	strSQL = strSQL & " left join MERFL.MER341F2 R on CTO.CPROR1=R.PLCPRO and CTO.CSUCR1=R.PLCSUC and CTO.COPER1=R.PLCOPE and CTO.NCTOR1=R.PLNCTO and CTO.ACOSR1=R.PLACOS left join MERFL.MER2I1F1 F on CTO.CFPAR1=F.CODIFP left join MERFL.MER132F1 M on CTO.COPER1=M.CODIOM"
	strSQL = strSQL & " left join MERFL.MER142F1 P on CTO.CPRDR1=P.CODIPC and CTO.AUXIR1=P.AUXIPC" 
	strSQL = strSQL & " left join MERFL.MER311FH MAT on CTO.CPROR1=MAT.CPRORH and CTO.CSUCR1=MAT.CSUCRH and CTO.COPER1=MAT.COPERH and CTO.NCTOR1=MAT.NCTORH and CTO.ACOSR1=MAT.ACOSRH " & strWhere & strgroup
    if (g_strCampoOrden = "") then g_strCampoOrden = " CTO.FCCTR1 asc"
    strSQL = strSQL & " order by " & g_strCampoOrden
	
	'response.write "<hr>la consulta dentro de include es " &  strSQL & "<hR>"
    Call GF_BD_AS400_2(g_rsContratos,oConn,"OPEN",strSQL)            
    if (g_rsContratos.eof) then
       initHeader = 0
    else
       initHeader = 1
    end if
End Function
'/**
' * Funcion: getNextHeader
' * Descripcion: Esta funcion lee la siguiente linea en los
' *              resultados obtenidos por initHeader.
' * Parametros:
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 01/07/2004
' */
Function getNextHeader()
if (g_rsContratos.eof) then
   getNextHeader=0
else
    getParametrosRepetidos()
   if isnull(g_intKgEntregados) then g_intKgEntregados = 0
   g_rsContratos.movenext
   getNextHeader=1
  end if
End Function
'------------------------------------------------------------------------------------------
'Se usa para evitar repeticion de codigo en distintas funciones de lectura de campos de record set
'Autor: Henzel Pavlo
Function  getParametrosRepetidos()
   g_intProducto= g_rsContratos("Producto")
   g_intSucursal= g_rsContratos("Sucursal")
   g_intOperacion= g_rsContratos("Operacion")
   g_intNumero= g_rsContratos("Numero")
   g_intCosecha= g_rsContratos("Cosecha")
   g_intFechaConc= g_rsContratos("FechaConc")
   g_intKCCOR= g_rsContratos("KCCOR")
   g_intKCVEN= g_rsContratos("KCVEN")
   g_intKilos= Cdbl(g_rsContratos("Kilos"))   
   g_intAnulaciones= g_rsContratos("Anulaciones")   
   if isnull(g_intAnulaciones) then g_intAnulaciones = 0   
   'Se corrigen los kilos contratados por la cantidad neta (descuento/suma de anulaciones y amplizaciones)
   g_intKilosNetos = g_intKilos + CDbl(g_intAnulaciones)
   g_intPrecioP= g_rsContratos("PrecioP")
   g_intPjeParcial= g_rsContratos("PjeParcial")
   g_intFechaFijaDesde= g_rsContratos("FechaFijaDesde")
   g_intFechaFijaHasta= g_rsContratos("FechaFijaHasta")
   g_intFechaEntDesde= g_rsContratos("FechaEntDesde")
   g_intFechaEntHasta= g_rsContratos("FechaEntHasta")
   g_intKilosMin = "0"
   if (cDbl(g_rsContratos("KilosMin")) <> 0) then g_intKilosMin= g_rsContratos("KilosMin")
   g_intKilosMax = "0"
   if (cDbl(g_rsContratos("KilosMax")) <> 0) then g_intKilosMax= g_rsContratos("KilosMax")
   g_intPuertoRecepcion= g_rsContratos("PuertoRecepcion")
   g_intPuertoDevolucion= g_rsContratos("PuertoDevolucion")
   g_chrMercPropia= g_rsContratos("MercPropia")
   g_chrMercConsigna= g_rsContratos("MercConsignacion")
   g_strCtoCorredor= g_rsContratos("CtoCorredor")
   g_chrTipoCto= g_rsContratos("TipoCto")
   g_chrMrcConfirma= g_rsContratos("MrcConfirma")
   g_strProcedencia= g_rsContratos("Procedencia")
   g_strCAProcedencia= g_rsContratos("CAProcedencia")
   g_strCPProcedencia= g_rsContratos("CPProcedencia")
   g_strKCPago= g_rsContratos("KCPago")
   g_intPrecioD= g_rsContratos("PrecioD")
   g_intTipoCambio= g_rsContratos("TipoCambio")
   g_intKgEntregados = g_rsContratos("SaldoEnt")
   g_intTransporte = g_rsContratos("Transporte")
   g_strCodigoPago= g_rsContratos("CodigoPago")
   g_chrMercConHumedad = g_rsContratos("MercConHumedad")
   g_intCamionesPactados = g_rsContratos("CamionesPactados")
   g_intFechaPago = g_rsContratos("FechaPago")
   g_strCondicionIVA = g_rsContratos("CondicionIVA")
   g_strRecibido = g_rsContratos("Recibido")
   g_strDiasPago = g_rsContratos("DiasPago")
   if isnull(g_intKgEntregados) then g_intKgEntregados = 0
End function
'------------------------------------------------------------------------------------------
Function GF_reset_Contrato()
   g_intProducto= ""
   g_intSucursal= ""
   g_intOperacion= ""
   g_intNumero= ""
   g_intCosecha= ""
   g_intFechaConc= ""
   g_intKCCOR= ""
   g_intKCVEN= ""
   g_intKilos= ""
   g_intAnulaciones= ""
   g_intPrecioP= ""
   g_intPjeParcial= ""
   g_intFechaFijaDesde= ""
   g_intFechaFijaHasta= ""
   g_intFechaEntDesde= ""
   g_intFechaEntHasta= ""
   g_intKilosMin = ""
   g_intKilosMax = ""
   g_intPuertoRecepcion= ""
   g_intPuertoDevolucion= ""
   g_chrMercPropia= ""
   g_strCtoCorredor= ""
   g_chrTipoCto= ""
   g_chrMrcConfirma= ""
   g_strProcedencia= ""
   g_intCPProcedencia= ""
   g_intCAProcedencia= ""
   g_strCodigoPago= ""
   g_intPrecioD= ""
   g_intTipoCambio= ""
End Function
'/**
' * Funcion: getKgEntregados
' * Descripcion: Esta funcion obtiene los Kg entregados
' *              para un determinado contrato.
' * Parametros : p_intProducto  [in] 2 digitos (producto)
' *              p_intSucursal  [in] 1 digito  (sucursal)
' *              p_intOperacion [in] 2 digitos (operacion)
' *              p_intNumero    [in] 5 digitos (secuencia)
' *              p_intCosecha   [in] 2 digitos (cosecha)
' *
' * Autor: Javier A. Scalisi
' * Fecha: 01/07/2004
' */
Function getKgEntregados(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha)
         Dim strSQL, oConn, rs
         Dim ret
         
         ret=0
         strSQL="Select sum(KGNER6) Kilos from MERFL.MER311F6 where"
         strSQL= strSQL & " CPROR6=" & p_intProducto & " and CSUCR6=" & p_intSucursal
         strSQL= strSQL & " and COPER6= " & p_intOperacion & " and NCTOR6=" & p_intNumero & "and ACOSR6=" & p_intCosecha
         call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
         if (not isNull(rs("Kilos"))) then ret = rs("Kilos")
         getKgEntregados = ret
End Function
'/**
' * Funcion: initDescarga
' * Descripcion: Esta funcion inicializa la lectura de la
' *              descarga de los contratos.
' * Parametros:
' * Valor Devuelto:
' * Si hay al menos un registro valido devuelve 1 sino 0.
' *
' * Autor: Eugenio Di Santo
' * Fecha: 05/08/2004
' */
Function initDescarga()
    Dim strWhere, strORKC, oConn, strSQL, strOrder, strGroup, fechaControl
            
     fechaControl = GF_DTEADD(mid(session("MmtoSistema"),1,8),-1,"a")
     fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,5,2)-1)*-1,"M")
     fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,7,2)-1)*-1,"D")     
    strORKC= session("KCOrganizacion")
        strGroup = " group by D.CPROR6, D.CSUCR6, D.COPER6, D.NCTOR6, D.ACOSR6, D.CDESR6, D.ACOPR6, D.PLANR6, D.FECDR6, C.CONCR1, C.CCORR1, C.CVENR1, D.CPORR6, D.RECIR6, S.MENSCS, D.NSANR6, D.MCONR6, D.GRADR6"
        strSQL = "Select D.CPROR6 as Producto, D.CSUCR6 as Sucursal, D.COPER6 as Operacion, D.NCTOR6 as Numero, D.ACOSR6 as Cosecha, D.CDESR6 as Puerto, D.ACOPR6 as PlanillaCos, D.PLANR6 as PlanillaNro, D.FECDR6 as FechaDescarga, C.CONCR1 as CtoCorredor, C.CCORR1 as KCCOR, C.CVENR1 as KCVEN, D.CPORR6 as CartaPorte, D.RECIR6 as ReciboNro, S.MENSCS as CdeEs, D.NSANR6 as SolicitudNro, sum(D.KGNER6) as KgDescarga, D.MCONR6 as MrcConforme, D.GRADR6 as AnalisisGdo, count(*) as Cantidad "
        strSQL = strSQL & "from (MERFL.MER311F6 D inner join MERFL.MER2D2F1 S on D.COSTR6=S.CONCCS) "
        strSQL = strSQL & "inner join MERFL.MER311F1 C on D.CPROR6=C.CPROR1 and D.CSUCR6=C.CSUCR1 and D.COPER6=C.COPER1 and D.NCTOR6=C.NCTOR1 and D.ACOSR6=C.ACOSR1 "
        strSQL = strSQL & "left join MERFL.MER311FH MAT on D.CPROR6=MAT.CPRORH and D.CSUCR6=MAT.CSUCRH and D.COPER6=MAT.COPERH and D.NCTOR6=MAT.NCTORH and D.ACOSR6=MAT.ACOSRH "
		strWhere = "where D.FECDR6 > '" & fechaControl & "' and (D.KGNER6 > 0) and ((D.COPER6 <> 04 and (C.CONFR1 = 'V' or C.CCORR1=5454)) or D.COPER6 = 04) and (C.CVENR1=" & strORKC & " or C.CCORR1=" & strORKC & " or MAT.CCORRH=" & strORKC & ")"    
    if (strORKC = "99999997") then strWhere = ""
    if (g_intProducto <> "") then Call mkWhere(strWhere, "D.CPROR6", g_intProducto,"=",1)
    if (g_intSucursal <> "") then Call mkWhere(strWhere, "D.CSUCR6", g_intSucursal,"=",1)
    if (g_intOperacion <> "") then Call mkWhere(strWhere, "D.COPER6", g_intOperacion,"=",1)
    if (g_intNumero <> "") then Call mkWhere(strWhere, "D.NCTOR6", g_intNumero,"=",1)
    if (g_intCosecha <> "") then Call mkWhere(strWhere, "D.ACOSR6", g_intCosecha,"=",1)
    if (g_intPuerto <> "") then Call mkWhere(strWhere, "D.CDESR6", g_intPuerto,"=",1)
    if (g_intPlanillaCos <> "") then Call mkWhere(strWhere, "D.ACOPR6", g_intPlanillaCos,"=",1)
    if (g_intPlanillaNro <> "") then Call mkWhere(strWhere, "D.PLANR6", g_intPlanillaNro,"=",1)
    if (g_intPlanillaSec <> "") then Call mkWhere(strWhere, "D.SECUR6", g_intPlanillaSec,"=",1)
    if (g_intFechaDescarga <> "") then strWhere = strWhere & GF_LIKE("D.FECDR6", g_intFechaDescarga)
    if (g_intSolicitudNro <> "") then Call mkWhere(strWhere, "D.NSANR6", g_intSolicitudNro,"=",1)
    if (g_intCartaPorte <> "") then Call mkWhere(strWhere,"D.CPORR6", g_intCartaPorte,"=",1)
    if (g_intFechaDesde <> "") then Call mkWhere(strWhere,"D.FECDR6", g_intFechaDesde,">=",1)
    if (g_intFechaHasta <> "") then Call mkWhere(strWhere,"D.FECDR6", g_intFechaHasta,"<=",1)
    if (g_strCampoOrden = "") then g_strCampoOrden = "D.FECDR6 asc"
    'call mkWhere(strWhere, "C.CONFR1","V","=",3)
        
    strSQL = strSQL & strWhere & strGroup & " order by " & g_strCampoOrden
    'response.write strSQL & "<br>"
    Call GF_BD_AS400_2(g_rsDescargas,oConn,"OPEN",strSQL)
    if (g_rsDescargas.eof) then
       initDescarga = 0
    else
       initDescarga = 1
    end if
End Function
'/**
' * Funcion: getNextDescarga
' * Descripcion: Esta funcion lee la siguiente linea en los
' *              resultados obtenidos por initDescarga.
' * Parametros:
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Eugenio Dio Santo
' * Fecha: 05/08/2004
' */
Function getNextDescarga()
if (g_rsDescargas.eof) then
   getNextDescarga=0
else
   g_intProducto     = g_rsDescargas("Producto")
   g_intSucursal     = g_rsDescargas("Sucursal")
   g_intOperacion    = g_rsDescargas("Operacion")
   g_intNumero       = g_rsDescargas("Numero")
   g_intCosecha      = g_rsDescargas("Cosecha")
   g_intFechaDescarga= g_rsDescargas("FechaDescarga")
   g_intKCCOR        = g_rsDescargas("KCCOR")
   g_intKCVEN        = g_rsDescargas("KCVEN")
   g_intKilosDescarga= cdbl(g_rsDescargas("KgDescarga"))
   g_strCtoCorredor  = g_rsDescargas("CtoCorredor")
   g_intPuerto       = g_rsDescargas("Puerto")   
   g_intPlanillaCos  = g_rsDescargas("PlanillaCos")
   g_intPlanillaNro  = g_rsDescargas("PlanillaNro")
   g_intCartaPorte   = g_rsDescargas("CartaPorte")
   g_intReciboNro    = g_rsDescargas("ReciboNro")
   g_intCdeEs        = g_rsDescargas("CdeEs")
   g_intCantidad     = g_rsDescargas("Cantidad")
   g_intSolicitudNro = g_rsDescargas("SolicitudNro")
   g_intAnalisisGdo  = g_rsDescargas("AnalisisGdo")
   g_chrMrcConforme  = g_rsDescargas("MrcConforme")
   'if (g_flgAgrupar <> true) then
   '     g_intPlanillaSec  = g_rsDescargas("PlanillaSec")
   'end if
   g_rsDescargas.movenext
   getNextDescarga=1
end if
End Function
'***************************************************************************************
function Gf_reset_Descargas()
   g_intProducto     = ""
   g_intSucursal     = ""
   g_intOperacion    = ""
   g_intNumero       = ""
   g_intCosecha      = ""
   g_intFechaDescarga= ""
   g_intKCCOR        = ""
   g_intKCVEN        = ""
   g_intKilos        = ""
   g_strCtoCorredor  = ""
   g_chrMrcConforme  = ""
   g_intPuerto       = ""
   g_intPlanillaCos  = ""
   g_intPlanillaNro  = ""
   g_intPlanillaSec  = ""
   g_intCartaPorte   = ""
   g_intReciboNro    = ""
   g_intSolicitudNro = ""
   g_intAnalisisGdo  = ""
   g_intCdeEs        = ""
end function
'/**
' * Funcion: initHeaderAnalisis
' * Descripcion: Esta funcion inicializa la lectura de la
' *              cabecera de los analisis.
' * Parametros:
' * Valor Devuelto:
' * Si hay al menos un registro valido devuelve 1 sino 0.
' *
' * Autor: Eugenio Di Santo
' * Fecha: 12/08/2004
' */
Function initHeaderAnalisis()
    Dim oConn, strSQL, fechaControl
    fechaControl = GF_DTEADD(mid(session("MmtoSistema"),1,8),-1,"a")
    fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,5,2)-1)*-1,"M")
    fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,7,2)-1)*-1,"D")
    strSQL= "Select CPORCA as CartaPorte, NSANCA as SolicitudNro, CPROCA as Producto, NROACA as Numero, FANACA as Fecha, KGMOCA as Kilos, COBECA as Bolsa, GRASCA as Grado, IMPACA as Costo "
    strSQL= strSQL & "from MERFL.MER591CA where FANACA > '" & fechaControl &  "' and CPORCA=" & g_intCartaPorte & " and NSANCA=" & g_intSolicitudNro

    'response.write strSQL
    Call GF_BD_AS400_2(g_rsAnalisis,oConn,"OPEN",strSQL)
    if (g_rsAnalisis.eof) then
       initHeaderAnalisis = 0
    else
       initHeaderAnalisis = 1
    end if
End Function
'/**
' * Funcion: getNextAnalisis
' * Descripcion: Esta funcion lee la siguiente linea en los
' *              resultados obtenidos por initHeaderAnalisis.
' * Parametros:
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Eugenio Dio Santo
' * Fecha: 12/08/2004
' */
Function getNextAnalisis()
if (g_rsAnalisis.eof) then
   getNextAnalisis=0
else
   g_intProducto        = g_rsAnalisis("Producto")
   g_intNumeroAnalisis  = g_rsAnalisis("Numero")
   g_intFechaAnalisis   = g_rsAnalisis("Fecha")
   g_intKilos           = g_rsAnalisis("Kilos")
   g_intBolsa           = g_rsAnalisis("Bolsa")
   g_intAnalisisGdo     = g_rsAnalisis("Grado")
   g_intCosto           = g_rsAnalisis("Costo")
   g_rsAnalisis.movenext
   getNextAnalisis=1
end if
End Function
'/**
' * Funcion: initHeaderDetAnalisis
' * Descripcion: Esta funcion inicializa la lectura de los
' *              detalles de un analisis.
' * Parametros:
' * Valor Devuelto:
' * Si hay al menos un registro valido devuelve 1 sino 0.
' *
' * Autor: Eugenio Di Santo
' * Fecha: 12/08/2004
' */
Function initHeaderDetAnalisis()
    Dim oConn, strSQL

    strSQL= "Select DA.COBEDA as Bolsa, DA.CPRODA as Producto, DA.NROADA as Numero, DA.FANADA as Fecha, CO.DESCAN as descconc, DA.COANDA as concepto, DA.VACADA as valor, DA.PREBDA as PjeRebaja, DA.PBONDA as PjeBonificacion"
    strSQL= strSQL & " from MERFL.MER591DA DA inner join MERFL.MER2E2F1 CO on DA.COANDA=CO.CONCAN"
    strSQL= strSQL & " where DA.COBEDA=" & g_intBolsa & " and DA.CPRODA=" & g_intProducto & " and DA.NROADA=" & g_intNumeroAnalisis & " and DA.FANADA=" & g_intFechaAnalisis
    if g_intConcepto <> "" then strSQl = strSQL & " and DA.COANDA=" & g_intConcepto
    'response.write strSQL
    Call GF_BD_AS400_2(g_rsDetAnalisis,oConn,"OPEN",strSQL)
    if (g_rsDetAnalisis.eof) then
       initHeaderDetAnalisis = 0
    else
       initHeaderDetAnalisis = 1
    end if
End Function
'/**
' * Funcion: getNextDetAnalisis
' * Descripcion: Esta funcion lee la siguiente linea en los
' *              resultados obtenidos por initHeaderDetAnalisis.
' * Parametros:
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Eugenio Dio Santo
' * Fecha: 12/08/2004
' */
Function getNextDetAnalisis()
if (g_rsDetAnalisis.eof) then
   getNextDetAnalisis=0
else
   g_strConceptoDs   = g_rsDetAnalisis("descconc")
   g_intConcepto     = g_rsDetAnalisis("concepto")
   g_intValor        = g_rsDetAnalisis("valor")
   g_intRebaja       = g_rsDetAnalisis("PjeRebaja")
   g_intBonif        = g_rsDetAnalisis("PjeBonificacion")
   g_rsDetAnalisis.movenext
   getNextDetAnalisis=1
end if
End Function
'/**
' * Funcion: GF_Contrato_Generar
' * Descripcion: Esta funcion genera el pdf correspondiente al
' *              contrato pasado por parametro.
' * Parametros:  p_intProducto  [in] 2 digitos (producto)
' *              p_intSucursal  [in] 1 digito  (sucursal)
' *              p_intOperacion [in] 2 digitos (operacion)
' *              p_intNumero    [in] 5 digitos (secuencia)
' *              p_intCosecha   [in] 2 digitos (cosecha)
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Eugenio Di Santo
' * Fecha: 27/08/2004
' */
Function GF_Contrato_Generar(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha)
   dim accion, imprimioCaratula
   accion = GF_Parametros7("ACCION","",6)

   Call GF_reset_Contrato()

   g_intProducto = p_intProducto
   g_intSucursal = p_intSucursal
   g_intOperacion = p_intOperacion
   g_intNumero = p_intNumero
   g_intCosecha = p_intCosecha
   GF_Contrato_Generar = False
    if accion="todos" or accion="contratos" then
        call GF_Print_Contrato(imprimioCaratula)
    end if   
    if accion="todos" or accion="descargas" then
      g_strCampoOrden = ""
      call GF_Print_ListadoDescargas(imprimioCaratula)
      g_strCampoOrden = ""
      call GF_reset_Descargas()
    end if
    if (imprimioCaratula = true) then GF_Contrato_Generar = True
end function
'/**
' * Funcion: GF_Print_Contrato()
' * Descripcion: Esta funcion imprime en el pdf solo los
' *              datos del contrato almacenado en las var
' *              globales
' * Parametros:
' * Valor Devuelto
' * Mientras haya resultados validos devuelve 1 sino 0.
' *
' * Autor: Eugenio Di Santo
' * Fecha: 30/08/2004
' */
Function GF_Print_Contrato(byref p_impCaratula)

   p_impCaratula = false
   if initHeader=1 then
       call getNextHeader()
       p_impCaratula = true
       Call GF_imprimirMembrete()
       Call GF_imprimirTitulo(100,"Contrato")
       'Se dibuja el recuadro
       Call GF_squareBox(Gbl_oPDF, 5, 140, 585, 700, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)

       call GF_Print_Header_Contrato(165)
       Call GF_H_SEPARADOR(Gbl_oPDF, 5, 210, 585)
       Call mostrarPartesInvolucradas(230)
       Call GF_H_SEPARADOR(Gbl_oPDF, 5, 300, 585)
       Call mostrarDatosMercaderias(320)
       Call GF_H_SEPARADOR(Gbl_oPDF, 5, 465, 585)
       Call mostrarDatosFijacion(485)
       Call GF_H_SEPARADOR(Gbl_oPDF, 5, 590, 585)
       Call mostrarDatosEntrega(610)
       Call GF_H_SEPARADOR(Gbl_oPDF, 5, 715, 585)
       Call mostrarPago(735)
   end if
end function   
'-------------------------------------------------------------------------------
Function GF_Print_Header_Contrato(p_y)
	dim strDs
	
         'Se escribe la informacion de la cabecera
         Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
         'Imprimo Nro. Contrato Toepfer
         Call GF_writeText(Gbl_oPDF,15, p_y, GF_Traducir("Contrato"), 0)
         Call GF_writeText(Gbl_oPDF,105, p_y, ":",0)
         Call GF_writeText(Gbl_oPDF,115, p_y, GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha),0)
         'Imprimo Fecha Concertacion
         Call GF_writeText(Gbl_oPDF,280, p_y, GF_Traducir("Fecha Concertacion"), 0)
         Call GF_writeText(Gbl_oPDF,415, p_y, ":", 0)
         Call GF_writeText(Gbl_oPDF,425, p_y, GF_FN2DTE(g_intFechaConc),0)

         p_y = p_y + 15
         'Imprimo Cto Corredor
         Call GF_writeText(Gbl_oPDF,15, p_y, GF_Traducir("Cto Corredor"), 0)
         Call GF_writeText(Gbl_oPDF,105, p_y, ":",0)
         Call GF_writeText(Gbl_oPDF,115, p_y, g_strCtoCorredor,0)

         'Imprimo Tipo Operacion
         Call GF_writeText(Gbl_oPDF,280, p_y, GF_Traducir("Operacion"), 0)
	 Call GF_writeText(Gbl_oPDF,350, p_y, ":", 0)
	 Call GF_MGC("MO",GF_nDigits(g_intOperacion,2),"",strDS)
	 Call GF_writeText(Gbl_oPDF,360,p_y, GF_Traducir(strDs),0)
End Function
'----------------------------------------------------------------------
function mostrarPartesInvolucradas(p_y)
	dim strDs
	
   'Subtitulo
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
   Call GF_writeText(Gbl_oPDF,15,p_y,GF_Traducir("Partes Involucradas"),0)
   
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 0)
   p_y = p_y + 30
   'Comprador
   Call GF_writeText(Gbl_oPDF,100,p_y,GF_Traducir("Comprador"),0)
   Call GF_writeText(Gbl_oPDF,170,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,180,p_y,GetDSEnterprise2(g_intKCCOR),0)
   p_y = p_y + 15
   'Vendedor
   Call GF_writeText(Gbl_oPDF,100,p_y,GF_Traducir("Vendedor"),0)
   Call GF_writeText(Gbl_oPDF,170,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,180,p_y,GetDSEnterprise2(g_intKCVEN),0)
end function
'----------------------------------------------------------------------
function mostrarDatosMercaderias(p_y)
	dim strDs
	dim unitDestino, intAnulaciones
	dim retValue
	
	unitDestino = GF_Parametros7("unidadDestino","",6)
	if unitDestino = "" then unitDestino = "1"
	
	'Subtitulo
	Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
        Call GF_writeText(Gbl_oPDF,15,p_y,GF_Traducir("Mercaderia"),0)
    
        Call GF_setFont(Gbl_oPDF,"Courier", 12, 0)
        p_y = p_y + 30
        'Producto
	Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Producto"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
	Call GF_MGC("AR",g_intProducto,0,strDs)
 	Call GF_writeText(Gbl_oPDF,160,p_y,GF_Traducir(left(strDs,23)),0)
        p_y = p_y + 15
        'Cosecha
	Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Cosecha"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
    if cint(g_intCosecha) > 95 then
        intcosecha = cInt(g_intCosecha) + 1900
    else
        intcosecha = cInt(g_intCosecha) + 2000
    end if
    Call GF_writeText(Gbl_oPDF,160,p_y,intCosecha,0)
        p_y = p_y + 15
        'Cantidad Contratada
	Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Contratada"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
	retValue = g_intKilos
	if (unitDestino = UNIDAD_TONELADAS) then retValue = g_intKilos/1000
	Call GF_writeText(Gbl_oPDF,160,p_y,retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDestino),0)
        p_y = p_y + 15
        'Cantidad Entregada
    Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Entregada"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
	retValue = g_intKgEntregados
	if (unitDestino = UNIDAD_TONELADAS) then retValue = g_intKgEntregados/1000	
	Call GF_writeText(Gbl_oPDF,160,p_y,retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDestino),0)
        p_y = p_y + 15
        'Procedencia
	Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Procedencia"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
	if len(strDs)>=20 then Call GF_setFont(Gbl_oPDF,"Courier", 10, 0)
	Call GF_writeText(Gbl_oPDF,160,p_y,left(g_strProcedencia,23),0)
        p_y = p_y + 15
        'Mercaderia Propia
	Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Merc. Propia"),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
	if g_chrMercPropia = "V" then
		strDs = "Si"
	else
		strDs = "No"
	end if
	Call GF_writeText(Gbl_oPDF,160,p_y,GF_Traducir(strDs),0)
        p_y = p_y + 15
        'Anulaciones
        strDs = "Ampliaciones"
        intAnulaciones = cDbl(g_intAnulaciones)
        if (intAnulaciones < 0) then
           strDs = "Anulaciones"
           intAnulaciones = intAnulaciones * -1
        end if
        Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir(strDs),0)
	Call GF_writeText(Gbl_oPDF,150,p_y,":",0)	
	retValue = intAnulaciones
	if (unitDestino = UNIDAD_TONELADAS) then retValue = intAnulaciones/1000			
	Call GF_writeText(Gbl_oPDF,160,p_y,retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDestino),0)
end function
'-------------------------------------------------------------------------------------
function mostrarDatosFijacion(p_y)
   dim unitDest
   dim retValue
   
   unitDest = GF_Parametros7("unidadDestino","",6)
   if unitDest = "" then unitDest = "1"
   'Subtitulo
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
   Call GF_writeText(Gbl_oPDF,15,p_y,GF_Traducir("Fijacion"),0)
   
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 0)
   p_y = p_y + 30
   if len(g_intFechaFijaDesde) = 8 then
      'Fecha Fija Desde
      Call GF_writeText(Gbl_oPDF,80,p_y,GF_Traducir("Desde"),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      Call GF_writeText(Gbl_oPDF,160,p_y,GF_FN2DTE(g_intFechaFijaDesde),0)
      
      p_y = p_y + 15
      'Fecha Fija Hasta
      Call GF_writeText(Gbl_oPDF,80,p_y,GF_Traducir("Hasta"),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      Call GF_writeText(Gbl_oPDF,160,p_y,GF_FN2DTE(g_intFechaFijaHasta),0)
      
      p_y = p_y + 15
      'Cantidad Minima
      Call GF_writeText(Gbl_oPDF,80,p_y,GF_Traducir("Cant(Min)"),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      
      retValue = g_intKilosMin
	  if (unitDest = UNIDAD_TONELADAS) then retValue = g_intKilosMin/1000
            
      Call GF_writeText(Gbl_oPDF,160,p_y,retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest),0)
      
      p_y = p_y + 15
      'Cantidad Maxima
      Call GF_writeText(Gbl_oPDF,80,p_y,GF_Traducir("Cant(Max)"),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      
      retValue = g_intKilosMax
	  if (unitDest = UNIDAD_TONELADAS) then retValue = g_intKilosMax/1000
	        
      Call GF_writeText(Gbl_oPDF,160,p_y,retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest),0)
   else
       p_y = p_y + 15
       Call GF_writeText(Gbl_oPDF,30,p_y,GF_Traducir("A este contrato no se le aplica fijacion") & ".",0)
   end if
end function
'-------------------------------------------------------------------------------------
function mostrarDatosEntrega(p_y)
   dim unitDest
   dim retValue
   dim strDs
   
   unitDest = GF_Parametros7("unidadDestino","",6)
   if unitDest = "" then unitDest = "1"
   'Subtitulo
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
   Call GF_writeText(Gbl_oPDF,15,p_y,GF_Traducir("Entrega"),0)
   
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 0)
   p_y = p_y + 30
   'Desde
   Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Desde"),0)
   Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,160,p_y,GF_FN2DTE(g_intFechaEntDesde),0)

   p_y = p_y + 15
   'Fecha Fija Hasta
   Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Hasta"),0)
   Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,160,p_y,GF_FN2DTE(g_intFechaEntHasta),0)
   

   'Puerto Entrega
   if cint(g_intPuertoRecepcion) > 0 then
      p_y = p_y + 15
      Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Puerto Entrega"),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      Call GF_MGC("PU",g_intPuertoRecepcion,0,strDs)
      'if len(strDs)>=20 then Call GF_setFont(Gbl_oPDF,"Courier", 10, 0)
      Call GF_writeText(Gbl_oPDF,160,p_y,left(strDs,23),0)
   end if
   
   'Puerto Devolucion
   if cint(g_intPuertoDevolucion) > 0 then
      p_y = p_y + 15
      Call GF_writeText(Gbl_oPDF,50,p_y,GF_Traducir("Puerto Devoluc."),0)
      Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
      Call GF_MGC("PU",g_intPuertoDevolucion,0,strDs)
      'if len(strDs)>=20 then Call GF_setFont(Gbl_oPDF,"Courier", 10, 0)
      Call GF_writeText(Gbl_oPDF,160,p_y,left(strDs,23),0)
   end if
end function
'---------------------------------------------------------------------------------------
function mostrarPago(p_y)

   'Subtitulo
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
   Call GF_writeText(Gbl_oPDF,15,p_y,GF_Traducir("Pago"),0)
   
   Call GF_setFont(Gbl_oPDF,"Courier", 12, 0)
   p_y = p_y + 30
   'Precio
   Call GF_writeText(Gbl_oPDF,55,p_y,GF_Traducir("Precio"),0)
   Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,160,p_y,"$ " & GF_EDIT_DECIMALS(cdbl(g_intPrecioP)*100, 2),0)
   p_y = p_y + 15
   'Parcial
   Call GF_writeText(Gbl_oPDF,55,p_y,GF_Traducir("Parcial"),0)
   Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
   Call GF_writeText(Gbl_oPDF,160,p_y,GF_EDIT_DECIMALS(cdbl(g_intPjeParcial)*100,2) & "%",0)
   p_y = p_y + 15
   'Forma de Pago
   Call GF_writeText(Gbl_oPDF,55,p_y,GF_Traducir("Forma de Pago"),0)
   Call GF_writeText(Gbl_oPDF,150,p_y,":",0)
   if len(strDs)>=20 then Call GF_setFont(Gbl_oPDF,"Courier", 10, 0)
   Call GF_writeText(Gbl_oPDF,160,p_y,GF_Traducir(left(g_strKCPago,23)),0)
end function
'----------------------------------------------------------------------------------------
function GF_imprimirMembrete()
   Call GF_squareBox(Gbl_oPDF, 5, 10, 585, 75, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
   'Se coloca el Logo
   Call GF_writeImage(Gbl_oPDF, Server.MapPath("..\Images\ACTILogoHBG.jpg"), 8, 15, 570, 60, 0)
end function
'----------------------------------------------------------------------------------------
Function GF_imprimirTitulo(p_y, p_titulo)
    Call GF_squareBox(Gbl_oPDF, 5, p_y, 585, 30, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
    call GF_setFont(Gbl_oPDF, "Courier", 14, 0)
    Call GF_writeTextAlign(Gbl_oPDF,0,p_y + 10,p_titulo,595,2)
end function
'----------------------------------------------------------------------------------------
Function GF_imprimirTituloApaisado(p_x, p_titulo, p_pag, p_totalPaginas)
    dim tit
    Call GF_squareBox(Gbl_oPDF, p_x, 5, 60, 840, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
    call GF_setFont(Gbl_oPDF, "Courier", 14, 8)
	call GF_writeVerticalText(Gbl_oPDF, p_x + 10, 840, GF_Traducir(p_titulo),840,2)
    call GF_setFont(Gbl_oPDF, "Courier", 12, 0)    
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 800, GF_Traducir("Contrato Nro"),800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 710, ":",800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 700, GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha),800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 45, 800, GF_Traducir("Cto Corredor"),800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 45, 710, ":",800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 45, 700, g_strCtoCorredor,800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 150, GF_Traducir("Pagina"),800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 100, ":",800,0)
    call GF_writeVerticalText(Gbl_oPDF, p_x + 30, 90, p_pag & "/" & p_totalPaginas,800,0)
    
end function
'----------------------------------------------------------------------------------------
Function GF_Print_Descarga()
    call GF_imprimirMembrete()
    call GF_imprimirTitulo(100,"Descarga")
    
    Call GF_squareBox(Gbl_oPDF, 5, 145, 585, 280, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
    Call GF_Print_Header_Descarga(160)
    Call GF_H_SEPARADOR(Gbl_oPDF, 5, 190, 585)
    Call mostrarPartesInvolucradas(220)
    Call GF_H_SEPARADOR(Gbl_oPDF, 5, 290, 585)
    call mostrarDetallesPlanilla(320)
    
    Call GF_Print_Analisis()
end Function
'----------------------------------------------------------------------------------------
function GF_Print_Header_Descarga(p_y)
    call GF_setFont(Gbl_oPDF, "Courier", 12, 8)

    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Contrato"),0)
    call GF_writeText(Gbl_oPDF, 140, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 150, p_y, g_intProducto & "-" & g_intSucursal & "-" & g_intOperacion & "-" & g_intNumero & "/" & g_intCosecha,0)

    call GF_writeText(Gbl_oPDF, 315, p_y, GF_Traducir("Fecha de Descarga"),0)
    call GF_writeText(Gbl_oPDF, 450, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 460, p_y, GF_FN2DTE(g_intFechaDescarga),0)
    
    p_y = p_y + 15
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Contrato Corredor"),0)
    call GF_writeText(Gbl_oPDF, 140, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 150, p_y, g_strCtoCorredor,0)

    call GF_writeText(Gbl_oPDF, 315, p_y, GF_Traducir("Carta de Porte"),0)
    call GF_writeText(Gbl_oPDF, 450, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 460, p_y, g_intCartaPorte,0)

end function
'----------------------------------------------------------------------------------------
Function mostrarDetallesPlanilla(p_y)
    dim strAux
    dim unitDest, retValue
    
    unitDest = GF_Parametros7("UnidadDestino","",6)
    if unitDest = "" then unitDest = "1"
    
    Call GF_setFont(Gbl_oPDF, "Courier", 12, 8)
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Detalles de la Planilla"),0)

    Call GF_setFont(Gbl_oPDF, "Courier", 12, 0)
    p_y = p_y + 30
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Nro de Recibo/Romaneo"),0)
    call GF_writeText(Gbl_oPDF, 170, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 180, p_y, g_intReciboNro,0)
    
    call GF_writeText(Gbl_oPDF, 315, p_y, GF_Traducir("Tipo Movimiento"),0)
    call GF_writeText(Gbl_oPDF, 470, p_y, ":",0)
    if (ucase(g_intCdeEs)="E") then
         strAux = "Salida"
    elseif (ucase(g_intCdeEs)="I") then
         strAux = "Entrada"
    end if
    call GF_writeText(Gbl_oPDF, 480, p_y, GF_Traducir(strAux),0)

    p_y = p_y + 15
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Mercaderia Conforme"),0)
    call GF_writeText(Gbl_oPDF, 170, p_y, ":",0)
    strAux = "Indefinido"
    if g_CHRMrcConforme="V" then
      strAux = "Si"
    elseif g_CHRMrcConforme="F" then
      strAux = "No"
    end if
    call GF_writeText(Gbl_oPDF, 180, p_y, GF_Traducir(strAux),0)
    
    call GF_writeText(Gbl_oPDF, 315, p_y, GF_Traducir("Solicitud de Analisis"),0)
    call GF_writeText(Gbl_oPDF, 470, p_y, ":",0)
    call GF_writeText(Gbl_oPDF, 480, p_y, g_intSolicitudNro,0)
    
    p_y = p_y + 15
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Cantidad Descargada"),0)
    call GF_writeText(Gbl_oPDF, 170, p_y, ":",0)
    
    retValue = g_intKilos
    if (unitDest = UNIDAD_TONELADAS) then retValue = g_intKilos/1000
	    
    call GF_writeText(Gbl_oPDF, 180, p_y, retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest),0)
    
    p_y = p_y + 15
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Puerto"),0)
    call GF_writeText(Gbl_oPDF, 170, p_y, ":",0)
    call GF_MGC("PU",g_intPuerto,0,strAux)
    call GF_writeText(Gbl_oPDF, 180, p_y, strAux,0)
end Function
'----------------------------------------------------------------------------------------
Function GF_Print_Analisis()

    call GF_imprimirMembrete()
	call GF_imprimirTitulo(440,"Analisis")

    Call GF_squareBox(Gbl_oPDF, 5, 480, 585, 363, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
    if initHeaderAnalisis()=1 then
        Call getNextAnalisis()
        Call GF_Print_Header_Analisis(500)
        Call GF_H_SEPARADOR(Gbl_oPDF, 5, 585, 585)
        Call GF_Print_DetallesAnalisis(605)
    else
    end if
    'Call GF_Print_Analisis()
end Function
'----------------------------------------------------------------------------------------
Function GF_Print_Header_Analisis(p_y)
    dim strAux
    dim retValue
    dim unitDest
    
    unitDest = GF_Parametros7("UnidadDestino","",6)
    if unitDest = "" then unitDest = "1"
    
    Call GF_setFont(Gbl_oPDF,"Courier",12,0)
    
    Call GF_writeText(Gbl_oPDF,15,p_y, GF_Traducir("Nro. de Analisis"),0)
    Call GF_writeText(Gbl_oPDF,150,p_y, ":",0)
    Call GF_writeText(Gbl_oPDF,160,p_y, g_intNumeroAnalisis,0)
    
    Call GF_writeText(Gbl_oPDF,315,p_y, GF_Traducir("Fecha de Analisis"),0)
    Call GF_writeText(Gbl_oPDF,450,p_y, ":",0)
    Call GF_writeText(Gbl_oPDF,460,p_y, GF_FN2DTE(g_intFechaAnalisis),0)
    
    p_y = p_y + 15
    Call GF_writeText(Gbl_oPDF,15,p_y, GF_Traducir("Grado del Analisis"),0)
    Call GF_writeText(Gbl_oPDF,150,p_y, ":",0)
    Call GF_writeText(Gbl_oPDF,160,p_y, g_intAnalisisGdo,0)

    Call GF_writeText(Gbl_oPDF,315,p_y, GF_Traducir("Costo del Analsis"),0)
    Call GF_writeText(Gbl_oPDF,450,p_y, ":",0)
    Call GF_writeText(Gbl_oPDF,460,p_y, GF_EDIT_DECIMALS(clng(g_intCosto), 2),0)

    p_y = p_y + 15
    Call GF_writeText(Gbl_oPDF,15,p_y, GF_Traducir("Producto"),0)
    Call GF_writeText(Gbl_oPDF,150,p_y, ":",0)
    call GF_MGC("AR",g_intProducto,0,strAux)
    Call GF_writeText(Gbl_oPDF,160,p_y, GF_Traducir(strAux),0)

    p_y = p_y + 15
    Call GF_writeText(Gbl_oPDF,15,p_y, GF_Traducir("Cantidad Analizada"),0)
    Call GF_writeText(Gbl_oPDF,150,p_y, ":",0)
    
    retValue = g_intKilos
    if (unitDest = UNIDAD_TONELADAS) then retValue = g_intKilos/1000
        
    Call GF_writeText(Gbl_oPDF,160,p_y, retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest),0)

    p_y = p_y + 15
    Call GF_writeText(Gbl_oPDF,15,p_y, GF_Traducir("Entidad"),0)
    Call GF_writeText(Gbl_oPDF, 150,p_y, ":",0)
    Call GF_MGC("ME", g_intBolsa, 0, strAux)
    Call GF_writeText(Gbl_oPDF,160,p_y, GF_Traducir(strAux),0)
    
end Function
'----------------------------------------------------------------------------------------
Function GF_Print_DetallesAnalisis(p_y)
    dim cantDetalles

    call GF_setFont(Gbl_oPDF, "Courier", 12, 8)
    call GF_writeText(Gbl_oPDF, 15, p_y, GF_Traducir("Detalles del Analisis"),0)
    
    'armo la tabla
    p_y = p_y + 30
    Call GF_squareBox(Gbl_oPDF, 15, p_y, 565, 187, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
    
    call GF_H_SEPARADOR(Gbl_oPDF, 24,p_y + 26,559)

    call GF_V_SEPARADOR(Gbl_oPDF, 255,p_y + 5, 175)
    call GF_V_SEPARADOR(Gbl_oPDF, 358,p_y + 5, 175)
    call GF_V_SEPARADOR(Gbl_oPDF, 461,p_y + 5, 175)

    p_y = p_y + 10
    call GF_writeTextAlign(Gbl_oPDF, 15,p_y,GF_Traducir("Concepto"),240,2)
    call GF_writeTextAlign(Gbl_oPDF, 258,p_y,GF_Traducir("Valor"),100,2)
    call GF_writeTextAlign(Gbl_oPDF, 361,p_y,GF_Traducir("Rebaja"),100,2)
    call GF_writeTextAlign(Gbl_oPDF, 464,p_y,GF_Traducir("Bonificacion"),120,2)
    
    p_y = p_y + 25
    call GF_setFont(Gbl_oPDF, "Courier", 12, 0)

    if initHeaderDetAnalisis()=1 then
        cantDetalles = 0
        while getNextDetAnalisis()= 1 and cantDetalles <= 8
            call GF_writeTextAlign(Gbl_oPDF, 15,p_y, g_intConcepto & " - "& GF_Traducir(g_strConceptoDs),240,2)
            call GF_writeTextAlign(Gbl_oPDF, 278,p_y,GF_EDIT_DECIMALS(cLng(g_intValor), 2) & "%",60,1)
            call GF_writeTextAlign(Gbl_oPDF, 381,p_y,GF_EDIT_DECIMALS(cLng(g_intRebaja), 2) & "%",60,1)
            call GF_writeTextAlign(Gbl_oPDF, 490,p_y,GF_EDIT_DECIMALS(cLng(g_intBonif), 2) & "%",60,1)
            p_y = p_y + 15
            cantDetalles = cantDetalles + 1
        wend
    end if
end Function
'------------------------------------------------------------------------------------------
Function GF_Print_ListadoDescargas(byref p_impCaratula)
    dim auxDs
    dim retValue, unitDestino
    dim p_x, pag, totalPaginas
    
    unitDest = GF_Parametros7("UnidadDestino","",6)
    if unitDest = "" then unitDest = "1"    
    if initDescarga()=1 then
        p_x = 596
        pag = 0
        while getNextDescarga=1        
            if p_x > 595 then
                pag = pag + 1
                totalPaginas = int(g_rsDescargas.recordcount / 26) + 1
               
                if (p_impCaratula = true) then
                    call GF_newPage(Gbl_oPDF)                                
                else
                    p_impCaratula = true
                end if
                
                call GF_imprimirTituloApaisado(5,"Descargas", pag, totalPaginas)
                
                Call GF_squareBox(Gbl_oPDF, 80, 5, 510, 840, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
                call GF_V_SEPARADOR(Gbl_oPDF, 105, 5, 840)
                call GF_setFont(Gbl_oPDF, "Courier",12,8)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 753, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 839,GF_Traducir("Fecha"),80,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 590, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 757,GF_Traducir("Puerto"),170,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 500, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 580,GF_Traducir("C. Porte"),70,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 410, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 510,GF_Traducir("Cantidad"),100,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 320, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 400,GF_Traducir("Romaneo"),70,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 245, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 310,GF_Traducir("Movim."),60,2)
                call GF_H_SEPARADOR(Gbl_oPDF, 80, 127, 510)
                call GF_writeVerticalText(Gbl_oPDF, 90, 240,GF_Traducir("Merc. Conforme"),100,2)
                call GF_writeVerticalText(Gbl_oPDF, 90, 127,GF_Traducir("Solic. Analisis"),127,2)
                call GF_setFont(Gbl_oPDF, "Courier",12,0)
                p_x = 115
            end if
            		
            call GF_writeVerticalText(Gbl_oPDF, p_x,839,GF_FN2DTE(g_intFechaDescarga),80,2)
            call GF_MGC("PU",g_intPuerto,0,auxDs)
            call GF_writeVerticalText(Gbl_oPDF, p_x,757,auxDs,170,2)
            call GF_writeVerticalText(Gbl_oPDF, p_x,580,g_intCartaPorte,70,2)
            
            retValue = g_intKilosDescarga
			if (unitDest = UNIDAD_TONELADAS) then retValue = g_intKilosDescarga/1000
                
            call GF_writeVerticalText(Gbl_oPDF, p_x, 500, retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest),80,1)
            call GF_writeVerticalText(Gbl_oPDF, p_x, 400, g_intReciboNro,70,2)
            if (ucase(g_intCdeEs)="E") then
                 auxDs = GF_Traducir("Salida")
            elseif (ucase(g_intCdeEs)="I") then
                 auxDs = GF_Traducir("Entrada")
            end if
            call GF_writeVerticalText(Gbl_oPDF, p_x, 310,auxDs,60,2)
            if g_CHRMrcConforme="V" then
              auxDs = GF_Traducir("Si")
            elseif g_CHRMrcConforme="F" then
              auxDs = GF_Traducir("No")
            end if
            call GF_writeVerticalText(Gbl_oPDF, p_x, 240,auxDs,100,2)
            call GF_writeVerticalText(Gbl_oPDF, p_x, 127,g_intSolicitudNro  ,127,2)
            p_x = p_x + 15
        wend
    else
        response.write g_intProducto & "-" & g_intSucursal & "-" & g_intOperacion & "-" & g_intNumero & "/"& g_intCosecha & "<br>"        
    end if    
end function
'------------------------------------------------------------------------------------------
Function getDescripcionOperacion(operacion)
	Dim strSQL, conn, rs
	
	strSQL="Select * from MERFL.MER132F1 where CODIOM=" & operacion
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	
	if (rs.eof) then
		getDescripcionOperacion = "ERROR"
	else
		getDescripcionOperacion = rs("DESCOM")
	end if
End Function
'------------------------------------------------------------------------------------------
Function getDescripcionProducto(codigo)
	strSQL="Select * from MERFL.MER112F1 where CODIPR=" & codigo
    call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    if (rs.eof) then
		getDescripcionProducto = "ERROR"
	else
		getDescripcionProducto = rs("DESCPR")
	end if
End Function
'------------------------------------------------------------------------------------------
Function getDescripcionTransporte(codigo)
	strSQL="Select * from MERFL.MER182F1 where CODITR=" & codigo
    call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    if (rs.eof) then
		getDescripcionTransporte = "ERROR"
	else
		getDescripcionTransporte = rs("DESCTR")
	end if
End Function
'------------------------------------------------------------------------------------------
Function getDescripcionDestino(codigo)
	strSQL="Select * from MERFL.MER192F1 where CODIDE=" & codigo
    call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    if (rs.eof) then
		getDescripcionDestino = "ERROR"
	else
		getDescripcionDestino = rs("DESCDE")
	end if
End Function
'------------------------------------------------------------------------------------------
'Descripcion de las entidades - Bolsa de Cereales
Function getDescripcionEntidad(codigo)
	strSQL="Select * from MERFL.MER2A2F1 where CODIBE=" & codigo
    call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    if (rs.eof) then
		getDescripcionEntidad = "ERROR"
	else
		getDescripcionEntidad = rs("DESCBE")
	end if
End Function
%>
