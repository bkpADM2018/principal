<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<%
'-------------------------------------------------------------------------------------------------------------------------
'Controla que la cantidad de cupos ingresadas cubran correctamente el rango establecido en el contrato
Function controlarCantidadCupos(p_ArrCupo, p_ArrFecha, p_NumeroPto, p_CdProveedor, p_CdProducto)
    Dim i,strSQL,rsCant,auxCantidad,rsNom
    flagControl = true
    i = 0
    logMig.info(" Iniciando control")
    while ((i <= UBound(p_ArrCupo))and(flagControl))
        if (CInt(p_ArrCupo(i)) > 0) then
           strSQL = "SELECT CASE WHEN SUM(CUCCCP) IS NULL THEN 0 ELSE SUM(CUCCCP) END AS CANTIDAD "&_
                    "FROM ( SELECT CUCODI,CUCPRO, CUCSUC, CUCOPE, CUNCTO, CUACOS,  CUFCCP, CUCDES,CUCCCP "&_
                    "       FROM   MERFL.MER517F1  "&_
                    "       WHERE  CUFCCP = "& p_ArrFecha(i) &" AND CUCDES = "& p_NumeroPto &" AND CUCOPE=04 AND CUCPRO=" & p_CdProducto & "  ) A  "&_
                    "INNER JOIN MERFL.MER311F1 B "&_
                    "   ON A.CUCPRO = B.CPROR1 AND A.CUCSUC = B.CSUCR1 AND A.CUCOPE = B.COPER1 AND A.CUNCTO = B.NCTOR1 AND A.CUACOS = B.ACOSR1 AND ( B.CVENR1 ="& p_CdProveedor &" OR B.CCORR1 = "& p_CdProveedor &")  "
           Call executeQuery(rsCant, "OPEN", strSQL)
           strSQL = "SELECT count(*) AS NOMINADOS "&_
                    "FROM  MERFL.TBLCUPOSNOMINADOS A "&_
                    "INNER JOIN MERFL.MER311F1 B "&_
                    "   ON A.IDPRODUCTO = B.CPROR1 AND A.IDSUCURSAL = B.CSUCR1 AND A.IDOPERACION = B.COPER1 AND A.NUMERO = B.NCTOR1 AND A.COSECHA = B.ACOSR1 AND ( B.CVENR1 ="& p_CdProveedor &" OR B.CCORR1 = "& p_CdProveedor &")  "&_
                    "WHERE FECHACUPO ="& p_ArrFecha(i) &" AND PUERTO="&p_NumeroPto &"   AND IDPRODUCTO = "&p_CdProducto
           Call executeQuery(rsNom, "OPEN", strSQL)
           auxCantidad = Cdbl(rsNom("NOMINADOS")) + Cdbl(p_ArrCupo(i))
           logMig.info(" ------> Fecha: "& GF_FN2DTE(p_ArrFecha(i)) )
           logMig.info(" --------------> Cantidad permitida: "&rsCant("CANTIDAD"))
           logMig.info(" --------------> Cantidad nominada: "& rsNom("NOMINADOS") )            
           logMig.info(" --------------> Cantidad a nominar : "&p_ArrCupo(i) )
           if (Cdbl(auxCantidad) > Cdbl(rsCant("CANTIDAD"))) then            
                flagControl = false
                logMig.info(" --------------> RESULTADO CONTROL : ERROR!")
           else
                logMig.info(" --------------> RESULTADO CONTROL : OK")
           end if
        end if
        i = i + 1
    wend    
    logMig.info(" Finalizando control")
    controlarCantidadCupos = flagControl
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function addNominacion(p_IdCupo,p_Fecha,p_NumeroPto,p_CdProducto,p_Sucursal,p_Operacion,p_Numero,p_Cosecha,p_CdCupo,p_CdCorredor,p_CdVendedor)
    logMig.info(" ----------------> "&p_CdCupo)
    Set sp_ret = executeSP(rs, "MERFL.TBLCUPOSNOMINADOS_INS", p_IdCupo &"||"& p_Fecha &"||"& p_NumeroPto &"||"& p_CdProducto &"||"& p_Sucursal &"||"& p_Operacion &"||"& p_Numero &"||"& p_Cosecha &"||"& p_CdCupo &"||"& p_CdCorredor &"||"& p_CdVendedor &"||"& Left(Session("Usuario"),10) &"||"& Session("MmtoDato"))
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function deleteNominacion(p_numeroPto, p_CdCorredor, p_CdVendedor, p_fechaDesde, p_fechaHasta, p_CdProducto, p_CdCupo)
    Dim strSQL ,rs
    logMig.info(" Cupo Liberado ")
    strSQL = "DELETE FROM MERFL.TBLCUPOSNOMINADOS WHERE PUERTO = "&p_numeroPto &_
             " AND FECHACUPO >= "& p_fechaDesde & " AND FECHACUPO <= "& p_fechaHasta &_
             " AND IDPRODUCTO=" & p_CdProducto
    logMig.info(" --------> Fecha desde: "& GF_FN2DTE(p_fechaDesde))
    logMig.info(" --------> Fecha hasta: "& GF_FN2DTE(p_fechaHasta))
    if (p_CdCorredor <> "") then 
        strSQL = strSQL & " AND IDCORREDOR = "&p_CdCorredor
        logMig.info(" --------> Corredor: "& p_CdCorredor)
    end if
    if (p_CdVendedor <> "") then 
        strSQL = strSQL & " AND IDVENDEDOR = "&p_CdVendedor    
        logMig.info(" --------> Vendedor: "& p_CdVendedor)
    end if
    if (p_CdCupo <> "") then
        strSQL = strSQL & " AND CODIGO = "& p_CdCupo
        logMig.info(" --------> Codigo de cupo: "& p_CdCupo)
    end if
    Call executeQuery(rs, "EXEC", strSQL)
End function 
'-------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------CORREDORES----------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'Administro y genero el codigo de corredor para los tres puertos
Function generarCodigoCorredor(p_cuitCorredor, p_dsCorredor, p_Pto, p_codigo)
    Dim auxCdCorredor
    
    'Compruebo para cada puerto si existe el cuit ingresado, si no existe lo agrego. Luego devuelvo el codigo del corredor para el puerto en cuestion (parametro p_Pto)
    'Arroyo:
    if (p_Pto = TERMINAL_ARROYO) and (p_codigo > 0) then
        generarCodigoCorredor = p_codigo
    else       
        'auxCdCorredor = getCdCorredorByCuit(TERMINAL_ARROYO, p_cuitCorredor)
        'if (Cdbl(auxCdCorredor) = 0) then 
        auxCdCorredor = p_codigo        
        if (p_codigo = 0) then
            auxCdCorredor = getNextCdCorredor(TERMINAL_ARROYO)
            Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_ARROYO)
        end if        
        if (p_Pto = TERMINAL_ARROYO) then generarCodigoCorredor = auxCdCorredor
    end if
    
    'Transito:
    if (p_Pto = TERMINAL_TRANSITO) and (p_codigo > 0) then
        generarCodigoCorredor = p_codigo
    else       
        'auxCdCorredor = getCdCorredorByCuit(TERMINAL_TRANSITO, p_cuitCorredor)
        'if (Cdbl(auxCdCorredor) = 0) then 
        auxCdCorredor = p_codigo
        if (p_codigo = 0) then
            auxCdCorredor = getNextCdCorredor(TERMINAL_TRANSITO)
            Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_TRANSITO)
        end if
        if (p_Pto = TERMINAL_TRANSITO) then generarCodigoCorredor = auxCdCorredor
    end if
    
    'Piedrabuena:
    if (p_Pto = TERMINAL_PIEDRABUENA) and (p_codigo > 0) then
        generarCodigoCorredor = p_codigo
    else       
        'auxCdCorredor = getCdCorredorByCuit(TERMINAL_PIEDRABUENA, p_cuitCorredor)
        'if (Cdbl(auxCdCorredor) = 0) then 
        auxCdCorredor = p_codigo
        if (p_codigo = 0) then
            auxCdCorredor = getNextCdCorredor(TERMINAL_PIEDRABUENA)
            Call addNewCorredor(auxCdCorredor, p_cuitCorredor, p_dsCorredor, TERMINAL_PIEDRABUENA)
        end if
        if (p_Pto = TERMINAL_PIEDRABUENA) then generarCodigoCorredor = auxCdCorredor
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Devuelve el codigo de corredor del puerto por su Cuit
'Function getCdCorredorByCuit(p_Pto,p_cuitCorredor)
'    Dim strSQL,rtrn
'    rtrn = 0
'    strSQL = "SELECT CDCORREDOR FROM CORREDORES WHERE RTRIM(NUCUIT) = '"& Trim(p_cuitCorredor) &"'"
'    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
'    if (not rs.Eof) then rtrn = rs("CDCORREDOR")
'    getCdCorredorByCuit = rtrn
'End function
'-------------------------------------------------------------------------------------------------------------------------
'Obtenemos el proximo Codigo de corredor libre para agregar
Function getNextCdCorredor(p_Pto)
    Dim strSQL 
    getNextCdCorredor = 100000
    strSQL = "Select NEXTID from "&_
             "(Select 99999 + ROW_NUMBER ( ) over (ORDER BY CDCORREDOR) NEXTID , *   from CORREDORES C where C.CDCORREDOR>=100000) T where NEXTID <>CDCORREDOR order by NEXTID"
    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then getNextCdCorredor = rs("NEXTID")
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Agrego un nuevo corredor a la planta
Function addNewCorredor(p_cdCorredor,p_cuitCorredor, p_dsCorredor, p_Pto)
    Dim strSQL
    strSQL = "INSERT INTO CORREDORES (CDCORREDOR, DSCORREDOR, CDTIPODOC, NUCUIT ) "&_
             "VALUES ("& p_cdCorredor &",'"& UCase(Trim(p_dsCorredor)) &"','C.U.I.T.','"& Trim(p_cuitCorredor) &"')"
    Call GF_BD_Puertos(p_Pto, rs, "EXEC", strSQL)
End function
'-------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------VENDEDORES----------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'Administro y genero el codigo de vendedor para los tres puertos
Function generarCodigoVendedor(p_cuitVendedor, p_dsVendedor, p_Pto, p_codigo)
    Dim auxCdVendedor
    
    'Compruebo para cada puerto si existe el cuit ingresado, si no existe lo agrego. Luego devuelvo el codigo del Vendedor para el puerto en cuestion (parametro p_Pto)
    'Arroyo:
    if (p_Pto = TERMINAL_ARROYO) and (p_codigo > 0) then
        generarCodigoVendedor = p_codigo
    else        
        'auxCdVendedor = getCdVendedorByCuit(TERMINAL_ARROYO, p_cuitVendedor)
        'if (Cdbl(auxCdVendedor) = 0) then 
        auxCdVendedor = p_codigo
        if (p_codigo = 0) then
            auxCdVendedor = getNextCdVendedor(TERMINAL_ARROYO)
            Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_ARROYO)
        end if
        if (p_Pto = TERMINAL_ARROYO) then generarCodigoVendedor = auxCdVendedor
    end if

    'Transito:
    if (p_Pto = TERMINAL_TRANSITO) and (p_codigo > 0) then
        generarCodigoVendedor = p_codigo
    else        
        'auxCdVendedor = getCdVendedorByCuit(TERMINAL_TRANSITO, p_cuitVendedor)
        'if (Cdbl(auxCdVendedor) = 0) then 
        auxCdVendedor = p_codigo
        if (p_codigo = 0) then
            auxCdVendedor = getNextCdVendedor(TERMINAL_TRANSITO)
            Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_TRANSITO)
        end if
        if (p_Pto = TERMINAL_TRANSITO) then generarCodigoVendedor = auxCdVendedor
    end if
    
    'Piedrabuena:
    if (p_Pto = TERMINAL_PIEDRABUENA) and (p_codigo > 0) then
        generarCodigoVendedor = p_codigo
    else        
        'auxCdVendedor = getCdVendedorByCuit(TERMINAL_PIEDRABUENA, p_cuitVendedor)
        'if (Cdbl(auxCdVendedor) = 0) then 
        auxCdVendedor = p_codigo
        if (p_codigo = 0) then
            auxCdVendedor = getNextCdVendedor(TERMINAL_PIEDRABUENA)
            Call addNewVendedor(auxCdVendedor, p_cuitVendedor, p_dsVendedor, TERMINAL_PIEDRABUENA)
        end if
        if (p_Pto = TERMINAL_PIEDRABUENA) then generarCodigoVendedor = auxCdVendedor
    end if
    
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Devuelve el codigo de Vendedor del puerto por su Cuit
'Function getCdVendedorByCuit(p_Pto,p_cuitVendedor)
'    Dim strSQL,rtrn
'    rtrn = 0
'    strSQL = "SELECT CDVENDEDOR FROM VENDEDORES WHERE RTRIM(NUDOCUMENTO) = '"& Trim(p_cuitVendedor) &"'"
'    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
'    if (not rs.Eof) then rtrn = rs("CDVENDEDOR")
'    getCdVendedorByCuit = rtrn
'End function
'-------------------------------------------------------------------------------------------------------------------------
'Obtenemos el proximo Codigo de Vendedor libre para agregar
Function getNextCdVendedor(p_Pto)
    Dim strSQL 
    getNextCdVendedor = 100000
    strSQL = "Select NEXTID from "&_
             "(Select 99999 + ROW_NUMBER ( ) over (ORDER BY CDVENDEDOR) NEXTID , *   from VENDEDORES C where C.CDVENDEDOR>=100000) T where NEXTID <>CDVendedor order by NEXTID"
    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL)
    if (not rs.Eof) then getNextCdVendedor = rs("NEXTID")
End Function
'-------------------------------------------------------------------------------------------------------------------------
'Agrego un nuevo Vendedor a la planta
Function addNewVendedor(p_cdVendedor,p_cuitVendedor, p_dsVendedor, p_Pto)
    Dim strSQL
    strSQL = "INSERT INTO VENDEDORES (CDVENDEDOR, DSVENDEDOR, CDTIPODOC, NUDOCUMENTO ) "&_
             "VALUES ("& p_cdVendedor &",'"& UCase(Trim(p_dsVendedor)) &"','C.U.I.T.','"& Trim(p_cuitVendedor) &"')"
    Call GF_BD_Puertos(p_Pto, rs, "EXEC", strSQL)
End function
'-------------------------------------------------------------------------------------------------------------------------
Function corteControlCupos(pRs,cuposDesde,cuposHasta,cuposActual,cantidad,totalCupo)
    corteControlCupos = false
    if (not pRs.Eof) then
        if ((Cdbl(cuposActual) >= Cdbl(cuposDesde))and(Cdbl(cuposActual) <= Cdbl(cuposHasta))) then
            if (Cdbl(totalCupo) < Cdbl(cantidad)) then  corteControlCupos = true
        end if
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function corteControlListaRangos(p_RsRangos, p_Cantidad)
    corteControlListaRangos = false
    if (not p_RsRangos.Eof) then
        if (Cdbl(p_Cantidad) > 0) then corteControlListaRangos = true
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------------
Function nominarCuentaCorriente(ByRef p_RsCup,p_RangoDesde,p_RangoHasta,p_Cantidad,p_CdCupo,p_CdProducto,p_Sucursal,p_Operacion,p_NroCto,p_Cosecha,p_NroPto,p_FechaCupo,p_CdCorredor,p_CdVendedor)
    Dim totalCupoAcumulado, saldoCupo
    totalCupoAcumulado = 0
    'El puntero inicia en el primer codigo que tiene el rango
    punteroCupo = p_RangoDesde
    while (corteControlCupos(p_RsCup,p_RangoDesde,p_RangoHasta,punteroCupo,p_Cantidad,totalCupoAcumulado))
        'Verifico si el primer codigo nominado que tenga (dentro del rango permitido) es mayor que el codigo de inicio del rango
        if ((Cdbl(p_RsCup("CODIGO")) > Cdbl(punteroCupo)) and (CInt(totalCupoAcumulado) < CInt(p_Cantidad)) and (Cdbl(punteroCupo) <= Cdbl(p_RangoHasta))) then
            Call addNominacion(p_CdCupo,p_FechaCupo,p_NroPto,p_CdProducto,p_Sucursal,p_Operacion,p_NroCto,p_Cosecha,punteroCupo,p_CdCorredor,p_CdVendedor)
            totalCupoAcumulado = Cint(totalCupoAcumulado) + 1
        else
            p_RsCup.MoveNext()
        end if
        punteroCupo = Cdbl(punteroCupo) + 1
    wend
    while((CInt(totalCupoAcumulado) < CInt(p_Cantidad)) and (Cdbl(punteroCupo) <= Cdbl(p_RangoHasta)))
        Call addNominacion(p_CdCupo,p_FechaCupo,p_NroPto,p_CdProducto,p_Sucursal,p_Operacion,p_NroCto,p_Cosecha,punteroCupo,p_CdCorredor,p_CdVendedor)
        totalCupoAcumulado = Cint(totalCupoAcumulado) + 1
        punteroCupo = Cdbl(punteroCupo) + 1
    wend
    'Obtenemos el saldo pendiente que tiene para nominar, si es 0 finalizo la nominacion correctamente, caso contrario analiza el proximo rango
    saldoCupo = CInt(p_Cantidad) - CInt(totalCupoAcumulado)
    nominarCuentaCorriente = saldoCupo
End function
'----------------------------------------------------------------------------------------------------------------------------
Function obtenerListaRangos(p_Fecha, p_NroPto, p_CdProveedor, p_CdProducto )
    Dim strSQL, rs
    logMig.info(" Obteniendo lista de rangos para la fecha "& GF_FN2DTE(p_Fecha))
    strSQL = "SELECT C.C5CODI,C.C5DSDE,C.C5HSTA,A.CUCPRO,A.CUCSUC,A.CUCOPE,A.CUNCTO,A.CUACOS,A.CUCDES,A.CUFCCP "&_
             "FROM ( SELECT CUCODI,CUCPRO, CUCSUC, CUCOPE, CUNCTO, CUACOS,  CUFCCP, CUCDES,CUCCCP  "&_
             "       FROM   MERFL.MER517F1  "&_
             "        WHERE  CUFCCP = "& p_Fecha &" AND CUCDES = "& p_NroPto &" AND CUCOPE = 04 AND CUCPRO=" & p_CdProducto & ") A  "&_
             "INNER JOIN MERFL.MER311F1 B   "&_
             "   ON A.CUCPRO = B.CPROR1 AND A.CUCSUC = B.CSUCR1 AND A.CUCOPE = B.COPER1 AND A.CUNCTO = B.NCTOR1 AND A.CUACOS = B.ACOSR1  "&_
    		 "   AND ( B.CVENR1 ="& p_CdProveedor &" OR B.CCORR1 ="& p_CdProveedor &")  "&_
             "INNER JOIN MERFL.MER517F5 C ON A.CUCODI = C.C5CODI "&_
             "ORDER BY C5DSDE,C5HSTA "
    Call executeQuery(rs, "OPEN", strSQL)
    Set obtenerListaRangos = rs
End function
'----------------------------------------------------------------------------------------------------------------------------
'Obtengo los cupos que hay nominados para un determinado rango de una fecha y puerto
Function obtenerCuposNominados(p_Fecha, p_NroPto, p_RangoDesde, p_RangoHasta, p_cdProducto)
    Dim strSQL, rs
    logMig.info(" ----> Obteniendo cupos nominados entre los rangos "&p_RangoDesde&" - "&p_RangoHasta)
    strSQL = "SELECT CODIGO "&_
             "FROM MERFL.TBLCUPOSNOMINADOS "&_
	         "WHERE FECHACUPO = "& p_Fecha &" AND PUERTO= "& p_NroPto &" AND CODIGO >= "& p_RangoDesde &" AND CODIGO <= "& p_RangoHasta & " AND IDPRODUCTO = "&p_cdProducto&_
             " ORDER BY CODIGO"
    Call executeQuery(rs, "OPEN", strSQL)
    Set obtenerCuposNominados = rs
End function
'----------------------------------------------------------------------------------------------------------------------------
Function cargarArrayCupos(p_FechaDesde, p_FechaHasta)
    Dim i,posicionesArray
    posicionesArray = GF_DTEDIFF(p_FechaDesde,p_FechaHasta,"D")
    redim arrCupo(posicionesArray)
    redim arrFecha(posicionesArray)
    i = 0
    while (i <= posicionesArray)
        arrCupo(i)  = GF_PARAMETROS7("cupo_" & i, 0, 6)
        arrFecha(i) = GF_DTEADD(p_FechaDesde, i, "D")
        i = i + 1
    wend
End function 
   '----------------------------------------------------------------------------------------------------------------------------
Function cargarLogNominacionCupos(p_Accion, p_FechaHasta,p_FechaDesde, p_NumeroPto,p_CdProveedor, p_CdProducto)
    Set logMig = new classLog
    call startLog(HND_FILE, MSG_INF_LOG)
    logMig.fileName = "NOMINACION-" & p_CdProveedor & "-" & Left(Session("MmtoDato"),8)
    logMig.info("****************************************** INICIA *********************************************************")
    logMig.info("-----> ACCION: "& UCase(p_Accion))
    logMig.info("-----> NUMERO PUERTO: "& p_NumeroPto)
    logMig.info("-----> PRODUCTO     : "& p_CdProducto)
    logMig.info("-----> FECHA DESDE  : "& GF_FN2DTE(p_FechaDesde))
    logMig.info("-----> FECHA HASTA  : "& GF_FN2DTE(p_FechaHasta))
    logMig.info("----------------------------------------------------------------------------------------------------")
End Function
'*****************************************************************************************************************************
'**************************************************** INICIO DE PAGINA *******************************************************
'*****************************************************************************************************************************
Dim cdProveedor, rs,fecha, i,cdVendedor,cdCorredor,cantidad,fechaDesde,fechaHasta,logMig,numeroPto,idCupo, arrCupo(),arrFecha()
Dim maxIndice, cdProducto

cdProveedor = GF_PARAMETROS7("cdProveedor",0,6)

'Se controla el acceso - Solo se permite elegir el proveedor por parametro si el usuario de la session es TOEPFER
if (session("KCOrganizacion") = "") then response.end
if (CLng(session("KCOrganizacion")) <> CLng(CD_TOEPFER)) then
    if (CLng(cdProveedor) <> CLng(session("KCOrganizacion"))) then
        response.end
    end if
end if


accion = GF_PARAMETROS7("accion","",6)
cdProducto = GF_PARAMETROS7("cdProducto",0,6)
fechaHasta = GF_PARAMETROS7("fechaHasta","",6)
fechaDesde = GF_PARAMETROS7("fechaDesde","",6)
cdCorredor = GF_PARAMETROS7("cdCorredor","",6)
cuitCorredor = GF_PARAMETROS7("cuitCorredor","",6)
dsCorredor = GF_PARAMETROS7("dsCorredor","",6)
cdVendedor = GF_PARAMETROS7("cdVendedor","",6)
cuitVendedor = GF_PARAMETROS7("cuitVendedor","",6)
dsVendedor = GF_PARAMETROS7("dsVendedor","",6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
cdCupo = GF_PARAMETROS7("cdCupo","",6)
numeroPto = getNumeroPuerto(g_strPuerto)
    
Call cargarLogNominacionCupos(accion, fechaHasta,fechaDesde, numeroPto,cdProveedor,cdProducto)

select case accion
    case ACCION_BORRAR
        Call deleteNominacion(numeroPto,cdCorredor,cdVendedor,fechaDesde,fechaHasta, cdProducto,cdCupo)
    case ACCION_GRABAR
        'Primero cargamos los array con los valores recibidos por parametros , estos array son globales
        Call cargarArrayCupos(fechaDesde,fechaHasta)
        'Controlo que la cantidad asignada sea correcta con el historico de nominaciones para el contrato y su rango (Desde-Hasta)
        if (controlarCantidadCupos(arrCupo,arrFecha,numeroPto,cdProveedor, cdProducto)) then            

            'Obtengo los codigo de vendedor y de corredor para grabar la nominacion
            cdVendedor = generarCodigoVendedor(cuitVendedor,dsVendedor,g_strPuerto, cdVendedor)
            cdCorredor = generarCodigoCorredor(cuitCorredor,dsCorredor,g_strPuerto, cdCorredor)
            
            for i = 0 to UBound(arrFecha)
                'Recorro los 10 valores ingresados para las fechas, solo se calcula cuando los kilos son mayores a 0
                if (CInt(arrCupo(i)) > 0) then
                    'Cantidad de cupos a nominar
                    cantidad = arrCupo(i)
                    'Primero obtengo una lista de rangos (Recordset) que tiene permitido el proveedor para un puerto y una fecha especifica
                    Set rsRangos = obtenerListaRangos(arrFecha(i), numeroPto, cdProveedor, cdProducto)
                    while(corteControlListaRangos(rsRangos, cantidad))
                        Set rsCup = obtenerCuposNominados(arrFecha(i), numeroPto, rsRangos("C5DSDE"), rsRangos("C5HSTA"),cdProducto)
                        cantidad = nominarCuentaCorriente(rsCup,rsRangos("C5DSDE"),rsRangos("C5HSTA"),cantidad,rsRangos("C5CODI"),rsRangos("CUCPRO"),rsRangos("CUCSUC"),rsRangos("CUCOPE"),rsRangos("CUNCTO"),rsRangos("CUACOS"),rsRangos("CUCDES"),rsRangos("CUFCCP"),cdCorredor,cdVendedor)
                        rsRangos.MoveNext()
                    wend
                end if
            next
            'Si grabo los datos correctamente envio el codigo de vendedor/comprador por si es nuevo alguno de los dos y fijar los valores visualmente en la pagina
            Response.Write RESPUESTA_OK &"|"& cdVendedor &"|"& cdCorredor
        else
           Response.Write "La cantidad no cubre los cupos" 
        end if
end select
logMig.info("****************************************** FINALIZA *******************************************************")
%>