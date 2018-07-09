<!--#include file="procedimientosMG.asp"-->
<!--#include file="procedimientosAS400.asp"-->
<!--#include file="procedimientosFechas.asp"-->
<!--#include file="procedimientosFormato.asp"-->
<!--#include file="procedimientosEmpresas.asp"-->
<!--#Include File="ExternalFunctions.ASP" -->
<!--#Include File="NroEnLetras.asp" -->
<%
const SEP_ENTRE_PARRAFOS = 5
const BOLETO_UPDATE = 0
const BOLETO_INSERT = 1
const MRCENVIO_F = 0
const MRCENVIO_X = 1
const MRCENVIO_N = 2
const MRCENVIO_T = 3
const MRCENVIO_V = 4
const CANT_ENV_AUX = 1
const ISTITULO = 0
const ISDESCRIPCION = 1
const ZONA_ROSARIO = 2

Dim intItem

'*****************************************************************************************************************
function GF_establecerValoresBoleto(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha, byref p_dicc)
    dim strSQL, rsContrato, aux, auxKilos, auxAnulaciones, auxEntity, auxDsEntity, auxArbEntity, auxDicEntity
    
    set p_dicc = server.createObject("Scripting.Dictionary")

    'strSQL = "select * from Contratos where c.Producto=" & p_intProducto & " and c.Sucursal=" & p_intSucursal & " and c.Operacion=" & p_intOperacion & " and c.Numero=" & p_intNumero & " and c.Cosecha=" & p_intCosecha
    strSQL = " Select C.CPROR1 as Producto, C.CSUCR1 as Sucursal, C.COPER1 as Operacion, C.NCTOR1 as Numero, C.ACOSR1 as Cosecha, C.FCCTR1 as FechaConc, C.CCORR1 as KCCOR, C.CVENR1 as KCVEN, C.KGCOR1 as Kilos, sum(B.KGCORB) as Anulaciones, case(M.MDOLOM) when 'F' then C.PRECR1*1 else C.PRECR1*C.TIPCR1 end as PrecioP, C.PORPR1 as PjeParcial, C.FDFIR1 as FechaFijaDesde, C.FHFIR1 as FechaFijaHasta, C.FDPER1 as FechaEntDesde, C.FHPER1 as FechaEntHasta,"
	strSQL = strSQL & " C.KGNFR1 as KilosMin, C.KGMFR1 as KilosMax, C.CDESR1 as PuertoRecepcion, C.DESTR1 as PuertoDevolucion, J.MCPDRJ as MercPropia, C.CONCR1 as CtoCorredor, C.TIPOR1 as TipoCto, C.CONFR1 as MrcConfirma, P.DESCPC as Procedencia, C.CPRDR1 as CPProcedencia, C.AUXIR1 as CAProcedencia, F.DESCFP as KCPago, case(M.MDOLOM) when 'F' then C.PRECR1/C.TIPCR1 else C.PRECR1*1 end as PrecioD, C.TIPCR1 as TipoCambio, C.KGRER1 as SaldoEnt, C.CTRAR1 as Transporte,"
	strSQL = strSQL & " C.CFPAR1 as CodigoPago, J.HUMERJ as MercConHumedad, K.CAPARK as CamionesPactados, C.FPACR1 as FechaPago, case(J.CDIVRJ) when 'C' then J.CDIVRJ else 'X' end as CondicionIVA, R.PLRECI as Recibido, C.DIAPR1 as DiasPago, J.MCONRJ as MercConsignacion, S.SIOGRANOS as CODIGOSIO "
	strSQL = strSQL & " from MERFL.MER311F1 C"
	strSQL = strSQL & " left join MERFL.MER311FB B on C.CPROR1=B.CPRORB and C.CSUCR1=B.CSUCRB and C.COPER1=B.COPERB and C.NCTOR1=B.NCTORB and C.ACOSR1=B.ACOSRB"
	strSQL = strSQL & " left join MERFL.MER311FJ J on C.CPROR1=J.CPRORJ and C.CSUCR1=J.CSUCRJ and C.COPER1=J.COPERJ and C.NCTOR1=J.NCTORJ and C.ACOSR1=J.ACOSRJ"
	strSQL = strSQL & " left join MERFL.MER311FK K on C.CPROR1=K.CPRORK and C.CSUCR1=K.CSUCRK and C.COPER1=K.COPERK and C.NCTOR1=K.NCTORK and C.ACOSR1=K.ACOSRK"
	strSQL = strSQL & " left join MERFL.MER341F2 R on C.CPROR1=R.PLCPRO and C.CSUCR1=R.PLCSUC and C.COPER1=R.PLCOPE and C.NCTOR1=R.PLNCTO and C.ACOSR1=R.PLACOS"
	strSQL = strSQL & " left join MERFL.MER311FM S on C.CPROR1=S.PRODUCTO and C.CSUCR1=S.SUCURSAL and C.COPER1=S.OPERACION and C.NCTOR1=S.NUMERO and C.ACOSR1=S.COSECHA"
	strSQL = strSQL & " left join MERFL.MER2I1F1 F on C.CFPAR1=F.CODIFP"
	strSQL = strSQL & " left join MERFL.MER132F1 M on C.COPER1=M.CODIOM"
	strSQL = strSQL & " left join MERFL.MER142F1 P on C.CPRDR1=P.CODIPC and C.AUXIR1=P.AUXIPC"
	strSQL = strSQL & " where C.CPROR1 =" & p_intProducto & " and C.CSUCR1 =" & p_intSucursal & " and C.COPER1 =" & p_intOperacion & " and C.NCTOR1 =" & p_intNumero & " and C.ACOSR1 = " & p_intCosecha
	strSQL = strSQL & " group by C.CPROR1, C.CSUCR1, C.COPER1, C.NCTOR1, C.ACOSR1, C.FCCTR1, C.CCORR1, C.CVENR1, C.KGCOR1, C.PORPR1, C.FDFIR1, C.FHFIR1, C.FDPER1, C.FHPER1, C.KGNFR1, C.KGMFR1, C.CDESR1, C.DESTR1, J.MCPDRJ, C.CONCR1, C.TIPOR1, C.CONFR1, P.DESCPC, C.CPRDR1, C.AUXIR1, F.DESCFP, M.MDOLOM, C.TIPCR1, C.KGRER1, C.CTRAR1, C.CFPAR1, J.HUMERJ, K.CAPARK, C.FPACR1, J.CDIVRJ, R.PLRECI, C.DIAPR1, J.MCONRJ, C.PRECR1, S.SIOGRANOS"
	strSQL = strSQL & " order by  C.FCCTR1 asc"
	
    'response.write strSQL
    call GF_BD_AS400_2(rsContrato, conn, "OPEN", strSQL)
    if not rsContrato.eof then
       'Datos del Contrato
        call p_dicc.add("Producto",cdbl(p_intProducto))
        call p_dicc.add("Sucursal",cdbl(p_intSucursal))
        call p_dicc.add("Operacion",cdbl(p_intOperacion))
        call p_dicc.add("Numero",cdbl(p_intNumero))
        call p_dicc.add("Cosecha",cdbl(p_intCosecha))
        aux = rsContrato("PuertoRecepcion")
        if isnull(aux) then aux = 0
        call p_dicc.add("PuertoRecepcion",cdbl(aux))
        aux = rsContrato("FechaConc")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaConc",cdbl(aux))
        aux = rsContrato("FechaEntDesde")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaEntDesde",cdbl(aux))
        aux = rsContrato("FechaEntHasta")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaEntHasta",cdbl(aux))
        aux = rsContrato("FechaFijaDesde")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaFijaDesde",cdbl(aux))
        aux = rsContrato("FechaFijaHasta")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaFijaHasta",cdbl(aux))
        aux = rsContrato("CPProcedencia")
        if isnull(aux) then aux = 0
        call p_dicc.add("CPProcedencia",cdbl(aux))
        aux = rsContrato("CAProcedencia")
        if isnull(aux) then aux = 0
        call p_dicc.add("CAProcedencia",cdbl(aux))
        call p_dicc.add("ProcedenciaProv",getProcProv(p_dicc("CPProcedencia"), p_dicc("CAProcedencia")))
        call p_dicc.add("DSProcedencia",getProcDs(p_dicc("CPProcedencia"), p_dicc("CAProcedencia")))
        'Datos del Vendedor
        aux = rsContrato("KCVEN")
        if isnull(aux) then aux = 0
        call p_dicc.add("KCVEN",cdbl(aux))
        aux = replace(getEnterpriseCUIT(cdbl(aux)), "-", "")
        call p_dicc.add("VEN_CUIT", aux)
        'Datos del Corredor
        aux = rsContrato("KCCOR")
        if isnull(aux) then aux = 0
        call p_dicc.add("KCCOR",cdbl(aux))
        aux = replace(getEnterpriseCUIT(cdbl(aux)), "-", "")
        call p_dicc.add("COR_CUIT", aux)
        'Datos del comprador (Toepfer)
        call p_dicc.add("KCCOMP", 99999997)
        call p_dicc.add("RSCOMP", GetDsEnterprise2(p_dicc("KCCOMP")))
        auxKilos = rsContrato("Kilos")
        auxAnulaciones = rsContrato("Anulaciones")
        if isnull(auxKilos) then auxKilos = 0
        if isnull(auxAnulaciones) then auxAnulaciones = 0
        aux = cDBl(auxKilos) + cDBl(auxAnulaciones)
        call p_dicc.add("Kilos", aux)
        aux = rsContrato("Anulaciones")
        call p_dicc.add("Anulaciones", aux)
        aux = rsContrato("KilosMin")
        if isnull(aux) then aux = 0
        call p_dicc.add("KilosMin", cdbl(aux))
        aux = rsContrato("KilosMax")
        if isnull(aux) then aux = 0
        call p_dicc.add("KilosMax", cdbl(aux))
        aux = CDbl(rsContrato("PrecioP"))
        if isnull(aux) then aux = 0
        'Elimino decimales basura
        if ((1-abs(aux - Int(aux))) < 0.01) then aux = Int(aux) + 1        
        call p_dicc.add("PrecioP", cdbl(aux))
        aux = CDbl(rsContrato("PrecioD"))    
        if isnull(aux) then aux = 0    
        'Elimino decimales basura
        if ((1-abs(aux - Int(aux))) < 0.01) then aux = Int(aux) + 1
        call p_dicc.add("PrecioD", cdbl(aux))
        if rsContrato("MercPropia") = "V" then
            call p_dicc.add("MercPropia", true)
        else
            call p_dicc.add("MercPropia", false)
        end if
        aux = rsContrato("TipoCto")
        call p_dicc.add("TipoCto", aux)
        aux = rsContrato("CodigoPago")
        call p_dicc.add("CodigoPago", aux)
        aux = rsContrato("KCPago")
        call p_dicc.add("KCPago", aux)
        aux = rsContrato("CtoCorredor")
        call p_dicc.add("CtoCorredor", aux)
        call p_dicc.add("CtoVendedor", aux)
        aux = rsContrato("Transporte")
        if isnull(aux) then aux = 0
        call p_dicc.add("Transporte", cdbl(aux))
        aux = rsContrato("MercConHumedad")
        call p_dicc.add("MercConHumedad", aux)
        aux = rsContrato("PjeParcial")
        call p_dicc.add("PjeParcial", aux)
        aux = rsContrato("CamionesPactados")
        if isnull(aux) then aux = 0
        call p_dicc.add("CamionesPactados", cdbl(aux))
        aux = rsContrato("FechaPago")
        if isnull(aux) then aux = 0
        call p_dicc.add("FechaPago", cdbl(aux))
        aux = getEntity(p_dicc("KCCOR"), 0, p_dicc("Sucursal"),0)
        call p_dicc.add("Entidad", aux)
        Call getDataEntity(p_dicc("Entidad"), auxDsEntity, auxArbEntity, auxDicEntity)
        Call p_dicc.add("Entidad_Descripcion",	auxDsEntity)
        Call p_dicc.add("Entidad_Arbitral",		auxArbEntity) 
        Call p_dicc.add("Entidad_Direccion",	auxDicEntity)
        aux = rsContrato("CondicionIVA")
        if isnull(aux) then aux = ""
        call p_dicc.add("CondicionIVA", aux)
        'Determino a quien se le paga
		if ((clng(p_intProducto) = 23) and (clng(p_intSucursal) = 0) and (clng(p_intOperacion) = 0) and (clng(p_intNumero) = 65794) and (clng(p_intCosecha) = 06)) then
            aux = p_dicc("KCVEN")
            call p_dicc.add("PagarA", aux)
        else
			strSQL = "select 1 from  PROVFL.AAD4CPP where GAOPR1=" & p_dicc("KCVEN")
			call GF_BD_AS400_2(rsProveedores, conn, "OPEN", strSQL)
			if (not rsProveedores.eof) or (cInt(p_dicc("KCCOR"))=0) then
			    aux = p_dicc("KCVEN")
			    call p_dicc.add("PagarA", aux)
			else
			    aux = p_dicc("KCCOR")
			    call p_dicc.add("PagarA", aux)
			end if
        end if
        aux = rsContrato("DiasPago")
        if isnull(aux) then aux = 0
        call p_dicc.add("DiasPago", aux)
        aux = rsContrato("CODIGOSIO")
        if isnull(aux) then aux = "-"
        call p_dicc.add("CODIGOSIO",aux)
        GF_establecerValoresBoleto = true
    else
        GF_establecerValoresBoleto = false
    end if
end function
'*****************************************************************************************************************
function getDsPort(p_KCPORT)
    dim strSQl, rs

    getDsPort = "#KCPORT invalido#"
    if isnumeric(p_KCPORT) then
        strSQL = "select DESCDE from MERFL.MER192F1 where CODIDE='" & p_KCPORT & "'"
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            getDsPort = Trim(rs("DESCDE"))
        end if
    end if
end function
'*****************************************************************************************************************
function getDsProduct(p_KC)
    dim rsMG, conn, strSQL
    
    getDsProduct = "#ERROR KC Producto#"
    'strSQL = "select * from mg where mg_km='AR' and mg_kc='" & p_KC & "'"
    strSQL = "Select * from MERFL.MER112F1 where CODIPR=" & p_KC
    call GF_BD_AS400_2(rsMG, conn, "OPEN", strSQL)    
    if not rsMG.eof then
        getDsProduct = Trim(rsMG("DESCPR"))        
    end if
end function
'*****************************************************************************************************************
function getDsTransport(p_kc)
    dim strTransDs
    if not isnull(p_kc) and isnumeric(p_kc) then
		select case (p_kc)
			case 1:		strTransDs = "Camiones"
			case 2:		strTransDs = "Vagones"
			case 3:		strTransDs = "Camiones/Vagones"
			case 4:		strTransDs = "Lanchones"
			case 5:		strTransDs = "Transferencias"
			case else:	strTransDs = "Otros"
		end select
		getDsTransport = strTransDs
    else
        getDsTransport = "#Dato no valido#"
    end if
end function
'*****************************************************************************************************************
function getProcProv(p_CodPostal, p_CodArea)

    if (p_CodPostal <> "") and (p_CodArea <> "") then
        if isnumeric(p_CodPostal) and isnumeric(p_CodArea)  then
            strSQL = "select PROVPC as Provincia from MERFL.MER142F1 where CODIPC=" & p_CodPostal & " and AUXIPC=" & p_CodArea
            call GF_BD_AS400_2(rsProc, conn, "OPEN", strSQL)
            if not rsProc.eof then
               getProcProv = rsProc("Provincia")
            else
               getProcProv = "#Procedencia no encontrada#"
            end if
        else
            getProcProv = "#Codigo Postal y de Area Erroneos#"
       end if
    else
        getProcProv = "#Codigo Postal y de Area No Migrados#"
    end if
end function
'*****************************************************************************************************************
function getDsProv(p_KcProv)
    dim auxDs, conn, strSQL, rs    
    strSQL="Select * from MERFL.MER1K2F1 where CODIPO='" & p_KcProv & "'"
    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)    
    auxDs = "???"
    if (not rs.eof) then auxDs = rs("DESCPO")
    getDsProv = auxDs
end function
'*****************************************************************************************************************
function getEntity(p_KCCOR, p_ProvProcedencia, p_sucursal, p_codigoPostalProcedencia)    
    if (CLng(p_KCCOR) = 12521) then
		'Si el proveedor en ADECO AGRO
		getEntity = "2"
    elseif (cint(getEnterpriseCP(p_KCCOR)) = 8000) then
        'Si el corredor esta domiciliado en Bahia Blanca
        getEntity = "3"
    elseif (ucase(p_ProvProcedencia) = "X") or  (ucase(p_ProvProcedencia) = "S") then
        'si la procedencia de la mercaderia es Cordoba o Rosario
        getEntity = "2"
    elseif ucase(p_ProvProcedencia) = "B" and cint(p_codigoPostalProcedencia) = 8000 then
        'Si la procedencia de la mercaderia es Buenos Aires y el destino tambien es Bs As
        'Tener en cuenta que el unico destino(puerto) en Bs As es Bahia Blanca, por eso pregunto
        'por el CP de Bahia
        getEntity = "1"
    elseif cint(p_Sucursal) = 1 then
        'Si la sucursal es Rosario
        getEntity = "2"
    else
        'Si no es ninguna la entidad es Buenos Aires
        getEntity = "1"
    end if
end function
'*****************************************************************************************************************
Function getDataEntity(idEntity, ByRef DsEntity, ByRef ArbEntity, ByRef DicEntity)
	dim strSQL, conn, rs

	getDataEntity = false
	DsEntity = "#Entidad no encontrada#"
	ArbEntity = "#Entidad no encontrada#"
	DicEntity = "#Direcci�n no encontrada#"
	
	strSQL = "Select * from MERFL.MER2A2F2 where CODIBD = " & idEntity
	call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	if (not rs.eof) then
		DsEntity = Trim(rs("DESCBD"))
		ArbEntity = Trim(rs("EARBBD"))
		DicEntity = Trim(rs("DIREBD"))
		getDataEntity = true
	end if
end function
'*****************************************************************************************************************
function getProcDs(p_CodPostal, p_CodArea)
	dim rsProc, conn, strSQL
    if (p_CodPostal <> "") and (p_CodArea <> "") then
        if isnumeric(p_CodPostal) and isnumeric(p_CodArea) then
            strSQL = "select DESCPC as Descripcion from MERFL.MER142F1 where CODIPC=" & p_CodPostal & " and AUXIPC=" & p_CodArea
            call GF_BD_AS400_2(rsProc, conn, "OPEN", strSQL)
            if not rsProc.eof then
                getProcDs = rsProc("Descripcion")
            else
                getProcDs = "#Procedencia no encontrada#"
            end if
        else
            getProcDs = "#Codigo Postal y de Area Erroneos#"
        end if
    else
        getProcDs = "#Codigo Postal y de Area No Migrados#"
    end if
end function
'*****************************************************************************************************************
Function getDsEntidad(p_valor, p_tipo)
	dim rtrn, rtrnBIS
	Select case (p_valor)
		case 1:
			rtrn = "B O L S A  D E  C E R E A L E S  D E  B S . A S ." 
			rtrnBIS = "Bolsa de cereales de Bs. As."
		case 2:
			rtrn = "B O L S A  D E  C E R E A L E S  D E  R O S A R I O"
			rtrnBIS = "Bolsa de cereales de Rosario"
		case 3:
			rtrn = "B O L S A  D E  C E R E A L E S  D E  B A H I A  B L A N C A"
			rtrnBIS = "Bolsa de cereales de Bahia Blanca"
		case 5:
			rtrn = "B O L S A  D E  C O M E R C I O  D E  S A N T A  F E"
			rtrnBIS = "Bolsa de cereales de Santa Fe"
		case else:
			rtrn = "B O L S A  D E  C O M E R C I O  ?"
			rtrnBIS = "Bolsa de cereales de ?"
	end Select
	if (p_tipo = ISTITULO) then
		getDsEntidad = rtrn
	else
		getDsEntidad = rtrnBIS
	end if
end Function
'*****************************************************************************************************************
function formatear(p_valor, p_tipo)

    select case ucase(p_tipo)
        case "CUIT": formatear = left(p_valor,2) & "-" & mid(p_valor,3,len(p_valor)-3) & "-" & right(p_valor,1)
        case "ANIO":
            formatear = "#ERROR ANIO#"
            if isnumeric(p_valor) then
                if cint(p_valor) < 30 then
                    formatear = 2000 + cint(p_valor)
                elseif cint(p_valor) > 1000 then
                    formatear = cint(p_valor)
                else
                    formatear = 1900 + cint(p_valor)
                end if
            end if
        case "TITULO2"
            auxVec = split(p_valor, " ")
            strExprFinal = ""
            for indicePalabra = lbound(auxVec) to ubound(auxVec)
                if (strExprFinal <> "") then strExprFinal = strExprFinal & space(3)
                strPalFinal = ""
                for indiceLetra = 1 to len(auxVec(indicePalabra))
                    if strPalFinal <> "" then strPalFinal = strPalFinal & " "
                    strPalFinal = strPalFinal & mid(auxVec(indicePalabra), indiceLetra, 1)
                next
                strExprFinal = strExprFinal & strPalFinal
            next
            formatear = strExprFinal
            response.write formatear
        case else: formatear = p_valor
    end select
end function
'*****************************************************************************************************************
sub GF_WM_diagonal(p_oPDF,p_xo,p_yo,p_cant)
    Dim i,Xo,Yo
    Xo=p_xo
    Yo=p_yo
    for i=1 to p_cant
        Call GF_writeImage(p_oPDF, server.mappath("images/watermark.jpg"), Xo, Yo, 70, 50, 1)
        Xo=Xo+90
        Yo=Yo+70
    next
end sub
'*****************************************************************************************************************
sub GF_dibujarMarcaDeAgua(p_oPDF)
    'Se arman diagonales.
    'Las diagonales se numeran de derecha a izquierda en forma crecientre.
    'DIAGONAL 1
    Call GF_WM_diagonal(p_oPDF,300,200,3)
    'DIAGONAL 2
    Call GF_WM_diagonal(p_oPDF,25,200,6)
    'DIAGONAL 3
    Call GF_WM_diagonal(p_oPDF,25,400,6)
    'DIAGONAL 4
    Call GF_WM_diagonal(p_oPDF,25,600,3)
end sub
'*****************************************************************************************************************
sub GF_dibujarEncabezado(p_oPDF, p_boleto)
    dim strTitulo, intLineaBase
    
    Call GF_setFont(p_oPDF, "Times",9,0)
    'Se dibuja el encabezado.        
    Call GF_squareBox(p_oPDF, 100, 20, 400, 100, 0, "#FFFFFF", "#000000", 1, 0)    
    Call GF_verticalLine(p_oPDF, 210, 20, 100)
    Call GF_verticalLine(p_oPDF, 300, 20, 100)
    Call GF_horizontalLine(p_oPDF, 100, 45, 200)
    Call GF_horizontalLine(p_oPDF, 100, 70, 200)
    Call GF_horizontalLine(p_oPDF, 100, 95, 200)    
    
    Call GF_writeText(p_oPDF, 120, 28, "Cto Vendedor:", 0)
    Call GF_writeText(p_oPDF, 215, 28, p_boleto("CtoVendedor"), 0)

    Call GF_writeText(p_oPDF, 120, 53, "Cto Comprador:", 0)
    Call GF_writeText(p_oPDF, 215, 53, GF_EDIT_Contrato(p_boleto("Producto"),p_boleto("Sucursal"), p_boleto("Operacion"), p_boleto("Numero"), p_boleto("Cosecha")), PDF_ALIGN_LEFT)

    Call GF_writeText(p_oPDF, 120, 78, "Cto Corredor:", 0)
    Call GF_writeText(p_oPDF, 215, 78, p_boleto("CtoCorredor"), 0)
    
    Call GF_writeText(p_oPDF, 120, 103, "C�digo SIO Granos:", 0)
    Call GF_writeText(p_oPDF, 215, 103, p_boleto("CODIGOSIO"), 0)
    
    call GF_setFont(p_oPDF, "Times",7,0)
    call GF_writeTextAlign(p_oPDF, 304, 110, "Espacio reservado para la " & getDsEntidad(p_boleto("Entidad"), ISDESCRIPCION), 190, PDF_ALIGN_CENTER)
    
    intLineaBase = 125
    call GF_setFont(p_oPDF, "Arial",16,10)
    call GF_writeTextAlign(p_oPDF, 1, intLineaBase, getDsEntidad(p_boleto("Entidad"), ISTITULO), 600, PDF_ALIGN_CENTER)

    call GF_setFont(p_oPDF, "Arial", 9, 0)
    intLineaBase = intLineaBase + 17
    call GF_writeTextAlign(p_oPDF, 1, intLineaBase, "(Formulario oficial emitido en orden a lo establecido por el Art.30 de la Ley de sellos Nacionales T.O. 1965 Decreto 9432/44 y",600,PDF_ALIGN_CENTER)
    intLineaBase = intLineaBase + 9
    call GF_writeTextAlign(p_oPDF, 1, intLineaBase, "modificatorias, y Art. 18 de la correspondiente Reglamentaci�m General, Decreto 3666/55 y modificatorias)", 600, PDF_ALIGN_CENTER)

    call GF_setFont(p_oPDF, "Arial", 10, 8)
    intLineaBase = intLineaBase + 10
    call GF_writeTextAlign(p_oPDF, 1, intLineaBase, "B O L E T O   D E   C O M P R A - V E N T A   D E   C E R E A L E S ,  O L E A G I N O S O S", 600, PDF_ALIGN_CENTER)
    intLineaBAse = intLineaBase + 10
    call GF_writeTextAlign(p_oPDF, 1, intLineaBase, "Y   D E M A S   P R O D U C T O S   D E   L A   A G R I C U L T U R A", 600, PDF_ALIGN_CENTER)

end sub
'*****************************************************************************************************************
sub GF_dibujarClausulas(p_oPDF, p_boleto)
    dim intLineaBase
    
    intLineaBase = 185
	intAnchoLinea = 560

    intAltoRenglon = 7
    intSeparacionEntreParrafos = 3
    intItem = 1
    strCamiones = "camiones"

    call GF_setFont(p_oPDF, "Times", 7, 0)

    select case cint(p_boleto("Operacion"))
        case 0,1,2,3,5:
            simboloMoneda = "$"
            precioTonelada = p_boleto("PrecioP")
            aclaracionMoneda = "pesos"
        case 6,9,10,11,12:
            simboloMoneda = "U$S"
            precioTonelada = p_boleto("PrecioD")
            aclaracionMoneda = "d�lares"
    end select

    if p_boleto("CamionesPactados")=1 then strCamiones = "cami�n"
    strParrafo = "<b>" & intItem & ".-</b> Los se�ores <b>" & GetDsEnterprise2(p_boleto("KCVEN")) & "</b> (CUIT:" & formatear(p_boleto("VEN_CUIT"),"CUIT") & "), domiciliado en " & getEnterpriseDIR(p_boleto("KCVEN"))& " C.P. " & getEnterpriseCP(p_boleto("KCVEN")) & " " & getEnterpriseLoc(p_boleto("KCVEN")) & " PCIA de " & Trim(getEnterpriseProv(p_boleto("KCVEN"))) & ", venden a <b>" & p_boleto("RSCOMP") & ",</b> domiciliada en Av del Libertador 350 10� PISO, Vicente Lopez (C.P. 1638), la cantidad de <b>" & GF_EDIT_INTEGER(p_boleto("Kilos")) & " (" & trim(NroEnLetras(p_boleto("Kilos"), true)) & ") Kgs "
	if ucase(p_boleto("CodigoPago")) = "X" then strParrafo = strParrafo & ", o el resultante de " & p_boleto("CamionesPactados") & " " & strCamiones & ", "
    strParrafo = strParrafo & "de " & lcase(getDsProduct(clng(p_boleto("Producto")))) & ", y dem�s condiciones  c�mara de la cosecha  <b>" & formatear(clng(p_boleto("Cosecha"))-1,"anio") & "-" & formatear(clng(p_boleto("Cosecha")),"anio")
    if ((clng(p_boleto("Operacion")) = 1) or (clng(p_boleto("Operacion")) = 10)) then
        strParrafo = strParrafo & ".(Precio a Fijar)</b>"
    else
        strParrafo = strParrafo & ",</b> al precio de <b>" & simboloMoneda & GF_EDIT_DECIMALS(precioTonelada*100,2) & " (" & trim(NroEnLetras(precioTonelada, true)) & " " &  Ucase(aclaracionMoneda) & ")</b>  la  tonelada  granel."
    end if   
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>" & intItem & ".-</b> Las entregas y recibos se efectuar�n en <b>Puerto " & getDsPort(p_boleto("PuertoRecepcion")) & "</b> - donde los vendedores deber�n remitir la mercader�a, en <b>" & lcase(getDsTransport(p_boleto("Transporte"))) & "</b> atracados y descargados desde el  <b>" & GF_FN2DTE(p_boleto("FechaEntDesde")) & "  y  hasta  el  " & GF_FN2DTE(p_boleto("FechaEntHasta")) & "</b>  inclusive, haci�ndose el recibo por el recibidor del comprador en destino."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)

    strParrafo = "Procedencia de la Mercader�a: <b>" & trim(p_boleto("DSProcedencia"))
    if cint(p_boleto("CPProcedencia")) <> 99 then
        strParrafo = strParrafo & ", " & ucase(getDsProv(p_boleto("ProcedenciaProv")))
    end if	
	
    strParrafo = strParrafo & ".</b>"
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    strParrafo = "La mercader�a entregada se liquidar� seg�n condici�n <b>CAMARA,</b> salvo las condiciones establecidas en el presente contrato."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    strParrafo = " Los vendedores declaran en forma expresa que la mercader�a: (marque lo que corresponda)"
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    call GF_dibujarOpcionesMercaderia(p_oPDF, intLineaBase)
	
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos    
    strParrafo = "<b>" & intItem & ".-</b> Los gastos correspondientes a servicios de acondicionamiento de la mercader�a ser�n descontados al vendedor o al corredor interviniente, seg�n las tarifas vigentes al momento de la descarga, salvo que sean acreditados al comprador dentro de las 72hs de recibido el servicio. En los contratos a fijar precio, las partes acuerdan que caso de que los servicios de acondicionamiento se realicen previo a la fijaci�n del precio, el comprador podr� fijar el precio, seg�n las condiciones contractuales, �nicamente sobre la cantidad de kilogramos entregados necesarios para cancelar los montos equivalentes al acondicionamiento adeudado, salvo que el vendedor acredite al comprador dicho gasto dentro de las 72hs de realizado el servicio."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)
    
    call armarClausulaFormaDePago(p_oPDF, p_boleto, intLineaBase, intItem, intAnchoLinea, intAltoRenglon, intSeparacionEntreParrafos)
    call armarClausulaCorredor(p_oPDF, p_boleto, intLineaBase, intItem, intAnchoLinea, intAltoRenglon, intSeparacionEntreParrafos)

    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>" & intItem & ".-</b> El presente contrato es INTRANSFERIBLE. Ninguna de las partes podr� ceder o transferir en forma alguna, ni total o parcialmente, ni el contrato ni los derechos y/u obligaciones emergentes del mismo."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)
    
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strSucursal = p_boleto("Entidad_Descripcion")
    strEntidadArbitrante = p_boleto("Entidad_Arbitral")
    if (strEntidadArbitrante = "?") then strEntidadArbitrante = p_boleto("Entidad_Descripcion")
    strParrafo = "<b>" & intItem & ".-</b> Todos los firmantes acuerdan que todas las divergencias, cuestiones o reclamos que surjan de o que se relacionen con cualquiera de las relaciones jur�dicas que se deriven de este contrato y entre cualesquiera de ellos, ser�n resueltas en forma definitiva por la " & strEntidadArbitrante & ". El tribunal actuar� como amigable componedor, con aplicaci�n de las Reglas y Usos del Comercio de Granos y del Reglamento de Procedimientos aprobado por Decreto 931/98 y/o sus futuras modificaciones, ampliaciones o normas complementarias."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)
    
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>" & intItem & ".-</b> Todos los firmantes declaran conocer y aceptar las Reglas y Usos del Comercio de Granos."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)
    
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    'levanto la direccion (SGDR) de la sec. de la camara arbitral de cereales de la mgdt
    'direccion = calle + altura + localidad [+ Provincia]
    direccion = p_boleto("Entidad_Direccion")
    strParrafo = "<b>" & intItem & ".-</b> A los efectos del presente boleto las partes constituyen domicilio especial en la " & p_boleto("Entidad_Descripcion") & ", " & direccion & ", donde se notificar�n v�lidamente todas las citaciones, providencias y resoluciones."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)

    call armarClausulaMercaderia(p_oPDF, p_boleto, intLineaBase, intItem, intAnchoLinea, intAltoRenglon, intSeparacionEntreParrafos)
		
	call armarClausulaAFIP(p_oPDF, p_boleto, intLineaBase, intItem, intAnchoLinea, intAltoRenglon, intSeparacionEntreParrafos)

    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>" & intItem & ".-</b> Operaci�n concertada el <b>" & GF_FN2DTE(p_boleto("FechaConc")) & ".</b>"
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    
    intItem = intItem + 1
    intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>" & intItem & ".-</b> El vendedor�deber� cumplir� con la Disposici�n �General del Servicio Nacional de Sanidad Vegetal N� �3/83, �sus consider�ndos y correspondientes Art. 1� ,2� ,3�� que se transcriben a continuaci�n como as� tambi�n con la los art�culos 134 ( Ex 123 ) y Art. 135� ( ex.124 )�� de la Ley Provincial de Santa Fe Nro. 10.703�. Disposici�n 3/83: ARTICULO 1�.- Proh�base el tratamiento con plaguicida fumigantes de los granos, productos y subproductos de cereales y oleaginosos, durante la carga de los mismos en camiones o vagones y durante el tr�nsito de �stos hasta su destino.ARTICULO 2�.- Todo cami�n o vag�n en el que se detecten restos de plaguicidas fumigantes sin descomponer o bien concentraciones elevadas de los mismos en el momento de la descarga, ser�n rechazados, debiendo cumplir o complementar seg�n corresponda, antes de ser descargado, el tiempo de exposici�n y de ventilaci�n que se indican: 96 horas y 6 horas respectivamente. ARTICULO 3�.- Las demoras y/o perjuicios econ�micos provocados por el uso inadecuado de estos plaguicidas correr�n por cuenta exclusiva de los remitentes."
	intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)    
	
	if ((p_boleto("Producto") = 17) and (p_boleto("FechaConc") >= "20130320")) then
	    intItem = intItem + 1
	    intLineaBase = intLineaBase + intSeparacionEntreParrafos
	    if (p_boleto("Cosecha") = 15) then
	        strParrafo = "<b>" & intItem & ".-</b> El vendedor reconoce que la mercader�a a ser entregada condice con la variedad de Cebada  Cervecera Declarada en el presente contrato,   la misma  debe mantener la Pureza Varietal al 95% de la declarada, no permitiendo mezclas con otras variedades al momento del almacenaje o carga de la misma.  La Pureza Varietal ser� determinada por la C�mara Arbitral correspondiente y en caso que  los an�lisis determinen un valor inferior al 95%, se aplicar� una rebaja del 15%."
	    else
	        strParrafo = "<b>" & intItem & ".-</b> El vendedor reconoce que la mercader�a a ser entregada condice con la variedad de Cebada  Cervecera Declarada en el presente contrato,   la misma  debe mantener la Pureza Varietal al 98% de la declarada, no permitiendo mezclas con otras variedades al momento del almacenaje o carga de la misma.  La Pureza Varietal ser� determinada por la C�mara Arbitral correspondiente y en caso que  los an�lisis determinen un valor inferior al 98%, se aplicar� una rebaja del 30%."
        end if	        
        intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)    
	end if
	
	
	if p_boleto("FechaConc") > "20120321" then
		intItem = intItem + 1
		intLineaBase = intLineaBase + intSeparacionEntreParrafos
		strParrafo = "<b>" & intItem & ".-</b> El comprador otorgar� el cupo con un c�digo alfanum�rico que obligatoriamente debe consignarse en el campo observaciones de cada Carta de Porte. En caso de que el vendedor remita camiones sin poseer cupo para la descarga, el comprador podr�, a su exclusiva opci�n, proceder a la descarga de los mismos, debiendo en tal caso el vendedor abonar al comprador U$S 10 (diez d�lares) por tonelada, en concepto de gastos extras por descargas no programadas."
		intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_JUSTIFY)    
	end if
		
    if (incluyeBiotecnologia(p_boleto("Producto"), p_boleto("Cosecha"), p_boleto("PuertoRecepcion"), p_boleto("FechaConc"))) then
        'Clausulas de biotecnologia        
        Call GF_dibujarNormasDeBiotecnologiaSoja(p_oPDF, p_boleto, intLineaBase, intItem, intAnchoLinea, intSeparacionEntreParrafos)
	end if
	
	intLineaBase = intLineaBase + intSeparacionEntreParrafos
    strParrafo = "<b>Observaciones: </b> Destino de la mercader�a del Presente Contrato: EXPORTACI�N. El comprador no act�a como comisionista consignatario en esta operaci�n."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    if ucase(p_boleto("MercConHumedad")="V") then
        strParrafo = "Se recibe con Humedad."
        intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, intAltoRenglon, PDF_ALIGN_LEFT)
    end if
    
    
        
    intLineaBase = armarObservacionesDinamicas(p_oPDF, 15, intLineaBase, strTexto, intAnchoLinea, intAltoRenglon, p_boleto)
    
    strParrafo = "Queda establecido que toda tasa, contribuci�n, impuesto provincial y/o municipal que grave la presente operaci�n estar� a cargo de la parte vendedora."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, 7, PDF_ALIGN_JUSTIFY)
    strParrafo = "En muestra de total conformidad las partes firman el boleto en Cinco ejemplares de un mismo tenor y a un solo efecto, uno para cada parte y el triplicado para ser registrado en la " & strSucursal & "."
    intLineaBase = GF_writeTextPlus(p_oPDF, 15, intLineaBase, strParrafo, intAnchoLinea, 7, PDF_ALIGN_JUSTIFY)    

        
    
    intLineaBase = intLineaBase + 5    
    Call armarFirmas(p_oPDF, p_boleto, intAnchoLinea, intLineaBase)
    Call armarPie(p_oPDF, p_boleto, intAnchoLinea, intLineaBase)

if err.number > 0 then
    call GF_closePDF(p_oPDF)
    response.write err.description
end if

end sub
'*****************************************************************************************************************
sub armarClausulaCorredor(p_oPDF, p_boleto, byref p_intLineaBase, byref p_intItem, p_intAnchoLinea, p_intAltoRenglon, p_intSeparacionEntreParrafos)
    dim strTexto

    if (p_boleto("KCCOR") <> "0") then
        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> Los se�ores <b>" & GetDsEnterprise2(p_boleto("KCCOR")) & "</b> act�an en la presente operaci�n como corredores, quedando facultados por los vendedores a endosar en su nombre recibos de mercader�a, ampliaciones y/o anulaciones, convenir eventuales pr�rrogas y a firmar en su nombre y representaci�n toda la documentaci�n necesaria para la instrumentaci�n o formalizaci�n del presente y en particular los formularios C 1116 A y B Res. Conj. SAGPyA� n� 456/03 y DGI n� 1593/03"
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intAnchoLinea,p_intAltoRenglon, PDF_ALIGN_JUSTIFY)
        'if (cstr(p_boleto("MercPropia")) = "true") then
        '    p_intItem = p_intItem + 1
        '    p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        '    strTexto =  "<b>" & p_intItem & ".-</b> El vendedor faculta al corredor a firmar en su nombre y representaci�n toda la documentaci�n necesaria para la instrumentaci�n o formalizaci�n del presente y en particular los formularios C 1116 A y B Res. Conj. SAGPyA� n� 456/03 y DGI n� 1593/03."
        '    p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intAnchoLinea,p_intAltoRenglon,3)
        'end if
    end if
    if ((cint(p_boleto("Operacion")) = 1) or (cint(p_boleto("Operacion")) = 10)) then
        'Boletos a fijar
        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> El precio de la mercader�a objeto del presente contrato, se fijar� a partir del <b>" & GF_FN2DTE(p_boleto("FechaFijaDesde")) & "</b> y a mas tardar hasta el <b>" & GF_FN2DTE(p_boleto("FechaFijaHasta")) & ".</b>"
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)

        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> Multa por incumplimiento <b>10%.</b>"
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)

        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> Fijaciones m�nimas diarias <b>" & p_boleto("KilosMin")/1000 & " TNS</b> y m�ximas diarias <b>" & p_boleto("KilosMax")/1000 & " TNS.</b>"
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)
        
        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> La Fijaci�n la comunicar� el vendedor al comprador el d�a elegido para la fijaci�n de precio. Dicha comunicaci�n deber� hacerse por medio fehaciente. La fijaci�n de los granos que se comercializar�n por el presente contrato se realizar� "
        if (Cint(p_boleto("Producto")) = 9) then
			strTexto = strTexto & "por mercado TOEPFER."
        else
			strTexto = strTexto & "en base al que registre La C�mara Arbitral de Cereales. En caso de que el d�a de la fijaci�n no se registrara cotizaci�n p�blica, las partes podr�n solicitar a la C�mara correspondiente la determinaci�n del precio al que habr� de liquidarse la fijaci�n para los casos en que se opte por la pizarra."
        end if
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intAnchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)

        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> En caso de cesaci�n de pagos, presentaci�n en o declaraci�n de quiebra, o presentaci�n en concurso preventivo de cualquiera de las partes, la otra podr� fijar el precio de la mercader�a pendiente de fijaci�n de acuerdo con lo establecido en la cl�usula 4�. A ese efecto, deber� comunicar su resoluci�n a la otra parte, por medio fehaciente, y se proceder� de inmediato a la liquidaci�n definitiva de la operaci�n."
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)

    end if

    if cstr(p_boleto("CondicionIVA")) = "C" then
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        p_intItem = p_intItem + 1
        strTexto = "<b>" & p_intItem & ".-</b> El vendedor declara que la mercader�a proviene de compraventas con pago en especies, por lo que de acuerdo a lo estipulado por el ART. 6 de la Reg. 2300, el presente contrato no se encuentra sujeto a retenci�n de IVA."
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)
    end if
end sub
'*****************************************************************************************************************
sub armarClausulaFormaDePago(p_oPDF, p_boleto, byref p_intLineaBase, byref p_intItem, p_intAnchoLinea,p_intAltoRenglon, p_intSeparacionEntreParrafos)
    dim intDias
    
    if cint(p_boleto("Operacion"))=1 or cint(p_boleto("Operacion"))=10 then
        ' Boletos a fijar
        if (cint(p_boleto("DiasPago")) <> 0) then
            intDias = cint(p_boleto("DiasPago"))
        elseif ((cint(p_boleto("Producto"))=17) and (cint(p_boleto("Operacion"))=10) and (cint(p_boleto("Cosecha"))=12)) then
			'Para cebada, operacion 10, cosecha 12
			intDias = 4
		elseif (cint(p_boleto("Operacion"))=10) then
            'Dependiendo si son en pesos o en dolares varia la cantidad de dias
            intDias = 2
        else
            intDias = 4
        end if
        strFormaPago = "a los  " & intDias & " d�as h�biles de cada fijaci�n"        
    elseif (ucase(p_boleto("CodigoPago")) = "A") or (ucase(p_boleto("CodigoPago")) = "X") then
        ' Contra entrega de mercader�a
        strFormaPago = "contra entrega de la mercader�a, factura y boleto"
        if len(p_boleto("FechaPago"))>7 then strFormaPago = strFormaPago & " el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)
    elseif ucase(p_boleto("CodigoPago")="I") then
        'Boleto con fecha cierta
        strFormaPago = "con fecha cierta"
        if len(p_boleto("FechaPago"))>7 then strFormaPago = strFormaPago & " el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)
        strFormaPago = strFormaPago & " contra mercader�a entregada"
    else
        strFormaPago = "contra "
        select case ucase(p_boleto("CodigoPago"))            
            case "Z", "E", "K", "T", "M":             	
            	if (p_boleto("CodigoPago") <> "M") then strFormaPago = strFormaPago & "CD, "            	
            	strFormaPago = strFormaPago & "c/gt�a y documentaci�n a entera satisfacci�n del comprador el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)            
            case "D": strFormaPago = strFormaPago & "carta de porte"
            case "R":
                strFormaPago = strFormaPago & "carta de garant�a del corredor"
                if len(p_boleto("FechaPago"))>7 then strFormaPago = strFormaPago & " el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)
            case "J":
                strFormaPago = strFormaPago & "carta de garant�a del corredor y C/D"
                if len(p_boleto("FechaPago"))>7 then strFormaPago = strFormaPago & " el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)
            case "H":
                strFormaPago = strFormaPago & "certificado de dep�sito"
                if len(p_boleto("FechaPago"))>7 then strFormaPago = strFormaPago & " el " & left(GF_FN2DTE(p_boleto("FechaPago")),10)
        end select
    end if
    p_intItem = p_intItem + 1
    p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos        
    strTexto = "<b>" & p_intItem & ".-</b> El pago se realizar� en Buenos Aires, <b>" & strFormaPago & ", el " & p_boleto("PjeParcial") & "%</b> a la orden irrevocable de " & getDsEnterprise2(p_boleto("PagarA")) & ", menos gastos por servicios de acondicionamiento, sellados, etc.,  de corresponder, y el saldo en la liquidaci�n final."
    p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)
    if cint(p_boleto("Operacion"))=9 then
        ' Operacion en U$S
        strTexto = "El pago se har� en pesos moneda argentina. A tal efecto, el vendedor facturar� el precio en pesos, moneda argentina, considerando el tipo de cambio comprador para el D�lar estadounidense fijado por el Banco de la Naci�n Argentina para divisas de exportaci�n conforme al producto objeto del contrato al cierre de las operaciones del d�a h�bil inmediato anterior a la fecha de la facturaci�n. El vendedor deber� emitir la factura siempre y cuando se encuentre la mercader�a descargada y en un plazo no mayor a las 72 horas h�biles de producida la descarga. El saldo en pesos pendiente de cada factura y la calidad resultante, quedar�n fijos y no sujetos a ajuste, actualizaci�n, inter�s o repotenciaci�n alguna, de ninguna naturaleza, por ninguna causa o concepto"
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)
    elseif cint(p_boleto("Operacion"))=10 then
        ' Operacion a fijar en U$S
        strTexto = "El pago se har� en pesos, moneda argentina. A tal efecto, el vendedor facturara el precio en pesos, moneda argentina, considerando el tipo de cambio comprador para el D�lar estadounidense fijado por el Banco de la Naci�n Argentina para divisas de exportaci�n conforme al producto objeto del contrato al cierre de las operaciones del d�a h�bil inmediato anterior a la fecha de la facturaci�n. El saldo en pesos pendiente de cada factura y la calidad resultante, quedaran fijos y no sujetos a ajuste, actualizaci�n, inter�s o repotenciacion alguna, de ninguna naturaleza, por ninguna causa o concepto."
        'strTexto = "El pago se har� en pesos moneda argentina. A tal efecto, el vendedor facturar� el precio en pesos, moneda argentina, considerando el MAT Bs. As. del d�a anterior a la fecha de facturaci�n. // Se puede fijar hasta 1/2 hora antes del cierre del MAT."
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_JUSTIFY)
    end if
end sub
'*****************************************************************************************************************
sub armarClausulaMercaderia(p_oPDF, p_boleto, p_intLineaBase, byref p_intItem, p_intAnchoLinea,p_intAltoRenglon, p_intSeparacionEntreParrafos)
    if cint(p_boleto("KCCOR")) > 0 then
        p_intItem = p_intItem + 1
        p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
        strTexto = "<b>" & p_intItem & ".-</b> "
        strTexto = strTexto & " " & getDsEnterprise2(p_boleto("KCCOR")) & " es intermediario obligado a actuar como agente de Ret. de Impuesto a las Gcias. en la presente operaci�n, conforme lo dispuesto por el Art. 2� Inc. b) de la R. G. 2118/2006 y de acuerdo a lo expresado en el contrato al cual esta documentaci�n corresponde."
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)
    end if
end sub
'*****************************************************************************************************************
sub armarClausulaAFIP(p_oPDF, p_boleto, p_intLineaBase, byref p_intItem, p_intAnchoLinea,p_intAltoRenglon, p_intSeparacionEntreParrafos)    
	p_intItem = p_intItem + 1
	p_intLineaBase = p_intLineaBase + p_intSeparacionEntreParrafos
	strTexto = "<b>" & p_intItem & ".-</b> "
	strTexto = strTexto & "El presente boleto ser� registrado ante la Afip en cumplimiento de la RG AFIP n� 2596 y sus modificatorias. Resulta condici�n esencial para la vigencia y perfeccionamiento de este boleto la obtenci�n de la Constancia establecida por dicha Resoluci�n General. Para el caso de no obtenerse la referida Constancia, el presente boleto quedar� sin efecto y se considerar� como no celebrado en los t�rminos del art. 548 primera parte del C�digo Civil, sin obligaci�n o responsabilidad alguna en cabeza del Comprador. Si dado el supuesto anterior, el Comprador ya hubiera recibido mercader�a, entonces la pondr� a disposici�n del Vendedor, siendo a cargo de este �ltimo los gastos ocasionados por descarga, almacenaje, acondicionamiento y retiro de la misma."
	p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intanchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)    
end sub
'*****************************************************************************************************************
function GF_dibujarOpcionesMercaderia(byref p_oPDF, byref p_intLineaBase)


    call GF_writeImage(p_oPDF, server.mapPath("images/checkbox.jpg"), 15, p_intLineaBase, 10, 10, 0)
    call GF_writeTextPlus(p_oPDF, 30, p_intLineaBase + 3, "es de propia producci�n", 100,15,PDF_ALIGN_LEFT)
    
    call GF_writeImage(p_oPDF, server.mapPath("images/checkbox.jpg"), 140, p_intLineaBase, 10, 10, 0)
    call GF_writeTextPlus(p_oPDF, 155, p_intLineaBase + 3, "es venta en consignaci�n por cuenta y orden de varios comitentes",300,15,PDF_ALIGN_LEFT)
    
    call GF_writeImage(p_oPDF, server.mapPath("images/checkbox.jpg"), 420, p_intLineaBase, 10, 10, 0)
    call GF_writeTextPlus(p_oPDF, 435, p_intLineaBase + 3, "no es de su propia producci�n",200,15,PDF_ALIGN_LEFT)

    p_intLineaBase = p_intLineaBase + 12
end function
'*****************************************************************************************************************
sub armarPie(p_oPDF, p_boleto, p_intAnchoLinea, p_intLineaBase)
    dim strText, strEntidad, strVendedor, strCorredor
    
    p_intLineaBase = 815    
    call GF_setFont(p_oPDF, "Arial", 6, 0)
    
    strEntidadAux = p_boleto("Entidad_Descripcion")
    strTexto = "La C�mara Arbitral de la " & strEntidadAux & " no intervendr� en ninguna cuesti�n que suscite como consecuencia del presente boleto si el mismo no est� registrado en la " & strEntidadAux & ". No se admitir�n en este boleto enmiendas, raspaduras ni agregados, si no est�n debidamente salvados. Los ejemplares en poder de las partes contratantes deber�n ser exactamente iguales al ejemplar que queda registrado en la " & strEntidadAux & "."
    call GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strTexto, p_intAnchoLinea, 7, PDF_ALIGN_LEFT)
end sub
'*****************************************************************************************************************
function armarObservacionesDinamicas(p_oPDF, p_x, p_intLineaBase, p_strTexto, p_intAnchoLinea, p_intAltoRenglon, p_boleto)
    dim strSQL, conn, rs, palabras, texto, linea
    'strSQL = "select Observacion from ObservacionesBoleto where Producto=" & p_boleto("Producto") & " and Sucursal=" & p_boleto("Sucursal") & " and Operacion=" & p_boleto("Operacion") & " and Numero=" & p_boleto("Numero") & " and Cosecha=" & p_boleto("Cosecha") & " order by Renglon asc"
    strSQL="Select * from MERFL.MER311FP where CPRORP=" & p_boleto("Producto") & " AND COPERP=" & p_boleto("Operacion") & " and CSUCRP=" & p_boleto("Sucursal") & " and NCTORP=" & p_boleto("Numero") & " and ACOSRP=" & p_boleto("Cosecha") & " order by NUMRRP"
    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
    longitudLinea=0
    linea = ""
    texto = ""
    while not rs.eof
        texto = texto & Trim(rs("OBSERP")) & " "
        rs.movenext
    wend
	call GF_setFont(p_oPDF, "Times", 7, 0)
    p_intLineaBase = GF_writeTextPlus(p_oPDF, p_x, p_intLineaBase, UCase(texto), p_intAnchoLinea, p_intAltoRenglon, PDF_ALIGN_LEFT)	
    armarObservacionesDinamicas = p_intLineaBase
end function
'*****************************************************************************************************************
Function armarFirmas(p_oPDF, p_boleto, p_intAnchoLinea, p_intLineaBase)

    call GF_setFont(p_oPDF, "Arial", 6, 0)

    strVendedor = trim(getDsEnterprise2(p_boleto("KCVEN")))

    dim wth
    wth = Int(p_oPDF.Metrics.GetTextWidth(strVendedor, pdf_currentFont, pdf_currentFontSize))
    x = cint(50 - wth/2)
    IF x <= 0 THEN x = 1
    call GF_writeTextAlign(p_oPDF, x, p_intLineaBase+7, strVendedor, 560, PDF_ALIGN_LEFT)

    if p_boleto("KCCOR")<>0 then
        strCorredor = trim(getDsEnterprise2(p_boleto("KCCOR")))
	    wth = Int(p_oPDF.Metrics.GetTextWidth(strCorredor, pdf_currentFont, pdf_currentFontSize))        
        x = cint(280 - wth/2)
        call GF_writeTextAlign(p_oPDF, x, p_intLineaBase, strCorredor, 560, PDF_ALIGN_LEFT)
    end if

    call GF_writeTextAlign(p_oPDF, 420, p_intLineaBase+7, p_boleto("RSCOMP"), 560, PDF_ALIGN_LEFT)

    call GF_setFont(p_oPDF, "Arial", 7, 0)

    p_intLineaBase = p_intLineaBase + 36
    call GF_writeTextAlign(p_oPDF, 15 , p_intLineaBase, ".....................................", 560, PDF_ALIGN_LEFT)
    if p_boleto("KCCOR")<>0 then call GF_writeTextAlign(p_oPDF, 245, p_intLineaBase, ".....................................", 560, 0)
    call GF_writeTextAlign(p_oPDF, 465, p_intLineaBase, ".....................................", 560, PDF_ALIGN_LEFT)

    p_intLineaBase = p_intLineaBase + 7
    call GF_writeTextAlign(p_oPDF, 16, p_intLineaBase, "           Vendedor         ", p_intAnchoLinea, PDF_ALIGN_LEFT)
    if p_boleto("KCCOR")<>0 then call GF_writeTextAlign(p_oPDF, 246, p_intLineaBase, "          Corredor         ", p_intAnchoLinea, 0)
    call GF_writeTextAlign(p_oPDF, 465, p_intLineaBase, "          Comprador        ", p_intAnchoLinea, PDF_ALIGN_LEFT)

    p_intLineaBase = p_intLineaBase + 9
    call GF_writeTextAlign(p_oPDF, 20, p_intLineaBase, "CUIT: " & formatear(p_boleto("VEN_CUIT"),"CUIT"), 560, PDF_ALIGN_LEFT)
    if p_boleto("KCCOR")<>0 then call GF_writeTextAlign(p_oPDF, 250, p_intLineaBase, "CUIT: " & formatear(p_boleto("COR_CUIT"),"CUIT"), 560, 0)
    call GF_writeTextAlign(p_oPDF, 470, p_intLineaBase, "CUIT: 30-62197317-3", 560, PDF_ALIGN_LEFT)
End Function
'*****************************************************************************************************************
function GF_generarBoleto(p_oPDF, p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha)
    dim oPDF, diccBoleto
		GF_generarBoleto = false
       if (GF_establecerValoresBoleto(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha, diccBoleto) = true) then 
           call GF_dibujarEncabezado(p_oPDF, diccBoleto)
           call GF_dibujarMarcaDeAgua(p_oPDF)
           call GF_dibujarClausulas(p_oPDF, diccBoleto)
           if (cInt(p_intProducto) = 9) then 
           'normas de calidad para colza (cod producto 9)
				Call GF_newPage(p_oPDF)
				Call GF_dibujarNormasDeCalidad(p_oPDF, diccBoleto)
		   end if
		   if ((cInt(p_intProducto) = 17) and ((cint(p_intOperacion) = 0) or (cint(p_intOperacion) = 9) or (cint(p_intOperacion) = 10) or (cint(p_intOperacion) = 01))) then 
           'normas de calidad para cebada (cod producto 17 y cod. operacion = 9)
				Call GF_newPage(p_oPDF)
				Call GF_dibujarNormasDeCalidadCebada(p_oPDF, diccBoleto,p_intCosecha, p_intOperacion)
		   end if
		   		   
           if err.number > 0 then
                response.write Err.Description
           else
                GF_generarBoleto = true
           end if
       else
            response.write "Contrato inv�lido"
       end if
end function
'*****************************************************************************************************************
'FUNCION PARA IMPRIMIR LA TABLA DE LLAVES DE SEGURIDAD.
'NO SE USA EN NINGUN LADO PUES NO DEBE NUNCA PODER GENERARSE SALVO QUE SE HAGA EL LLAMADO A MANO!!!
Function GF_PRINT_SEC_TABLE(p_oPDF)

    'Set sPDF = GF_createPDF(Server.MapPath("temp\SecTable.pdf"))
    call GF_newPage(Gbl_oPDF)
    Call GF_writeImage(p_oPDF, server.mapPath("images/H_Series_Key.jpg"),0,30,39,37,0)
    'Call GF_closePDF(sPDF)
End Function
'*****************************************************************************************************************
function generarPDF(p_Producto, p_Sucursal, p_Operacion, p_Numero, p_Cosecha, p_KCVEN)
    dim oPDF, fs, strPath

    on error resume next

    'Establezco la ruta y el nombre del PDF a crear
    strResto = replace(replace(GF_EDIT_CONTRATO(p_producto, p_Sucursal, p_Operacion, p_Numero, p_Cosecha) & " " & replace(GetDsEnterprise2(p_KCVEN),".","") & ".pdf","/","-"), ",", "")
    strPath = Server.mapPath("temp\") & "\" & strResto

    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
'    response.write strPath & "<br>"
    If fs.FileExists(strPath) Then
        call fs.deleteFile(strPath, true)
    end if
    if Err.Number <> 0 and Err.Number <> -2147217900 then 'Para q no tire error de la SessionHeader
        call GP_ENVIAR_MAIL("Error Envio Autom. de mails", Err.Number & CHR(13) & CHR(10) & Err.Description & CHR(13) & CHR(10) & CHR(13) & CHR(10) & strPath & chr(13) & chr(10) & "No puede borrar el archivo", strToepferDenomination & " <" & SENDER_MERCADERIAS & ">","santij@toepfer.com;scalisij@toepfer.com")
        'response.write Err.Description & "no pudo crear - "
        generarPDF = false
    else           
        set fs = nothing
        set oPDF = GF_createPDF(strPath)
        ret = GF_GenerarBoleto(oPDF, p_Producto, p_Sucursal, p_Operacion, p_Numero, p_Cosecha)
        generarPDF = ret
        Call GF_closePDF(oPDF)
    end if
end function
'*****************************************************************************************
sub enviarMailBoleto(byref p_rs, p_cantEnv)
    dim strDestinatario, strAsunto, strPathAttachment, ORKC
    dim vecMails(1), vecMailsToepfer(1)

    'Establezco si el destinatario es el corredor o el vendedor
    if cdbl(p_rs("KCCOR")) > 0 then
        ORKC = p_rs("KCCOR")
    else
        ORKC = p_rs("KCVEN")
    end if
    'completo los datos del mail
    strAsunto = "Boleto Toepfer Negocio " & GF_EDIT_Contrato(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha"))
    strPathAttachment = Server.mapPath("temp/") & "\" & replace(replace(GF_EDIT_Contrato(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha")) & " " & replace(GetDsEnterprise2(p_rs("KCVEN")),".","") & ".pdf","/","-"), ",", "")
    strToepferDenomination = GetDsEnterprise2("99999997")
    strBody = "Se adjunta al presente mail el boleto de compra/venta correspondiente al contrato " & GF_EDIT_Contrato(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha")) & "." & chr(13) & chr(10) & chr(13) & chr(10)
    strBody = strBody & "                  " & strToepferDenomination

    'Busco los mails del destinatario y envio
    call obtenerMailBoletos(ORKC, vecMails)

    'Para ponerlo en Prod hay q sacar la linea de abajo (q me manda el mail a mi),
    'poner que elija los boletos a partir de una determinada fecha y
    'una vez enviado el boleto que le ponga la marca de enviado
    if ((not isnull(vecMails(0))) and (vecMails(0) <> "")) then strDestinatario = vecMails(0) & "; "
    if ((not isnull(vecMails(1))) and (vecMails(1) <> "")) then strDestinatario = strDestinatario & vecMails(1) & ";"            
    
    if (strDestinatario <> "") then
		Call actualizarBoleto(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha"), MRCENVIO_T, session("MomentoSistema"), 0, BOLETO_UPDATE)
        Call GP_ENVIAR_MAIL_ATTACHMENT(strAsunto, strBody,strToepferDenomination & " <" & SENDER_MERCADERIAS & ">", strDestinatario, strPathAttachment)
        'Pongo la marca de enviado a V(erdadero)
        Call actualizarBoleto(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha"), MRCENVIO_V, session("MomentoSistema"), p_cantEnv, BOLETO_UPDATE)	%>
        <tr>
            <td align=left style="left-padding:10px;">
                Enviado a: <%=strDestinatario%>
            </td>
        </tr>
        <% Call writeLog("INF", "Boleto " & GF_EDIT_CONTRATO(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha")) & " enviado." )
	else    
        call writeLog("WRN", "Boleto " & GF_EDIT_Contrato(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha")) & " no enviado por falta de destinatarios")
        Call actualizarBoleto(p_rs("Producto"), p_rs("Sucursal"), p_rs("Operacion"), p_rs("Numero"), p_rs("Cosecha"), MRCENVIO_N, session("MomentoSistema"), 0, BOLETO_UPDATE)%>    
        <tr>
            <td align=left style="left-padding:10px;">No se ha podido enviar el boleto, debido a que no se han establecido las direcciones de mail donde enviarlos.</td>
        </tr>
        <tr>
            <td align=left style="left-padding:10px;">Cargue las direcciones de E-Mail haciendo click <a href="datosPersonales.asp">aqui</a>.</td>
        </tr>            		
	<% end if
end sub
'***************************************************************************************
'est6a funcion crea o actualiza un boleto en la ControlBoleto
'si es insert se solicita el numero de cotrato completo
'si es update se solicita que algun campo a actualizar tenga valor
Function actualizarBoleto(p_producto, p_sucursal, p_operacion, p_numero, p_cosecha, p_MrcEnvio, p_mmtoEnvio, p_CantEnvio, p_tipo)
	Dim strSql, rs, conn, aux
	aux = false
	if (p_tipo = BOLETO_INSERT) then
		if (p_producto <> "" and p_sucursal <> "" and p_operacion <> "" and p_numero <> "" and p_cosecha <> "") then
			strSQL = "insert into ControlBoleto values (" & p_producto & ", " & p_sucursal & ", " & p_operacion & ", " & p_numero & ", " & p_cosecha & ", '" & get_mrcEnvio(p_MrcEnvio) & "', " & p_mmtoEnvio & ", " & p_CantEnvio & ")"
			call GF_BD_AS400(rs, conn, "EXEC", strSQL)
			aux = true
		end if
	elseif (p_tipo = BOLETO_UPDATE) then
		if (p_MrcEnvio <> "" or p_mmtoEnvio <> "") then
			if (p_CantEnvio = 0) then
				strSQL = "update ControlBoleto set mrcEnvioBoleto = '" & get_mrcEnvio(p_MrcEnvio) & "', MmtoEnvioAutomatico=" & p_mmtoEnvio & " where producto=" & p_producto & " and sucursal=" & p_sucursal & " and operacion=" & p_operacion & " and numero=" & p_numero & " and cosecha=" & p_cosecha
			else
				strSQL = "update ControlBoleto set mrcEnvioBoleto = '" & get_mrcEnvio(p_MrcEnvio) & "', Cant_env =" & p_CantEnvio & ", MmtoEnvioAutomatico=" & p_mmtoEnvio & " where producto=" & p_producto & " and sucursal=" & p_sucursal & " and operacion=" & p_operacion & " and numero=" & p_numero & " and cosecha=" & p_cosecha
			end if
			call GF_BD_AS400(rs, conn, "EXEC", strSQL)
			aux = true
		end if
	end if
    actualizarBoleto = aux
end Function
'***************************************************************************************
'esta funci�n toma la constante que recivio como parametro de marca de envio para insertar o actualizar la base de daos
'y devuleve el valor correspondiente
Function get_mrcEnvio(p_MrcEnvio)
	Dim mrcEnvioAux
	Select case p_MrcEnvio
		case MRCENVIO_F: mrcEnvioAux = "F"
		case MRCENVIO_X: mrcEnvioAux = "X"
		case MRCENVIO_N: mrcEnvioAux = "N"
		case MRCENVIO_T: mrcEnvioAux = "T"
		case MRCENVIO_V: mrcEnvioAux = "V"
		case else: mrcEnvioAux = "F"
	end Select
	get_mrcEnvio = mrcEnvioAux
end Function
'***************************************************************************************
Function GF_dibujarNormasDeBiotecnologiaSoja(p_oPDF, p_boleto, ByRef p_intLineaBase, ByRef p_intItem, p_intAnchoLinea, p_intSeparacionEntreParrafos)
    dim aux, strParrafo
    
        p_intItem = p_intItem + 1		
	    p_intLineaBase = p_intLineaBase +  p_intSeparacionEntreParrafos        
        strParrafo = "<b>" & p_intItem & ".-</b> El vendedor acepta que el grano de soja ser� analizado y en caso de detectarse la presencia de tecnolog�as patentadas se le descontar�, de corresponder, el importe de la regal�a correspondiente por cuenta y orden del propietario de la tecnolog�a o de quien �ste designe. Toda controversia derivada de la aplicaci�n de esta cl�usula ser� resuelta con el propietario de la tecnolog�a patentada por la C�mara Arbitral de Cereales de la jurisdicci�n donde se registre el presente documento o de donde el cargamento es entregado. El tribunal elegido actuar� como amigable componedor, con aplicaci�n de las Reglas y Usos del Comercio de Granos y del Reglamento de Procedimientos aprobado por Decreto 931/98 y/o sus normas complementarias. La ejecuci�n del laudo arbitral se efectuar� ante los Tribunales Ordinarios de la Ciudad Aut�noma de Buenos Aires."
        p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strParrafo, p_intAnchoLinea, 7, PDF_ALIGN_JUSTIFY)	                
        
        'Cierro la pagina actual.
        'Call GF_setFont(p_oPDF, "Arial", 10, 0)
        'strParrafo = "-------- Contin�a en la proxima pagina --------"
        'Call GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strParrafo, p_intAnchoLinea, 10, PDF_ALIGN_CENTER)
        'Call armarPie(p_oPDF, p_boleto, p_intAnchoLinea, p_intLineaBase)        
        
        'Call GF_newPage(p_oPDF)
        'Call GF_dibujarEncabezado(p_oPDF, p_boleto)
        'Call GF_dibujarMarcaDeAgua(p_oPDF)
        'Call GF_setFont(p_oPDF, "Times", 7, 0)
        'p_intLineaBase = 185
	
        'Clausulas de biotecnologia - 2da parte, la 1ra esta al final de las calusulas comunes.    
        'p_intItem = p_intItem + 1
        'p_intLineaBase = p_intLineaBase +  p_intSeparacionEntreParrafos        
        'strParrafo = "<b>" & p_intItem & ".-</b> Como excepci�n, para las entregas de grano de soja que se realicen en la campa�a 2014/2015, la cl�usula precedente s�lo ser� de aplicaci�n a los cargamentos de grano de soja cuyo origen sea las provincias de Salta, Chaco, Jujuy, Catamarca, Tucum�n, Formosa y Santiago del Estero  y los departamentos de San Justo, 9 de Julio, General Obligado, San Crist�bal, Vera y San Javier de la Provincia de Santa Fe."
        'p_intLineaBase = GF_writeTextPlus(p_oPDF, 15, p_intLineaBase, strParrafo, 560, 7, PDF_ALIGN_JUSTIFY)            
End Function
'***************************************************************************************
'funcion que arma la hoja de normas de calidad para la comercializacion de colza (cod Prod = 9)
Function GF_dibujarNormasDeCalidad(p_oPDF, p_boleto)
    Call GF_WM_diagonal(p_oPDF,40,20,6)
    Call GF_WM_diagonal(p_oPDF,300,20,3)
    Call GF_WM_diagonal(p_oPDF,40,220,6)
    Call GF_WM_diagonal(p_oPDF,40,420,6)
    Call GF_WM_diagonal(p_oPDF,40,620,3)
	Call dibujarBoxNormas(p_oPDF)
	Call writeInfo(p_oPDF, p_boleto)
	Call writeFirmas(p_oPDF, p_boleto)
end Function
'***************************************************************************************
'funcion que arma la hoja de normas de calidad para la comercializacion de cebada (cod Producto = 17)
Function GF_dibujarNormasDeCalidadCebada(p_oPDF, p_boleto,p_Cosecha,p_operacion)
    Call GF_WM_diagonal(p_oPDF,40,20,6)
    Call GF_WM_diagonal(p_oPDF,300,20,3)
    Call GF_WM_diagonal(p_oPDF,40,220,6)
    Call GF_WM_diagonal(p_oPDF,40,420,6)
    Call GF_WM_diagonal(p_oPDF,40,620,3)
	Call dibujarBoxNormasCebada(p_oPDF)
	Call writeInfoCebada(p_oPDF, p_boleto,p_Cosecha,p_operacion)
	Call writeFirmas(p_oPDF, p_boleto)
end Function
'***************************************************************************************
Function dibujarBoxNormas(p_oPDF)'es para colza (cod prod = 9)
	'se arma el cuadro de la info
	Call GF_horizontalLine(p_oPDF, 40, 20, 510)
	Call GF_horizontalLine(p_oPDF, 40, 70, 510)
	Call GF_horizontalLine(p_oPDF, 40, 100, 510)
	Call GF_horizontalLine(p_oPDF, 40, 180, 510)
	Call GF_horizontalLine(p_oPDF, 40, 280, 510)
	Call GF_horizontalLine(p_oPDF, 40, 420, 510)
	Call GF_horizontalLine(p_oPDF, 40, 450, 510)
	Call GF_horizontalLine(p_oPDF, 40, 540, 510)
	Call GF_horizontalLine(p_oPDF, 40, 600, 510)
	Call GF_verticalLine(p_oPDF, 40, 20, 580)
	Call GF_verticalLine(p_oPDF, 550, 20, 580)
	Call GF_verticalLine(p_oPDF, 140, 70, 530)
	Call GF_verticalLine(p_oPDF, 220, 70, 530)
	Call GF_verticalLine(p_oPDF, 310, 70, 530)
	Call GF_verticalLine(p_oPDF, 410, 70, 530)
	'se arma el cuadro de firmas
	Call GF_horizontalLine(p_oPDF, 40, 690, 510)
	Call GF_horizontalLine(p_oPDF, 40, 720, 510)
	Call GF_horizontalLine(p_oPDF, 40, 790, 510)
	Call GF_horizontalLine(p_oPDF, 40, 830, 510)
	Call GF_verticalLine(p_oPDF, 40, 690, 140)
	Call GF_verticalLine(p_oPDF, 210, 690, 140)
	Call GF_verticalLine(p_oPDF, 380, 690, 140)
	Call GF_verticalLine(p_oPDF, 550, 690, 140)
end Function
'***************************************************************************************
Function dibujarBoxNormasCebada(p_oPDF)
	'se arma el cuadro de la info
	Call GF_horizontalLine(p_oPDF, 40, 20, 510)
	Call GF_horizontalLine(p_oPDF, 40, 70, 510)
	Call GF_horizontalLine(p_oPDF, 40, 100, 510)
	Call GF_horizontalLine(p_oPDF, 40, 160, 510)
	Call GF_horizontalLine(p_oPDF, 40, 250, 510)
	Call GF_horizontalLine(p_oPDF, 40, 330, 510)
	Call GF_horizontalLine(p_oPDF, 40, 410, 510)
	Call GF_horizontalLine(p_oPDF, 40, 490, 510)
	Call GF_horizontalLine(p_oPDF, 40, 540, 510)
	Call GF_horizontalLine(p_oPDF, 40, 600, 510)
	Call GF_horizontalLine(p_oPDF, 40, 660, 510)
	Call GF_verticalLine(p_oPDF, 40, 20, 640)
	Call GF_verticalLine(p_oPDF, 550, 20, 640)
	Call GF_verticalLine(p_oPDF, 120, 70, 590)
	Call GF_verticalLine(p_oPDF, 180, 70, 590)
	Call GF_verticalLine(p_oPDF, 260, 70, 590)
	Call GF_verticalLine(p_oPDF, 350, 70, 470)
	Call GF_verticalLine(p_oPDF, 350, 600, 60)
	Call GF_verticalLine(p_oPDF, 450, 70, 530)
	
	'se arma el cuadro de firmas
	Call GF_horizontalLine(p_oPDF, 40, 690, 510)
	Call GF_horizontalLine(p_oPDF, 40, 720, 510)
	Call GF_horizontalLine(p_oPDF, 40, 790, 510)
	Call GF_horizontalLine(p_oPDF, 40, 830, 510)
	Call GF_verticalLine(p_oPDF, 40, 690, 140)
	Call GF_verticalLine(p_oPDF, 210, 690, 140)
	Call GF_verticalLine(p_oPDF, 380, 690, 140)
	Call GF_verticalLine(p_oPDF, 550, 690, 140)
end Function
'***************************************************************************************
Function writeInfo(p_oPDF, p_boleto) 'es para colza (09)
	dim aux
	'texto info
    call GF_setFont(p_oPDF, "Times",10,0)
    'titulo
	aux = "NORMAS  DE  CALIDAD  PARA  LA  COMERCIALIZACION  DE  " & Ucase(getDsProduct(clng(p_boleto("Producto")))) & "  NORMA  VIII"
	Call GF_writeTextPlus(p_oPDF, 40, 22, aux, 510, 20, PDF_ALIGN_CENTER)
	aux = "ALFRED  C.  TOEPFER  INTERNATIONAL  ARGENTINA  SRL"
	Call GF_writeTextAlign(p_oPDF, 40, 40, aux, 510, PDF_ALIGN_CENTER)
	aux = "ANEXO  DE  BOLETO  DE  COMPRAVENTA"
	Call GF_writeTextAlign(p_oPDF, 40, 58, aux, 510, PDF_ALIGN_CENTER)
    'cuadro
    call GF_setFont(p_oPDF, "Times",9,8)
	aux = "RUBROS"
	Call GF_writeTextAlign(p_oPDF, 40, 80, aux, 100, PDF_ALIGN_CENTER)
	aux = "BASES"
	Call GF_writeTextAlign(p_oPDF, 140, 80, aux, 80, PDF_ALIGN_CENTER)
	aux = "TOLERANCIA  DE"
	Call GF_writeTextAlign(p_oPDF, 220, 75, aux, 90, PDF_ALIGN_CENTER)
	aux = "RECIBO"
	Call GF_writeTextAlign(p_oPDF, 220, 85, aux, 90, PDF_ALIGN_CENTER)
	aux = "BONIFICACIONES"
	Call GF_writeTextAlign(p_oPDF, 310, 80, aux, 100, PDF_ALIGN_CENTER)
	aux = "REBAJAS"
	Call GF_writeTextAlign(p_oPDF, 410, 80, aux, 140, PDF_ALIGN_CENTER)
    call GF_setFont(p_oPDF, "Times",9,0)
    '1 CONTENIDO DE MATERIA GRASA
	aux = "CONTENIDO DE MATERIA GRASA S.S.S Y L (1)"
	Call GF_writeTextPlus(p_oPDF, 50, 120, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "43 %"
	Call GF_writeTextAlign(p_oPDF, 140, 130, aux, 80, PDF_ALIGN_CENTER)
	aux = "Para valores superiores a 43% a raz�n de 1% por cada por ciento o fracci�n proporcional."
	Call GF_writeTextPlus(p_oPDF, 315, 105, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Para valores inferiores a 43% a raz�n de 1% por cada por ciento o fracci�n proporcional."
	Call GF_writeTextPlus(p_oPDF, 415, 105, aux, 130, 10, PDF_ALIGN_LEFT)
	'2 ACIDES DE LA MATERIA GRASA
	aux = "ACIDES DE LA MATERIA GRASA"
	Call GF_writeTextPlus(p_oPDF, 50, 215, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "1,0 %"
	Call GF_writeTextAlign(p_oPDF, 140, 220, aux, 80, PDF_ALIGN_CENTER)
	aux = "1,5 %"
	Call GF_writeTextAlign(p_oPDF, 220, 220, aux, 90, PDF_ALIGN_CENTER)
	aux = "NO CORRESPONDE"
	Call GF_writeTextPlus(p_oPDF, 310, 220, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "Para valores superiores a 1% y hasta 1,5% a raz�n de 2,5% por cada por ciento o fracci�n proporcional.<br> <br>Para valores superiores a 1,5% a raz�n de 5% por cada por ciento o fracci�n proporcional."
	Call GF_writeTextPlus(p_oPDF, 415, 185, aux, 130, 10, PDF_ALIGN_LEFT)
	'3 CUERPOS EXTRA�OS
	'if ((p_boleto("Cosecha") = 13) and ((p_boleto("Operacion") = 9) or (p_boleto("Operacion") = 10))) then
	if (p_boleto("Cosecha") >= 13) then
		aux = "CUERPOS EXTRA�OS (4)"
		Call GF_writeTextPlus(p_oPDF, 40, 340, aux, 100, 10, PDF_ALIGN_CENTER)
		aux = "4 %"
		Call GF_writeTextAlign(p_oPDF, 220, 340, aux, 90, PDF_ALIGN_CENTER)
		aux = "NO CORRESPONDE"
		Call GF_writeTextPlus(p_oPDF, 310, 340, aux, 100, 50, PDF_ALIGN_CENTER)
		aux = "Hasta la tolerancia de recibo 4% a raz�n de 1% por cada por ciento o fracci�n proporcional.<br> <br>Para valores superiores a 4% ser� rechazo."
		Call GF_writeTextPlus(p_oPDF, 415, 285, aux, 130, 10, PDF_ALIGN_LEFT)
	else
		aux = "CUERPOS EXTRA�OS "
		Call GF_writeTextPlus(p_oPDF, 40, 340, aux, 100, 10, PDF_ALIGN_CENTER)
		aux = "5 %"
		Call GF_writeTextAlign(p_oPDF, 220, 340, aux, 90, PDF_ALIGN_CENTER)
		aux = "NO CORRESPONDE"
		Call GF_writeTextPlus(p_oPDF, 310, 340, aux, 100, 50, PDF_ALIGN_CENTER)
		aux = "Hasta la tolerancia de recibo 5% a raz�n de 1% por cada por ciento o fracci�n proporcional.<br> <br>Para valores superiores a 5% a raz�n de 1,5% por cada por ciento o fracci�n proporcional."
		Call GF_writeTextPlus(p_oPDF, 415, 285, aux, 130, 10, PDF_ALIGN_LEFT)			
	end if
	'4 HUMEDAD
    call GF_setFont(p_oPDF, "Times",9,0)
	aux = "HUMEDAD"
	Call GF_writeTextPlus(p_oPDF, 40, 430, aux, 100, 10, PDF_ALIGN_CENTER)
    call GF_setFont(p_oPDF, "Times",9,0)
	aux = "8,5 %"
	Call GF_writeTextAlign(p_oPDF, 140, 430, aux, 80, PDF_ALIGN_CENTER)
	aux = "NO CORRESPONDE"
	Call GF_writeTextPlus(p_oPDF, 310, 430, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "(2)"
	Call GF_writeTextPlus(p_oPDF, 415, 430, aux, 130, 10, PDF_ALIGN_CENTER)
	'5 ACIDO ERUCICO
	aux = "ACIDO ERUCICO"
	Call GF_writeTextPlus(p_oPDF, 40, 490, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "2,0 %"
	Call GF_writeTextAlign(p_oPDF, 220, 490, aux, 90, PDF_ALIGN_CENTER)
	aux = "NO CORRESPONDE"
	Call GF_writeTextPlus(p_oPDF, 310, 490, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "Para contenidos superiores a 2% a raz�n de 2 puntos por cada por ciento o fracci�n proporcional."
	Call GF_writeTextPlus(p_oPDF, 415, 455, aux, 130, 10, PDF_ALIGN_LEFT)
	'6 GLUCOSINOLATOS
	aux = "GLUCOSINOLATOS"
	Call GF_writeTextPlus(p_oPDF, 40, 560, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "20,0 (3)"
	Call GF_writeTextAlign(p_oPDF, 220, 560, aux, 90, PDF_ALIGN_CENTER)
	aux = "NO CORRESPONDE"
	Call GF_writeTextPlus(p_oPDF, 310, 560, aux, 100, 10, PDF_ALIGN_CENTER)
	aux = "Para contenidos superiores a    <br>20 a raz�n de 1 puntos por cada micromol en exceso."
	Call GF_writeTextPlus(p_oPDF, 415, 545, aux, 130, 10, PDF_ALIGN_CENTER)
    'texto llamados
	call GF_setFont(p_oPDF, "Times",10,8)
	aux = "LIBRE  DE  INSECTOS  Y/O  ARACNIDOS  VIVOS"
	Call GF_writeTextAlign(p_oPDF, 50, 620, aux, 450, PDF_ALIGN_LEFT)
	call GF_setFont(p_oPDF, "Times",7,0)
	aux = "(1)  Sobre sustancia seca y limpia."
	Call GF_writeTextAlign(p_oPDF, 60, 635, aux, 450, PDF_ALIGN_LEFT)
	aux = "(2) "
	Call GF_writeTextAlign(p_oPDF, 60, 645, aux, 450, PDF_ALIGN_LEFT)
	'if ((p_boleto("Cosecha") = 13) and ((p_boleto("Operacion") = 9) or (p_boleto("Operacion") = 10))) then
	if (p_boleto("Cosecha") >= 13) then
		aux = "Cuando la mercader�a exceda la base de humedad (8,5%) se acondicionar� con los gastos de secada a convenir."
	else
		aux = "Cuando la mercader�a exceda la base de humedad (8,5%) se descontara la merma correspondiente de acuerdo a las tablas establecidas por el IASCAV y que forman parte de la presente Norma de Clasificaci�n."
	end if
	Call GF_writeTextPlus(p_oPDF, 74, 645, aux, 430, 10, PDF_ALIGN_LEFT)
	aux = "(3)  En micromoles por gramo de grano base 8,5% de humedad."	
	'if ((p_boleto("Cosecha") = 13) and ((p_boleto("Operacion") = 9) or (p_boleto("Operacion") = 10))) then
	if (p_boleto("Cosecha") >= 13) then
		Call GF_writeTextAlign(p_oPDF, 60, 655, aux, 450, PDF_ALIGN_LEFT)
		aux = "(4)  Libre de granos de ma�z y soja."
		Call GF_writeTextAlign(p_oPDF, 60, 665, aux, 450, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(p_oPDF, 60, 665, aux, 450, PDF_ALIGN_LEFT)
	end if
end Function
'***************************************************************************************
Function writeInfoCebada(p_oPDF, p_boleto,p_cosecha,p_operacion) '(producto 17)
	dim aux
	'texto info
    call GF_setFont(p_oPDF, "Times",10,0)
    'titulo
	aux = "NORMAS  DE  CALIDAD  PARA  LA  COMERCIALIZACION  DE  " & Ucase(getDsProduct(clng(p_boleto("Producto")))) & " CERVECERA"
	Call GF_writeTextPlus(p_oPDF, 40, 22, aux, 510, 20, PDF_ALIGN_CENTER)
	aux = "ALFRED  C.  TOEPFER  INTERNATIONAL  ARGENTINA  SRL"
	Call GF_writeTextAlign(p_oPDF, 40, 40, aux, 510, PDF_ALIGN_CENTER)
	aux = "ANEXO  DE  BOLETO  DE  COMPRAVENTA"
	Call GF_writeTextAlign(p_oPDF, 40, 58, aux, 510, PDF_ALIGN_CENTER)
    'cuadro
    call GF_setFont(p_oPDF, "Times",9,8)
	aux = "RUBROS"
	Call GF_writeTextAlign(p_oPDF, 40, 80, aux, 80, PDF_ALIGN_CENTER)
	aux = "BASES"
	Call GF_writeTextAlign(p_oPDF, 120, 80, aux, 60, PDF_ALIGN_CENTER)
	aux = "TOLERANCIA  DE"
	Call GF_writeTextAlign(p_oPDF, 180, 75, aux, 80, PDF_ALIGN_CENTER)
	aux = "RECIBO"
	Call GF_writeTextAlign(p_oPDF, 180, 85, aux, 80, PDF_ALIGN_CENTER)
	aux = "BONIFICACIONES"
	Call GF_writeTextAlign(p_oPDF, 260, 80, aux, 90, PDF_ALIGN_CENTER)
	aux = "REBAJAS"
	Call GF_writeTextAlign(p_oPDF, 350, 80, aux, 100, PDF_ALIGN_CENTER)
	aux = "OBSERVACIONES"
	Call GF_writeTextAlign(p_oPDF, 450, 80, aux, 100, PDF_ALIGN_CENTER)
    call GF_setFont(p_oPDF, "Times",9,0)
    '1 Capacidad germinativa
	aux = "CAPACIDAD GERMINATIVA"
	Call GF_writeTextPlus(p_oPDF, 50, 110, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "98% MIN"
	Call GF_writeTextAlign(p_oPDF, 120, 110, aux, 60, PDF_ALIGN_CENTER)
	aux = "Hasta el 95%"
	Call GF_writeTextAlign(p_oPDF, 180, 110, aux, 80, PDF_ALIGN_CENTER)
	aux = "Sin Bonificaciones"
	Call GF_writeTextAlign(p_oPDF, 260, 110, aux, 80, PDF_ALIGN_CENTER)
	aux = "Valores inferiores al 98% y hasta 95% a raz�n de 1% por cada por ciento � fracci�n."
	Call GF_writeTextPlus(p_oPDF, 355, 110, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Por debajo de 95%, ser� de rechazo."
	Call GF_writeTextPlus(p_oPDF, 455, 110, aux, 90, 10, PDF_ALIGN_LEFT)
	'2 Granos quebrados, partidos, pelados, da�ados y materias extra�as
	aux = "GRANOS QUEBRADOS, PARTIDOS, PELADOS, DA�ADOS Y MATERIAS EXTRA�AS"
	Call GF_writeTextPlus(p_oPDF, 50, 170, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "2,0% MAX"
	Call GF_writeTextAlign(p_oPDF, 120, 170, aux, 60, PDF_ALIGN_CENTER)
	aux = "Hasta el 4,0% maximo"
	Call GF_writeTextPlus(p_oPDF, 180, 170, aux, 80, 10, PDF_ALIGN_CENTER)
	if (cdbl(p_Cosecha) < 17) then
	    aux = "Para valores inferiores a 2,0% a raz�n de 1% por cada por ciento o fracci�n proporcional"
	    Call GF_writeTextPlus(p_oPDF, 265, 170, aux, 70, 10, PDF_ALIGN_LEFT)
    else	    
	    aux = "Sin Bonificaciones"
	    Call GF_writeTextAlign(p_oPDF, 260, 170, aux, 80, PDF_ALIGN_CENTER)
    end if	        	
	aux = "Para valores superiores a 2,0% y hasta el 4%, a raz�n de 1% por cada por ciento o fracci�n proporcional"
	Call GF_writeTextPlus(p_oPDF, 355, 170, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Por encima del 4%, la mercader�a ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 455, 170, aux, 90, 10, PDF_ALIGN_LEFT)
	'3 Granos con carbon
	aux = "GRANOS CON CARBON"
	Call GF_writeTextPlus(p_oPDF, 50, 260, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "-"
	Call GF_writeTextAlign(p_oPDF, 120, 260, aux, 60, PDF_ALIGN_CENTER)
	aux = "0,2% MAXIMO"
	Call GF_writeTextPlus(p_oPDF, 180, 260, aux, 80, 10, PDF_ALIGN_CENTER)
	aux = "Sin Bonificaciones"
	Call GF_writeTextAlign(p_oPDF, 260, 260, aux, 80, PDF_ALIGN_CENTER)
	aux = "Para valores superiores a 0,20% y hasta el 0,4%,ser� a raz�n de 1% por cada por ciento o fracci�n proporcional"
	Call GF_writeTextPlus(p_oPDF, 355, 260, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Mercaderia que exceda el 0,4%, ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 455, 260, aux, 90, 10, PDF_ALIGN_LEFT)
	'4 Granos picados
	aux = "GRANOS PICADOS"
	Call GF_writeTextPlus(p_oPDF, 50, 340, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "-"
	Call GF_writeTextAlign(p_oPDF, 120, 340, aux, 60, PDF_ALIGN_CENTER)
	aux = "0,5% MAXIMO"
	Call GF_writeTextPlus(p_oPDF, 180, 340, aux, 80, 10, PDF_ALIGN_CENTER)
	aux = "Sin Bonificaciones"
	Call GF_writeTextAlign(p_oPDF, 260, 340, aux, 80, PDF_ALIGN_CENTER)
	aux = "Para valores superiores a 0,5% y hasta el 1%, a raz�n de 1% por cada por ciento o fracci�n proporcional"
	Call GF_writeTextPlus(p_oPDF, 355, 340, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Mercaderia que exceda el 1%, ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 455, 340, aux, 90, 10, PDF_ALIGN_LEFT)
	'5 Calibre (bajo zaranda de 2,2 mm)
	aux = "CALIBRE (bajo zaranda de 2,2 mm.)"
	Call GF_writeTextPlus(p_oPDF, 50, 420, aux, 60, 10, PDF_ALIGN_LEFT)
	aux = "-"
	Call GF_writeTextAlign(p_oPDF, 120, 420, aux, 60, PDF_ALIGN_CENTER)
	aux = "3,0% MAXIMO"
	Call GF_writeTextPlus(p_oPDF, 180, 420, aux, 80, 10, PDF_ALIGN_CENTER)
	if (cdbl(p_Cosecha) < 17) then
	    aux = "Para valores inferiores a 3, a raz�n de 1% por cada por ciento o fracci�n proporcional"
    else
        aux = "Sin Bonificaciones"
    end if	    
	Call GF_writeTextPlus(p_oPDF, 265, 420, aux, 80, 10, PDF_ALIGN_LEFT)
	aux = "Para valores superiores a 3,0% y hasta el 4%, a raz�n de 1% por cada por ciento o fracci�n proporcional"
	Call GF_writeTextPlus(p_oPDF, 355, 420, aux, 90, 10, PDF_ALIGN_LEFT)
	aux = "Mercaderia que exceda el 4%, ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 455, 420, aux, 90, 10, PDF_ALIGN_LEFT)
	'6 Calibre (sobre zaranda de 2,5 mm)
	aux = "CALIBRE (sobre zaranda de 2,5 mm.)"
	Call GF_writeTextPlus(p_oPDF, 50, 500, aux, 60, 10, PDF_ALIGN_LEFT)
	aux = "-"
	Call GF_writeTextAlign(p_oPDF, 120, 500, aux, 60, PDF_ALIGN_CENTER)
	aux = "85% MINIMO"
	Call GF_writeTextPlus(p_oPDF, 180, 500, aux, 80, 10, PDF_ALIGN_CENTER)
	aux = "Sin Bonificaciones"
	Call GF_writeTextAlign(p_oPDF, 260, 500, aux, 80, PDF_ALIGN_CENTER)
	aux = "Mercaderia que est� por debajo del 85%, ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 355, 500, aux, 90, 10, PDF_ALIGN_LEFT)	
	'7 PROTEINA (sobre sustancia seca)
	aux = "PROTEINA (sobre sustancia seca)"
	Call GF_writeTextPlus(p_oPDF, 50, 550, aux, 70, 10, PDF_ALIGN_LEFT)
	aux = "M�nimo 9%"
	if ((cdbl(p_Cosecha) = 15) or (cdbl(p_Cosecha) = 16)) then aux = "M�nimo 9.5%"
	if (cdbl(p_Cosecha) = 17) then aux = "M�nimo 10%"
	Call GF_writeTextAlign(p_oPDF, 120, 550, aux, 60, PDF_ALIGN_CENTER)
	aux = "M�ximo 12%"
	if ((cdbl(p_Cosecha) = 15) or (cdbl(p_Cosecha) = 16)) then aux = "M�ximo 12.5%"	
	if (cdbl(p_Cosecha) = 17) then aux = "M�ximo 12%"	
	Call GF_writeTextPlus(p_oPDF, 180, 550, aux, 80, 10, PDF_ALIGN_CENTER)
	Call GF_setFont(p_oPDF, "Times",7,0)
	aux = "Liquidaci�n.<br>"
	if (cdbl(p_Cosecha) = 15) then	    	
			aux = aux & "- Del 12,1% al 12,5% se liquidara al 97% del valor fijado.<br>"
			aux = aux & "- Del 10% hasta 12% se liquidara al 100% del valor fijado.<br>"					
			aux = aux & "- Del 9,5% al 9,9% se liquidara al 95% del valor fijado.<br>"			
	end if	
	if (cdbl(p_Cosecha) = 16) then
			aux = aux & "- Del 10% al 12% se liquidara al 100% del valor fijado.<br>"
			aux = aux & "- Del 9.5% hasta 9.9% se liquidara al 97% del valor fijado.<br>"					
			aux = aux & "- Del 12.1% al 12.5% se liquidara al 97% del valor fijado.<br>"			
	end if	
	if (cdbl(p_Cosecha) = 17) then	    	
			aux = aux & "Del 10% al 12% se liquidara al 100% del valor fijado.<br>"			
	end if	
	Call GF_writeTextPlus(p_oPDF, 265, 545, aux, 180, 10, PDF_ALIGN_LEFT)
	call GF_setFont(p_oPDF, "Times",9,0)
	'8 HUMEDAD
	aux = "HUMEDAD"
	Call GF_writeTextPlus(p_oPDF, 50, 610, aux, 60, 10, PDF_ALIGN_LEFT)
	aux = "12,5% m�ximo"
	Call GF_writeTextPlus(p_oPDF, 180, 610, aux, 80, 10, PDF_ALIGN_CENTER)
	if (cdbl(p_Cosecha) < 17) then	    	
	    aux = "Para valores inferiores a 12%,a raz�n de 1,2% por cada por ciento o fracci�n proporcional"
    else
        aux = "Sin Bonificaciones"
    end if 	    
	Call GF_writeTextPlus(p_oPDF, 265, 605, aux, 85, 10, PDF_ALIGN_LEFT)
	aux = "Mercader�a que exceda el 12,5% de humedad, ser� de rechazo"
	Call GF_writeTextPlus(p_oPDF, 355, 610, aux, 190, 10, PDF_ALIGN_LEFT)
			
    'texto llamados
	call GF_setFont(p_oPDF, "Times",8,8)
	aux = "LIBRE  DE  INSECTOS  Y/O  ARACNIDOS  VIVOS."	
	if ((cdbl(p_Cosecha) = 15) or (cdbl(p_Cosecha) = 16)) then aux = aux & " LIBRE  DE  DICLORVOS  (NO  DEBE  ESTAR  FUMIGADA  CON  LIQUIDO  DDVP)"		
	Call GF_writeTextAlign(p_oPDF, 50, 665, aux, 450, PDF_ALIGN_LEFT)	
	if (((cdbl(p_Cosecha) = 12) or (cdbl(p_Cosecha) = 13)) and (p_operacion = 9)) then 
		aux = "LIBRE  DE  DICLORVOS  (NO  DEBE  ESTAR  FUMIGADA  CON  LIQUIDO  DDVP)"
		Call GF_writeTextAlign(p_oPDF, 50, 675, aux, 450, PDF_ALIGN_LEFT)	
	end if	
	if (cdbl(p_Cosecha) >= 15) then
	    aux = "PUREZA VARIETAL: MIN 95%"
		Call GF_writeTextAlign(p_oPDF, 50, 675, aux, 450, PDF_ALIGN_LEFT)	
	end if
end Function
'***************************************************************************************
Function writeFirmas(p_oPDF, p_boleto)
	dim aux
	'texto firmas
	call GF_setFont(p_oPDF, "Times",10,0)
	aux = left(trim(getDsEnterprise2(p_boleto("KCVEN"))), 60)
	if len(aux) > 28 then
		Call GF_writeTextPlus(p_oPDF, 45, 690, aux, 160, 10, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(p_oPDF, 40, 700, aux, 170, PDF_ALIGN_CENTER)
	end if
	aux = left(trim(getDsEnterprise2(p_boleto("KCCOR"))), 50)
	if len(aux) > 28 then
		Call GF_writeTextPlus(p_oPDF, 215, 690, aux, 160, 10, PDF_ALIGN_LEFT)
	else
		Call GF_writeTextAlign(p_oPDF, 210, 700, aux, 170, PDF_ALIGN_CENTER)
	end if
	aux = "ALFRED  C.  TOEPFER  INT. ARG."
	Call GF_writeTextAlign(p_oPDF, 380, 700, aux, 170, PDF_ALIGN_CENTER)
	aux = ".............................."
	Call GF_writeTextAlign(p_oPDF, 40, 775, aux, 170, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(p_oPDF, 210, 775, aux, 170, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(p_oPDF, 380, 775, aux, 170, PDF_ALIGN_CENTER)
	aux = "Vendedor"
	Call GF_writeTextAlign(p_oPDF, 40, 795, aux, 170, PDF_ALIGN_CENTER)
	aux = "Corredor"
	Call GF_writeTextAlign(p_oPDF, 210, 795, aux, 170, PDF_ALIGN_CENTER)
	aux = "Comprador"
	Call GF_writeTextAlign(p_oPDF, 380, 795, aux, 170, PDF_ALIGN_CENTER)
	aux = "C.U.I.T."
	Call GF_writeTextAlign(p_oPDF, 40, 805, aux, 170, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(p_oPDF, 210, 805, aux, 170, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(p_oPDF, 380, 805, aux, 170, PDF_ALIGN_CENTER)
	aux = formatear(p_boleto("VEN_CUIT"),"CUIT")
	Call GF_writeTextAlign(p_oPDF, 40, 815, aux, 170, PDF_ALIGN_CENTER)
	aux = formatear(p_boleto("COR_CUIT"),"CUIT")
	if (p_boleto("KCCOR") <> 0) then Call GF_writeTextAlign(p_oPDF, 210, 815, aux, 170, PDF_ALIGN_CENTER)
	aux = "30-62197317-3"
	Call GF_writeTextAlign(p_oPDF, 380, 815, aux, 170, PDF_ALIGN_CENTER)
end Function
'***************************************************************************************
Function procesarBoleto(rs, cant_env)
	dim vecMails(10)

	response.write "Negocio:" & GF_EDIT_CONTRATO(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha"))) & "<br><br>"
	'Obtengo los mails. Si no los tiene seteados, no genero el pdf siquiera
	if cdbl(rs("KCCOR"))>0 then
		call obtenerMailBoletos(cdbl(rs("KCCOR")), vecMails)
	else
		call obtenerMailBoletos(cdbl(rs("KCVEN")), vecMails)
	end if
	if vecMails(0)<>"" or vecMails(1)<>"" then
		ret = generarPDF(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha")), cdbl(rs("KCVEN")))
		if ret then
			response.write "PDF generado<br>"
			enviarMailBoleto rs, cant_env
			if err.number <> 0 then
				response.write "Error enviarMailBoleto<br><br>"
				'Pongo la marca de enviado a X(error)
				Call actualizarBoleto(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha")), MRCENVIO_X, session("MomentoSistema"), cant_env, BOLETO_UPDATE)
				call GF_LogError(Err, "Pagina: procedimientoBoletos.asp. Error envio Mail")
			else
				response.write "Boleto enviado<br><br>"
			end	if
		else
			call writeLog("ERR", "Boleto :" & GF_EDIT_CONTRATO(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha"))) & " imposible de generar.")
			response.write "imposible generar PDF<br><br>"
			'Pongo la marca de enviado a X(error)
			Call actualizarBoleto(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha")), MRCENVIO_X, session("MomentoSistema"), cant_env, BOLETO_UPDATE)
		end if
	else
		response.write "No tiene direcciones de mail para enviar<br><br>"
		Call actualizarBoleto(cdbl(rs("Producto")), cdbl(rs("Sucursal")), cdbl(rs("Operacion")), cdbl(rs("Numero")), cdbl(rs("Cosecha")), MRCENVIO_N, session("MomentoSistema"), cant_env, BOLETO_UPDATE)
	end if
end Function
'***************************************************************************************
Function incluyeBiotecnologia(p_producto, p_cosecha, p_destino, p_fechaCto)
    Dim ret, rsDestino, myZona
    
    ret = false
    'Call executeSP(rsDestino, "MERFL.MER192F1_GET_BY_CODIDE", p_destino)    
    'myZona = 0
    'if (not rsDestino.eof) then myZona = CInt(rsDestino("ZONADE"))    
    if ((CInt(p_producto) = 23) and (CInt(p_cosecha) >= 16)) then ret = true    
    incluyeBiotecnologia = ret
    
End Function    
%>