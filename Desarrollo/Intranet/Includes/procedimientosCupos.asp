<%
Const MOSTRAR_PUBLICADOS = 3
Const MOSTRAR_ENVIADOS = 2
Const MOSTRAR_NO_ENVIADOS = 1

Const SENDER_CUPOS_BA = "LogisticaToepfer.AR@toepfer.com"
Const SENDER_CUPOS_RO = "TradingRos.AR@Toepfer.com"

Const SIN_CORREDOR = 15
Const MERCADO_A_TERMINO = 5454

const PUERTO_PIEDRABUENA = 91
const MUELLE_BAHIABLANCA = 64

Const CUPO_CANCELADO = 0
Const CUPO_PROVISORIO = 1 'Aun no publicado al cliente, recien cargado.
Const CUPO_OTORGADO = 2 'Ya publicado al cliente
Const CUPO_NOMINADO = 3 'El destinatario ya indicò corredor y vendedor.
Const CUPO_PUBLICADO_AFIP = 11 'El cupo fue dado de alta en el sistema STOP-Q
Const CUPO_ACTIVADO_AFIP = 12 'La AFIP le asigno un CTG al cupo.
Const CUPO_DESCARGADO_AFIP = 13 'La AFIP reporta que el cupo ya descargo
Const CUPO_ARRIBADO_AFIP = 15 'La AFIP reporta que el camion arribo a planta.

Const CUPOS_MAX_DISPONIBLES = "QTMAXCUPOS" ' Cantidad maxima de cupos disponibles en la terminal para un dia

Const CUPOS_FECHA_BASE = 20180101

const OPERACION_PRESTAMO_DEVOLUCION = 4
'******************************************************************************
'Funciones nuevas
function obtenerListaPuertos()
	Dim strSQL, rs,con
	strSQL = "select CODIDE as IDPUERTO, DESCDE as DSPUERTO from MERFL.MER192F1 order by CODIDE"
	call GF_BD_AS400_2(rs,con,"OPEN",strSQL)
	Set obtenerListaPuertos = rs
end function
'******************************************************************************
'Definicion de constantes para el PDF de cupos
Const PDF_CUPOS_Separacion_lineas = 20
Const PDF_CUPOS_Separacion_cupos = 15
Const PDF_CUPOS_Separacion_total_fecha = 10
Const PDF_CUPOS_Tipo_letra = "Arial"
Const PDF_CUPOS_Tamano_letra = 10
Const CODIGO_POR_DEFECTO = "X00000000"
'******************************************************************************
function GF_generarPDFCuposProveedores(ByVal p_rs, ByVal p_KCPRO, ByVal p_strPath, p_strError)
    dim oFS, oPDF
    
    set oFS = Server.CreateObject("Scripting.FileSystemObject")
    if (oFS.FileExists(p_strPath)) then
        oFS.deleteFile(p_strPath)
    end if

    set oPDF = GF_createPDF(p_strPath)
    call dibujarEncabezado(oPDF, p_KCPRO)
    call dibujarListadoCupos(oPDF, p_rs, p_KCPRO)
    Call GF_closePDF(oPDF)
    
end function
'******************************************************************************
sub dibujarEncabezado(p_oPDF, p_KCPRO)
     'Se dibuja el recuadro
     Call GF_squareBox(p_oPDF, 5, 10, 585, 125, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
     Call GF_squareBox(p_oPDF, 5, 142, 585, 685, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
     'Se coloca el Logo
     Call GF_writeImage(p_oPDF, Server.MapPath("Images\kogge256.jpg"), 520, 60, 75, 60, 0)
     'Se escribe la informacion de la cabecera
     Call GF_setFont(p_oPDF,"ARIAL", 16,8)
     'Call GF_writeText(p_oPDF,15, 23, GF_Traducir("Informe de Cupos"), 0)
     call GF_writeTextAlign(p_oPDF, 15, 23, GF_Traducir("Informe de Cupos"), 570, 2)
     Call GF_setFont(p_oPDF,"Courier", 12, 8)
     Call GF_writeText(p_oPDF,15, 60, getDsEnterprise2(99999997), 0)
     'Call GF_writeText(p_oPDF,15, 75, GF_Traducir("1"), 0)
     Call GF_writeText(p_oPDF,15, 90, getDsEnterprise2(p_KCPRO), 0)
     'Call GF_writeText(p_oPDF,240, 90, GF_Traducir("3"), 0)
end sub
'******************************************************************************
sub dibujarListadoCupos(ByRef p_oPDF, ByRef rs, ByVal p_KCPRO)
    dim intLineaBase, intLineaCupos

    'Aca es donde realizo el corte de control
    intLineaBase = 127
    intPagina = 1
    call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 0)
    acumuladoCamionesProveedor = 0
    while not rs.eof
        puertoAnterior = rs("Puerto")
        puertoAnteriorDS = getPortDescription(puertoAnterior)
        while esMismo(rs, "Puerto", puertoAnterior)
            productoAnterior = rs("Producto")
            call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
            intLineaBase = intLineaBase + PDF_CUPOS_Separacion_lineas
            strTexto = "Puerto : " & puertoAnteriorDs & " - " & getDsProducto(productoAnterior)
            call GF_writeTextAlign(p_oPDF, 15, intLineaBase, strTexto, 570, 0)
            acumuladoCamionesProducto = 0
            while esMismo(rs, "Producto", productoAnterior) and esMismo(rs, "Puerto", puertoAnterior)
                vendedorAnterior = rs("Vendedor")
                acumuladoCamionesPersona = 0
                while esMismo(rs, "Vendedor", vendedorAnterior) and esMismo(rs, "Producto", productoAnterior) and esMismo(rs, "Puerto", puertoAnterior)
                    sucursalAnterior = rs("Sucursal")
                    operacionAnterior = rs("Operacion")
                    numeroAnterior = rs("Numero")
                    cosechaAnterior = rs("Cosecha")
                    call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
                    intLineaBase = intLineaBase + PDF_CUPOS_Separacion_lineas
                    strTexto = "Para : " & getDsEnterprise2(vendedorAnterior) & " - " & rs("CtoCorredor") & " - " & GF_EDIT_CONTRATO(productoAnterior,sucursalAnterior,operacionAnterior,numeroAnterior,cosechaAnterior)
                    call GF_writeTextAlign(p_oPDF, 30, intLineaBase, strTexto, 570, 0)
                    contadorFechasPorContrato = 0
                    acumuladoCamionesContrato = 0
                    while esMismoContrato(rs, productoAnterior, sucursalAnterior, operacionAnterior, numeroAnterior, cosechaAnterior)
                        fechaAnterior = rs("Fecha")
                        call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
                        call GF_writeTextAlign(p_oPDF, 75, intLineaBase + PDF_CUPOS_Separacion_cupos, "Fecha : " & GF_FN2DTE(fechaAnterior), 520, 0)
                        contadorFechasPorContrato = contadorFechasPorContrato + 1
                        acumuladoCamionesFecha = 0
                        while esMismo(rs, "Fecha", fechaAnterior) and esMismoContrato(rs, productoAnterior, sucursalAnterior, operacionAnterior, numeroAnterior, cosechaAnterior)
                            codigoAnterior = rs("numcupo")
                            call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
                            intLineaBase = intLineaBase + PDF_CUPOS_Separacion_cupos
                            cantidadCamiones = rs("Camiones")
                            if not esCodigoPorDefecto(codigoAnterior) then
                                call GF_writeTextAlign(p_oPDF, 180, intLineaBase,  codigoAnterior, 100, 0)
                                call GF_writeTextAlign(p_oPDF, 225, intLineaBase,  "a", 30, 2)
                            end if
                            rs.movenext
                            acumuladoCamionesFecha = acumuladoCamionesFecha + cantidadCamiones
                            while esConsecutivo(rs, codigoAnterior)
                                cantidadCamiones = cantidadCamiones + rs("Camiones")
                                acumuladoCamionesFecha = acumuladoCamionesFecha + rs("Camiones")
                                codigoAnterior = rs("numcupo")
                                rs.movenext
                            wend
                            if not esCodigoPorDefecto(codigoAnterior) then call GF_writeTextAlign(p_oPDF, 250, intLineaBase,  codigoAnterior, 70, 0)
                            if cantidadCamiones > 1 then
                                strDenominacionCamiones = "camiones"
                            else
                                strDenominacionCamiones = "camión"
                            end if
                            call GF_writeTextAlign(p_oPDF, 280, intLineaBase,  cantidadCamiones, 40, 1)
                            call GF_writeTextAlign(p_oPDF, 325, intLineaBase,  strDenominacionCamiones, 70, 0)
                        wend
                        acumuladoCamionesContrato = acumuladoCamionesContrato + acumuladoCamionesFecha
                    WEND
                    acumuladoCamionesPersona = acumuladoCamionesPersona + acumuladoCamionesContrato
                    call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
                    intLineaBase = intLineaBase + PDF_CUPOS_Separacion_lineas
                    call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 8)
                    call GF_writeTextAlign(p_oPDF, 60, intLineaBase, "Total Contrato", 520, 0)
                    call GF_writeTextAlign(p_oPDF, 60, intLineaBase, "___________________________________________________________________", 520, 0)
                    call GF_writeTextAlign(p_oPDF, 330, intLineaBase,  acumuladoCamionesContrato, 50, 1)
                    call GF_writeTextAlign(p_oPDF, 385, intLineaBase,  "camiones", 70, 0)
                    call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 0)
                    intLineaBase = intLineaBase + PDF_CUPOS_Separacion_total_fecha
                wend
                acumuladoCamionesProducto = acumuladoCamionesProducto + acumuladoCamionesContrato
            wend
            acumuladoCamionesProveedor = acumuladoCamionesProveedor + acumuladoCamionesPersona
            call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
            intLineaBase = intLineaBase + PDF_CUPOS_Separacion_lineas
            call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 8)
            call GF_writeTextAlign(p_oPDF, 30, intLineaBase, "Total " & getDsEnterprise2(vendedorAnterior), 520, 0)
            call GF_writeTextAlign(p_oPDF, 30, intLineaBase, "________________________________________________________________________________________________", 590, 0)
            call GF_writeTextAlign(p_oPDF, 460, intLineaBase,  acumuladoCamionesPersona, 50, 1)
            call GF_writeTextAlign(p_oPDF, 515, intLineaBase,  "camiones", 70, 0)
            call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 0)
            intLineaBase = intLineaBase + PDF_CUPOS_Separacion_lineas
        wend
    wend
    call verificarPagina(p_oPDF, intLineaBase, p_KCPRO, intPagina)
    intLineaBase = intLineaBase + PDF_CUPOS_Separacion_total_fecha
    Call GF_writeImage(p_oPDF, Server.MapPath("Images\marco_t_l_fondoBlanco.jpg"), 0, 780, 8, 24, 0)
    Call GF_writeImage(p_oPDF, Server.MapPath("Images\marco_t_r_fondoBlanco.jpg"), 587, 780, 8, 24, 0)
    Call GF_writeImage(p_oPDF, Server.MapPath("Images\marco_r1_c2_fondoBlanco.jpg"), 8, 795, 579, 8, 0)
    call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 8)
    call GF_writeTextAlign(p_oPDF, 30, 805, "Total " & getDsEnterprise2(p_KCPRO), 520, 0)
    call GF_writeTextAlign(p_oPDF, 460, 805,  acumuladoCamionesProveedor, 50, 1)
    call GF_writeTextAlign(p_oPDF, 515, 805,  "camiones", 70, 0)
    call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 0)

    call GF_setFont(p_oPDF, "Arial", 9, 0)
    call GF_writeTextAlign(p_oPDF, 30, 830, "Página " & intPagina, 555, 1)
end sub
'******************************************************************************
function esMismo(ByRef p_rs, ByVal p_strCampo, ByVal p_strValor)

    esMismo = false
    if not p_rs.eof then
        if p_rs(p_strCampo) = p_strValor then
            esMismo = true
        end if
    end if
end function
'******************************************************************************
function esMismoContrato(ByRef p_rs, ByVal p_producto, ByVal p_sucursal, ByVal p_operacion, ByVal p_numero, ByVal p_cosecha)
    if esMismo(p_rs, "Producto", p_producto) and esMismo(p_rs, "Sucursal", p_sucursal) and esMismo(p_rs, "Operacion", p_operacion) and esMismo(p_rs, "Numero", p_numero) and esMismo(p_rs, "Cosecha", p_cosecha) then
        esMismoContrato = true
    else
        esMismoContrato = false
    end if
end function
'******************************************************************************
function esConsecutivo(ByRef p_rs, ByVal p_strValor)

    esConsecutivo = false
    if not p_rs.eof then
        if clng(right(p_rs("Codigo"), 8)) = (clng(right(p_strValor, 8)) + 1) then
            esConsecutivo = true
        end if
    end if
end function
'*****************************************************************************************************************
sub verificarPagina(ByRef p_oPDF, ByRef p_intLineaBase, p_KCPRO, ByRef p_intPagina)
    if p_intLineaBase > 780 then
        call GF_setFont(p_oPDF, "Arial", 9, 0)
        call GF_writeTextAlign(p_oPDF, 30, 830, "Página " & p_intPagina, 555, 1)
        p_intPagina = p_intPagina + 1
        call GF_newPage(p_oPDF)
        call dibujarEncabezado(p_oPDF, p_KCPRO)
        p_intLineaBase = 130
        call GF_setFont(p_oPDF, PDF_CUPOS_Tipo_letra, PDF_CUPOS_Tamano_letra, 0)
    end if
end sub
'*****************************************************************************************************************
sub enviarMailCuposProveedor(p_KCPRO, p_Producto, p_Sucursal, p_Operacion, p_Numero, p_Cosecha, p_FechaDesde, p_FechaHasta, p_pathAttachment)
    dim vecMails(10), conn, strSQL, strDestinatarios, cant
    
    strDsProveedor = getDsEnterprise2(p_KCPRO)
    strToepferDenomination = getDsEnterprise2(99999997)
    'completo los datos del mail
    strAsunto = "Cupos Toepfer"
    strPathAttachment = p_pathAttachment

    strBody = "Se adjuntan los cupos asigandos para " & strDsProveedor
    if p_Producto > 0 then
        strBody = strBody & " correspondientes al negocio " & GF_Edit_Contrato(p_Producto, p_Sucursal, p_Operacion, p_Numero, p_Cosecha)
    end if
    strBody = strBody & " a partir de la fecha " & GF_FN2DTE(p_FechaDesde)
    if cint(p_FechaHasta)>0 then
        strBody = strBody & " hasta la fecha " & GF_FN2DTE(p_FechaHasta)
    else
        strBody = strBody & " en adelante."
    end if

    cant = obtenerMailCuposProveedor(p_KCPRO, vecMails)
    while (cant > 0)
        cant= cant-1
        strDestinatarios = strDestinatarios & vecMails(cant) & "; "
    wend
    'if vecMails(0)<>"" then strDestinatarios = strDestinatarios & vecMails(0) & "; "
    'if vecMails(1)<>"" then strDestinatarios = strDestinatarios & vecMails(1) & "; "
    'response.write strDestinatarios
    'response.end
    'strDestinatarios = "scalisij@toepfer.com;"
    call GP_ENVIAR_MAIL_ATTACHMENT(strAsunto, strBody,strToepferDenomination & " <" & SENDER_MERCADERIAS & ">",strDestinatarios, strPathAttachment)
end sub
'*****************************************************************************************************************
function getPortDescription(p_KCPort)
    strSQl = "select DESCDE from MERFL.MER192F1 where CODIDE='" & p_KCPort & "'"
    'response.write strSQL
    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
    if not rs.eof then
        getPortDescription = trim(rs("DESCDE"))
    else
        getPortDescription = "#KC Puerto no valido#"
    end if
end function
'*****************************************************************************************************************
function esCodigoPorDefecto(p_Codigo)
    if p_Codigo = CODIGO_POR_DEFECTO then
        esCodigoPorDefecto = true
    else
        esCodigoPorDefecto = false
    end if
end function
'------------------------------------------------------------------------------------------	
function getDsAbrProductoParaCodigoCupo(pCdProducto, pCdPuerto)
dim strSQL, rsAbr, rtrn
if cint(pCdPuerto) = 91 then
	rtrn = left(getDsProduct(pCdProducto),1)
else
	rtrn = "XX"
	strSQL = " SELECT A8BGTX FROM MERFL.MER112F1 WHERE CODIPR = " & pCdProducto
	Call GF_BD_COMPRAS(rsAbr, oConn, "OPEN", strSQL)
	if not rsAbr.eof then
		rtrn = left(trim(rsAbr("A8BGTX")),2)
	end if
end if	
getDsAbrProductoParaCodigoCupo = rtrn 		
end function
'-----------------------------------------------------------------------------------------------------------------
function getLetraCupo(pCodigo)
dim rtrn
rtrn = "?"
    select case ucase(pCodigo) 
        case DBSITE_TRANSITO
	        rtrn = "T"
        case DBSITE_BAHIA
	        rtrn = ""
        case DBSITE_ARROYO
	        rtrn = "A"
    end select 	
getLetraCupo = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
function getLetraCupo2(pCodigo)
dim rtrn
rtrn = "?"
    select case ucase(pCodigo) 
        case DBSITE_TRANSITO
	        rtrn = "T"
        case DBSITE_BAHIA
	        rtrn = "P"
        case DBSITE_ARROYO
	        rtrn = "N"
    end select 	
getLetraCupo2 = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
sub grabarMails(eMails, idProveedor)
'borrar mails anteriores
Dim strSQL, rs, conn

	strSQL = "delete from toepferdb.tblmailsmercaderias where idproveedor = " & idProveedor
	Call executeQuery(rs, "EXEC", strSQL)
'grabar mails, que se pasaron por parametro, cuando el usuario pulso el boton Guardar mails.
dim arrayMails	
	eMails = Replace(eMails,CHR(13), "") 
	arrayMails = split(eMails, ";")
	for i = 0 to ubound(arrayMails)
		Call guardarMailTipo(idProveedor, MAIL_CUPO, i+1, arrayMails(i))
	next
end sub
'-----------------------------------------------------------------------------------------------
function getStringMailsProveedor(idProveedor)
	dim cantMails, mails(10)
	dim strMailsAux
	cantMails = obtenerMailTipo(idProveedor, MAIL_CUPO, mails)
	for i=0 to cantMails - 1
		if i < cantMails - 1 then
			strMailsAux = strMailsAux & trim(mails(i)) & ";"
		else 'es el ultimo en la lista
			strMailsAux = strMailsAux & trim(mails(i))
		end if
	next
	getStringMailsProveedor = trim(strMailsAux)
end function
'----------------------------------------------------------------------------------------------
Function defineCdCorredor(pCuitCliente, pCdCorredor)
    Dim myCorredor
    myCorredor=pCdCorredor        
    if (CDbl(pCuitCliente) = CDbl(CUIT_TOEPFER)) then
        if (CLng(myCorredor)=0) then
            myCorredor = SIN_CORREDOR   
        else
            if (CLng(myCorredor) = SIN_CORREDOR) then myCorredor=0            
        end if        
    end if
    defineCdCorredor = myCorredor
End Function

%>
