<%

'/**
' * Funcion: Descargar
' * Descripcion: Descarga el archivo indicado al cliente.
' * Parametros:  strPath  [in] Archivo, con el path completo
' * Autor: Javier A. Scalisi
' * Fecha: 30/11/2005
' */

Function Descargar(strPath)

Dim fs, a

set fs = Server.CreateObject("Scripting.FileSystemObject")
Set a = fs.GetFile(strPath)
Response.Buffer = True
Response.Clear

Response.AddHeader "Content-Disposition", "attachment; filename=" & a.name
Response.AddHeader "Content-Length", a.size
Response.CharSet = "UTF-8"
Response.ContentType = "application/octet-stream"


Set s = Server.CreateObject("ADODB.Stream")
s.Open
s.Type=1
s.LoadFromFile(strPath)

Response.BinaryWrite s.Read
Response.Flush

s.Close
Set s = Nothing

End Function

'/**
' * Funcion: GF_GENERAR_DESCARGAS
' * Descripcion: Crea un archivo de texto con toda la informacion
' *              de las descargas, para el proveedor y las fechas indicadas.
' * Parametros:  p_dteFechaDesde  [in] Fecha en formato DD/MM/AAAA
' *              p_dteFechaHasta  [in] Fecha en formato DD/MM/AAAA
' * Autor: Javier A. Scalisi
' * Fecha: 30/11/2005
' */
Function GF_GENERAR_DESCARGAS(p_dteFechaDesde, p_dteFechaHasta)

    Dim strPath, strRegistro
    Dim arch, fs, intIndex

    'Se configura la lectura de descargas
    g_strCampoOrden = "D.FECDR6 asc"
    g_intFechaDesde = "'" & GF_DTE2FN(p_dteFechaDesde) & "'"
    g_intFechaHasta = "'" & GF_DTE2FN(p_dteFechaHasta) & "'"
    'Se crea el archivo a generar
    Randomize()
    strPath = Server.mapPath("temp\D" & "-" & GF_DTE2FN(p_dteFechaDesde) & "-" & GF_DTE2FN(p_dteFechaHasta) & "-" & Int(100 * Rnd()) & ".txt")
    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
'    response.write strPath & "<br>"
    If fs.FileExists(strPath) Then
        call fs.deleteFile(strPath, true)
    end if
    'Se crea nuevamente.
    Set arch = fs.CreateTextFile(strPath)
    intIndex=0
    if (initDescarga()) then
        while (getNextDescarga())
            intIndex=1
            'Se arma el registro con la informacion de la descarga
            'Fecha Descarga
            strRegistro = g_intFechaDescarga
            'Contrato Toepfer
            strRegistro = strRegistro & GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha)
            'Contrato Proveedor
            strRegistro = strRegistro & GF_nChars(g_strCtoCorredor,15," ",CHR_FWD)
            'Carta de Porte
            strRegistro = strRegistro & GF_nChars(g_intCartaPorte,12," ",CHR_FWD)
            'Kg. Descargados
            strRegistro = strRegistro & GF_nChars(g_intKilosDescarga,10," ",CHR_FWD)
            'Marca Conforme
            strRegistro = strRegistro & g_chrMrcConforme
            'Nro Solicitud Analisis
            strRegistro = strRegistro & GF_nChars(g_intSolicitudNro,8," ",CHR_FWD)
            'Grado
            strRegistro = strRegistro & g_intAnalisisGdo
            arch.WriteLine(strRegistro)
        wend
    end if
    arch.close
    if (intIndex > 0) then
        GF_GENERAR_DESCARGAS = strPath
    else
        GF_GENERAR_DESCARGAS = ""
    end if
End Function

'/**
' * Funcion: GF_GENERAR_PAGOS
' * Descripcion: Crea un archivo de texto con toda la informacion
' *              de los pagos, para el proveedor y las fechas indicadas.
' * Parametros:  p_dteFechaDesde  [in] Fecha en formato DD/MM/AAAA
' *              p_dteFechaHasta  [in] Fecha en formato DD/MM/AAAA
' * Autor: Javier A. Scalisi
' * Fecha: 30/11/2005
' */
Function GF_GENERAR_PAGOS(p_dteFechaDesde, p_dteFechaHasta)

    Dim oConn, rsCab, rsDet, strSQL, strRegistro, strPath
    Dim intFechaDesde, intFechaHasta, strKC, strAux, strDS
    Dim intIndex
    
    intFechaDesde = GF_DTE2FN(p_dteFechaDesde)
    intFechaHasta = GF_DTE2FN(p_dteFechaHasta)
    strKC=session("KCOrganizacion")
    
    'Se crea el archivo a generar
    Randomize()
    strPath = Server.mapPath("temp\P" & "-" & intFechaDesde & "-" & intFechaHasta & "-" & Int(100 * Rnd()) & ".txt")
    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
'    response.write strPath & "<br>"
    If fs.FileExists(strPath) Then
        call fs.deleteFile(strPath, true)
    end if
    'Se crea nuevamente.
    Set arch = fs.CreateTextFile(strPath)

    'Leo las ordenes de pago que deben agregarse al archivo.
    'strSQL="Select * from cor_PGCAB where FechaPago>='" & intFechaDesde & "' and FechaPago <='" & intFechaHasta & "' and (KCCOR=" & strKC & " or KCVEN=" & strKC & ") order by FechaPago"
    strSQL="Select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, WCIMME as importeMerc, WCIMIV as ImporteIVA, WCPOME as KCMERC, WCPOIV as KCIVA, WCCONT as Contrato "
	strSQL= strSQL & "from TESFL.TES960F1 where WCFPAG >= '" & intFechaDesde & "' and WCFPAG <= '" & intFechaHasta & "' and (WCNPRO=" & strKC & " or WCPRET=" & strKC & ") order by WCFPAG"
	'Response.Write strSQL
	GF_BD_AS400_2 rsCab,oConn,"OPEN",strSQL
	intIndex = 0
	while (not rsCab.eof)
         'Obtengo el detalle y lo incluyo en el archivo.
         Call GF_DET_LEER(rsDet,rsCab("FechaPago"),rsCab("TipoCbte"),rsCab("Minuta"))
         while (not rsDet.eof)
            if (((trim(rsDet("KCDetalle")) <> "W") and (trim(rsDet("MarcaAnulacion")) = "F")) or (trim(rsDet("KCDetalle")) = "W"))then
                intIndex = 1
                'Fecha de Pago
                strRegistro = rsDet("FechaPago")
                'Tipo Comprobante
                strRegistro = strRegistro & GF_nChars(trim(rsDet("TipoCbte")),3," ",CHR_FWD)
                'Nro Comprobante
                strRegistro = strRegistro & GF_nChars(trim(rsCab("CbteProveedor")),12,"0",CHR_FWD)
                'Numero de Orden de Pago
                strRegistro = strRegistro & GF_nChars(trim(rsCab("Orden")),6,"0",CHR_FWD)
                'Debito/Credito
                strRegistro = strRegistro & GF_nChars(trim(rsDet("DBCR")),1,"0",CHR_FWD)
                'Codigo Detalle
                strAux = rsDet("KCDetalle")
                if (trim(rsDet("KCDetalle")) = "W") then
                    strAux = trim(rsDet("KCDetalle")) & trim(rsDet("MRCPago"))
                end if
                strRegistro = strRegistro & GF_nChars(strAux,2," ",CHR_FWD)
                'Descripcion Detalle
                Call GF_MGC("MC",strAux,"",strDs)
                strRegistro = strRegistro & GF_nChars(strDS,50," ",CHR_BCK)
                'Forma de Pago
                strRegistro = strRegistro & GF_nChars(trim(rsDet("KCPago")),2," ",CHR_FWD)
                'Desc. Forma de Pago
                Call GF_MGC("CP",trim(rsDet("KCPago")),"",strDs)
                strRegistro = strRegistro & GF_nChars(strDS,50," ",CHR_BCK)
                'Importe.
                strRegistro = strRegistro & GF_nChars(cdbl(rsDet("ImportePesos"))*100,10,"0",CHR_FWD)
                arch.WriteLine(strRegistro)                
            end if
            rsDet.MoveNext
         wend
         rsCab.MoveNext
	wend
    arch.close
    if (intIndex > 0) then
        GF_GENERAR_PAGOS = strPath
    else
        GF_GENERAR_PAGOS = ""
    end if
End Function

'/**
' * Funcion: GF_GENERAR_ANALISIS
' * Descripcion: Crea un archivo de texto con toda la informacion
' *              de los analisis, para el proveedor y las fechas indicadas.
' * Parametros:  p_dteFechaDesde  [in] Fecha en formato DD/MM/AAAA
' *              p_dteFechaHasta  [in] Fecha en formato DD/MM/AAAA
' * Autor: Eugenio D. Di Santo
' * Fecha: 15/12/2005
' */
Function GF_GENERAR_ANALISIS(p_dteFechaDesde, p_dteFechaHasta)

    Dim strPath, strRegistro
    Dim arch, fs, intIndex

    'Se crea el archivo a generar
    Randomize()
    strPath = Server.mapPath("temp\A" & "-" & GF_DTE2FN(p_dteFechaDesde) & "-" & GF_DTE2FN(p_dteFechaHasta) & "-" & Int(100 * Rnd()) & ".txt")
    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
'    response.write strPath & "<br>"
    If fs.FileExists(strPath) Then
        call fs.deleteFile(strPath, true)
    end if
    'Se crea nuevamente.
    Set arch = fs.CreateTextFile(strPath)
    strSQl = "select C.CCORR1 as KCCOR, C.CVENR1 as KCVEN, CA.FANACA as Fecha, CA.CPORCA as CartaPorte, CA.NSANCA as SolicitudNro, CA.NROACA as Numero, CA.COBECA as Bolsa, CA.CPROCA as Producto, CA.KGMOCA as Kilos, DA.COANDA as Concepto, SA.DESCAN as DescConc, DA.VACADA as Valor, DA.PREBDA as PjeRebaja, DA.PBONDA as PjeBonificacion, CA.IMPACA as Costo"
	strSQl = strSQL & " from ((((MERFL.MER591DA DA inner join MERFL.MER2E2F1 SA on DA.COANDA=SA.CONCAN) inner join MERFL.MER591CA CA on DA.COBEDA=CA.COBECA and DA.CPRODA=CA.CPROCA and DA.FANADA=CA.FANACA and DA.NROADA=CA.NROACA) inner join MERFL.MER311F6 D on CA.CPORCA=D.CPORR6) inner join MERFL.MER311F1 C on D.CPROR6=C.CPROR1 and D.CSUCR6=C.CSUCR1 and D.COPER6=C.COPER1 and D.NCTOR6=C.NCTOR1 and D.ACOSR6=C.ACOSR1)"
	strSQl = strSQL & " where C.CONFR1 = 'V' and CA.FANACA between '" & GF_DTE2FN(p_dteFechaDesde) & "' and '" & GF_DTE2FN(p_dteFechaHasta) & "' and (C.CCORR1 = " & session("KCOrganizacion") & " or C.CVENR1 = " & session("KCOrganizacion") & ")"
	strSQl = strSQL & " order by CA.FANACA, CA.CPORCA, CA.NSANCA, CA.NROACA, CA.COBECA, DA.COANDA"

    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
    if not rs.eof then
        GF_GENERAR_ANALISIS = strPath
        while not rs.eof
            BolsaDs = ""
            ProductoDs = ""
            ConceptoDs = ""
            'Fecha Analisis
            registro = rs("Fecha")
            'Carta de Porte
            registro = registro & GF_nChars(rs("CartaPorte"),12," ",CHR_FWD)
            'Nro. Solicitud Analisis
            registro = registro & GF_nChars(rs("SolicitudNro"),8," ",CHR_FWD)
            'Numero de Analisis
            registro = registro & GF_nChars(rs("Numero"),8," ",CHR_FWD)
            'Bolsa
            registro = registro & GF_nChars(rs("Bolsa"),2," ",CHR_FWD)
            'Descripcion Bolsa
            call GF_MGC("ME", rs("Bolsa"), 0 , BolsaDS)
            registro = registro & GF_nChars(BolsaDs,30," ",CHR_BCK)
            'Descripcion Producto
            call GF_MGC("AR", rs("Producto"), 0 , ProductoDS)
            registro = registro & GF_nChars(ProductoDs,30," ",CHR_BCK)
            'Kg. Descargados
            registro = registro & GF_nChars(rs("Kilos"),10," ",CHR_FWD)
            'Concepto
            registro = registro & GF_nChars(rs("Concepto"),2," ",CHR_FWD)
            'Descripcion del Concepto
            registro = registro & GF_nChars(rs("DescConc"),30," ",CHR_BCK)
            'Valor resultado del Analisis
            registro = registro & GF_nChars(GF_EDIT_DECIMALS(CDbl(rs("Valor"))*100, 2),6," ",CHR_FWD)
            'Porcentaje Rebaja
            registro = registro & GF_nChars(GF_nDigits(rs("PjeRebaja"),5),5," ",CHR_FWD)
            'Porcentaje Bonificacion
            registro = registro & GF_nChars(GF_nDigits(rs("PjeBonificacion"),5),5," ",CHR_FWD)
            'Costo
            registro = registro & GF_nChars(rs("Costo"),8," ",CHR_FWD)
            arch.WriteLine(registro)
            rs.movenext
        wend
    else
        GF_GENERAR_ANALISIS = ""
    end if
    arch.close
End Function%>
