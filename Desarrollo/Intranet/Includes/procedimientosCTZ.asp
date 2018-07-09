<%
Const CTZ_PENDIENTE = "0"	'Recien Cargada
Const CTZ_EN_FIRMA  = "1"	'Con al menos una firma
Const CTZ_FIRMADA   = "2"	'Con todas las firmas requeridas
Const CTZ_FACTURADA = "F"	'Factura cargada
Const CTZ_ANULADA   = "R"	'Anulada
Const CTZ_EN_AJUSTE = "A"	'En proceso de Ajuste

'Secuencias de las firmas.
Const PIC_FIRMA_RESPONSABLE 		= 0
Const PIC_FIRMA_GTE_PUERTO 	 		= 4      
Const PIC_FIRMA_GTE_SECTOR 	 		= 5
Const PIC_FIRMA_GTE_COMPRAS	 		= 6
Const PIC_FIRMA_SUP_PUERTOS	 		= 7
Const PIC_FIRMA_CONTROLLER	 		= 8
Const PIC_FIRMA_DIRECCION	 		= 9
'Const CTZ_FIRMA_SOLICITANTE = 0
'Const CTZ_FIRMA_RESPONSABLE = 1
'Const CTZ_FIRMA_SUPERVISOR	= 2
'Const CTZ_FIRMA_SUP_PUERTOS = 4
'Const CTZ_FIRMA_DIRECCION	= 5

Const CTZ_NO_USER = "NAU"

Const PIC_TYPE_PURCHASE_SMALL   = 1
Const PIC_TYPE_PURCHASE_MEDIUM  = 2
Const PIC_TYPE_PURCHASE_X_MEDIUM = 3
Const PIC_TYPE_PURCHASE_LARGE   = 4

Const CTZ_SIN_ARCHIVOS = -1

Const CTZ_ITEM_DIFF_CAMBIO = 11508  'Item para pagar diferencias de cambio.

dim ctz_IdCotizacion, ctz_IdPedido, ctz_IdProveedor, ctz_importePesos, ctz_importeDolares, ctz_FecEntrega, ctz_Observaciones, ctz_idObra, ctz_TipoCambio
dim ctz_det_IdArticulo,ctz_det_ArticuloCantidad,ctz_det_ArticuloIdUnidad,ctz_det_IdArea,ctz_det_IdDetalle,ctz_det_Facturado, ctz_det_ImportePesos, ctz_det_TipoCambio, ctz_det_ImporteDolares, ctz_det_ImportePesosFacturado, ctz_det_ImporteDolaresFacturado
dim ctz_AjusteTotalPesos, ctz_AjusteTotalDolares, ctz_AjusteObservaciones, ctz_IdAjuste, ctz_AjusteIdArticulo, ctz_det_Estado
dim ctz_det_ImportePesosCredito, ctz_det_ImporteDolaresCredito, ctz_cdMoneda, ctz_IdContrato, ctz_IdDivision, ctz_docCode
'---------------------------------------------------------------------------------------------
'Función responsable por dejar el esquema de firmas de un PIC según las reglas definidas por la empresa.
Function addCTZFirmas(idCotizacion, cdSolicitante, cdAutorizante)

    Dim picType, auxUser
    
    Call readCTZ(idCotizacion)    
    
    esExpo = (ctz_idDivision = getDivisionID(CODIGO_EXPORTACION))
    'Para los puertos se deja en cero para que se rija por el rol de la firma.
	picType = getPICAuthorizationType(ctz_IdPedido, ctz_IdContrato, ctz_IdProveedor, CDbl(ctz_importeDolares)/100, MONEDA_DOLAR)				
	' Solicitante
	Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_RESPONSABLE, cdSolicitante)
	'Autorizante
	Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_GTE_SECTOR, cdAutorizante)
	'-- Gerente de Compras
    auxUser = ""    
    if (picType <> PIC_TYPE_PURCHASE_SMALL) then 
        rolFirmaSolici = getRolFirma(cdSolicitante, SEC_SYS_COMPRAS)
        if (rolFirmaSolici <> FIRMA_ROL_GTE_COMPRAS) then auxUser = FIRMA_NO_USER
    end if        
    Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_GTE_COMPRAS, auxUser)		    
    '-- Coordinador de Puertos 
    auxUser = ""
    if (((picType = PIC_TYPE_PURCHASE_LARGE) or (picType = PIC_TYPE_PURCHASE_X_MEDIUM)) and (not esExpo)) then auxUser = FIRMA_NO_USER
    Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_SUP_PUERTOS, auxUser)            
    '-- Controller
    auxUser = ""
    if (((picType = PIC_TYPE_PURCHASE_LARGE) or (picType = PIC_TYPE_PURCHASE_X_MEDIUM)) and (esExpo)) then auxUser = FIRMA_NO_USER
    Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_CONTROLLER, auxUser)                                    
    '-- Director
    auxUser = ""
    if (picType = PIC_TYPE_PURCHASE_LARGE) then auxUser = FIRMA_NO_USER
    Call adminCTZFirmas(ctz_IdCotizacion, PIC_FIRMA_DIRECCION, auxUser)
End Function
'---------------------------------------------------------------------------------------------
'Función que agregas las firmas del PIC.
Function adminCTZFirmas(pIdCotizacion, pSecuencia, pCdUsuario)
	Dim rs, strSQL, conn
			
	'Cargo los nuevos firmantes
	strSQL="Select * from TBLCTZFIRMAS where IDCOTIZACION=" & pIdCotizacion & " and SECUENCIA=" & pSecuencia
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (rs.eof) then
		if (pCdUsuario <> "") then 
            strSQL="Insert into TBLCTZFIRMAS(IDCOTIZACION, SECUENCIA, CDUSUARIO) values (" & pIdCotizacion & ", " & pSecuencia & ", '" & pCdUsuario & "')"			
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
        end if
	else
		if (pCdUsuario <> "") then
			strSQL="Update TBLCTZFIRMAS SET CDUSUARIO='" & pCdUsuario & "', FECHAFIRMA = null, HKEY = null where IDCOTIZACION=" & pIdCotizacion & " and SECUENCIA= " & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		else
			strSQL="Delete from TBLCTZFIRMAS where IDCOTIZACION=" & pIdCotizacion & " and SECUENCIA=" & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	end if
	
End Function
'---------------------------------------------------------------------------------------------
'Función responsable por dejar el esquema de firmas de un AJUSTE DE PIC según las reglas definidas por la empresa.
Function addAJUCTZFirmas(idCotizacion, idAjuste, cdSolicitante, cdAutorizante)

    Dim picType, auxUser
    
    Call readCTZ(idCotizacion)
    
    esExpo = (ctz_idDivision = getDivisionID(CODIGO_EXPORTACION))
    'Para los puertos se deja en cero para que se rija por el rol de la firma.
	picType = getPICAuthorizationType(ctz_IdPedido, ctz_IdContrato, ctz_IdProveedor, CDbl(ctz_importeDolares)/100, MONEDA_DOLAR)
	' Solicitante - Se copia del PIC Origen
	Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_RESPONSABLE, cdSolicitante)
	'Autorizante
	Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_GTE_SECTOR, cdAutorizante)
	'-- Gerente de Compras
    auxUser = ""
    if (picType <> PIC_TYPE_PURCHASE_SMALL) then 
        rolFirmaSolici = getRolFirma(cdSolicitante, SEC_SYS_COMPRAS)
        if (rolFirmaSolici <> FIRMA_ROL_GTE_COMPRAS) then auxUser = FIRMA_NO_USER
    end if        
    Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_GTE_COMPRAS, auxUser)		
    '-- Coordinador de Puertos 
    auxUser = ""
    if (((picType = PIC_TYPE_PURCHASE_LARGE) or (picType = PIC_TYPE_PURCHASE_X_MEDIUM)) and (not esExpo)) then auxUser = FIRMA_NO_USER
    Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_SUP_PUERTOS, auxUser)                                
    '-- Controller
    auxUser = ""
    if (((picType = PIC_TYPE_PURCHASE_LARGE) or (picType = PIC_TYPE_PURCHASE_X_MEDIUM)) and (esExpo)) then auxUser = FIRMA_NO_USER
    Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_CONTROLLER, auxUser)
    '-- Director
    auxUser = ""
    if (picType = PIC_TYPE_PURCHASE_LARGE) then auxUser = FIRMA_NO_USER
    Call adminAJUCTZFirmas(idAjuste, PIC_FIRMA_DIRECCION, auxUser)
End Function
'---------------------------------------------------------------------------------------------
'Función que agregas las firmas del AJUSTE DE PIC.
Function adminAJUCTZFirmas(pIdAjuste, pSecuencia, pCdUsuario)
	Dim rs, strSQL, conn
			
	'Cargo los nuevos firmantes
	strSQL="Select * from TBLCTZAJUSTESFIRMAS where IDAJUSTE=" & pIdAjuste & " and SECUENCIA=" & pSecuencia
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (rs.eof) then
		if (pCdUsuario <> "") then 
            strSQL="Insert into TBLCTZAJUSTESFIRMAS(IDAJUSTE, SECUENCIA, CDUSUARIO) values (" & pIdAjuste & ", " & pSecuencia & ", '" & pCdUsuario & "')"			
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
        end if
	else
		if (pCdUsuario <> "") then
			strSQL="Update TBLCTZAJUSTESFIRMAS SET CDUSUARIO='" & pCdUsuario & "', FECHAFIRMA = null, HKEY = null where IDAJUSTE=" & pIdAjuste & " and SECUENCIA= " & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		else
			strSQL="Delete from TBLCTZAJUSTESFIRMAS where IDAJUSTE=" & pIdAjuste & " and SECUENCIA=" & pSecuencia
            Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		end if
	end if
	
End Function
'---------------------------------------------------------------------------------------------
Function dibujarComboGte(pCdUsr, pCdAutorizante) 
    Dim rs, rsAut, conn
    
    'Se obtienen los usuarios con permiso para firmar el PIC como gerentes
    Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", "Select * from VWPARAMSFIRMAUSUARIO where NombreUsuario = '" & pCdUsr & "'") 
    if (not rs.eof) then
        Call executeQueryDb(DBSITE_SQL_INTRA, rsAut, "OPEN", "EXEC VWPARAMSFIRMAUSUARIO_GET_AUTORIZANTES " & rs("IdSector") & ", " & rs("NivelFirmaSector") & ", " & rs("NivelFirmaMgmt"))
        if (not rsAut.eof) then
%>
            <select id="cmbUsrAut" name="cmbUsrAut" onchange="seleccionAutorizante()">
                <option value="" /> - Seleccione -
<%          while (not rsAut.eof)                   
                if (pCdUsr <> rsAut("NombreUsuario")) then%>
                <option value="<% =rsAut("NombreUsuario") %>" <% if (pCdAutorizante = rsAut("NombreUsuario")) then response.write "selected" %> /> <% =UCase(rsAut("Apellido") & ", " & rsAut("Nombre"))  %>
<%              end if
                rsAut.MoveNext()
            wend            
%>        
            </select>
<%  
        else
            response.Write "No existen autorizantes definidos para el sector " & rs("IdSector")
        end if
    else
        response.Write "No existen autorizantes definidos para el solicitante indicado."
    end if
End Function
'---------------------------------------------------------------------------------------------
sub delCTZItems(pIdCotizacion)
dim strSQL, rs, conn, rsDel, connDel
	strSQL="Select * from TBLCTZDETALLE where idCotizacion = " & pIdCotizacion
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not rs.eof then
		strSQL = "Delete from TBLCTZDETALLE where IDCOTIZACION = " & pIdCotizacion & " and IDARTICULO <> " & CTZ_ITEM_DIFF_CAMBIO
		Call executeQueryDB(DBSITE_SQL_INTRA, rsDel, "EXEC", strSQL)	
	end if	
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)		
end sub
'---------------------------------------------------------------------------------------------
function addCTZItems(pIdCotizacion, pIdArticulo, pCantidad, pUnidad, pBudgetArea, pBudgetDetalle, pImportePesos, pImporteDolares, pTipoCambio)
'on error resume next
dim strSQL, rs, conn, rsIns, connIns, auxFec, cant, impoP, impD
addCTZItems = true

	strSQL="Select * from TBLCTZDETALLE where idCotizacion = " & pIdCotizacion & " and idarticulo=" & pIdArticulo & " and IDAREA=" & pBudgetArea & " and IDDETALLE=" & pBudgetDetalle
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if rs.eof then
		strSQL = "Insert Into TBLCTZDETALLE(IDCOTIZACION, IDARTICULO, CANTIDAD, IDUNIDAD, IDAREA, IDDETALLE, IMPORTEPESOS, IMPORTEDOLARES, TIPOCAMBIO, FACTURADO, IMPORTEPESOSFACTURADO, IMPORTEDOLARESFACTURADO, CANTIDADCREDITO, IMPORTEPESOSCREDITO, IMPORTEDOLARESCREDITO)"
		strSQL = strSQL & " VALUES(" & pIdCotizacion & "," & pIdArticulo & "," & pCantidad & "," & pUnidad & "," & pBudgetArea & ", " & pBudgetDetalle & ", " & pImportePesos & "," & pImporteDolares & "," & pTipoCambio & ", 0, 0, 0, 0, 0, 0)"
		Call executeQueryDB(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
		Call grabarImporteUltimaCompra(pIdCotizacion, pIdArticulo, session("MmtoDato"), CLng(pImportePesos/pCantidad), CLng(pImporteDolares/pCantidad))
	end if			
if err.number > 0 then addCTZItems = false
end function
'------------------------------------------------------------------------------------------
'Esta función graba el precio al que se compró un articulo, es el ultimo precio a nivel empresa.
'Es decir el precio ignora la división y simplemente refleja el valor al que se adquirió el articulo.
'Parametros
'	idPIC			= ID de la ultima compra
'	idArticulo		= ID del articulo al que se actualizará el precio.
'	mmto			= Momento de carga del PIC
'	importePesos	= Importe unitario en pesos.
'	importeDolares	= Importe unitario en dolares.
Function grabarImporteUltimaCompra(idPic, idArticulo, mmto, importePesos, importeDolares)
	Dim strSQL, conn, rs
	
	strSQL="Select * from TBLARTICULOS where IDARTICULO=" & idArticulo
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		'El articulo existe.
		if (CLng(rs("IDPIC")) <= idPIC) then
			'El precio nuevo es d euna compra mas reciente o bien es una modificación de la ultima compra.
			strSQL="Update TBLARTICULOS set MMTOULTIMACOMPRA=" & mmto & ", VLUPESOSULTIMACOMPRA=" & importePesos & " , VLUDOLARESULTIMACOMPRA=" & importeDolares & ", IDPIC=" & idPic & " where IDARTICULO=" & idArticulo
			Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
function addCTZCabecera(byref pIdCotizacion, pIdObra, pIdPedido, pIdProveedor, pFecEntrega, pObservaciones, pImportePesos, pImporteDolares, pTipoCambio, pIdDivision, pMoneda, pContrato)
'on error resume next
dim strSQL, rs, conn, rsIns, connIns, auxFec
addCTZCabecera = true

auxFec = pFecEntrega
if pFecEntrega = "" then auxFec = "0"	

strSQL="Select * from TBLCTZCABECERA where IDCOTIZACION = " & pIdCotizacion
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if rs.eof then
	strSQL = "Insert Into TBLCTZCABECERA (IDOBRA, IDPEDIDO, IDPROVEEDOR, FECHAENTREGA, OBSERVACIONES, IMPORTEPESOS, IMPORTEDOLARES, TIPOCAMBIO, MOMENTO, CDUSUARIO, CDMONEDA, IDDIVISION, ESTADO, IDCONTRATO, IMPRESIONES)"
	strSQL = strSQL & " VALUES(" & pIdObra & "," & pIdPedido & "," & pIdProveedor & "," & auxFec & ",'" & pObservaciones & "', " & pImportePesos & "," & pImporteDolares & "," & pTipoCambio & "," & session("MmtoSistema") & ",'" & session("Usuario") & "', '" & pMoneda & "', " & pIdDivision & ", '" & CTZ_PENDIENTE & "', " & pContrato & ", 0)"
	Call executeQueryDB(DBSITE_SQL_INTRA, rsIns, "EXEC", strSQL)
	Call executeQueryDB(DBSITE_SQL_INTRA, rsMax, "OPEN", "Select max(IDCOTIZACION) as Maximo from TBLCTZCABECERA")
	pIdCotizacion = CLng(rsMax("Maximo")) 
else
	strSQL = "Update TBLCTZCABECERA set IDOBRA=" & pIdObra & ", TIPOCAMBIO=" & pTipoCambio & ", IMPORTEPESOS=" & pImportePesos & ", IMPORTEDOLARES=" & pImporteDolares & ", IDPROVEEDOR=" & pIdProveedor & ", FECHAENTREGA=" & auxFec & ", OBSERVACIONES='" & pObservaciones & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoSistema") & ", IDDIVISION=" & pIdDivision & ", CDMONEDA='" & pMoneda & "', ESTADO='" & CTZ_PENDIENTE & "', IDCONTRATO=" & pContrato & " where IDCOTIZACION = " & pIdCotizacion
	Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "UPDATE", strSQL)
end if
end function
'---------------------------------------------------------------------------------------------
sub updCTZEstado(pIdCotizacion, pNuevoEstado)
dim strSQL, rs, conn, rsExec, connExec
strSQL="Select * from TBLCTZCABECERA where IDCOTIZACION = " & pIdCotizacion
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then
	strSQL = "Update TBLCTZCABECERA set ESTADO='" & pNuevoEstado & "' where IDCOTIZACION = " & pIdCotizacion
	Call executeQueryDb(DBSITE_SQL_INTRA, rsExec, "UPDATE", strSQL)
end if
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
end sub
'---------------------------------------------------------------------------------------------
function readCTZ(pIdCotizacion)
'on error resume next
dim strSQL, rs, conn
strSQL="Select * from TBLCTZCABECERA where IDCOTIZACION = " & pIdCotizacion
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
ctz_docCode = "PIC"
if not rs.eof then	
	ctz_IdCotizacion	 = rs("IDCOTIZACION")
	ctz_IdPedido		 = rs("IDPEDIDO")
	ctz_idObra			 = rs("IDOBRA")
	ctz_IdProveedor		 = rs("IDPROVEEDOR")
	ctz_importePesos	 = rs("IMPORTEPESOS")
	ctz_importeDolares	 = rs("IMPORTEDOLARES")
	ctz_FecEntrega		 = rs("FECHAENTREGA")
	ctz_Observaciones	 = rs("OBSERVACIONES")
	ctz_IdContrato	     = rs("IDCONTRATO")
	ctz_IdDivision	     = rs("IDDIVISION")
    if (CLng(ctz_IdContrato) > 0) then ctz_docCode = "CEC"
else
	ctz_IdCotizacion	 = 0
	ctz_IdPedido		 = 0
	ctz_idObra			 = 0
	ctz_IdProveedor		 = 0
	ctz_importePesos	 = ""
	ctz_importeDolares	 = ""
	ctz_FecEntrega		 = ""
	ctz_Observaciones	 = ""
    ctz_IdContrato       = 0
    ctz_IdDivision       = 0
end if
end function
'-------------------------------------------------------------------
'Funcion que permite leer una linea del detalle de un pic. 
Function readCTZDetail(pIdCotizacion, pIdArticulo, pIdArea, pIdDetalle)
	Dim strSQL, rs,conn

	strSQL = "SELECT c.idcotizacion,c.idpedido,c.idobra,c.idproveedor,c.importepesos,c.importedolares,c.FECHAENTREGA, c.idContrato,C.observaciones,d.idarticulo,d.cantidad,d.idunidad,d.idarea,d.iddetalle,d.facturado, d.ImportePesos as ImportePesosDetail, d.ImporteDolares as ImporteDolaresDetail, d.ImportePesosFacturado, d.ImporteDolaresFacturado, d.ImportePesosCredito, d.ImporteDolaresCredito, c.tipocambio, c.cdmoneda"
	strSQL = strSQL & " FROM tblctzdetalle d "
	strSQL = strSQL & " INNER JOIN tblctzcabecera c ON d.idcotizacion = c.idcotizacion "
	strSQL = strSQL & " WHERE d.IDCOTIZACION=" & pIdCotizacion & " AND IDARTICULO=" & pIdArticulo 
	strSQL = strSQL & " AND IDAREA=" & pIdArea & " AND IDDETALLE=" & pIdDetalle
	'Response.Write strSQL
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	ctz_docCode = "PIC"
	if not rs.eof then	
		ctz_IdCotizacion	 = CLng(rs("IDCOTIZACION"))
		ctz_IdPedido		 = CLng(rs("IDPEDIDO"))
		ctz_idObra			 = CLng(rs("IDOBRA"))
		ctz_IdProveedor		 = CLng(rs("IDPROVEEDOR"))
		ctz_importePesos	 = CDbl(rs("IMPORTEPESOS"))
		ctz_importeDolares	 = CDbl(rs("IMPORTEDOLARES"))
		ctz_FecEntrega		 = rs("FECHAENTREGA")
		ctz_Observaciones	 = rs("OBSERVACIONES")
		ctz_cdMoneda		 = rs("CDMONEDA")
		ctz_det_IdArticulo		 = rs("IDARTICULO")
		ctz_det_ArticuloCantidad = CDbl(rs("CANTIDAD"))
		ctz_det_ArticuloIdUnidad = CLng(rs("IDUNIDAD"))
		ctz_det_IdArea			 = CLng(rs("IDAREA"))
		ctz_det_IdDetalle		 = CLng(rs("IDDETALLE"))
		ctz_det_Facturado		 = CDbl(rs("FACTURADO"))
		ctz_det_TipoCambio		 = CDbl(rs("TIPOCAMBIO"))
		ctz_det_ImportePesos 	 = CDbl(rs("IMPORTEPESOSDETAIL"))
		ctz_det_ImporteDolares	 = CDbl(rs("IMPORTEDOLARESDETAIL"))
		ctz_det_ImportePesosFacturado	= CDbl(rs("IMPORTEPESOSFACTURADO"))
		ctz_det_ImporteDolaresFacturado	= CDbl(rs("IMPORTEDOLARESFACTURADO"))
		ctz_det_ImportePesosCredito	= CDbl(rs("IMPORTEPESOSCREDITO"))
		ctz_det_ImporteDolaresCredito	= CDbl(rs("IMPORTEDOLARESCREDITO"))
		ctz_IdContrato	     = CLng(rs("IDCONTRATO"))
		if (CLng(ctz_IdContrato) > 0) then ctz_docCode = "CEC"
	else
		ctz_IdCotizacion	 = 0
		ctz_IdPedido		 = 0
		ctz_idObra			 = 0
		ctz_IdProveedor		 = 0
		ctz_importePesos	 = 0
		ctz_importeDolares	 = 0
		ctz_FecEntrega		 = ""
		ctz_Observaciones	 = ""
		ctz_det_IdArticulo		 = 0
		ctz_det_ArticuloCantidad = 0
		ctz_det_ArticuloIdUnidad = 0
		ctz_det_IdArea			 = 0
		ctz_det_IdDetalle		 = 0
		ctz_det_Facturado		 = 0
		ctz_det_ImportePesos	 = 0
		ctz_det_ImporteDolares	 = 0
		ctz_det_ImportePesosFacturado   = 0
		ctz_det_ImporteDolaresFacturado	= 0
		ctz_det_ImportePesosCredito	= 0
		ctz_det_ImporteDolaresCredito	= 0 
		ctz_det_TipoCambio = 0
	end if
End Function
'-------------------------------------------------------------------
Function getCotizacionesObra(idObra) 
	Dim strSQL, rs, myWhere, conn
	strSQL = "select * from tblctzcabecera where idobra = " & idObra & " order by momento"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	Set getCotizacionesObra = rs
End function
'------------------------------------------------------------------------------------------------
Function AnularPIC(pIdPedido)
	Dim strSQL, conn, rs
	strSQL = "update TBLCTZCABECERA set ESTADO = '" & CTZ_ANULADA & "' where idPedido = " & pIdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	GFG - Guido Fonticelli
' Fecha: 	22/10/10
' Objetivo:	
'			Obtener los articulos que tienen una diferencia mayor al 10% en su valor con
'			respecto a su ultimo cierre contable
' Parametros:
'			pMoneda 	[char] 	Moneda con la cual se obtendra el resultado
'			vArtID		[array]	Vector con los ids de los articulos			
'			vArtPrecioPesos	[array] Vector con los precios de los articulos	en pesos	
'			vArtPrecioDolares	[array] Vector con los precios de los articulos	en dolares
'			vArtCant	[array] Vector con la cantidad de cada articulo
'			fecEntrega	[int] 	Fecha limite para el cierre contable
'			idDivision	[int] 	
' Devuelve:
'			Diccionario
' Modificacion: CNA - Ajaya Nahuel
' Fecha:		01/03/2013		
'--------------------------------------------------------------------------------------------------
Function controlarPrecioArticulo(pMoneda,vArtID,vArtPrecioPesos,vArtPrecioDolares,vArtCant,idDivision, idPic)
	dim articulos,dicArt,precioCierre,diferencia, rs
	dim vPrecios, mensaje, dicRtrn,i,strPrecio, fechaCierre, pos
	
	Set dicRtrn = Server.CreateObject("Scripting.Dictionary")
	Set dicArt = Server.CreateObject("Scripting.Dictionary")
	
	i = 0
	for each articulo in vArtID
			if ( CDbl(vArtCant(i)) <>0 ) then
				if (articulo <> 0) then articulos = articulos & articulo & ","
			end if
			i = i +1
	next
	articulos = left(articulos,len(articulos)-1)

	
	vPrecios = vArtPrecioPesos
	if (pMoneda = MONEDA_DOLAR) then vPrecios = vArtPrecioDolares
	
	'set dicArt = obtenerPreciosArticulos(articulos,idDivision,pMoneda,fecEntrega)
	articulos = listaControlPrecio(articulos)
	Set rs = getUltimaCompra(idDivision, articulos, idPic)
	dicArt.removeAll
	while not rs.EoF
		if (not dicArt.Exists(cstr(rs("IDARTICULO")))) then
			dicArt.add trim(rs("IDARTICULO")), trim(rs("VLUPESOS")) & ";" & Left(rs("MOMENTO"),8)
		end if
		rs.MoveNext
	wend
	
	if (dicArt.count > 0) then
		for i = 0 to ubound(vArtID)			
			if ( vArtCant(i)<>0 ) then
				if (dicArt.Exists(trim(vArtID(i)))) then
					pos = InStr(1,Trim(dicArt.Item(trim(vArtID(i)))),";")
					precioCierre = Left(Trim(dicArt.Item(trim(vArtID(i)))), pos - 1)
					fechaCierre = Right(dicArt.Item(trim(vArtID(i))), Len(Trim(dicArt.Item(trim(vArtID(i))))) - pos)					
					diferencia = ( cdbl(vPrecios(i))/cdbl(vArtCant(i))  ) - cdbl(precioCierre)
					porcentaje = 0
					if (precioCierre > 0) then porcentaje =  round((100 * diferencia) / precioCierre) 
					if (porcentaje > 10) then
						if (not dicRtrn.Exists(vArtID(i)))then 
							strPrecio =  "Precio Ult. Compra: " & getSimboloMoneda(pMoneda) &"&nbsp;"& GF_EDIT_DECIMALS(precioCierre,2) & "| Precio unitario actual:" & getSimboloMoneda(pMoneda) &"&nbsp;"& GF_EDIT_DECIMALS(cdbl(vPrecios(i))/cdbl(vArtCant(i)),2) & "| Diferencia: " & getSimboloMoneda(pMoneda) &"&nbsp;"& GF_EDIT_DECIMALS(diferencia,2) &"&nbsp;  ("& porcentaje & " %) | Fecha Ult. Compra: " & "&nbsp;" & GF_FN2DTE(fechaCierre)
							dicRtrn.Add vArtID(i),strPrecio							
						end if
					end if
				end if
			end if
		next
	end if	
	Set controlarPrecioArticulo = dicRtrn
End Function
'---------------------------------------------------------------------------------------------
function readAjusteCTZ(pIdCotizacion, pIdArticulo)
'on error resume next
dim strSQL, rs, conn
strSQL="Select * from TBLCTZAJUSTES where IDCOTIZACION = " & pIdCotizacion
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then	
	ctz_IdAjuste			= rs("IDAJUSTE")
	ctz_AjusteIdArticulo	= rs("IDARTICULO")
	ctz_AjusteTotalPesos	= rs("IMPORTEPESOS")
	ctz_AjusteTotalDolares	= rs("IMPORTEDOLARES")
	ctz_AjusteObservaciones = rs("OBSERVACIONES")
else
	ctz_IdAjuste = 0
	ctz_AjusteIdArticulo = 0
	ctz_AjusteTotalPesos   = 0
	ctz_AjusteTotalDolares = 0
	ctz_AjusteObservaciones = ""
end if
end function
'---------------------------------------------------------------------------------------------
function existeAjusteCotizacion(pIdCotizacion)
existeAjusteCotizacion = false
dim strSQL, rs, conn
strSQL="Select * from TBLCTZAJUSTES where IDCOTIZACION = " & pIdCotizacion & " AND APLICADO='" & TIPO_AFIRMACION & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then existeAjusteCotizacion = true	
end function
'---------------------------------------------------------------------------------------------
function existeAjusteCotizacionArticulo(pIdCotizacion, pIdArticulo, pArea, pDetalle)
existeAjusteCotizacionArticulo = false
dim strSQL, rs, conn
strSQL="Select * from TBLCTZAJUSTES where IDCOTIZACION = " & pIdCotizacion & " AND IDARTICULO=" & pIdArticulo & " AND IDAREA=" & pArea & " AND IDDETALLE=" & pDetalle & " AND APLICADO='" & TIPO_NEGACION & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then existeAjusteCotizacionArticulo = true	
end function
'---------------------------------------------------------------------------------------------
sub loadTotalesAjustesCotizacion(pIdCotizacion, byref pTotalPesos, byref pTotalDolares)
dim strSQL, rs, conn
pTotalPesos = 0
pTotalDolares = 0
strSQL="Select SUM(IMPORTEPESOS) AS TOTALPESOS, SUM(IMPORTEDOLARES) AS TOTALDOLARES from TBLCTZAJUSTES where IDCOTIZACION = " & pIdCotizacion & " AND APLICADO='" & TIPO_AFIRMACION & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then 
	pTotalPesos = getSimboloMoneda(MONEDA_PESO) & " " & GF_EDIT_DECIMALS(rs("TOTALPESOS"),2)
	pTotalDolares = getSimboloMoneda(MONEDA_DOLAR) & " " & GF_EDIT_DECIMALS(rs("TOTALDOLARES"),2)	
end if	
end sub
'---------------------------------------------------------------------------------------------
function getAjusteCotizacionObservacion(pIdCotizacion)
dim rtrn
dim strSQL, rs, conn
rtrn = ""
strSQL="Select TOP 1 OBSERVACIONES from TBLCTZAJUSTES where IDCOTIZACION = " & pIdCotizacion & " AND APLICADO='" & TIPO_AFIRMACION & "'"
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
if not rs.eof then rtrn = rs("OBSERVACIONES")	
getAjusteCotizacionObservacion = rtrn
end function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli - GFG
' Fecha: 27/12/2011
' Objetivo:
'			Obtener el maximo nro de secuencia de los archivos asociados al pic
' Parametros:
'			[int]	pIdCotizacion
' Devuelve:
'			[int]	maxima secuencia
' Modificaciones:
'			--/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function getCantFilePic(pIdCotizacion)
	Dim strSQL,rs,rtrn
	strSQL = "select max(fileno) maximo from TBLCTZBINARYFILES where idcotizacion = " & pIdCotizacion
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	rtrn =CTZ_SIN_ARCHIVOS
	if (not isnull(rs("maximo"))) then rtrn = cdbl(rs("maximo"))
	getCantFilePic = rtrn
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli - GFG
' Fecha: 28/12/2011
' Objetivo:
'			Guarda en la base de datos el archivo pasado por parametro
' Parametros:
'			[int]	pIdCotizacion
'			[str]	filePath
' Devuelve:
'			nada
' Modificaciones:
'			--/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function picFile2Binary(pIdcotizacion,filePath)
	
	Dim rs, strSQL, extension,fileno,fileName
	
	fileno = 0
	if (pIdcotizacion <> 0) then
		fileno = getCantFilePic(pIdcotizacion) + 1
	end if
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	extension = fso.GetExtensionName(server.MapPath(filePath))
    fileName = fso.getfilename(server.MapPath(filePath))
	fileName = left(fileName,InStrRev(filename,".")-1) 'le quito la extencion
	
	Call FileName2DbName(filename)
		
	strSQL = "select IDCOTIZACION,FILENO,FILEBIN,SIGNATURE,NAME,EXT from TBLCTZBINARYFILES where 1=0"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	rs.AddNew
	rs("IDCOTIZACION") = pIdcotizacion
	rs("FILENO") = fileno
	rs("SIGNATURE") = generateFileSignature(filePath)
	rs("EXT") = extension
	rs("NAME") = fileName
	rs("FILEBIN") = readBinaryFile(server.MapPath(filePath))

	rs.Update

End Function
'-------------------------------------------------------------------------------------------------------------'
' Nombre Funcion:		  cargaPICMensajeNDA(pIdCotizacion)
' Objetivo:
'			Obtener diferentes datos de una cotizacion, sea sus proveedores, fechas, usuarios; con el proposito de completar con estos datos 
' 			un mensaje que indique el alta de una nota de Aceptacion
' Parametros:
'			[int]	pIdCotizacion
' Devuelve:
'			[recordset]
' Autor: Ajaya Nahuel - CNA
' Fecha: 26/03/2012
' Modificaciones: 	
'				--/--/-- - XXX
 '--------------------------------------------------------------------------------------------------
Function cargaPICMensajeNDA(pidCotizacion)			 
	Dim rs, strSQL
	strSQL="SELECT ctz.idproveedor,ctz.fechaentrega, prov.nomemp as dsempresa,ctz.CDUSUARIO,ctz.MOMENTO "
	strSQL = strSQL & "FROM TBLCTZCABECERA ctz "
	strSQL = strSQL & "INNER JOIN [Database].[dbo].MET001A prov "
	strSQL = strSQL & "    ON ctz.idproveedor = prov.nroemp "	
	strSQL = strSQL & " WHERE  ctz.IDCOTIZACION = " & pidCotizacion
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
	Set cargaPICMensajeNDA = rs	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Scalisi, Javier alejandro
' Fecha: 	20/12/2012
' Parametros: idArtículo
' Devuelve: Determina si el artículo indicado debe incluirse en el control de AFEs del PIC.
'-------------------------------------------------------------------------------------------------
Function incluirArticuloControlAFE(pIdArticulo)
	Dim strSQL, rs
	
	incluirArticuloControlAFE = false
	
	strSQL=	"Select * from TBLARTCATEGORIAS CAT inner join TBLARTICULOS ART on CAT.IDCATEGORIA=ART.IDCATEGORIA " & _
			" where IDARTICULO=" & pIdArticulo & _
			" and TIPOCATEGORIA not in ('" & TIPO_CAT_IMPUESTOS & "', '" & TIPO_CAT_FONDO_REPARO & "')"
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if (not rs.eof) then incluirArticuloControlAFE = true
	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: 	Javier Scalisi
' Fecha: 	14/08/2013
'Función que verifica si el artículo indicado es elegible para un PIC.
Function esArticuloElegiblePIC(idArticulo)
    esArticuloElegible = false
    if (idArticulo <> CTZ_ITEM_DIFF_CAMBIO) then esArticuloElegiblePIC = true
End Function
'--------------------------------------------------------------------------------------------------
'Funcion encargada de totalizar las compras directas realizadas a un proveedor en un período dado de tiempo.
'Parametros
'Autor: Javier A. Scalisi
'Fecha: 17/07/2014
Function totalizarComprasDirectasProveedor(idProveedor, idDivision, mmtoDesde, mmtoHasta, cdMoneda, ByRef retCant, ByRef retImporte)
    Dim sp_ret
        
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA , rs, "TBLCTZCABECERA_GET_TOTAL_CD_BY_PROV_MMTO", idProveedor & "||" & idDivision & "||" & mmtoDesde & "||" & mmtoHasta & "$$CANT||DOL||PES")
    'response.Write sp_ret(SP_DSERROR) & "<br>"    
    retCant = CLng(rs("CANTIDAD"))
    retImporte = 0
    if (retCant > 0) then
        if (cdMoneda = MONEDA_PESO) then
            retImporte = CDbl(rs("MONTO_PESOS"))
        else
            retImporte = CDbl(rs("MONTO_DOLARES"))
        end if
    end if
End Function
'-----------------------------------------------------------------------------------------------
Function getPICAuthorizationType(pIdPedido, pIdContrato, pIdProveedor, pImporte, pMoneda)
    Dim limiteFirmaSupervisor, limiteFirmaDireccion, limiteFirmaGte
    Dim myImporte, unidadCD
    
    getPICAuthorizationType = PIC_TYPE_PURCHASE_SMALL  
    if (pIdContrato = 0) then 
        if (pIdPedido = 0) then    
            'Es compraDirecta
		    'Verifico si el proveedor esta pre-autorizado.
		    Call executeProcedureDb(DBSITE_SQL_INTRA , rsProv, "TBLPROVEEDORESCD_GET_BY_IDPROVEEDOR", pIdProveedor)
		    if (rsProv.eof) then			
                'Se traen los limites de firma para uso y controles.
                limiteFirmaGte = CDbl(getValorNorma("VLMAXGS"))
                limiteFirmaSupervisor = CDbl(getValorNorma("VLMAXCD"))
                limiteFirmaDireccion = CDbl(getValorNorma("VLMAXSP"))
                unidadCD = getUnidadNorma("VLMAXSP")
                    
                'Se transforma el importe a la moneda de la regal de auditoria. (Importe con 2 decimales!)
                myImporte = pImporte
	            if (pMoneda <> unidadCD) then
		            if (pMoneda = MONEDA_PESO) then	
			            myImporte = round(myImporte / getTipoCambio(MONEDA_DOLAR, ""),2)
		            else
			            myImporte = round(myImporte * getTipoCambio(MONEDA_DOLAR, ""),2)
		            end if
	            end if	
            	
	            'Controlo el importe contra los limites
	            if (Cdbl(myImporte) < CDbl(limiteFirmaSupervisor)) then getPICAuthorizationType = PIC_TYPE_PURCHASE_MEDIUM
	            if (Cdbl(myImporte) <= CDbl(limiteFirmaGte)) then getPICAuthorizationType = PIC_TYPE_PURCHASE_SMALL	        
	            if ((Cdbl(myImporte) >= Cdbl(limiteFirmaSupervisor)) and (myImporte < limiteFirmaDireccion)) then getPICAuthorizationType = PIC_TYPE_PURCHASE_X_MEDIUM
	            if (Cdbl(myImporte) >= Cdbl(limiteFirmaDireccion)) then getPICAuthorizationType = PIC_TYPE_PURCHASE_LARGE			            				            
	        end if
        end if
    End if
End Function    
%>