<%
dim pct_idPedido, pct_FechaInicio, pct_FechaCierre, pct_cdSolicitante, pct_dsSolicitante, pct_idDivision, pct_dsDivision
dim pct_idEstado, pct_usuario, pct_momento, pct_dsPedido, pct_emailProveedor
dim pct_idProveedor, pct_dsProveedor, pct_rsProveedores, pct_hayCotizacion, pct_idProveedorElegido, pct_dsProveedorElegido
dim pct_CantProveedores, pct_ProveedorActual, pct_cdPedido, pct_cdUsuarioAdmin,pct_dsUsuarioAdmin, pct_usuarioCarga, pct_momentoCarga
dim pct_hayCabecera, pct_pathCotizacion, pct_idSector
dim pct_idObra, pct_observaciones, pct_tipoCompra, pct_tituloPedido, pct_idDivisionObra
Dim pct_idArea, pct_idDetalle, pct_Extension

'Constante
Const PCT_BINARY_SPECIFICATION = -2
Const PCT_BINARY_CONDITIONS = -1

Const MAX_SIGNATURE_SIZE = 819200	'800 KB

Const ACCION_PCT_COTIZAR = "cotizar"
Const ACCION_PCT_APERTURA = "apertura"
Const ACCION_PCT_RETIRARSE = "NO_COTIZA"

Const SERVICIO_DEFAULT = "1"

'/* ESTADOS DE UN PCT */
Const ESTADO_PCT_PENDIENTE=10	'Pedido recien cargado
Const ESTADO_PCT_AUTORIZADO=11	'Pedido aprobado por solicitante
Const ESTADO_PCT_PUBLICADO=15	'Pedido enviado a los proveedores.
Const ESTADO_PCT_COTIZADO=20	'Todos los proveedores enviaron su cotizacion o bien se cumplio el plazo
Const ESTADO_PCT_ABIERTO=21		'Se ha realizado la apertura de sobres
Const ESTADO_PCT_EN_ANALISIS=22	'Se cargo la planilla comparativa y aun no se registan firmas.
Const ESTADO_PCT_EN_FIRMA_AC=23	'Se han registrado algunas de las firmas requeridas para adjudicar.
Const ESTADO_PCT_ADJUDICADO=30	'Se ha completado la firma del Analisis Comparativo.
Const ESTADO_PCT_APROBADO=50	'Con todas las firmas sobre el Pedido Interno (Cotizacion Elegida)
Const ESTADO_PCT_CANCELADO=60	'Algun firmante rechazo el proyecto

Const TIPO_PCT_CONCURSO = "C"
Const TIPO_PCT_COMPARATIVA = "P"

Const TEXTO_EXTENSION_PEDIDO = "#EXTENSION"
Const ACCION_ARCHIVO_LECTURA = "LEER"

const ACCION_IMPRIMIR_NDA = 1
const ACCION_ENVIAR_NDA = 2
const ACCION_LEER_NDA = 4
const ACCION_ACTUALIZAR_NDA = 5

'/*** POLIZAS DE CAUCION (PDC)  ***/
'----------ESTADOS--------------
const ESTADO_PDC_PENDIENTE = 1
const ESTADO_PDC_RECIBIDA  = 2
const ESTADO_PDC_VENCIDA   = 3
const ESTADO_PDC_DEVUELTA  = 4
const ESTADO_PDC_ANULADA   = 5
'------------TIPO--------------
const TIPO_PDC_POR_ADELANTO  = "A"
'------------ SEPARADORES CRC ---------------
Const SEPARATOR_CRC_PROVEEDOR = "%$"
Const SEPARATOR_CRC_PEDIDO = "#?"
'---------------------------------------------------------------------------------------------
Function initHeader(pIdPedido)
	Call clearHeader()		
	if (isFormSubmit()) then
		Call initHeaderParams(pIdPedido)
	else 
		if (pIdPedido > 0) then 
			Call initHeaderDB(pIdPedido)					
		else		
			Call initHeaderNuevo()			
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderNuevo() 
	pct_FechaInicio = left(GF_VERFECHADATO(),10)
	pct_FechaCierre = left(GF_VERFECHADATO(),10)
	pct_cdUsuarioAdmin = session("Usuario")	
	
	pct_dsUsuarioAdmin = getUserDescription(pct_cdUsuarioAdmin)	
	pct_momentoCarga = session("MmtoSistema")
	pct_usuarioCarga = session("Usuario")
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderParams(pIdPedido)
	dim strSQL, rs, km, kc, rsObra
	
	pct_idPedido = pIdPedido	
	pct_FechaInicio = GF_PARAMETROS7("issuedate","",6)
	pct_FechaCierre = GF_PARAMETROS7("closingdate","",6)
	pct_cdSolicitante = GF_PARAMETROS7("cdSolicitante","",6)
	pct_dsSolicitante = getUserDescription(pct_cdSolicitante)
	pct_cdPedido = GF_PARAMETROS7("cdPedido","",6)
	pct_idObra = GF_PARAMETROS7("idobra",0,6)	
	pct_idArea = GF_PARAMETROS7("cmbIdArea",0,6)		
	pct_idDetalle = GF_PARAMETROS7("cmbIdDeta",0,6)			
	pct_idDivision = GF_PARAMETROS7("idDivision",0,6)	
	strSQL = "Select * from TBLDATOSOBRAS where IDOBRA=" & pct_idObra		
	Call executeQueryDb(DBSITE_SQL_INTRA, rsObra, "OPEN", strSQL)
	if (not rsObra.eof) then pct_idDivisionObra = rsObra("IDDIVISION")	
	pct_dsDivision = getDivisionDS(pct_idDivision)
	pct_dsPedido = GF_PARAMETROS7("description","",6)	
	pct_tituloPedido = GF_PARAMETROS7("titulo","",6)				
	pct_idEstado = CInt(GF_PARAMETROS7("idEstado",0,6))
	if (idEstado = 0) then idEstado = ESTADO_PCT_PENDIENTE	 		
	pct_observaciones = GF_PARAMETROS7("observaciones","",6)
	pct_usuario = session("Usuario")
	pct_cdUsuarioAdmin = GF_PARAMETROS7("cdUsuarioAdmin","",6)	
	pct_dsUsuarioAdmin = getUserDescription(pct_cdUsuarioAdmin)
	pct_usuarioCarga = GF_PARAMETROS7("cdUsuarioCarga","",6)
	pct_momentoCarga = GF_PARAMETROS7("momentoCarga","",6)
	pct_momento = session("MmtoSistema")
	if (pct_idPedido = 0) then	pct_momento = pct_momentoCarga	
	pct_tipoCompra = TIPO_PCT_COMPARATIVA
	pct_Extension = GF_PARAMETROS7("extension","",6)
	
	
	pct_idProveedorElegido = 0
	pct_dsProveedorElegido = ""
	pct_hayCabecera = True
	
End Function
'---------------------------------------------------------------------------------------------
Function initHeaderDB(pIdPedido)
	dim strSQL, rs 
	
	pct_hayCabecera = False
	strSQL="select * from TBLPCTCABECERA where IDPEDIDO=" & pIdPedido
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		Call pctCargarVariables(rs)	
		pct_hayCabecera = True
	end if		
End Function
'---------------------------------------------------------------------------------------------
Function pctCargarVariables(rs)
	Dim km
	
	pct_idPedido = rs("IDPEDIDO")
	pct_FechaInicio = GF_FN2DTE(rs("FECHAINICIO"))
	pct_FechaCierre = GF_FN2DTE(rs("FECHACIERRE"))
	pct_cdSolicitante = rs("CDSOLICITANTE")
	pct_dsSolicitante = getUserDescription(pct_cdSolicitante)
	pct_dsPedido = rs("DSPEDIDO")
	pct_tituloPedido = rs("TITULO")				
	pct_idObra = rs("IDOBRA")
	pct_idArea = rs("IDAREA")
	pct_idDetalle = rs("IDDETALLE")
	if (pct_idObra = "") then pct_idObra = 0		
	pct_idDivision = rs("IDDIVISION")
	pct_dsDivision = getDivisionDS(pct_idDivision)
	pct_cdPedido = rs("CDPEDIDO")	
	pct_idSector = 0
	if (isNumeric(rs("IDSECTOR"))) then pct_idSector = CLng(rs("IDSECTOR"))
	pct_idEstado = rs("ESTADO")
	pct_observaciones = rs("OBSERVACIONES")
	pct_usuario = rs("CDUSUARIO")
	pct_usuarioCarga = rs("CDUSRCARGA")
	pct_cdUsuarioAdmin = rs("CDUSRADMIN")
	pct_dsUsuarioAdmin = getUserDescription(pct_cdUsuarioAdmin)
	pct_momento = rs("MOMENTO")
	pct_momentoCarga = rs("MMTOCARGA")	
	pct_idProveedorElegido = rs("IDPROVEEDOR")
	if (isNull(pct_idProveedorElegido)) then pct_idProveedorElegido = 0
	pct_dsProveedorElegido = getDescripcionProveedor(pct_idProveedorElegido)
	pct_tipoCompra = rs("TIPOCOMPRA")
    pct_Extension = rs("EXTENSIBLE")
    if (isNull(pct_Extension)) then pct_Extension = TIPO_NEGACION
End Function
'---------------------------------------------------------------------------------------------
Function initProveedores()
	if (isFormSubmit()) then		
		pct_CantProveedores = GF_PARAMETROS7("cantProveedores",0,6)		
		pct_ProveedorActual=0
		initProveedores = true
	else
		initProveedores = initProveedoresDB()
	end if
End Function
'---------------------------------------------------------------------------------------------
Function initProveedoresDB()
	dim strSQL, rs, km, kc
	
	initProveedoresDB = false	
	if (pct_hayCabecera) then		
		strSQL="select * from TBLPCTPROVEEDORES where IDPEDIDO=" & pct_idPedido
		call executeQueryDb(DBSITE_SQL_INTRA, pct_rsProveedores, "OPEN", strSQL)
		if (not pct_rsProveedores.eof) then		
			initProveedoresDB = true
		end if
	end if
End Function
'---------------------------------------------------------------------------------------------
Function readNextProveedor()
	Call clearProveedor()
	if (isFormSubmit()) then
		readNextProveedor = readNextProveedorParams()
	else
		readNextProveedor = readNextProveedorDB()
	end if
End Function
'---------------------------------------------------------------------------------------------
Function readNextProveedorParams()
	Dim ret 
	ret = false
	while ((pct_ProveedorActual < pct_CantProveedores) and (not ret))
		pct_idProveedor = GF_PARAMETROS7("supplier" & pct_ProveedorActual,0,6)
		if (pct_idProveedor > 0) then		
			pct_dsProveedor = getDescripcionProveedor(pct_idProveedor)				
			pct_emailProveedor = obtenerMail(pct_idProveedor)
			ret = true
		end if
		pct_hayCotizacion = false		
		pct_pathCotizacion = ""
		pct_ProveedorActual=pct_ProveedorActual + 1
	wend
	readNextProveedorParams = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextProveedorDB()	
	Dim strSQL, rs
	
	readNextProveedorDB = false	
	if (not pct_rsProveedores.eof) then
		pct_idProveedor = pct_rsProveedores("IDPROVEEDOR")
		pct_dsProveedor = getDescripcionProveedor(pct_idProveedor)
		'Averiguo si tiene cotizaciones presentadas
		Set rs = getCotizaciones(pct_idPedido, pct_idProveedor)
		pct_hayCotizacion = false		
		if (not rs.eof) then 			
			pct_hayCotizacion = true
			pct_pathCotizacion = rs("PATHCOTIZACION")
		end if	
		pct_emailProveedor = obtenerMail(pct_idProveedor)
		pct_rsProveedores.MoveNext()
		readNextProveedorDB = true
	end if	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos del pedido de cotizaci�n.
Function controlarPedidoCotizacion(idPedido)
	
	Dim tmp, cantProv, provs, nrmName
	
	Set provs=Server.CreateObject("Scripting.Dictionary")
	controlarPedidoCotizacion = false		
	Call initHeader(idPedido)	
	if (controlarHeader()) then
		'Se controlan los proveedores
		tmp = true		
		if (initProveedores()) then				
			cantProv = 0
			while ((readNextProveedor()) and (tmp))
				tmp = controlarProveedorPCT()			
				if (not provs.Exists(pct_idProveedor)) then
					provs.add pct_idProveedor, 1
				else
					tmp=false
					setError(PROVEEDOR_REPETIDO)
				end if
				cantProv = cantProv + 1				
			wend			
		end if				
		if (tmp) and (cantProv < cint(getValorNorma("MINPRCP"))) then
			'Si hay menos de 3 proveedores y no hay observaciones, falta justificacion.
			if (pct_observaciones = "") then 
				setError(FALTA_JUSTIFICACION)
				tmp=false
			end if
		end if							
		controlarPedidoCotizacion = tmp
	end if
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de la cabecera cargada.
Function controlarHeader() 
	
	Dim myFechaCierreNueva
	
	controlarHeader = false
	if (pct_idDivision <> SIN_DIVISION) then
		if (pct_idObra <> 0) then			
			if (pct_idDivision <> pct_idDivisionObra) then setError(DIVISION_PCT_DIFF_OBRA)			
		end if
		if (not hayError()) then		
			if (isAdmin(pct_idDivision) or isUser(pct_idDivision)) then
				if (pct_cdSolicitante <> "") then		
					if (GF_CONTROL_PERIODO_2(pct_FechaInicio, pct_FechaCierre) = 0) then					
						if (pct_tituloPedido <> "") then								
							'Si el control anterior result� OK se procede a controlar que no se est� modificando la fecha de cierre m�s alla de 10 d�as de la fecha original si ya cerro o 15 si a�n no cerr�.
							
							'El control sobre la fecha de cierre intenta garantizar que el pedido solo se estire una vez 
							'y como m�ximo 10/15 d�as corredios desde su fecha de vencimiento original dependiendo de si ya cerr� o no.
							'As� como esta implementado el control no cumplir�a el objetivo, por que si bien se controla que no hayana pasado
							'mas de 10 d�as de vencido y que no se elija una fecha de m�s de 10 d�as en el futuro, al modificarse se pierde el valor original.
							if (pct_IdPedido > 0) then
								'Si el pedido ya estaba vencido pero se modific� la fecha de cierre se debe validar el cambio.
								strSQL = "select FECHACIERRE from TBLPCTCABECERA where IDPEDIDO = "& pct_IdPedido 									
								Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
								myFechaCierreNueva = CDbl(GF_DTE2FN(pct_FechaCierre))
								if (CDbl(rs("FECHACIERRE")) <> myFechaCierreNueva) then
								    if (GF_DTEDIFF(rs("FECHACIERRE"), Left(session("MmtoSistema"),8),"D") > 0) then
								        'Ya venci�n								    
									    maxDias = 10									    
                                    else
                                        'Todavia esta publicado.
                                        maxDias = 15
                                    end if                                        									    
                                    if (GF_DTEDIFF(rs("FECHACIERRE"), myFechaCierreNueva,"D") < 15) then
									    controlarHeader = true									
								    else
									    setError(PCT_LIMITE_CIERRE)
								    end if                                                             
                                    									    
								else								    
									controlarHeader = true
								end if							
							else
								controlarHeader = true
							end if							
						else
							setError(FALTA_TITULO)
						end if			
					else
						setError(PERIODO_ERRONEO)
					end if
				else
					setError(SOLICITANTE_NO_EXISTE)
				end if	
			else
				setError(USUARIO_NO_AUTORIZADO)
			end if
		end if
	else
		setError(DIVISION_NO_EXISTE)
	end if
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un proveedor.
Function controlarProveedorPCT()
	controlarProveedorPCT = controlarProveedor(pct_idProveedor)
End Function
'---------------------------------------------------------------------------------------------
'Devuelve la cantidad de cotizaciones que tiene un pedido
Function getCantidadCotizaciones(pIdPedido) 	
	Dim rs, strSQL, rtrn
	rtrn = 0
	strSQL="select count(*) as Cantidad from TBLPCTPROVEEDORES where IDPEDIDO=" & pIdPedido
	'Response.Write strsql
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		if not isnull(rs("Cantidad")) then rtrn = rs("Cantidad")
	end if
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
	getCantidadCotizaciones = rtrn
End Function
'---------------------------------------------------------------------------------------------
function getCantidadCotizacionesRecibidas(idPedido) 
	dim strSQL, rs
	getCantidadCotizacionesRecibidas = 0
	strSQL = "Select count(*) Cantidad from (Select Distinct(IDPROVEEDOR) from TBLPCTCOTIZACIONES where IDPEDIDO=" & idPedido & ") as TABLA"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then getCantidadCotizacionesRecibidas = rs("Cantidad")
end function
'---------------------------------------------------------------------------------------------
Function hayEspecifTecnica(idPedido)
	Dim strSQL, rs
	
	hayEspecifTecnica=false
	strSQL = "Select IDPEDIDO from TBLPCTBINARYFILES where IDPEDIDO= " & idPedido & " and FILENO=" & PCT_BINARY_SPECIFICATION
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		hayEspecifTecnica = true	
	end if
End Function
'---------------------------------------------------------------------------------------------
Function hayCondParticulares(idPedido)
	Dim strSQL, rs
	
	hayCondParticulares=false
	strSQL = "Select IDPEDIDO from TBLPCTBINARYFILES where IDPEDIDO= " & idPedido & " and FILENO=" & PCT_BINARY_CONDITIONS
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		hayCondParticulares = true	
	end if
	
End Function
'---------------------------------------------------------------------------------------------
Function grabarFormulario() 
	pct_idPedido = grabarHeader()
	Call grabarProveedores()
	grabarFormulario = pct_idPedido
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeader()
		
	if (pct_idPedido = 0) then		
		grabarHeader = grabarHeaderInsert()
	else			
		Call grabarHeaderUpdate()
		grabarHeader = pct_idPedido
	end if			
	'Se graban los datos binarios de los archivos adjuntos.	
	Call pctGrabarBinarios(pct_idPedido)	
End Function
'---------------------------------------------------------------------------------------------
'Compara un archivo con uno ya almacenado y permite saber si se modific�.
'pPath: Path al archivo nuevo a comparar.
'idFile: ID del archivo almacenado para el pedido (E. Tecnica, Reglamento o ID CTZ)
Function isFileModified(pPath, idPedido, idFile)	
	Dim strSQL, rs, signature, fso
	
	'Obtengo la firma del archivo almacenado.
	strSQL = "Select SIGNATURE from TBLPCTBINARYFILES where IDPEDIDO= " & idPedido & " and FILENO=" & idFile
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	isFileModified = false	
	if (pPath <> "") then		
		Set fso = CreateObject("Scripting.FileSystemObject")
		if (fso.FileExists(pPath)) then	'Si el archivo no existe, se ignora el proc. y nada cambia.
			'Creo la firma del archivo a comparar.
			if (not rs.eof) then
				if (not isNull(rs("SIGNATURE"))) then
					signature = generateFileSignature(pPath)
					if (rs("SIGNATURE") <> signature) then isFileModified = true				
				else
					isFileModified = true
				end if			
			else
				isFileModified = true
			end if
		end if
	else
		'No haya archivo elegido
		if (not rs.eof) then isFileModified = true
	end if
End Function
'---------------------------------------------------------------------------------------------
Function pctGrabaArchivo(idPedido, pPath, fileno)
	Dim path, strSQL, rs
	
	pctGrabaArchivo = false
	'Se cambio el archivo.	
	strSQL = "Delete from TBLPCTBINARYFILES where IDPEDIDO=" & idPedido & " and FILENO=" & fileno
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	if (pPath <> "") then
		Call pctFile2Binary(idPedido, fileno, pPath)
		pctGrabaArchivo = true
	end if	
	
End Function
'---------------------------------------------------------------------------------------------
'Graba los archivos adjuntados a un PCT en la BD.
Function pctGrabarBinarios(idPedido)
	Dim path, fileEspecifTecnica, fileReglamento
		
	Set fso = CreateObject("Scripting.FileSystemObject")
	fileEspecifTecnica = GF_PARAMETROS7("etFile","",6)	
	path = ""
	if (fileEspecifTecnica <> "") then path = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & fileEspecifTecnica
	if (isFileModified(path, idPedido, PCT_BINARY_SPECIFICATION)) then		
		if (pctGrabaArchivo(idPedido, path, PCT_BINARY_SPECIFICATION)) then fso.DeleteFile(path)
	end if	
	fileReglamento = GF_PARAMETROS7("rgFile","",6)	
	path = ""
	if (fileReglamento <> "") then path = server.MapPath(".") & "\" & PATH_COMPRAS_TEMP & "\" & fileReglamento
	if (isFileModified(path, idPedido, PCT_BINARY_CONDITIONS)) then		
		if (pctGrabaArchivo(idPedido, path, PCT_BINARY_CONDITIONS)) then fso.DeleteFile(path)
	end if	
	Set fso = nothing
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderInsert()
	Dim strSQL, rs, dte, idPedido, cdPedido, estado
	pct_cdPedido = getIdPedidoCompleto()
	
	dte = Left(GF_DTE2FN(pct_FechaInicio),4) & "0101"
	strSQL = "Select MAX(IDPEDIDO) IDPEDIDO from TBLPCTCABECERA"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (isNull(rs("IDPEDIDO"))) then 
		idPedido=0
	else
		idPedido=rs("IDPEDIDO")
	end if
	idPedido= idPedido+1
	estado = ESTADO_PCT_PENDIENTE
	if (pct_cdSolicitante = session("Usuario")) then estado = ESTADO_PCT_AUTORIZADO
	strSQL= "Insert into TBLPCTCABECERA(IDPEDIDO,IDOBRA, CDSOLICITANTE, FECHAINICIO, FECHACIERRE, DSPEDIDO, TITULO, IDDIVISION, ESTADO, OBSERVACIONES, CDUSUARIO, MOMENTO, CDPEDIDO, TIPOCOMPRA, CDUSRCARGA, MMTOCARGA, CDUSRADMIN,IDAREA,IDDETALLE,EXTENSIBLE) values(" & idPedido
	strSQL = strSQL & ", " & pct_idObra & ", '" & pct_cdSolicitante & "', " & GF_DTE2FN(pct_FechaInicio) & ", " & GF_DTE2FN(pct_FechaCierre) & ", '" & pct_dsPedido & "','" & UCASE(pct_tituloPedido) & "', " & pct_idDivision & ", " & estado
	strSQL = strSQL & ", '" & pct_observaciones & "', '" & session("Usuario") & "', " & session("MmtoSistema") & ", '" & pct_cdPedido & "','" & pct_tipoCompra & "', '" & session("Usuario") & "', " & session("MmtoSistema") & ",'" & pct_cdUsuarioAdmin & "',"&pct_idArea&","&pct_idDetalle&",'"&TIPO_AFIRMACION&"')"
	'Response.Write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	pct_idPedido = idPedido
	grabarHeaderInsert = pct_idPedido
End Function
'---------------------------------------------------------------------------------------------
Function grabarHeaderUpdate()
	Dim strSQL, rs
	
	strSQL = "Update TBLPCTCABECERA set IDOBRA = " & pct_idObra & ", CDSOLICITANTE='" & pct_cdSolicitante & "', FECHAINICIO=" & GF_DTE2FN(pct_FechaInicio)
	strSQL = strSQL & ", FECHACIERRE=" & GF_DTE2FN(pct_FechaCierre) & ", DSPEDIDO='" & pct_dsPedido & "', TITULO='" & pct_tituloPedido & "', IDDIVISION=" & pct_idDivision & ", ESTADO=" & pct_idEstado
	strSQL = strSQL & ", OBSERVACIONES='" & pct_observaciones & "', MOMENTO=" & session("MmtoSistema") & ", TIPOCOMPRA='" & pct_tipoCompra & "', CDUSUARIO='" & session("Usuario") & "', CDUSRADMIN='" & pct_cdUsuarioAdmin & "'"
	strSQL = strSQl & ", IDAREA = " & pct_idArea & ", IDDETALLE = " & pct_idDetalle &", EXTENSIBLE = '"&pct_Extension&"'"
	strSQL = strSQL & " where IDPEDIDO=" & pct_idPedido
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	
End Function
'---------------------------------------------------------------------------------------------
Function cancelarPedido(idPedido, motivo)

	Call initHeaderDB(idPedido)
	pct_idEstado = ESTADO_PCT_CANCELADO
	pct_Extension = TIPO_NEGACION
	pct_Observaciones = pct_observaciones & " - PEDIDO CANCELADO: " & motivo
	Call grabarHeaderUpdate()		
	
End Function
'---------------------------------------------------------------------------------------------
Function grabarProveedores()
	Dim strSQL, rs, xxx
			
	'Borro los proveedores viejos
	strSQL= "Delete from TBLPCTPROVEEDORES where IDPEDIDO=" & pct_idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	'Guardo los proveedores nuevos
	Call initProveedores()			
	while (readNextProveedor())			
		strSQL= "Insert into TBLPCTPROVEEDORES(IDPEDIDO, IDPROVEEDOR) values(" & pct_idPedido & ", " & pct_idProveedor & ")"
		'Response.Write strSQL
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	wend		
End Function
'---------------------------------------------------------------------------------------------
Function getCotizaciones(idPedido, idProveedor)
	Dim strSQL, rs	
	strSQL = "Select * from TBLPCTCOTIZACIONES where IDPEDIDO=" & idPedido & " and IDPROVEEDOR=" & idProveedor
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getCotizaciones = rs	
End Function
'---------------------------------------------------------------------------------------------
Function getUsuariosApertura(pIdPedido)
	Dim strSQL, rs, rtrn
	
	strSQL = "Select * from TBLPCTFIRMASAPERTURA where IDPEDIDO=" & pIdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof
		rtrn = rtrn & rs("CDUSUARIO") & "-"
		rs.movenext
	wend		
	if (LEN(rtrn) > 0) then rtrn = left(rtrn, len(rtrn)-1)
	getUsuariosApertura = ucase(rtrn)
End Function
'---------------------------------------------------------------------------------------------
'Para utilizar este metodo se debe ejecutar primero initHeader para inicializar los datos del PCT
Function getIdPedidoCompleto() 
	Dim cdDivision, rsObra, rsDivision, strSQL, dte, idDiv, rs
	
	'Determino la division a la que pertenece el pedido.
	cdDivision = ALL
	idDiv = pct_idDivision
	strSQL="Select * from TBLDIVISIONES where IDDIVISION=" & idDiv
	Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
	if (not rsDivision.eof) then cdDivision = rsDivision("CDDIVISION")
	'Obtengo la porcion de la fecha que compone el codigo.
	dte = Right(Left(GF_DTE2FN(pct_FechaInicio),4),2)
	strSQL="Select * from TBLNUMERACION where CLAVE='" & cdDivision & "_" & dte & "' and PREFIJO='" & PREFIX_PCT & "'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then
		pct_idPedido = clng(rs("Valor")) + 1
		strsql = "Update TBLNUMERACION set VALOR=" & pct_idPedido & " where CLAVE = '" & cdDivision & "_" & dte & "' and PREFIJO='" & PREFIX_PCT & "'"	
        Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "UPDATE", strSQL)
	else
		pct_idPedido = 1
		strSQL = "Insert into TBLNUMERACION values('" & PREFIX_PCT & "','" & cdDivision & "_" & dte & "',1)"
        Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "EXEC", strSQL)
	end if
	if (cdDivision = ALL) then cdDivision = "ALL"
	getIdPedidoCompleto = cdDivision & "-" & GF_nDigits(pct_idPedido, 3) & "-" & dte
	
End Function
'---------------------------------------------------------------------------------------------
'	Se determina el estado que tendr�a que tener el pedido.
Function actualizarEstado(rs)
	Dim hayCambio, sUsr
	
	Call pctCargarVariables(rs)
	
	if ((pct_idEstado >= ESTADO_PCT_PUBLICADO) and _
		(pct_idEstado < ESTADO_PCT_ABIERTO)) then 
			estado = ESTADO_PCT_PUBLICADO
			'Si ya concluyo el per�odo dispuesto, se da por cerrado el pedido.
			if (GF_DTEDIFF(session("MmtoSistema"), GF_DTE2FN(pct_FechaCierre), "D") < 0) then
				'Ya paso la fecha de cierre.
				hayCambio = true
			else
				if (not (hayCotizacionesPendientes(pct_idPedido))) then
					if (hayCotizacionesMinimas(pct_idPedido)) then	hayCambio = true
				end if
			end if
            if (hayCambio) then	estado = ESTADO_PCT_COTIZADO
			if (estado <> pct_idEstado) then	
				pct_idEstado = estado	
				'Antes de actualizar el estado del pedido se toma el usuario de la session para que no quede �ste como responsable del cambio.
				'Esto se hace as� dado que el usuario del sistema no esta enterado de este proceso. (Tarea: INT-974)
				sUsr = session("Usuario")
				session("Usuario") = "UPDSTATUS"
				Call grabarHeaderUpdate()
				session("Usuario") = sUsr
			end if			
	end if				
End Function
'---------------------------------------------------------------------------------------------
Function actualizarExtensible()	
		
	if (pct_Extension = TIPO_AFIRMACION) then
		'Si aun figura como extensible, se revisa que no hayan pasado mas de 8 dias del vencimiento.
		if (GF_DTEDIFF(GF_DTE2FN(pct_FechaCierre), session("MmtoSistema"), "D") > 8) then 
			pct_Extension = TIPO_NEGACION
			sUsr = session("Usuario")
			session("Usuario") = "UPDSTATUS"
			Call grabarHeaderUpdate()
			session("Usuario") = sUsr
		end if
	end if		
	
End Function
'---------------------------------------------------------------------------------------------
'Devuelve verdadero si alguno de los proveedores no ha cargado su cotizacion
Function hayCotizacionesPendientes(pIdPedido) 	
	Dim rs, strSQL, rtrn, flag
	rtrn = false
	flag = 0
	strSQL = "SELECT * FROM TBLPCTPROVEEDORES P LEFT JOIN TBLPCTCOTIZACIONES C ON P.IdPedido = C.IdPedido AND P.IdProveedor = C.IdProveedor where P.IdPedido=" & pIdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while not rs.eof and flag = 0
		if isnull(rs("idCotizacion")) then 
			rtrn = true
			flag = 1
		end if	
		rs.movenext
	wend		
    hayCotizacionesPendientes = rtrn
End Function
'---------------------------------------------------------------------------------------------
'Devuelve verdadero si el numero de cotizaciones presentadas es mayor o igual al minimo requerido
Function hayCotizacionesMinimas(pIdPedido) 	
	Dim rs, strSQL, rtrn, cantidad, numMin, cdNorma
	rtrn = false
	flag = 0
	strSQL = "Select count(*) as Cantidad from (Select Distinct(IDPROVEEDOR) from TBLPCTCOTIZACIONES where IDPEDIDO=" & pIdPedido & " and PATHCOTIZACION <> '" & ACCION_PCT_RETIRARSE & "') as TABLA"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if isNull(rs("Cantidad")) then
		Cantidad = 0
	else
		Cantidad = cint(rs("Cantidad"))
	end if
	numMin = getValorNorma("MINPRCP") 
	'Response.Write "<br>" & cantidad & " - " & numMin
	if cantidad => cint(numMin) then rtrn = true
	hayCotizacionesMinimas = rtrn
End Function
'---------------------------------------------------------------------------------------------
'Borra todas las variables del header
Function clearHeader()
	pct_idPedido = 0
	pct_cdPedido = ""
	pct_FechaInicio = ""
	pct_FechaCierre = ""
	pct_cdSolicitante = ""
	pct_idDivision = 0	
	pct_dsDivision = ""
	pct_idObra = 0
	pct_dsPedido = ""
	pct_tituloPedido = ""
	pct_idEstado = ESTADO_PCT_PENDIENTE
	pct_usuario = ""
	pct_cdUsuarioAdmin = ""
	pct_usuarioCarga = ""
	pct_momento = ""
	pct_observaciones = ""	
	pct_idProveedorElegido = 0
	pct_dsProveedorElegido = ""
    pct_Extension = TIPO_AFIRMACION
	pct_hayCabecera = false
End function
'---------------------------------------------------------------------------------------------
Function clearProveedor()
	pct_idProveedor = 0
	pct_dsProveedor = ""
	pct_pathCotizacion = ""	
End Function
'---------------------------------------------------------------------------------------------
function getComments(pId)
dim strSQL, rs, conn, rtrn
rtrn = ""
strSQL="Select * from TOEPFERDB.TBLPCTCOMENTARIOS where IDPEDIDO=" & pId
Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
while not rs.eof
		rtrn = rtrn & GF_TRADUCIR(rs("COMENTARIO")) & "<br>"
	rs.movenext
wend	
Call GF_BD_COMPRAS(rs, conn, "CLOSE", strSQL)
getComments = rtrn
end function
'---------------------------------------------------------------------------------------------
sub loadSectorEmpleado(pCdEmpleado, byref pIdSector, byref pDsSector)
dim strSQL, rs, conn
	if (pCdEmpleado <> "") then
		strSQL = "select * from WFProfesional where CDUSUARIO='" & pCdEmpleado & "'"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			pIdSector = rs("SectorKr")
			pDsSector = rs("Sector")
		else
			pIdSector = 0
			pDsSector = ""
		end if	
	end if
end sub
'----------------------------------------------------------------
'Suma los AFE que se aprobaron para el pedido.
Function calcularPresupuestoPedido(idMoneda, idPedido)
	Dim strSQL, rs, conn, importe
	
	importe = 0
	strSQL = "Select sum(IMPORTEPESOS) IMPORTEPESOS, sum(IMPORTEDOLARES) IMPORTEDOLARES from TBLDATOSAFE where CONFIRMADO='A' and IDPEDIDO=" & idPedido
	'response.write strSQL
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then 
		if (rs("IMPORTEPESOS") <> "") then
			importe = cdbl(rs("IMPORTEPESOS"))
			if (idMoneda = MONEDA_DOLAR) then importe = cdbl(rs("IMPORTEDOLARES"))
		end if
	end if
	if (not IsNumeric(importe)) then importe = 0	
	calcularPresupuestoPedido = importe	
End Function
'------------------------------------------------------------------
Function pctSaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function
'------------------------------------------------------------------
Function pctFile2Binary(idPedido, fileno, filePath)
  
  Dim rs, strSQL, extension

  Set fso = CreateObject("Scripting.FileSystemObject")
  extension = fso.GetExtensionName(filePath)
    
  strSQL = "Select IDPEDIDO, EXT, FILENO, FILEBIN, SIGNATURE from TBLPCTBINARYFILES Where 1=0"
  Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
  rs.AddNew
  rs("IDPEDIDO") = idPedido
  rs("EXT") = extension
  rs("FILENO") = fileno
  rs("FILEBIN") = readBinaryFile(filePath)
  if (fso.GetFile(filePath).Size < MAX_SIGNATURE_SIZE) then rs("SIGNATURE") = generateFileSignature(filePath)
  rs.Update
  
End Function
'------------------------------------------------------------------
Function buildFileName(idPedido, fileno, ext)
	Dim rs, strSQL, fileName, fileType
	
	strSQL = "Select CDPEDIDO from TBLPCTCABECERA where idPedido=" & idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then			
			if (fileno < 0) then 
				fileName = PREFIX_PCT
				fileType= "ET"				
				if (fileno = PCT_BINARY_CONDITIONS) then fileType = "CP"
			else
				fileName = PREFIX_CTZ
				fileType = fileno
			end if
			fileName = fileName & "-" & rs("CDPEDIDO") & "-" & fileType & "." & ext			
	end if
	buildFileName = fileName
End Function
'------------------------------------------------------------------
Function pctBinary2File(idPedido, fileno, filePath)
	Dim rs, strSQL, fileName, ret
		
	Set fso = CreateObject("Scripting.FileSystemObject")		
	ret = ""
	strSQL = "Select * from TBLPCTBINARYFILES where idPedido=" & idPedido & " and FILENO=" & fileno
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then								
			ret = filePath &"\"& buildFileName(idPedido, fileno, rs("EXT"))
			if (not fso.FileExists(ret)) then Call pctSaveBinaryData(ret, rs("FILEBIN"))								
	end if
	pctBinary2File = ret
End Function
'-------------------------------------------------------------------
'Verifica si tiene autoridad para cambiar cosas en el PCT
Function checkControlPCT()
	Dim ret		
	ret = false
	if ((not isAdmin(pct_idDivision)) and (not isAuditor(pct_idDivision)))then
		if (((pct_cdSolicitante = session("Usuario")) and (pct_idEstado <= ESTADO_PCT_APROBADO)) or _
			((pct_usuario = session("Usuario")) and (pct_idEstado < ESTADO_PCT_AUTORIZADO)) or _
			(pct_usuarioAdmin = session("Usuario"))) then 
			ret = true
		end if
	else
		ret = true
	end if
	checkControlPCT = ret
End Function
'------------------------------------------------------------------
Function findIdPedidoByCode(cdPedido)
	Dim rs, strSQL, ret
	
	strSQL="select IDPEDIDO from TBLPCTCABECERA where CDPEDIDO=" & cdPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	ret = 0
	if (not rs.eof) then ret = rs("IDPEDIDO")			
	findIdPedidoByCode = ret 
End Function
'-------------------------------------------------------------------------------------------------
Function obtenerImporteCotizacionElegida(p_idPedido, pMoneda,tipoCambio)
	Dim strSQL, conn, rs,rtrn,myTipoCambio,myImporte
	

	if (tipoCambio = "") then
		myTipoCambio = getTipoCambio(pMoneda, SEC_SYS_COMPRAS)
	else
		myTipoCambio = tipoCambio
	end if

	strSQL =          "SELECT pcpD.IMPORTE,pcpD.CDMONEDA "
	strSQL = strSQL & "FROM   tblpctcabecera pctCab "
	strSQL = strSQL & "       INNER JOIN TBLPCPDETALLE pcpD "
	strSQL = strSQL & "       ON     pcpD.idpedido    = pctCab.idpedido "
	strSQL = strSQL & "       AND    pcpD.idproveedor = pctCab.idproveedor "	
	strSQL = strSQL & "WHERE pctCab.IDPEDIDO = " & p_idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	rtrn = 0
	if (not rs.EoF) then 
		myImporte = CDbl(rs("IMPORTE"))
		if (pMoneda = rs("CDMONEDA")) then 
			rtrn = myImporte
		else
			if (pMoneda = MONEDA_DOLAR) then				
				rtrn = round(myImporte / (myTipoCambio),0)
			else
				rtrn = round(myImporte * (myTipoCambio) ,0)
			end if
		end if
	end if
	'Response.Write rtrn & "<br>"
	obtenerImporteCotizacionElegida = rtrn
	
End Function
'------------------------------------------------------------------------------------------
'Devuelve el total de los PIC registrados para el pedido y el proveedor indicado sin contar el PIC indicado.
'idProveedor es opcional (si se para 0 se ignora)
sub loadImporteAcumuladoPIC(idPedido, idCotizacion, idProveedor, pSoloAprobados, byref importePesos, byRef importeDolares)
dim rs, strSQL, ret
importePesos = 0
importeDolares = 0
strSQL = "select sum(IMPORTEPESOS) as ImportePesos, sum(IMPORTEDOLARES) as ImporteDolares from TBLCTZCABECERA where IDCOTIZACION<>" & idCotizacion & " AND IDPEDIDO=" & idPedido 
if (Cdbl(idProveedor) > 0) then strSQL = strSQL & " AND IDPROVEEDOR=" & idProveedor 
if (pSoloAprobados) then 
	'Se toman solo los gastos aprobados o ya facturados
	strSQL = strSQL & " and estado in ('" & CTZ_FIRMADA & "', '" & CTZ_FACTURADA & "')"
else
	'SE toman todos los gastos cargados y no anulados
	strSQL = strSQl & " AND estado <> '" & CTZ_ANULADA & "'"
end if
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not isNull(rs("ImportePesos")) then importePesos = CDbl(rs("ImportePesos"))
	if not isNull(rs("ImporteDolares")) then importeDolares = CDbl(rs("ImporteDolares"))
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", strSQL)
end sub
'-----------------------------------------------------------------------------------------------
'A partir de la planilla comparativa determina la moneda en la que se ingresa el gasto.
Function obtenerGanadorPlanilla(ByRef pMoneda, ByRef pImporte, byref pIdProveedor)
	
	Dim strSQL, rsPCP, rsPCT, rsCTZ, con
	
	pMoneda = MONEDA_PESO	'Se asume.				
	if ((pct_idPedido <> "") and (pct_idPedido <> 0)) then 
		'Obtengo de la planilla el importe y la moneda que determina el precio aceptado
		strSQL = "Select * from TBLPCPDETALLE where IDPEDIDO=" & pct_idPedido & " and IDPROVEEDOR=" & pct_idProveedorElegido
		Call executeQueryDb(DBSITE_SQL_INTRA, rsPCP, "OPEN", strSQL)
		if (not rsPCP.eof) then 
			'Tomo la moneda y el encuentro de acuerdo a lo indicado en la planilla.
			pMoneda = rsPCP("CDMONEDA")						
			pImporte = CDbl(rsPCP("IMPORTE"))			
			pIdProveedor = CDbl(rsPCP("IDPROVEEDOR"))			
		end if
	end if
End Function

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
' Autor: 	Ajaya Nahuel - CNA
' Fecha: 	20/12/2011
' Objetivo:	
'			Verifica si fue abierto alguna vez el archivo de un pedido, por medio de los campos agregados en TBLPCTCOTIZACIONES
' Parametros:
'			pIdPedido[int]
' Devuelve: 
' 			TRUE si no fue abierto o si el pedido no tiene cotizaciones
'			FALSE si ya fue abierto
Function verificarAperturaArchivo(pIdPedido)
	Dim strSQL, rtrn, rs 	
	rtrn = false
	strSQL = "Select FECHALECTURA from TBLPCTCOTIZACIONES where IDPEDIDO=" & pIdPedido & " and FECHALECTURA is not null"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	rtrn = rs.EoF
	verificarAperturaArchivo = rtrn	
End Function
'-------------------------------------------------------------------------------------------------------------'
' Nombre Funcion:		  cargaMensajeNDA(pIdPedido)
' Objetivo:
'			Obtener diferentes datos de un pedido, sea sus proveedores, fechas, usuarios; con el proposito de completar con estos datos 
' 			un mensaje que indique el alta de una nota de Aceptacion. 
' Parametros:
'			[int]	pIdPedido, pIdCotizacion
' Devuelve:
'			[recordset]
' Autor: ???
' Fecha: 20/03/2012
' Modificaciones: 	Ajaya Nahuel - CNA
'			12/11/2012 
 '--------------------------------------------------------------------------------------------------
Function cargaMensajeNDA(pIdPedido,pIdCotizacion)			 	
	Dim strSQL, rs, myWhere, flagCotizacion	
	flagCotizacion = false
	if(pIdCotizacion > 0)then flagCotizacion =true
	strSQL = " select ctz.idcotizacion,ctz.idpedido,ctz.idproveedor,prov.NOMEMP AS dsempresa,pct.cdusradmin,pct.titulo,pct.cdpedido,pct.tipocompra,pct_coti.fechapresentacion  " 
	strSQL = strSQL & " from tblctzcabecera ctz "
	if(flagCotizacion)then
		strSQL = strSQL & " left JOIN tblpctcabecera pct ON ctz.idpedido = pct.idpedido "
	else
		strSQL = strSQL & "	inner JOIN tblpctcabecera pct ON ctz.idpedido = pct.idpedido  and ctz.idproveedor  = pct.idproveedor "
	end if
	strSQL = strSQL & " left join (Select IDPedido, IDProveedor, Max(fechapresentacion) fechapresentacion "
    strSQL = strSQL & "            from TBLPCTCOTIZACIONES group by IDPedido, IDProveedor) pct_coti "
	strSQL = strSQL & " ON ctz.idPedido = pct_coti.idPedido and ctz.IdProveedor = pct_coti.IdProveedor "
	strSQL = strSQL & " Left JOIN [Database].[dbo].MET001A prov ON ctz.idproveedor = prov.nroemp "
	strSQL = strSQL & " where ctz.idpedido = " &  pIdPedido 
	if(flagCotizacion)then	strSQL = strSQL & "	and  ctz.idcotizacion = " & pIdCotizacion
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set cargaMensajeNDA = rs
End Function

'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 20/03/2012
'Nombre Funcion: 			getRsNotaAceptacion(IdNDA)
' Objetivo:
'			Lee los registros de una determinada NDA de la tabla TBLNOTACEPTACION 
' Parametros:
'			[int]	idNDA
' Devuelve:
'			[recordset]
' Modificaciones:
'			26/04/2012 - CNA
'--------------------------------------------------------------------------------------------------
Function getRsNotaAceptacion(pIdNDA)
	dim strSQL, rs
	strSQL = "Select * from tblnotaceptacion where IDNDA = "& pIdNDA	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getRsNotaAceptacion = rs
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 16/04/2013
' Nombre:	readPDC
' Objetivo:
'			Lee los registros de poliza de Caucion 
' Parametros:
'			[int]	 pIdPDC
'			[string] pCdPedido 
'			[string] pNroPoliza
'			[int]	 pIdAseguradora
'			[int]	 pTomador
'			[int]	 pMonto
'			[int]	 pIdEstado
'			[int]	 pIdDivision
'			[string] pOrder
'			[char]	 pMoneda
'			[string] pImporteAprox
'			[string] pFecha
' Devuelve:
'			[recordset]
'--------------------------------------------------------------------------------------------------
Function readPDC(pIdPDC, pCdPedido, pNroPoliza, pIdAseguradora, pTomador, pIdEstado, pIdDivision, pOrder, pImporte, pMoneda, pImporteAprox, pFecha, isVencida)
	dim strSQL, rs
	call buscarFiltrosPDC(myWhere, pNroPoliza, pIdAseguradora, pTomador, pIdEstado, pImporte, pMoneda, pImporteAprox, pFecha, isVencida)	
	strSQL = strSQL & "			SELECT POL.*,				"
	strSQL = strSQL & "				   PCT.CDPEDIDO,		" 
	strSQL = strSQL & "				   PCT.IDDIVISION,		" 
	strSQL = strSQL & "				   SEC.DSASEGURADORA,	" 
	strSQL = strSQL & "				   EMP.NOMEMP AS DSEMPRESA		" 	
	strSQL = strSQL & "			FROM (						" 
	strSQL = strSQL & "				   SELECT IDPDC,		" 
	strSQL = strSQL & "						  IDPEDIDO,	    " 
	strSQL = strSQL & "						  TIPO,			" 
	strSQL = strSQL & "						  NROPOLIZA,	" 
	strSQL = strSQL & "						  IDASEGURADORA," 
	strSQL = strSQL & "						  TOMADOR,		"	 
	strSQL = strSQL & "						  IMPORTE,		" 
	strSQL = strSQL & "						  CDMONEDA,		"	 
	strSQL = strSQL & "						  VENCIMIENTO,  " 
	strSQL = strSQL & "						  ESTADO,		" 
	strSQL = strSQL & "						  CDUSUARIO,	" 
	strSQL = strSQL & "						  MMTO			" 
	strSQL = strSQL & "				   FROM TBLPOLIZASCAUCION		 " 
	strSQL = strSQL &				   myWhere
	strSQL = strSQL & "				   ) AS POL				" 	
	strSQL = strSQL & "			INNER JOIN TBLPCTCABECERA PCT		 "
	strSQL = strSQL & "				ON PCT.IDPEDIDO = POL.IDPEDIDO			 "
	strSQL = strSQL & "			LEFT JOIN TBLPDCASEGURADORAS SEC	 "
	strSQL = strSQL & "				ON SEC.IDASEGURADORA = POL.IDASEGURADORA "
	strSQL = strSQL & "         LEFT JOIN [Database].[dbo].MET001A EMP			 "
	strSQL = strSQL & "				ON EMP.NROEMP = POL.TOMADOR			 "
	if ((pCdPedido <> "0") and (pCdPedido <> "")) then Call mkWhere(auxWhere, "PCT.CDPEDIDO", UCASE(pCdPedido), "like", 0)
	if (pIdDivision > 0) then Call mkWhere(auxWhere, "PCT.IDDIVISION", pIdDivision, "=", 1)
	strSQL = strSQL & auxWhere & pOrder	
    'Response.Write strSQL
    
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set readPDC = rs
End Function
'--------------------------------------------------------------------------------------------------
Function buscarFiltrosPDC(ByRef myWhere,pNroPoliza, pIdAseguradora, pTomador, pIdEstado, pImporte, pMoneda, pImporteAprox, pFecha, isVencida)
	Dim auxMoneda
	if (pIdPDC > 0) then Call mkWhere(myWhere, "IDPDC", pIdPDC, "=", 1)	
	if (pNroPoliza <> "") then Call mkWhere(myWhere, "NROPOLIZA", Trim(Ucase(pNroPoliza)), "LIKE", 3)
	if (pIdAseguradora > 0) then Call mkWhere(myWhere, "IDASEGURADORA", pIdAseguradora, "=", 1)
	if (pIdEstado = 0)then		
		if(isVencida)then 
			Call mkWhere(myWhere, "ESTADO", ESTADO_PDC_VENCIDA, "=", 1)
		else
			Call mkWhere(myWhere, "ESTADO", ESTADO_PDC_RECIBIDA, "<=", 1)	
		end if
	else
		if((not isVencida)and(pIdEstado = ESTADO_PDC_VENCIDA))then pIdEstado = 0
		if((isVencida)and(pIdEstado <> ESTADO_PDC_VENCIDA)and(pIdEstado <> 0))then pIdEstado = 0
		Call mkWhere(myWhere, "ESTADO", pIdEstado, "=", 1)
	end if
	if (pTomador > 0)then Call mkWhere(myWhere, "TOMADOR", pTomador, "=", 1)
	if ((pFecha <> "")and(pFecha <> 0))then Call mkWhere(myWhere, "VENCIMIENTO", pFecha, "=", 3)
	if (pMoneda <> "")then Call	mkWhere(myWhere, "CDMONEDA", pMoneda, "=", 3)			
	if (pImporte > 0) then
		if (isnumeric(pImporte)) then
			select case pImporteAprox
				case "Mayor"
					Call mkWhere(myWhere, "IMPORTE", pImporte * 100,">",1)
				case "Menor"
					Call mkWhere(myWhere, "IMPORTE", pImporte * 100,"<",1)
				case "Igual"
					Call mkWhere(myWhere, "IMPORTE", pImporte * 100,"=",1)
			end select
		end if
	end if			
	buscarFiltrosPDC = myWhere
End function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 17/04/2013
' Nombre Funcion: 			addPolizaCaucion(idPedido,Tipo, Monto, Moneda, Momento, Usuario, Estado)
' Objetivo:
'			Agrega un nuevo registro de Poliza de Caucion, y devuelve su ID
' Parametros:
'			[int]	  idPedido
'			[char]	  Tipo
'			[decimal] Monto
'			[char]	  Moneda
'			[string]  Momento
'			[string]  Usuario
'			[int]	  Estado					
' Devuelve:	[int]	  idPoliza nueva
'--------------------------------------------------------------------------------------------------
Function addPolizaCaucion(idPedido,Tipo, Monto, Moneda, Momento, Usuario, Estado)
	Dim strSQL, rs
	strSQL = " INSERT INTO TBLPOLIZASCAUCION (IDPEDIDO, TIPO, IMPORTE, CDMONEDA, ESTADO, CDUSUARIO, MMTO)" 
	strSQL = strSQL & " VALUES ("& idPedido &",'"& Tipo &"',"& Monto &",'"& Moneda &"',"& Estado &",'"& Usuario &"',"& Momento &")"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	strSQL = " SELECT MAX(IDPDC) AS ULTIMOPDC FROM TBLPOLIZASCAUCION "
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if(not rs.eof)then rtrn = Cdbl(rs("ULTIMOPDC"))
	addPolizaCaucion = rtrn
End function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 17/04/2013
' Nombre Funcion: 			updatePolizaCaucion(pIdPDC, pNroPoliza, pIdAseguradora, pTomador, pImporte, pVencimiento, pEstado)
' Objetivo:
'			Actualiza los nuevos datos de una Poliza de Caucion 
' Parametros:
'			[int]	  pIdPDC
'			[string]  pNroPoliza
'			[int]	  pIdAseguradora
'			[int]	  pTomador
'			[int]	  pImporte
'			[string]  pVencimiento
'			[int]	  pEstado					
' Devuelve:		-
'--------------------------------------------------------------------------------------------------
Function updatePolizaCaucion(pIdPDC, pNroPoliza, pIdAseguradora, pTomador, pImporte, pVencimiento, pEstado)
	Dim strSQL, rs
	strSQL = " UPDATE TBLPOLIZASCAUCION SET NROPOLIZA = '"& Trim(Ucase(pNroPoliza)) &"',IDASEGURADORA = "& pIdAseguradora &" ," 
	strSQL = strSQL & " TOMADOR = " & pTomador & ", IMPORTE = " & pImporte & ", VENCIMIENTO = " & pVencimiento & ", ESTADO = " & pEstado & ", MMTO = "& session("MmtoSistema")
	strSQL = strSQL & " WHERE IDPDC = " & pIdPDC	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 17/04/2013
' Nombre Funcion: 	updateSaldoPolizaCaucion(pIdPDC, pSaldo, pMoneda)
' Objetivo:
'			Actualiza el saldo de una Poliza de Caucion
' Parametros:
'			[int]	  pIdPDC
'			[int]	  pImporte
' Devuelve:		-
'--------------------------------------------------------------------------------------------------
Function updateSaldoPolizaCaucion(pIdPDC, pSaldo, pMoneda)
	Dim strSQL, rs
	strSQL = " UPDATE TBLPOLIZASCAUCION SET IMPORTE = " & pSaldo & ", CDMONEDA = '" & pMoneda & "' , MMTO = "& session("MmtoSistema")
	strSQL = strSQL & " WHERE IDPDC = " & pIdPDC
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End Function
'-------------------------------------------------------------------------------------------
Function generarSeccionServicioCompras(pIdProveedor, pIdPedido, pPermiso)
    Dim strMsg
    
    strMsg = "---------------NO MODIFIQUE POR DEBAJO DE ESTA LINEA---------------" & vbCrLf & vbCrLf
    strMsg = strMsg & "SV:"& SISTEMA_COMPRAS &"|PR:"& pIdProveedor &"|P:"& pIdPedido &"|F:" & pPermiso & vbCrLf
    strMsg = strMsg & Trim(MD5(generarCRCByPCT(pIdProveedor,pIdPedido,pPermiso))) & vbCrLf & vbCrLf
    strMsg = strMsg & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf  & vbCrLf  & vbCrLf
    
    generarSeccionServicioCompras = strMsg
    
End Function
'-------------------------------------------------------------------------------------------
Function generarCRCByPCT(pIdProveedor,pidPedido,pPermiso)
    generarCRCByPCT = SISTEMA_COMPRAS & pIdProveedor & SEPARATOR_CRC_PROVEEDOR & pidPedido & SEPARATOR_CRC_PEDIDO & pPermiso
End Function
'-------------------------------------------------------------------------------------------
%>