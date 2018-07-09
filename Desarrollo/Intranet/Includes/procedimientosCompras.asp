<!--#include file="procedimientosUnificador.asp"-->
<%

Const SISTEMA_COMPRAS = "COMPRAS" 

Const PREFIX_PCT = "PCT"
Const PREFIX_CTZ = "PIC"
Const PREFIX_OBR = "OBR"
Const PREFIX_PCP = "PCP"
Const PREFIX_AFE = "AFE"
Const PREFIX_CTC = "CTC"
Const PREFIX_FAC = "FAC"
Const PREFIX_NDB = "NDB"
Const PREFIX_NCR = "NCR"

Const CBTE_PROVEEDORES_FAC = "1"
Const CBTE_PROVEEDORES_NDB = "2"
Const CBTE_PROVEEDORES_NCR = "3"
Const CATEGORIA_MANTENIMIENTO = "S"
Const CATEGORIA_COMUN = "N"

Const TIPO_MONEDA_DOLAR = "US$"
Const TIPO_MONEDA_PESO = "$"

' CONSTANTES DE LA SECCION LISTAS DE CORREOS '
' A MEDIDA QUE SE VALLA CREANDO UN CODIGO DE LISTA DISTINTO SE DEBE CREAR LA CONSTANTE
Const LISTA_PCP_PROV_GANADOR = "LISTA-PCP"
  

Const PIC_TEXTO_DETALLE_PRESUPUESTO = "#PRESUPUESTO#"

Dim oConn, ccCompras, ccListaDivisionAdmin

'--------------------------------------------------------------------------------------
Function validarPasaporteCompras(pIdPedido, pIdTitular, pasaporte)
	Dim result, payLoad
	Dim fields		
	result = false		
	if (validarPasaporte(SISTEMA_COMPRAS, pIdTitular, pasaporte) = PASSPORT_INF_VALID) then		
		'La carga es solo el id del proveedor.
		if (retrievePayload(pasaporte, payLoad)) then
			if (CLng(payLoad("IDPEDIDO")) = CLng(pIdPedido)) then	result = true
		end if
	end if		
	validarPasaporteCompras = result
End Function
'--------------------------------------------------------------------------------------
'El principal de una empresa se obtiene pasando el id del usuario en 0
Function obtenerMail(idEmpresa)
	Dim result, rs, strSQL
	
	result = ""
	
	strSQL = "Select * from TBLMAILSCOMPRAS where IDEMPRESA= " & idEmpresa
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then result = rs("EMAIL")
	obtenerMail = result
	
End Function

'--------------------------------------------------------------------------------------
'Se controla si el usuario posee acceso al cargo.
Function comprasControlAccesoCM(pCargo)
	Call initAccessInfo(pCargo)
End Function
'--------------------------------------------------------------------------------------
'Permite chequear el acceso a la información segun el codigo de la misma.
Function checkPointAcceso(pIdDivision) 	
	checkPointAcceso = (InStr(ccListaDivisionAdmin, pIdDivision) > 0)		
End Function
'--------------------------------------------------------------------------------------
'Devuelve la lista de divisiones en las que tiene permiso de A, Y o U
Function getListaCargosAdmin()
dim rtrn, chrE, chrP, chrA, chrT
	chrE = ccValorListaDivision(1)
	chrA = ccValorListaDivision(2)
	chrP = ccValorListaDivision(3)
	chrT = ccValorListaDivision(4)
	if chrE = SEC_A or chrE = SEC_Y or chrE = SEC_U then rtrn = rtrn & "," & getDivisionID(CODIGO_EXPORTACION)
	if chrA = SEC_A or chrA = SEC_Y or chrA = SEC_U then rtrn = rtrn & "," & getDivisionID(CODIGO_ARROYO)
	if chrP = SEC_A or chrP = SEC_Y or chrP = SEC_U then rtrn = rtrn & "," & getDivisionID(CODIGO_PIEDRABUENA)
	if chrT = SEC_A or chrT = SEC_Y or chrT = SEC_U then rtrn = rtrn & "," & getDivisionID(CODIGO_TRANSITO)			
	if len(rtrn) > 0 then
		getListaCargosAdmin = right(rtrn,len(rtrn)-1)
	else
		getListaCargosAdmin = ""
	end if	
End Function
'--------------------------------------------------------------------------------------
function puedeCrear()
	if isInList(SIN_DIVISION, SEC_U) or isInList(SIN_DIVISION, SEC_A) then puedeCrear = true
end function
'---------------------------------------------------------------------------------------------------------------------
function isAdminInAny()
dim listaDeCargos, cargosSplitted, rtrn, index
rtrn = false
listaDeCargos = getListaCargosAdmin()
cargosSplitted = split(listaDeCargos,",")
for index=0 to ubound(cargosSplitted)
	if isAdmin(cargosSplitted(index)) then
		rtrn = true
		exit for
	end if	
next
isAdminInAny = rtrn
end function
'--------------------------------------------------------------------------------------
'Obtiene el nombre del archivo a generar.
'pOrigen: Path completo de origen, con nombre del archivo actual. (C:\origen\file.ext)
'pDestino: path del destino final del archivo, sin nombre de archivo (C:\destino\)
Function getFilename(prefix, pOrigen, pDestino)
	Dim hayError, filename, nbrName, fso, ext, base
	Dim seed
	
	Set fso = CreateObject("Scripting.FileSystemObject")			
	ext = fso.GetExtensionName(pOrigen)
	baseName = prefix & "-" & session.SessionId & "-"
	hayError = True
	nbrName = 0		
	while ((hayError) and (nbrName < 1000))		
		filename = baseName & nbrName & "." & ext
		if (not fso.FileExists(pDestino & filename)) Then hayError = False		
        nbrName = nbrName + 1		
    wend
    getFilename = filename
	
	Set fso = nothing
	
End Function
'--------------------------------------------------------------------------------------
'Arma los path para los diferentes archivosd de un pedido de cotizacion.
Function setPaths(id, file, ByRef path, ByRef pathWeb)
	Dim cPath, webApp
	
	if (file <> "") then		
		cPath = PATH_COMPRAS_FINAL
		if (id = 0) then cPath = PATH_COMPRAS_TEMP
		path = server.mappath(".") & "\" & cPath & "\" & file				
		webApp = Request.ServerVariables("SCRIPT_NAME")		
		webApp = Left(webApp, InStrRev(webApp, "/"))		
		pathWeb = "http://" & Request.ServerVariables("SERVER_NAME") & webApp & Replace(cPath, "\", "/") & "/" & file
	else
		path = ""
		pathWeb = ""
	end if	
	
End Function
'---------------------------------------------------------------------------------------------
'Controla los datos de un proveedor.
Function controlarProveedor(pIdProveedor)	
	Dim desc, strSQL, conn, rs, cuitEmpresa
	
	controlarProveedor = false
	'Se valida que el proveedor existe y que esta habilitado para operar.
	desc = getDescripcionProveedor(pIdProveedor)
	if (desc <> "") then
		'cuitEmpresa = getCUITProveedor(pIdProveedor)
		'strSQL = "Select * from TBLESTADOEMPRESAS where CUIT=" & cuitEmpresa
        'Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		'if (rs.eof) then
			controlarProveedor = true
		'else
		'	setError(PROVEEDOR_NO_RECOMENDADO)
		'end if
	else
		setError(PROVEEDOR_NO_EXISTE)
	end if
End Function
'--------------------------------------------------------------------------------------------------------
Function getUnidadArticulo(idArticulo, idUnidad, cdUnidad, dsUnidad)
	Dim strSQL, conn, rs
	
	idUnidad = 0
	cdUnidad = ""
	dsUnidad = ""	
	strSQL="Select A.IDUNIDAD IDUNIDAD, B.CDUNIDAD CDUNIDAD, B.DSUNIDAD DSUNIDAD from  TBLARTICULOS A inner join TBLUNIDADES B on A.IDUNIDAD=B.IDUNIDAD where A.IDARTICULO=" & idArticulo
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then
		idUnidad = rs("IDUNIDAD")
		cdUnidad = rs("CDUNIDAD")
		dsUnidad = rs("DSUNIDAD")
	end if
End Function
'--------------------------------------------------------------------------------------------------------
' Función:	readBinaryFile
'-----
'se desconoce el autor y la fecha de creación de la funcion
'esta se encontraba en PROCEDIMIENTOS PCT y fue transferida para 
'poder tener un uso general, ya que sera usada tambien por los CTC (Contratos)
' Modifico:	JPS - Santi Juan Pablo
' Fecha: 	08/02/11
'-----
' Objetivo:	Convertir un archivo a binario para se guardado en la base de datos.
' Parametros:
'			FileName 	[path] 	Path donde esta alojado el archivo
' Devuelve:	valor binario del archivo
'--------------------------------------------------------------------------------------------------------
Function readBinaryFile(FileName)
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To get binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream
  BinaryStream.Open
  
  'Load the file data from disk To stream object
  BinaryStream.LoadFromFile FileName

  'Open the stream And get binary data from the object
  readBinaryFile = BinaryStream.Read
End Function
'------------------------------------------------------------------

%>