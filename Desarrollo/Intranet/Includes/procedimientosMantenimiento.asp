<%
'Variable globales para la cabecera de una OT
dim SM_idOrder
dim SM_nroOrder
dim SM_dsOrder
dim SM_idDivision
dim SM_idActiveEquipment
dim SM_scheduledDate
dim SM_startDate
dim SM_finishedDate
dim SM_date
dim SM_maintenanceType
dim SM_cdState
dim SM_cdApplicant
dim SM_dsApplicant
dim SM_cdExecutedBy
dim SM_dsExecutedBy
dim SM_orderType
dim SM_idResponsableCompany
dim SM_dsResponsableCompany
dim SM_cdUser
dim SM_moment
dim SM_observations
dim SM_idObra
dim SM_idBudgetArea
dim SM_idBudgetDetalle

dim SM_cdActivation
dim SM_dsActivation
dim SM_dsSector
dim SM_activeCode

'Variable globales para las Tasks de una OT
dim SM_nroTask
dim SM_dsTask
dim SM_doneTask
dim SM_taskQuantity
dim SM_ActualTask
dim SM_taskQuantityDB

'Variable globales para los repuestos de una OT
dim SM_idItem
dim SM_dsItem
dim SM_nroItem
dim SM_programQuantityItem
dim SM_realQuantityItem
dim SM_idPMItem
dim SM_itemQuantity
dim SM_actualItem
dim SM_itemQuantityDB

dim SM_OTFrequencyUnit
dim SM_OTFrequencyQuantity

Const SM_TASK_DONE_YES = "S"
Const SM_TASK_DONE_NO = "N"
'Recordset para tareas
dim rsTasks
'Recordset para repuestos
dim rsItems

Const PREFIX_ODT = "ODT"
'Constantes de Estados
Const STATE_ALLS = 0
Const STATE_STAND_BY = 1
Const STATE_STARTED = 2
Const STATE_FINISHED = 3
Const STATE_CANCELED = 9

'Constantes de Tipo de Mantenimiento
Const MAIN_TYPE_ALLS = "T"
Const MAIN_TYPE_PREVENTIVE = "P"
Const MAIN_TYPE_PREDICTIVE = "D"
Const MAIN_TYPE_CORRECTIVE = "C"
'Constantes de Tipo de ORDEN
Const ORDER_TYPE_ALLS = "T"
Const ORDER_TYPE_MECHANICAL = "M"
Const ORDER_TYPE_ELECRONIC = "E"
Const ORDER_TYPE_CIVIL = "C"
Const ORDER_TYPE_SECURITY = "S"
Const ORDER_TYPE_OPERATIVE = "O"
Const ORDER_TYPE_SYSTEM = "Y"
'Constantes de Frecuencia
Const ORDER_FREQ_UNIQUE = "U"
Const ORDER_FREQ_DAY = "d"
Const ORDER_FREQ_WEEK = "w"
Const ORDER_FREQ_MONTH = "m"
Const ORDER_FREQ_YEAR = "y"

Const ORDER_FREQ_ENABLED = "1"
Const ORDER_FREQ_DISABLED = "2"

Function addParam(p_strKey,p_strValue,ByRef p_strParam)
    if (not isEmpty(p_strValue)) then
       if (isEmpty(p_strParam)) then
          p_strParam = "?"
       else
          p_strParam = p_strParam & "&"
       end if
       p_strParam = p_strParam & p_strKey & "=" & p_strValue
    end if
End Function
'------------------------------------------------------------------------------------------------------------------------
function checkDatosActivacion(cdActivacion, dsActivacion)
dim rtrn
rtrn = true
if len(cdActivacion)<3 then setError(SM_CODIGO_EQUIPO_INCORRECTO)
if len(dsActivacion) < 1 then setError(SM_DESCRIPCION_EQUIPO_INCORRECTO)
if (hayError) then rtrn = false
checkDatosActivacion = rtrn
end function
'------------------------------------------------------------------------------------------------------------------------
sub activarEquipo(idEquipo, idDivision, idSector, idUbicacion, cdActivacion, dsActivacion, cdActivoFijo)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_INS", idEquipo & "||" & idDivision & "||" & idSector & "||" & cdActivacion & "||" & dsActivacion & "||" & cdActivoFijo & "||" & ESTADO_ACTIVO & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
sub modificarEquipoActivo(idEquipoActivo, idDivision, idSector, idUbicacion, cdActivacion, dsActivacion, cdActivoFijo)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_UPD", idEquipoActivo & "||" & idDivision & "||" & idSector & "||" & cdActivacion & "||" & dsActivacion & "||" & cdActivoFijo & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
sub desactivarEquipo(idEquipoActivo, pState)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_UPD_STATE_BY_ID", idEquipoActivo & "||" & pState & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
function existeEquipoActivo(idEquipoActivo, idDivision, cdActivoFijo, cdActivacion)
Dim rs, rtrn
rtrn = false
call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_COUNT_BY_CDACTIVATION", idEquipoActivo & "||" & idDivision & "||" & cdActivacion)
if rs("QUANTITY")<> 0 then 
	rtrn = true
else
	if len(cdActivoFijo)>0 then
		call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_COUNT_BY_CDACTIVECODE", idEquipoActivo & "||" & cdActivoFijo)
		if rs("QUANTITY")<> 0 then rtrn = true	
	end if	
end if	
existeEquipoActivo = rtrn	
end function
'------------------------------------------------------------------------------------------------------------------------
function tieneOTActiva(idEquipoActivo)
dim rtrn
rtrn = false
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_GET_BY_PARAMETERS", "0|| || || ||0||" & idEquipoActivo & "||0||" & ORDER_TYPE_ALLS & "||" & MAIN_TYPE_ALLS & "||0||0||0||1||||1||0")
	if not rs.eof then
		if rs("CDSTATE") = STATE_STAND_BY or rs("CDSTATE") = STATE_STARTED then rtrn = true
	end if
tieneOTActiva = rtrn
end function
'------------------------------------------------------------------------------------------------------------------------
sub agregarComponente(idEquipo, idEquipoActivo, txtComponente, idGrupo)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMCOMPONENT_INS", idEquipo & "||" & idEquipoActivo & "||" & txtComponente & "||" & idGrupo & "||" & ESTADO_ACTIVO & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
sub modificarComponente(idComponente, dsComponente)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMCOMPONENT_UPD", idComponente & "||" & dsComponente & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
sub adminComponente(idComponente, idGrupo, pState)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMCOMPONENT_UPD_STATE_BY_ID", idComponente & "||" & idGrupo & "||" & pState & "||" & session("Usuario") & "||" & session("MmtoDato"))
end sub
'------------------------------------------------------------------------------------------------------------------------
function existeComponente(idEquipo, idEquipoActivo, idComponente, txtComponente, idGrupo)
dim rsComp, rtrn
	rtrn = false
	call executeProcedureDb(DBSITE_SQL_INTRA, rsComp, "TBLSMCOMPONENT_GET_COUNT_BY_DS", idComponente & "||" & idEquipo & "||" & idEquipoActivo & "||" & idGrupo & "||" & txtComponente)
	if rsComp("QUANTITY") <> 0 then rtrn = true
	existeComponente = rtrn
end function
'--------------------------------------------------------------------------------------------------
Function getFiles(pId, pType)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMBINARYFILE_GET_BY_ID", pId & "||" & pType)
	Set getFiles = rs
End Function
'--------------------------------------------------------------------------------------------------
Function smFile2Binary(pId, pType, filePath)
	Dim rs, strSQL,fileName
	fileno = 0
	if (clng(pId) <> 0) then
		fileno = getCantFileEquipo(pId, pType)
	end if	
	Set fso = CreateObject("Scripting.FileSystemObject")
	extension = fso.GetExtensionName(server.MapPath(filePath))
    fileName = fso.getfilename(server.MapPath(filePath))
	fileName = left(fileName,InStrRev(filename,".")-1) 'le quito la extension
	Call FileName2DbName(filename)

	strSQL = "SELECT ID, TYPE, FILENO, NAME, EXT, BINARYFILE FROM TBLSMBINARYFILE where 1=0"
	Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)  
	rs.AddNew
	rs("ID") = pId
	rs("TYPE") = pType 
	rs("FILENO") = fileno
	rs("NAME") = fileName
	rs("EXT") = extension
	rs("BINARYFILE") = readBinaryFile(server.MapPath(filePath))	
	rs.Update
End Function
'--------------------------------------------------------------------------------------------------
Function getCantFileEquipo(pIdEquipoDefault, pType)
	Dim strSQL,rs,rtrn
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMBINARYFILE_GET_NEXT_FILENO", pIdEquipoDefault & "||" & pType)
	rtrn = cdbl(rs("MAXIMO"))
	getCantFileEquipo = rtrn
End Function
'--------------------------------------------------------------------------------------------------
sub deleteFile(pIdEquipoDefault, pType, pFileNo)
	Dim strSQL,rs,rtrn
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMBINARYFILE_DEL", pIdEquipoDefault & "||" & pType & "||" & pFileNo)
end sub
'--------------------------------------------------------------------------------------------------
function getImageByExt(pExtension)
dim myImage
	select case UCASE(pExtension)
		case "DOC", "DOCX"
			myImage = "doc.gif"
		case "XLS", "XLSX"
			myImage = "excel.gif"
		case "PDF"
			myImage = "pdf.gif"
		case "GIF","JPG","PNG","JPEG", "BMP", "TIF"
			myImage = "document_image-16x16.png"
		case "TXT"
			myImage = "Bloc_Notes-16x16.png"
		case "ZIP", "RAR"
			myImage = "Winrar-16x16.png"
		case else
			myImage = "1p.gif"
	end select		
	getImageByExt = "<image src='images/" & myImage & "'>"
end function


'--------------------------------------------------------------------------------------------------
Function getOTFiles(pIdOT)
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTBINARYFILE_GET_BY_ID", pIdOT)
	Set getOTFiles = rs
End Function
'--------------------------------------------------------------------------------------------------
Function OTFile2Binary(pIdOT, filePath)
	Dim rs, strSQL,fileName
	fileno = 0
	if (clng(pIdOT) <> 0) then
		fileno = getCantOTFiles(pIdOT)
	end if	
	Set fso = CreateObject("Scripting.FileSystemObject")
	extension = fso.GetExtensionName(server.MapPath(filePath))
    fileName = fso.getfilename(server.MapPath(filePath))
	fileName = left(fileName,InStrRev(filename,".")-1) 'le quito la extension
	Call FileName2DbName(filename)

	strSQL = "SELECT ID, FILENO, NAME, EXT, BINARYFILE FROM TBLSMOTBINARYFILE where 1=0"
	Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)  
	rs.AddNew
	rs("ID") = pIdOT
	rs("FILENO") = fileno
	rs("NAME") = fileName
	rs("EXT") = extension
	rs("BINARYFILE") = readBinaryFile(server.MapPath(filePath))	
	rs.Update
End Function
'--------------------------------------------------------------------------------------------------
Function getCantOTFiles(pIdOT)
	Dim strSQL,rs,rtrn
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTBINARYFILE_GET_NEXT_FILENO", pIdOT)
	rtrn = cdbl(rs("MAXIMO"))
	getCantOTFiles = rtrn
End Function
'--------------------------------------------------------------------------------------------------
sub deleteOTFile(pIdOT, pFileNo)
	Dim strSQL,rs,rtrn
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTBINARYFILE_DEL", pIdOT & "||" & pFileNo)
end sub


'-----------------------------------------------------------------------------------------------------------------------------
Function controlarEquipo(idEquipo, cdEquipo, dsEquipo)
	Dim strSQL, rs, conn
	if len(cdEquipo)=0 then
		call setError(SM_CODIGO_EQUIPO_INCORRECTO)
	else
		if (existeCodigoEquipo(idEquipo,cdEquipo)) then call setError(SM_CODIGO_EQUIPO_EXISTENTE)		
	end if	
	if len(dsEquipo)=0 then call setError(SM_DESCRIPCION_EQUIPO_INCORRECTO)
	if not hayError then controlarEquipo = RESPUESTA_OK
End Function
'-----------------------------------------------------------------------------------------
Function grabarEquipo(idEquipo, cdEquipo, dsEquipo, cdEstado)
	Dim strSQL, rs, conn
	if (idEquipo = 0) then
		call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_INS", ucase(cdEquipo) & "||" & dsEquipo & "||" & cdEstado)
	else
		call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_UPD", idEquipo & "||" & ucase(cdEquipo) & "||" & dsEquipo & "||" & cdEstado)
	end if
End Function
'-----------------------------------------------------------------------------------------
function tieneActivaciones(pIdEquipo)
Dim strSQL, rs, conn, rtrn
rtrn = false
call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_PARAMETERS", "0||" & pIdEquipo & "|| ||0|| || || ||0||")
if not rs.eof then rtrn = RESPUESTA_OK
tieneActivaciones = rtrn
end function
'-----------------------------------------------------------------------------------------
function existeCodigoEquipo(pIdEquipo, pCdEquipo)
Dim strSQL, rs, conn, rtrn
rtrn = false
call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_GET_COUNT_BY_CD", pIdEquipo & "||" & ucase(pCdEquipo))
if rs("QUANTITY")<> 0 then rtrn = RESPUESTA_OK
existeCodigoEquipo = rtrn
end function
'------------------------------------------------------------------
sub clearHeader()
SM_nroOrder = ""
SM_dsOrder = ""
SM_idDivision = 0
SM_idActiveEquipment = 0
SM_scheduledDate = ""
SM_startDate = ""
SM_finishedDate = ""
SM_date = ""
SM_maintenanceType = 0
SM_cdState = 0
SM_cdApplicant = 0
SM_dsApplicant = ""
SM_cdExecutedBy = 0
SM_dsExecutedBy = ""
SM_orderType = 0
SM_idResponsableCompany = 0
SM_dsResponsableCompany = ""
SM_cdUser = ""
SM_moment = 0
SM_observations = ""
SM_idObra = 0
SM_idBudgetArea = 0
SM_idBudgetDetalle = 0
SM_OTFrequencyUnit = "U"
SM_OTFrequencyQuantity = 0
end sub
'------------------------------------------------------------------
function readHeaderOT(pIdOT)
dim strSQL, ret
Call clearHeader()
ret = false
SM_idOrder = pIdOT
if (isFormSubmit()) then
	ret = readHeaderOtParams()
else 
	if (SM_idOrder > 0) then
		ret = readHeaderOtDB()
	end if
end if		
readHeaderOT = ret
end function
'------------------------------------------------------------------
function readHeaderOtDB()
dim rtrn
rtrn = false
call executeProcedureDb(DBSITE_SQL_INTRA, rsHeaderOT, "TBLSMORDER_GET_BY_ID", SM_idOrder)
if not rsHeaderOT.eof then
	'SM_idOrder = pIdOrder
	SM_nroOrder = rsHeaderOT("nroOrder")
	SM_dsOrder = rsHeaderOT("dsOrder")
	SM_idDivision = rsHeaderOT("IDDIVISION")
	SM_idActiveEquipment = rsHeaderOT("IDACTIVEEQUIPMENT")
	SM_scheduledDate = GF_FN2DTE(rsHeaderOT("SCHEDULEDDATE"))
	SM_startDate = GF_FN2DTE(rsHeaderOT("STARTDATE"))
	SM_finishedDate = GF_FN2DTE(rsHeaderOT("FINISHEDDATE"))
	SM_maintenanceType = rsHeaderOT("MAINTENANCETYPE")
	SM_cdState = rsHeaderOT("CDSTATE")
	SM_cdApplicant = rsHeaderOT("CDAPPLICANT")
	SM_dsApplicant = getUserDescription(SM_cdApplicant)
	SM_cdExecutedBy = rsHeaderOT("CDEXECUTEDBY")
	SM_dsExecutedBy = getUserDescription(SM_cdExecutedBy)
	SM_orderType = rsHeaderOT("ORDERTYPE")
	SM_idResponsableCompany = rsHeaderOT("IDRESPONSABLECOMPANY")
	SM_dsResponsableCompany = rsHeaderOT("DSEMPRESA")
	SM_cdUser = rsHeaderOT("CDUSER")
	SM_moment = rsHeaderOT("MOMENT")
	SM_observations = rsHeaderOT("OBSERVATIONS")
	SM_idObra = rsHeaderOT("IDOBRA")
	SM_idBudgetArea = rsHeaderOT("IDBUDGETAREA")
	SM_idBudgetDetalle = rsHeaderOT("IDBUDGETDETALLE")
	SM_OTFrequencyUnit = rsHeaderOT("UNIT")
    'A partir de la primera ocurrencia para una OT planificada (repetitivo) su valor vendr� null, debido a que la unidad y la cantidad estar�n asignadas a la OT maestra
    if (isNull(rsHeaderOT("UNIT"))) then SM_OTFrequencyUnit = ORDER_FREQ_UNIQUE
	SM_OTFrequencyQuantity = rsHeaderOT("QUANTITY")
	rtrn = true
end if
readOTDb = rtrn
end function
'------------------------------------------------------------------
function readHeaderOtParams()
	SM_nroOrder = GF_PARAMETROS7("SM_nroOrder","",6)
	SM_dsOrder = GF_PARAMETROS7("SM_dsOrder","",6)
	SM_idDivision = GF_PARAMETROS7("SM_idDivision",0,6)
	SM_idActiveEquipment = GF_PARAMETROS7("SM_idActiveEquipment",0,6)
	SM_scheduledDate = GF_PARAMETROS7("SM_scheduledDate","",6)
	SM_startDate =  GF_PARAMETROS7("SM_startDate","",6)
	SM_finishedDate = GF_PARAMETROS7("SM_finishedDate","",6)
	SM_maintenanceType = GF_PARAMETROS7("SM_maintenanceType","",6)
	if(SM_maintenanceType = "") then SM_maintenanceType = "T"
	SM_cdState = GF_PARAMETROS7("SM_cdState","",6)
	SM_cdApplicant = GF_PARAMETROS7("SM_cdApplicant","",6)
	if(SM_cdApplicant <> "")then SM_dsApplicant = GF_PARAMETROS7("SM_dsApplicant","",6)
	'SM_cdExecutedBy = GF_PARAMETROS7("SM_cdExecutedBy","",6)
	'if(SM_cdExecutedBy <> "")then SM_dsExecutedBy = GF_PARAMETROS7("SM_dsExecutedBy","",6)
	SM_orderType = GF_PARAMETROS7("SM_orderType","",6)
	if(SM_orderType = "") then SM_orderType = "T"
	SM_idResponsableCompany = GF_PARAMETROS7("SM_idResponsableCompany","",6)
	if(SM_idResponsableCompany <> "")then SM_dsResponsableCompany = GF_PARAMETROS7("SM_dsResponsableCompany","",6)
	SM_cdUser = GF_PARAMETROS7("SM_cdUser","",6)
	SM_moment = GF_PARAMETROS7("SM_moment",0,6)
	SM_observations = GF_PARAMETROS7("SM_observations","",6)
	SM_idObra = GF_PARAMETROS7("idObra",0,6)
	SM_idBudgetArea = GF_PARAMETROS7("idBudgetArea",0,6)
	SM_idBudgetDetalle = GF_PARAMETROS7("idBudgetDetalle",0,6)
	SM_OTFrequencyUnit = GF_PARAMETROS7("SM_OTFrequencyUnit","",6)
	SM_OTFrequencyQuantity = GF_PARAMETROS7("SM_OTFrequencyQuantity",0,6)
end function

'------------------------------------------------------------------
sub saveOT()
dim strSQL, rs, myNroOrder
aux_SM_scheduledDate = SM_scheduledDate
if isdate(aux_SM_scheduledDate) then aux_SM_scheduledDate = GF_DTE2FN(aux_SM_scheduledDate)
if SM_idOrder > 0 then
	rtrn = SM_idOrder
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_UPD", SM_idOrder & "||" & SM_dsOrder & "||" & SM_idDivision & "||" & SM_idActiveEquipment & "||" & aux_SM_scheduledDate & "||" & SM_maintenanceType & "||" & SM_cdState & "||" & SM_cdApplicant & "||" & SM_orderType & "||" & SM_idResponsableCompany & "||" & SM_idObra & "||" & SM_idBudgetArea & "||" & SM_idBudgetDetalle & "||" & session("Usuario") & "||" & session("MmtoDato") & "||" )
else
	rtrn = 0
	SM_nroOrder = getNumeracionOT(SM_idActiveEquipment)
	SM_cdState = STATE_STAND_BY 
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_INS", SM_nroOrder & "||" & SM_dsOrder & "||" & SM_idDivision & "||" & SM_idActiveEquipment & "||" & aux_SM_scheduledDate & "||" & SM_maintenanceType & "||" & SM_cdState & "||" & SM_cdApplicant & "||" & SM_orderType & "||" & SM_idResponsableCompany & "||" & SM_idObra & "||" & SM_idBudgetArea & "||" & SM_idBudgetDetalle & "||" & session("Usuario") & "||" & session("MmtoDato") & "||" )
	call executeProcedureDb(DBSITE_SQL_INTRA, rsMax, "TBLSMORDER_GET_MAX_ID","" )
	SM_idOrder = rsMax("MAXID")
end if
'Guardar Periodicidad
if SM_OTFrequencyUnit <> ORDER_FREQ_UNIQUE then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTFREQUENCY_INS", SM_idOrder & "||" & SM_OTFrequencyUnit & "||" & SM_OTFrequencyQuantity & "||" & ORDER_FREQ_ENABLED & "||" & session("Usuario") & "||" & session("MmtoDato") & "||")
end if

'Guardar Detalles
call initTasksOT()
while readNextTaskOt()
	SM_doneTask = SM_TASK_DONE_NO 
	saveOtTasks()
wend
'Guardar Repuestos
Call initItemsOT()
while readNextItemOt()
    SM_idPMItem = 0
    SM_realQuantityItem = 0
    saveOtItems()
wend

end sub
'----------------------------------------------------------------------------
sub updateOtStatus()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_UPD_STS", SM_idOrder & "||" & SM_cdState & "||" & SM_date & "||" & editText4DB(SM_observations) & "||" & SM_cdExecutedBy & "||" & session("Usuario") & "||" & session("MmtoDato"))
	if SM_cdState = STATE_STARTED then
        'Verifico si se necesita o no agregar el registro proximo de la planificacion
        if SM_OTFrequencyUnit <> ORDER_FREQ_UNIQUE and SM_OTFrequencyUnit <> "" then
            'Actualiza la fecha de inicio de la primera OT generada
            call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTEXECUTIONS_INS", SM_idOrder & "||" & SM_date & "||" & SM_idOrder & "||N")
            'Insertar la proxima ocurrencia (fecha planificada) de la primera repeticion de la planificacion de la OT Maestra
            call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTEXECUTIONS_INS", SM_idOrder & "||" & GF_DTE2FN(getNextExecution(GF_STANDARIZAR_FECHA_RTRN(GF_STANDARIZAR_FECHA_RTRN(date())),SM_OTFrequencyUnit, SM_OTFrequencyQuantity)) & "||0||N")
		else
			'Ver si corresponde a una planificacion
			call executeProcedureDb(DBSITE_SQL_INTRA, rsNext, "TBLSMOTEXECUTIONS_GET_BY_IDOTGENERATED", SM_idOrder)
			if not rsNext.eof then
				call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTEXECUTIONS_INS", rsNext("IDOT") & "||" & GF_DTE2FN(getNextExecution(GF_FN2DTE(SM_date), rsNext("UNIT"), rsNext("QUANTITY"))) & "||0||N")	
			end if	
		end if
	end if
end if
end sub
'----------------------------------------------------------------------------
sub updateSceduledOtStatus()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMOTFREQUENCY_UPD_STS", SM_idOrder & "||" & SM_cdState & "||" & session("Usuario") & "||" & session("MmtoDato"))
end if
end sub
'----------------------------------------------------------------------------
function getNextExecution(pLastExec, pUnit, pQuantity)
dim rtrn, myInterval
	if not isnull(pUnit) and not isnull(pQuantity) then
		'Es la ot maestra
		select case pUnit
			case ORDER_FREQ_DAY 
				myInterval = "d"
			case ORDER_FREQ_WEEK 
				myInterval = "ww"	
			case ORDER_FREQ_MONTH 
				myInterval = "m"	
			case ORDER_FREQ_YEAR 
				myInterval = "yyyy"	
			case else
				myInterval = "ERROR"
		end select		
		rtrn = GF_STANDARIZAR_FECHA_RTRN(DateAdd(myInterval, pQuantity, pLastExec))
	end if	
getNextExecution = rtrn
end function
'----------------------------------------------------------------------------
sub saveOtTasks()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDERTASK_INS", SM_idOrder & "||" & SM_nroTask & "||" & SM_dsTask & "||" & SM_doneTask)
end if
end sub
'----------------------------------------------------------------------------
sub deleteOtTask()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDERTASK_DEL", SM_idOrder & "||" & SM_nroTask)
end if
end sub
'----------------------------------------------------------------------------
sub udpateOtItemsPM()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDERITEM_UPD_PM", SM_idOrder & "||" & SM_idPMItem)
end if
end sub
'----------------------------------------------------------------------------
sub saveOtItems()
if SM_idOrder > 0 then
	if SM_programQuantityItem = "" then SM_programQuantityItem = 0
	if SM_realQuantityItem = "" then SM_realQuantityItem = 0
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDERITEM_INS", SM_idOrder & "||" & SM_nroItem & "||" & SM_idItem & "||" & SM_programQuantityItem & "||" & SM_realQuantityItem & "||" & SM_idPMItem)
end if
end sub
'----------------------------------------------------------------------------
sub deleteOtItem()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDERITEM_DEL", SM_idOrder & "||" & SM_nroItem)
end if
end sub
'----------------------------------------------------------------------------
sub cancelPmOt()
if SM_idOrder > 0 then
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMORDER_UPD_STS", SM_idOrder & "||" & SM_cdState & "||" & SM_date & "||" & editText4DB(SM_observations) & "||" & SM_cdExecutedBy & "||" & session("Usuario") & "||" & session("MmtoDato"))
end if
end sub
'----------------------------------------------------------------------------
Function getNumeracionOT(idActiveEquipment)
	Dim strSQL, oConn, rs,rtrn, nr, clave, myDivision
	strSQL = "SELECT DIV.CDDIVISION FROM TBLSMACTIVEEQUIPMENT SAE INNER JOIN TBLDIVISIONES DIV " & _ 
			 "ON SAE.IDDIVISION=DIV.IDDIVISION WHERE SAE.IDACTIVEEQUIPMENT=" & idActiveEquipment
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not rs.eof then 
		cdDivision = rs("CDDIVISION")
		clave= PREFIX_ODT & "_" & cdDivision & "_" & Right(year(now()),2)
		strSQL="Select * from TBLNUMERACION where PREFIJO='" & PREFIX_ODT & "' and CLAVE='" & clave & "'"
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			nr = CLng(rs("VALOR")) + 1
			strsql = "Update TBLNUMERACION set VALOR=" & nr & " where CLAVE = '" & clave & "' and PREFIJO='" & PREFIX_ODT & "'"
		else
			nr=1
			strSQL = "Insert into TBLNUMERACION values('" & PREFIX_ODT & "','" & clave & "', 1)"
		end if	
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		getNumeracionOT = Right(year(now()),2) & "-" & cdDivision & "-" & GF_nDigits(nr, 6)	
	else
		getNumeracionOT = "ERROR"
	end if	
End Function
'------------------------------------------------------------------
function checkItemsOT()
dim rtrn, myAux
rtrn = true
'Guardar Repuestos
call initItemsOT()
do while readNextItemOt()
	if instr(myAux, "," & SM_idItem & ",") > 0 then 
		setError(SM_ART_REPETIDOS)
		exit do
	end if	
	myAux = myAux & "," & SM_idItem & ","
	
loop
if (hayError) then rtrn = false
checkItemsOT = rtrn
end function
'------------------------------------------------------------------
function checkHeaderOT()
dim rtrn
rtrn = true
if len(SM_dsOrder) < 1 then setError(SM_OT_FALTA_TITULO)
if SM_idActiveEquipment = 0 then setError(SM_OT_FALTA_EQUIPO)
if SM_maintenanceType = MAIN_TYPE_ALLS then setError(SM_OT_FALTA_TIPO_MANT)
if SM_orderType = ORDER_TYPE_ALLS then setError(SM_OT_FALTA_TIPO_ORDEN)
if SM_cdApplicant = "" then setError(SM_OT_FALTA_SOLICITANTE)
if SM_idObra = 0 then setError(SM_OT_FALTA_OBRA)
if SM_idBudgetArea = 0 then setError(SM_OT_FALTA_BUDGET)
if SM_idBudgetDetalle = 0 then setError(SM_OT_FALTA_BUDGET)
if SM_scheduledDate = "" then 
	setError(SM_OT_FALTA_FECHA_PROG)
else	
	if isdate(SM_scheduledDate) then SM_scheduledDate = GF_DTE2FN(SM_scheduledDate)
	if SM_scheduledDate < GF_DTE2FN(GF_STANDARIZAR_FECHA_RTRN(date())) then setError(SM_OT_FALTA_FECHA_PROG_VIEJA)
end if	
if SM_OTFrequencyUnit = ORDER_FREQ_UNIQUE then 
	SM_OTFrequencyQuantity = 0 
else
	if SM_OTFrequencyQuantity = 0 then setError(SM_OT_FALTA_FREQUENCIA)
end if	
if (hayError) then rtrn = false
checkHeaderOT = rtrn
end function
'------------------------------------------------------------------
function getDsMaintenanceType(pCdType)
dim rtrn
select case pCdType
	case MAIN_TYPE_PREVENTIVE
		rtrn = "Preventivo"
	case MAIN_TYPE_PREDICTIVE 
		rtrn = "Predictivo"
	case MAIN_TYPE_CORRECTIVE
		rtrn = "Correctivo"
	case else 
		rtrn = "Error"	
end select	
getDsMaintenanceType = rtrn	
end function
'------------------------------------------------------------------
function getDsOrderType(pCdType)
dim rtrn
select case pCdType
	case ORDER_TYPE_MECHANICAL
		rtrn = "Mec�nico"
	case ORDER_TYPE_ELECRONIC 
		rtrn = "El�ctrico"
	case ORDER_TYPE_CIVIL
		rtrn = "Civil"
	case ORDER_TYPE_SECURITY 
		rtrn = "Seguridad"
	case ORDER_TYPE_OPERATIVE 
		rtrn = "Operativa"
	case ORDER_TYPE_SYSTEM 
		rtrn = "Sistema"				
	case else 
		rtrn = "Error"	
end select	
getDsOrderType = rtrn	
end function
'------------------------------------------------------------------
function getDsStateScheduled(pCdType)
dim rtrn
select case CINT(pCdType)
	case CINT(ORDER_FREQ_ENABLED)
		rtrn = "Activa"
	case CINT(ORDER_FREQ_DISABLED)
		rtrn = "Inactiva"
	case else 
		rtrn = "Error"	
end select	
getDsStateScheduled = rtrn	
end function
'------------------------------------------------------------------
function getDsState(pCdType)
dim rtrn
select case pCdType
	case STATE_STAND_BY
		rtrn = "Programada"
	case STATE_STARTED 
		rtrn = "Iniciada"
	case STATE_FINISHED
		rtrn = "Finalizada"
	case STATE_CANCELED
		rtrn = "<font style='FONT-SIZE: 10px;' color='red'><b>Cancelada</b></font>"
	case else 
		rtrn = "Error"	
end select	
getDsState = rtrn	
end function
'---------------------------------------------------------------------------------------------
function getFrequency(pUnit, pQuantity)
dim rtrn
rtrn = "Cada " & pQuantity
select case pUnit
	case ORDER_FREQ_DAY 
		rtrn = rtrn & " d�as."
	case ORDER_FREQ_WEEK
		rtrn = rtrn & " semanas."
	case ORDER_FREQ_MONTH 
		rtrn = rtrn & " meses."
	case ORDER_FREQ_YEAR 
		rtrn = rtrn & " a�os."
	case else 
		rtrn = "Error"	
end select	
getFrequency = GF_Traducir(rtrn)
end function
'---------------------------------------------------------------------------------------------
Function initItemsOT()
	SM_actualItem=0
	initItemsOT = true
	if (cDbl(SM_idOrder) > 0) then
		initItemsOT = initItemsOTDB()
	end if
End Function
'---------------------------------------------------------------------------------------------
function initItemsOTDB()
	dim strSQL, rs
	initItemsOTDB = false
	call executeProcedureDb(DBSITE_SQL_INTRA, rsItems, "TBLSMORDERITEM_GET_BY_ID", SM_idOrder)
	if not rsItems.eof then initItemsOTDB = true
end function
'---------------------------------------------------------------------------------------------
Function initTasksOT()
	SM_ActualTask=0
	initTasksOT = true
	if (CLng(SM_idOrder) > 0) then
		initTasksOT = initTasksOTDB()
	end if
End Function
'---------------------------------------------------------------------------------------------
function initTasksOTDB()
	dim strSQL, rs
	initTasksOTDB = false
	call executeProcedureDb(DBSITE_SQL_INTRA, rsTasks, "TBLSMORDERTASK_GET_BY_ID", SM_idOrder)
	if not rsTasks.eof then initTasksOTDB = true
end function
'---------------------------------------------------------------------------------------------
Function readNextTaskOt()
dim ret
	Call clearTask()
	if (isFormSubmit() and accion<> "") then
		ret = readNextTaskOtParams()
	else 
		if (cint(SM_idOrder) > 0) then
			ret = readNextTaskOtDB()
		end if
	end if		
	readNextTaskOt = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextTaskOtDB()
	dim strSQL, rs, km
	readNextTaskOtDB = false	
	if not rsTasks.eof then
		SM_nroTask = rsTasks("NROTASK")
		SM_dsTask = rsTasks("DSTASK")
		SM_doneTask = rsTasks("DONE")		
		readNextTaskOtDB = true
		rsTasks.MoveNext()
	end if	
SM_ActualTask = SM_ActualTask + 1	
End Function
'---------------------------------------------------------------------------------------------
Function readNextTaskOtParams()
	dim strSQL, rs, ret
	ret = false		
	SM_ActualTask = SM_ActualTask + 1
	SM_dsTask = GF_PARAMETROS7("SM_dsTask" & SM_ActualTask, "",6)
	if (SM_dsTask <> "") then
		SM_nroTask = GF_PARAMETROS7("SM_nroTask" & SM_ActualTask, 0,6)
		SM_doneTask = GF_PARAMETROS7("SM_doneTask" & SM_ActualTask, "",6)
		ret = true
	end if
	readNextTaskOtParams = ret
End Function
'---------------------------------------------------------------------------------------------
Function clearTask()
	SM_nroTask = 0
	SM_dsTask = ""
	SM_doneTask = "N"
End Function
'---------------------------------------------------------------------------------------------
Function readNextItemOt()
dim ret
	Call clearItem()
	ret = false
	if (isFormSubmit() and accion<> "") then
		ret = readNextItemOtParams()
	else 
		if (cint(SM_idOrder) > 0) then
			ret = readNextItemOtDB()
		end if
	end if		
	readNextItemOt = ret
End Function
'---------------------------------------------------------------------------------------------
Function readNextItemOtDB()
	dim strSQL, rs, km
	readNextItemOtDB = false	
	if not rsItems.eof then
		SM_nroItem = rsItems("NROITEM")
		SM_idItem = rsItems("IDITEM")
		SM_dsItem = rsItems("DSITEM")
		SM_programQuantityItem = rsItems("PROGRAMQUANTITY")		
		SM_idPMItem = rsItems("IDPM")
		if SM_cdState = STATE_STARTED then 'Ir a buscar PMs asociados
			'PM asociados directamente a la OT			
			'Pendiente! Modificar carga de vales!
			'PM generado por la OT
			SM_realQuantityItem = getCumplimientoOTItem(SM_idPMItem, SM_idItem)
			SM_realQuantityItem = Cdbl(SM_programQuantityItem) - Cdbl(SM_realQuantityItem)
		else
			SM_realQuantityItem = rsItems("REALQUANTITY")
		end if	
		readNextItemOtDB = true
		rsItems.MoveNext()
	end if	
SM_actualItem = SM_actualItem + 1	
End Function
'---------------------------------------------------------------------------------------------
function getCumplimientoOTItem(pIdPM, pIdItem)
dim strSQL, rsArticulos, rtrn
rtrn = 0 
strSQL = "select * from TBLPMDETALLE where IDPEDIDO=" & pIdPM & " AND IDARTICULO=" & pIdItem
call executeQueryDb(DBSITE_SQL_INTRA, rsArticulos, "OPEN", strSQL)
if (not rsArticulos.eof) then rtrn = rsArticulos("SALDO")
getCumplimientoOTItem = rtrn 
end function
'---------------------------------------------------------------------------------------------
Function readNextItemOtParams()
	dim strSQL, rs, ret
	ret = false		
	SM_actualItem = SM_actualItem + 1
	SM_dsItem = GF_PARAMETROS7("SM_dsItem" & SM_actualItem, "",6)
	SM_idItem = GF_PARAMETROS7("SM_idItem" & SM_actualItem, 0,6)
	if (trim(SM_dsItem) <> "") then
		SM_nroItem = GF_PARAMETROS7("SM_nroItem" & SM_actualItem, 0,6)
		SM_realQuantityItem = GF_PARAMETROS7("SM_realQuantityItem" & SM_actualItem, "",6)
		SM_programQuantityItem = GF_PARAMETROS7("SM_programQuantityItem" & SM_actualItem, "",6)
		SM_idPMItem = GF_PARAMETROS7("SM_idPmItem" & SM_actualItem, "",6)
		ret = true
	end if
	readNextItemOtParams = ret
End Function
'---------------------------------------------------------------------------------------------
Function clearItem()
		SM_nroItem = 0
		SM_idItem = 0
		SM_dsItem = ""
		SM_realQuantityItem = 0
		SM_programQuantityItem = 0		
		SM_idPMItem = 0
End Function
'---------------------------------------------------------------------------------------------
function getMaxFromDivision(pIdDivision)
Dim strSQL, rs, rtrn
rtrn = "ERR"
strSQL= "SELECT MIN(IDALMACEN) AS ALMACEN_DEFAULT FROM TBLALMACENES WHERE IDDIVISION=" & pIdDivision
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then rtrn = rs("ALMACEN_DEFAULT")
getMaxFromDivision = rtrn
end function
%>

