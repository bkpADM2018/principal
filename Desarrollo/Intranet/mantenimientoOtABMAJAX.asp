<!--#include file="Includes/procedimientosALmacenes.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<%
Call initAccessInfo(RES_OT_SM)
'tipoOperacion
'DT = Eliminar Task
'UO = Modificar Estado OT
'DI = Eliminar Item

tipoOperacion = GF_PARAMETROS7("tipoOpr", "",6)
idOrder = GF_PARAMETROS7("idOrder", 0,6)
'cdEstado = GF_PARAMETROS7("cdEstado", 0,6)

if tipoOperacion = "UO" then
	SM_idOrder = idOrder
	SM_cdState_New = GF_PARAMETROS7("cdEstado", 0,6)
	if SM_cdState_New = STATE_STARTED then
		'Generar el PM
		call readHeaderOT(SM_idOrder)
		PM_idAlmacen = getMaxFromDivision(SM_idDivision)
		PM_idObra = SM_idObra
		PM_FechaSolicitud = day(date()) & "/" & month(date()) & "/" & year(date())
		PM_FechaRequerido = PM_FechaSolicitud
		Response.Write PM_FechaSolicitud & "---" & PM_FechaRequerido
		PM_idAlmacenDest = 0
		PM_idBudgetArea = SM_idBudgetArea
		PM_idBudgetDetalle = SM_idBudgetDetalle
		PM_comentario = "Generado automaticamente por la Orden de Trabajo Nro: " & SM_nroOrder
		PM_idSector = 0
		PM_cdSolicitante = SM_cdApplicant
		idPM = grabarHeaderPMInsert()	
		call initItemsOT()
		while readNextItemOt()
			call grabarPMDetalle(idPM, SM_idItem, CDbl(SM_programQuantityItem), 0)
		wend		
	end if	
	'Actualizar el nro de PM de los items
	SM_idPMItem = idPM
	call udpateOtItemsPM
	SM_cdState = SM_cdState_New
	SM_date = GF_DTE2FN(day(date()) & "/" & month(date()) & "/" & year(date()))
	SM_observations = GF_PARAMETROS7("SM_observations", "",6)
	call updateOtStatus()
elseif tipoOperacion = "DT" then
	SM_idOrder = idOrder
	SM_nroTask = GF_PARAMETROS7("nroTask", 0,6)
	SM_typeTask = GF_PARAMETROS7("typeTask", 0,6)
	call deleteOtTask()
elseif tipoOperacion = "DI" then
	SM_idOrder = idOrder
	SM_nroItem = GF_PARAMETROS7("nroItem", 0,6)
	SM_typeItem = GF_PARAMETROS7("typeItem", 0,6)
	call deleteOtItem()
end if	

if hayError() then
%>
		<tr>
		<td colspan="2"><% call showErrors() %></td>
		</tr>
<%
end if
if err.number <> 0 then
%>
		<tr>
		<td colspan="2"><%=err.number & " - " & err.Description%></td>
		</tr>
<%
end if
%>