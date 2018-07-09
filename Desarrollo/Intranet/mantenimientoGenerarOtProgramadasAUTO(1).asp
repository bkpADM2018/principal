<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%
'TAREA 1748
dim strSQL, rsList, diffDays, nextExecution
dim strTitle, strBody, strMailFrom, strMailTo, G_NroOrden
'Obtener todos las OT planificadas que requieren que se les genere una OT recurrente
call executeSP(rsList, "TOEPFERDB.TBLSMOTEXECUTIONS_GET_BY_IDOTGENERATED", 0)
while not rsList.eof 
		'Se toma una semana como tiempo prudencial a la generacion de una OT recurrente
		diffDays = Clng(rsList("NEXTEXECUTIONEX")) - clng(GF_DTE2FN(DateAdd("d",7,Date())))
		if diffDays<=0 then
			'Implica que se alcanzo la fecha estimada de ejecucion de la proxima ocurrencia 
			'Se debe crear una OT copia de la master
			myOtGenerated = copyMasterOT(rsList("IDOT"), rsList("NEXTEXECUTIONEX"))
			'Se guarda en la tabla de ejecuciones, la ot generada
			call updateOtGenerated(rsList("IDOT"), myOtGenerated)
		end if
	rsList.movenext
wend	
if G_NroOrden <> "" then
	'Enviar Mail
	strTitle = "Sistema de Mantenimiento - Planificaciones"
	strBody = "Detalle de las Ordenes de Trabajo generadas a partir de la planificación:" & G_NroOrden
	strMailFrom = "bacarinie@toepfer.com"
	strMailTo = "bacarinie@toepfer.com"
	Call GP_ENVIAR_MAIL(strTitle, strBody, strMailFrom, strMailTo)	
end if
'-----------------------------------------------------
function copyMasterOT(pIdOtSource, pScheduledDate)
dim idOtGenerated, rsOT

idOtGenerated = 0
SM_idOrder = pIdOtSource
'Leer cabecera de OT maestra
call readHeaderOtDB()
'Generar nuevo numero
SM_nroOrder = getNumeracionOT(SM_idActiveEquipment)
SM_cdState = STATE_STAND_BY 
G_NroOrden =  vbcrlf & G_NroOrden & "Nro de Orden: " & SM_nroOrder & vbcrlf & "Titulo: " & SM_dsOrder & vbcrlf & "Fecha Programada: " & GF_FN2DTE(pScheduledDate) & vbcrlf & vbcrlf
call executeSP(rs, "TOEPFERDB.TBLSMORDER_INS", SM_nroOrder & "||" & SM_dsOrder & "||" & SM_idDivision & "||" & SM_idActiveEquipment & "||" & pScheduledDate & "||" & SM_maintenanceType & "||" & SM_cdState & "||" & SM_cdApplicant & "||" & SM_orderType & "||" & SM_idResponsableCompany & "||" & SM_idObra & "||" & SM_idBudgetArea & "||" & SM_idBudgetDetalle & "||" & session("Usuario") & "||" & session("MmtoDato") & "||" )

'Obtener id de la OT recientemente generada
call executeSP(rsMax, "TOEPFERDB.TBLSMORDER_GET_MAX_ID","" )
idOtGenerated = rsMax("MAXID")

'Para el detalle de las OT se va a leer de la maestra y se va a grabar en la generada con lo cual se va a ir cambiando
'el valor de la variable global SM_idOrder dependiendo de la accion que se va a realizar

'Leer detalles de OT Maestra
SM_idOrder = cLng(pIdOtSource)
call initTasksOTDB()
while readNextTaskOtDB()
	SM_doneTask = SM_TASK_DONE_NO 
	'guardar nuevo detalle en id de OT generada
	SM_idOrder = cLng(idOtGenerated)
	SM_nroTask = 0
	saveOtTasks()
	'Se vuelve a poner la OT maestra para que continue leyendo sus items
	SM_idOrder = cLng(pIdOtSource)
wend


SM_idOrder = cLng(pIdOtSource)
call initItemsOTDB()
while readNextItemOtDB()
	SM_idPMItem = 0
	SM_realQuantityItem = 0
	SM_idOrder = cLng(idOtGenerated)
	SM_nroItem = 0
	saveOtItems()
	SM_idOrder = cLng(pIdOtSource)
wend

copyMasterOT = idOtGenerated
end function
'-----------------------------------------------------
sub updateOtGenerated(pMasterOT, pGeneratedOT)
	call executeSP(rs, "TOEPFERDB.TBLSMOTEXECUTIONS_INS", pMasterOT & "||0||" & pGeneratedOT & "||S")
end sub
%>