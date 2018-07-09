<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientosVales.asp"-->
<%
'----------------------------------------------------------------------------------------------------------------------
'Devuelve el Id de Ajuste de un Draft Survey
Function getIdAjusteByDraft(pIddraft)
	Dim strSQL, rtrn
	rtrn = "" 
	strSQL = "SELECT IDAJUSTE FROM TBLAJUSTES WHERE IDORIGEN = " & pIddraft & " AND CDAJUSTE = '"& AJUSTE_DRAFT_SURVEY &"'"
	Call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)
	if not rs.Eof Then rtrn = rs("IDAJUSTE")
	getIdAjusteByDraft = rtrn
end Function	
'----------------------------------------------------------------------------------------------------------------------
Dim g_strPuerto,g_cdAviso, strClass, g_cdPorducto,cont, g_idDraft, g_fechaDS, strFecha, flagControl, g_accion,g_kilosDs
Dim g_fechaBza, g_kilosBza, g_archivo, g_difKilos, msg, g_idAjuste

g_strPuerto = GF_Parametros7("Pto","",6)
g_accion    = GF_Parametros7("accion","",6)
g_cdAviso   = GF_Parametros7("cdAviso",0,6)
g_cdPorducto= GF_Parametros7("cdProducto",0,6)
g_fechaDS   = GF_Parametros7("fechaDraft","",6)
g_kilosDs   = GF_Parametros7("kilosDraft",2,6)
g_fechaBza  = GF_Parametros7("fechaBza","",6)
g_fechaBza  = Year(g_fechaBza) & GF_nDigits(Month(g_fechaBza),2) & GF_nDigits(Day(g_fechaBza),2)
g_kilosBza  = GF_Parametros7("kilosBza",0,6)
g_kilosBzaToepfer  = GF_Parametros7("kgBzaToepfer",0,6)
g_archivo   = GF_Parametros7("archivo","",6)
g_idDraft   = GF_Parametros7("idDraft",0,6)
g_kilosCosecha = GF_Parametros7("kilosCosecha",0,6)

if (g_accion = ACCION_GRABAR) Then		
	if(g_kilosDS = 0) Then	msg = "<li>Debe ingresar los kilos.</li>"
	if(g_kilosCosecha > 0)then 
		if(g_kilosDs < g_kilosCosecha) Then	msg = "<li>Debe desasignar kilos con Cosecha y luego cargue el Draft Survey.</li>" & msg
	end if	
	if msg = "" then
		if(g_idDraft > 0)then
			'Actualiza
			Call updateDraftSurvey(g_idDraft,g_fechaDS,g_kilosDs,g_fechaBza,g_kilosBza, g_kilosBzaToepfer)
			g_difKilos  = g_kilosBza - g_kilosDs
			'Si se actualiza un draft se debera actualizar el ajuste, su estado pasa a no autorizado (esperando las firmas)
            Call updateAjustePuerto(AJUSTE_DRAFT_SURVEY, g_idDraft, g_fechaDS, g_difKilos,AJUSTE_ESTADO_NOAUTORIZADO)			
			g_idAjuste = getIdAjusteByDraft(g_idDraft)
            'Si se actualiza un draft se debera borrar las autorizaciones del ajuste para que lo firmen nuevamente y el estado del ajuste vuelve al inicial
            Call executeSP_Puertos(rs, g_strPuerto, "TBLAJUSTESFIRMAS_DEL_BY_IDAJUSTE", g_idAjuste)
		else
			'Nuevo
			g_idDraft = grabarDraftSurvey(g_cdAviso, g_cdPorducto, g_kilosDs, g_fechaDS, g_fechaBza, g_kilosBza, g_kilosBzaToepfer)
			if(g_archivo <> "")Then Call saveDraftAttach(g_archivo,g_idDraft)			
			g_difKilos  = g_kilosDs - g_kilosBza
			Call executeSP_Puertos(rs, g_strPuerto, "TBLAJUSTES_INS", AJUSTE_DRAFT_SURVEY &"||"& g_idDraft &"||"& g_cdPorducto &"||"& g_difKilos &"||0||0||0||0||0||"& AJUSTE_ESTADO_NOAUTORIZADO &"||"& g_fechaDS &"||"& g_fechaDS &"||"& Session("Usuario") &"||"& Session("MmtoDato") &"||"& getIdDivision(g_strPuerto))
            strSQL = "SELECT MAX(IDAJUSTE) AS IDAJUSTE FROM TBLAJUSTES"
	        Call GF_BD_Puertos(g_strPuerto, rsAjs, "OPEN",strSQL)
	        if (not rsAjs.Eof) Then g_idAjuste = rsAjs("IDAJUSTE")
		end if
		msg = RESPUESTA_OK
	else
		msg = "<u>ATENCION:</u><ul>" & msg & "</ul>"	
	end if
else if (g_accion = ACCION_BORRAR) Then
		Call deleteDraftSurvey(g_idDraft)
		g_idAjuste = getIdAjusteByDraft(g_idDraft)
		if (g_idAjuste <> "") then
            'Modifico a estado Cancelado el ajuste que contiene el Draft
            Call executeSP_Puertos(rs, g_strPuerto, "TBLAJUSTES_UPD_ESTADO_BY_PARAMETERS", "||"& g_idAjuste &"||"& AJUSTE_ESTADO_CANCELADO)
		    Call executeSP_Puertos(rs, g_strPuerto, "TBLAJUSTESFIRMAS_DEL_BY_IDAJUSTE", g_idAjuste)
		end if
		msg = RESPUESTA_OK
	 end if
end if
Response.Write msg


%>





