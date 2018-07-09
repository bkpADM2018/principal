<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_INV_SM)
Dim strSQL, rs, conn
dim idEquipo, idEquipoActivo, idComponente, idGrupo, dsComponente, tipoOperacion, rtrnErr
'tipoOperacion
'A = Alta
'B = Baja
'M = Modificacion
'H = Habilitar
idEquipo = GF_PARAMETROS7("idEquipo", 0,6)
idEquipoActivo = GF_PARAMETROS7("idEquipoActivo", 0,6)
idGrupo = GF_PARAMETROS7("idGrupo", 0,6)
idComponente = GF_PARAMETROS7("idComponente", 0,6)
dsComponente = GF_PARAMETROS7("dsComponente", "",6)
tipoOperacion = GF_PARAMETROS7("tipoOperacion", "",6)

if tipoOperacion = "A" then
	if trim(dsComponente) = "" then
		rtrnErr = setError(SM_COMPONENTE_DESC_REQ)
	else	
		if not existeComponente(idEquipo, idEquipoActivo, idComponente, dsComponente, idGrupo) then
			call agregarComponente(idEquipo, idEquipoActivo, dsComponente, idGrupo)
		else
			rtrnErr = setError(SM_COMPONENTE_YA_EXISTE)
		end if
	end if
elseif tipoOperacion = "M" then
	if trim(dsComponente) = "" then
		rtrnErr = setError(SM_COMPONENTE_DESC_REQ)
	else	
		if not existeComponente(idEquipo, idEquipoActivo, idComponente, dsComponente, idGrupo) then
			call modificarComponente(idComponente, dsComponente)
		else
			rtrnErr = setError(SM_COMPONENTE_YA_EXISTE)
		end if
	end if	
elseif tipoOperacion = "B" then
	if not tieneOTActiva(idComponente) then
		call adminComponente(idComponente, idGrupo, ESTADO_BAJA)
	else
		rtrnErr = setError(SM_COMPONENTE_EXISTE_EN_OT)
	end if	
elseif tipoOperacion = "H" then
	call adminComponente(idComponente, idGrupo, ESTADO_ACTIVO)
end if	
'------------------------------------------------------------------------------------------------------------------------
if hayError() then
	Response.Write showMessages()
end if
%>