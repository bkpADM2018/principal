<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<%
Call initAccessInfo(RES_INV_SM)
Dim strSQL, rs, conn
dim idEquipo, idEquipoActivo, idDivision, idSector, idUbicacion, tipoOperacion, cdActivoFijo, rtrnErr

'tipoOperacion
'A = Activar Equipo
'M = Modificar Equipo Activo
'H = Habilitar Equipo Activo
'B = Desactivar/Deshabilitar Equipo Activo

idEquipo = GF_PARAMETROS7("idEquipo", 0,6)
idEquipoActivo = GF_PARAMETROS7("idEquipoActivo", 0,6)
idDivision = GF_PARAMETROS7("idDivision", 0,6)
idSector = GF_PARAMETROS7("idSector", 0,6)
idUbicacion = GF_PARAMETROS7("idUbicacion", 0,6)
cdActivacion = trim(UCASE(GF_PARAMETROS7("cdActivacion", "",6)))
dsActivacion = trim(GF_PARAMETROS7("dsActivacion", "",6))
cdActivoFijo = trim(UCASE(GF_PARAMETROS7("cdActivoFijo", "",6)))
tipoOperacion = GF_PARAMETROS7("tipoOperacion", "",6)

if tipoOperacion = "A" then
	if checkDatosActivacion(cdActivacion, dsActivacion) then
		if not existeEquipoActivo(idEquipoActivo, idDivision, cdActivoFijo, cdActivacion) then
			call activarEquipo(idEquipo, idDivision, idSector, idUbicacion, cdActivacion, dsActivacion, cdActivoFijo)
		else
			rtrnErr = setError(SM_EQUIPO_ACTIVO_YA_EXISTE)
		end if
	end if	
elseif tipoOperacion = "M" then
	if checkDatosActivacion(cdActivacion, dsActivacion) then
		if not existeEquipoActivo(idEquipoActivo, idDivision, cdActivoFijo, cdActivacion) then
			call modificarEquipoActivo(idEquipoActivo, idDivision, idSector, idUbicacion, cdActivacion, dsActivacion, cdActivoFijo)
		else
			rtrnErr = setError(SM_EQUIPO_ACTIVO_YA_EXISTE)
		end if
	end if
elseif tipoOperacion = "H" then
	cdActivacion = right(cdActivacion,len(cdActivacion)-2)
	if not existeEquipoActivo(idEquipoActivo, idDivision, cdActivoFijo, cdActivacion) then
		call desactivarEquipo(idEquipoActivo, ESTADO_ACTIVO)
	else
		rtrnErr = setError(SM_EQUIPO_ACTIVO_YA_EXISTE)
	end if	
elseif tipoOperacion = "B" then
	if not tieneOTActiva(idEquipoActivo) then
		call desactivarEquipo(idEquipoActivo, ESTADO_BAJA)
	else
		rtrnErr = setError(SM_EQUIPO_ACTIVO_EXISTE_EN_OT)
	end if	
end if	

if hayError() then
	Response.Write showMessages

end if
if err.number <> 0 then
%>
		<tr>
		<td colspan="2"><%=err.number & " - " & err.Description%></td>
		</tr>
<%
end if
%>