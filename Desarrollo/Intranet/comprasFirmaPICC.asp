<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosunificador.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<% 
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pMinuta, pFecha, pTipoCbte, pEvento, pUsuario)
	Dim rs
    
	set rs_Ret = executeSP(rs, "MERFL.MER301F1_UPD_FIRMAR_MINUTA", pMinuta &"||"& pFecha &"||"& pTipoCbte &"||"& pEvento &"||"& Session("Usuario") &"||"& Session("MmtoDato"))
    'Si actulizo el estado correctamente, agrego la firma
    if (rs_Ret(SP_IDERROR) = ESTADO_ACTIVO) then
        Call executeSP(rs, "MERFL.TBLMINUTASFIRMAS_INS", pMinuta &"||"& pFecha &"||"& pEvento &"||"& Session("Usuario") &"||"& Session("MmtoDato"))
        registrarFirma = RESPUESTA_OK		
	end if

End Function
'---------------------------------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds	
	ret = false	
	if (HK_isKeyReady()) then
		strSQL = "Select * from TOEPFERDB.TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
        Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true
		else
			gCdUsuario = ""
		end if
	end if		
	leerRegistroFirmas = ret
End Function
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim minuta,gCdUsuario,gsecuencia,respuesta,fecha,tipoCbte,evento

minuta = GF_PARAMETROS7("minuta", "", 6)
fecha = GF_PARAMETROS7("fecha", "", 6)
tipoCbte = GF_PARAMETROS7("tipoCbte", "", 6)
evento = GF_PARAMETROS7("evento", "", 6)
    
Call GP_CONFIGURARMOMENTOS()

respuesta = LLAVE_NO_CORRESPONDE
if (minuta <> "")and(fecha <> "") then
	if (leerRegistroFirmas()) then
		respuesta = registrarFirma(minuta,fecha,tipoCbte,evento,gCdUsuario)
	end if
else
	respuesta = CODIGO_VACIO
end if	
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
