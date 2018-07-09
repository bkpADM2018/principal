<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosunificador.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="../Includes/procedimientosseguridad.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosHKEY.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<% 
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pIdFactura, pUsuario)
	Dim rs
		
	Call executeSP(rs, "TFFL.TF100F1_UPD_FIRMAR_CBTE", pIdFactura & "||" & pUsuario)
	Call executeSP(rs, "TFFL.TF105F1_INS", pIdFactura & "||" & pUsuario & "||" & session("MmtoDato"))
	
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
Dim id,gCdUsuario,gsecuencia,respuesta

id = GF_PARAMETROS7("id", "", 6)

Call GP_CONFIGURARMOMENTOS()
respuesta = LLAVE_NO_CORRESPONDE
if (id <> "") then
	if (leerRegistroFirmas()) then			
		respuesta = RESPUESTA_OK			
		Call registrarFirma(id,gCdUsuario)
	end if
else
	respuesta = CODIGO_VACIO	
end if	
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
