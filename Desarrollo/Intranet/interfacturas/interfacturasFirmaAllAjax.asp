<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosunificador.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="../Includes/procedimientosseguridad.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosHKEY.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<% 
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pConcepto, pUsuario)
	Dim rs, rsX		
	
    Call executeSP(rs, "TFFL.TF100F1_GET_CBTES_A_FIRMAR", pUsuario & "||1||0$$totalRegistros")    
    while (not rs.eof)
        if (rs("LDLYCD") = pConcepto) then
	        Call executeSP(rsX, "TFFL.TF100F1_UPD_FIRMAR_CBTE", rs("FCRGNR") & "||" & pUsuario)
	        Call executeSP(rsX, "TFFL.TF105F1_INS", rs("FCRGNR") & "||" & pUsuario & "||" & session("MmtoDato"))
        end if	        
	    rs.MoveNext()
    wend	    
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
Dim cd,gCdUsuario,gsecuencia,respuesta

cd = GF_PARAMETROS7("cd", "", 6)

Call GP_CONFIGURARMOMENTOS()
respuesta = LLAVE_NO_CORRESPONDE
if (cd <> "") then
	if (leerRegistroFirmas()) then			
		respuesta = RESPUESTA_OK			
		Call registrarFirma(cd,gCdUsuario)
	end if
else
	respuesta = CODIGO_VACIO	
end if	
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
