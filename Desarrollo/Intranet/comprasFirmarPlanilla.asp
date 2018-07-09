<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPCP.asp"-->
<%
Dim usr, dPass, idPedido, ret, errCode, gCdUsuario,secuencia
'----------------------------------------------------------------------------------------------------------------------
Function checkOperacion(pIdPedido, pSecuencia, pCdUsuario)			
	dim valor, rs, strSQL, conn,auxUser, rolUsuario, ctrl
		
	checkOperacion=ERROR_AUTENTICACION
	if (pCdUsuario = session("Usuario")) then
		strSQL = "Select F.*, C.IDSECTOR from TBLPCPFIRMAS F inner join TBLPCTCABECERA C on C.IDPEDIDO=F.IDPEDIDO where F.IDPEDIDO=" & pIdPedido & " and F.SECUENCIA=" & pSecuencia
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then 
		    ctrl = false
		    rolUsuario = CInt(getRolFirma(pCdUsuario, SEC_SYS_COMPRAS)) 
		    if ((pSecuencia = PCP_FIRMA_RESPONSABLE) and (rs("CDUSUARIO") = gCdUsuario))        then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_MIEMBRO1)    and (rs("CDUSUARIO") = gCdUsuario))        then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_MIEMBRO2)    and (rs("CDUSUARIO") = gCdUsuario))        then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_MIEMBRO3)    and (rs("CDUSUARIO") = gCdUsuario))        then ctrl = true
		    if (pSecuencia = PCP_FIRMA_GTE_SECTOR) then 
		        if (isBossOf(pCdUsuario, rs("IDSECTOR"))) then ctrl = true
            end if		        
		    if ((pSecuencia = PCP_FIRMA_GTE_PUERTO)  and (rolUsuario= FIRMA_ROL_RESP_PUERTO))   then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_GTE_COMPRAS) and (rolUsuario= FIRMA_ROL_GTE_COMPRAS))   then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_SUP_PUERTOS) and (rolUsuario= FIRMA_ROL_SUP_PUERTO))    then ctrl = true
		    if ((pSecuencia = PCP_FIRMA_DIRECCION)   and (rolUsuario= FIRMA_ROL_DIRECTOR))      then ctrl = true		
		    if (ctrl) then checkOperacion = RESPUESTA_OK
        end if		
	end if
		
End Function
'------------------------------------------------------------------
Function registarFirma(pIdPedido, pCdUsuario, llave)
	Dim strSQL, conn, rs
	
	strSQL="Update TBLPCPFIRMAS SET FECHAFIRMA=" & session("MmtoDato") & ", HKEY='" & llave & "', CDUSUARIO='" & UCase(pCdUsuario) & "' where IDPEDIDO=" & pIdPedido & " and SECUENCIA=" & secuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	
End Function
'------------------------------------------------------------------
' Autor  :
'			Ajaya Nahuel - CNA
' Fecha	 :	
'			27-03-2014
' Nombre :
'			NotificarPlanillaCompleta
' Objetivo:
'			Se encarga de enviar mail notificando que se completaron todas las firmas de la Planilla
'			Los que reciben el mail son:
'				1- Licitaciones (Toepfer)
'				2- Responsables de Seguridad (usuario cargados en una lista de correo)
'					Para identificar a una lista se requiere:	
'						-->	DIVISION: Representa a la division de donde proviene la Planilla	
'						-->	CODIGO: Representa a la seccion, en este caso es Planilla Comparativa
'					De esta manera se controla dinamicamente el envio de mail para una determinada lista
'---------------------------------------------------------------------------------------------------------------------
Function NotificarPlanillaCompleta()
	Dim strMsg, idUsuario, ds, emailToepfer,strTitle	
	'Preparo la descripcion del mail para ambos casos
	strMsg = "Se ha adjudicado el pedido de precio " & pct_cdPedido & " - " & pct_tituloPedido & vbCrLf & vbCrLf
	strMsg = strMsg & "El proveedor seleccionado es " &  pct_idProveedorElegido &" - "& Trim(pct_dsProveedorElegido) &" ("& getCUITProveedor(pct_idProveedorElegido) &") "& vbCrLf
	strTitle = GF_TRADUCIR("Sistema de Compras Web - Nueva Adjudicacion - Pedido " & pct_cdPedido )
	emailToepfer = obtenerMail(CD_TOEPFER)
	if (emailToepfer <> "") then
		'1)Manda el mail a licitaciones.
		Call GP_ENVIAR_MAIL(strTitle,strMsg,emailToepfer,emailToepfer)
		'2)Manda el mail a los responsables de Seguriadad de la division.
		Set rs = getListMail(pct_idDivision, LISTA_PCP_PROV_GANADOR)		
		while not rs.eof
			if(Len(Trim(rs("EMAIL"))) > 0)then Call GP_ENVIAR_MAIL(strTitle,strMsg,emailToepfer, Trim(rs("EMAIL")))
			rs.MoveNext()
		wend
	end if	
End Function
'------------------------------------------------------------------
Function revisarEstadoPedido(pIdPedido)

Dim rsApertura, strSQL, conn

if (checkPlanillaFinalizada(pIdPedido)) then	
	'Se cambia el estado del pedido
	strSQL="Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_ADJUDICADO & " where IDPEDIDO=" & pct_idPedido
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	Call NotificarPlanillaCompleta()
else
	if (pct_idEstado <> ESTADO_PCT_EN_FIRMA_AC) then
		strSQL="Update TBLPCTCABECERA set ESTADO=" & ESTADO_PCT_EN_FIRMA_AC & " where IDPEDIDO=" & pct_idPedido
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	end if
end if

End Function
'------------------------------------------------------------------
Function checkPlanillaFinalizada(pIdPedido)
	Dim ret, conn, strSQL, rs

	ret = false
	strSQL = "Select * from TBLPCPFIRMAS where IDPEDIDO=" & pIdPedido & " and HKEY is NULL"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (rs.eof) then ret = true
	checkPlanillaFinalizada = ret 
	
End Function
'------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	
	gCdUsuario = ""				
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then			
			gCdUsuario = rs("CDUSUARIO")
			ret = true
		end if
	end if
	leerRegistroFirmas = ret
End Function
'**************************************
'***	COMIENZO DE LA PAGINA		***
'**************************************
'Se llama por AJAX para verificar una credencial
idPedido = GF_PARAMETROS7("idPedido",0,6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
Call GP_CONFIGURARMOMENTOS()

ret = LLAVE_NO_CORRESPONDE
if (leerRegistroFirmas()) then	
	Call initHeader(idPedido)
	'1º - Se controla que el usuario tenga permiso para la operación		
	ret = checkOperacion(idPedido, secuencia, gCdUsuario)	
	if (ret = RESPUESTA_OK) then				
		Call registarFirma(idPedido, gCdUsuario, HK_readKey())
		Call revisarEstadoPedido(idPedido)		
	end if
end if
Call HK_sendResponse(ret)
%>