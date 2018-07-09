<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosunificador.asp"-->
<!--#include file="Includes/procedimientosparametros.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<% 

Function actualizarDetalleTrimestralPartidaPresupuestaria(pIdObra, pIdAreaOrigen, pIdDetalleOrigen, pIdAreaDestino, pIdDetalleDestino, pImportePesos, pImporteDolares, pTipoCambio, pPeriodo)
    Dim strSQL,rsUpd
    'Actualizo primero la Partida Origen
    strSQL = "UPDATE TBLBUDGETOBRASDETALLE SET psbudget = psbudget - " & pImportePesos & ", dlbudget = dlbudget - " & pImporteDolares &_
	         " WHERE idarea = " & pIdAreaOrigen & " AND iddetalle = " & pIdDetalleOrigen & " AND periodo = " & pPeriodo & " and idobra =" & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "UPDATE", strSQL)
	
    'Actualizo la Partida Destino
    strSQL = "Select * from TBLBUDGETOBRASDETALLE WHERE idarea = " & pIdAreaDestino & " AND iddetalle = " & pIdDetalleDestino & " AND periodo = " & pPeriodo & " and idobra =" & pIdObra
	Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
	if (not rs2.EoF) then
	    strSQL = "UPDATE TBLBUDGETOBRASDETALLE SET psbudget = psbudget + " & pImportePesos & ", dlbudget = dlbudget + " & pImporteDolares &_
		         " WHERE idarea = " & pIdAreaDestino & " AND iddetalle = " & pIdDetalleDestino & " AND periodo = " & pPeriodo & " and idobra =" & pIdObra
		Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "UPDATE", strSQL)
	else
		strSQL = "INSERT INTO TBLBUDGETOBRASDETALLE (idobra,idarea,iddetalle,periodo,psbudget,dlbudget,tipocambio,cdusuario,momento) "&_
		         " VALUES ("& pIdObra &","& pIdAreaDestino &","& pIdDetalleDestino &","& pPeriodo &","& pImportePesos &","& pImporteDolares &","& pTipoCambio &",'"& session("USUARIO") &"',"& session("MmtoSistema") &")"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs2, "EXEC", strSQL)
	end if

End function
'---------------------------------------------------------------------------------------------------------------------
Function actualizarDetallePartidaPresupuestaria(pIdObra, pIdAreaOrigen, pIdDetalleOrigen, pIdAreaDestino, pIdDetalleDestino, pImportePesos, pImporteDolares, pTipoCambio)
    Dim strSQL,saldoPesos,saldoDolares,rsUpd
    
    'modifico los importes del budget origen solo si es una REASIGNACION (se quita el importe de la partida origen)
	if ((Cdbl(pIdAreaOrigen) <> 0)and(Cdbl(pIdDetalleOrigen) <> 0)) then
        strSQL = "UPDATE TBLBUDGETOBRAS SET PSBUDGET =  PSBUDGET - " & pImportePesos & ", DLBUDGET = DLBUDGET - " & pImporteDolares &_
	             " WHERE idobra = "& pIdObra &" and idarea = " & pIdAreaOrigen & " AND iddetalle = " & pIdDetalleOrigen
        Call executeQueryDb(DBSITE_SQL_INTRA, rsUpd, "UPDATE", strSQL)
    end if
    'modifico los importe del budget del destino, puede ser una REASIGNACION (se agrega a la nueva partida el importe) o un AJUSTE (puede ser un importe que agrege o quite a una partida sin quitarle a otra)
    strSQL = "UPDATE TBLBUDGETOBRAS SET PSBUDGET =  PSBUDGET + " & pImportePesos & ", DLBUDGET = DLBUDGET + " & pImporteDolares &_
	         " WHERE idobra = "& pIdObra &" and idarea = " & pIdAreaDestino & " AND iddetalle = " & pIdDetalleDestino
    Call executeQueryDb(DBSITE_SQL_INTRA, rsUpd, "UPDATE", strSQL)
    
	
End function
'---------------------------------------------------------------------------------------------------------------------
Function actualizarImportePartidaPresupuestaria(pRsRea)
    Dim tipoFormularioObra

    'Modificar los valores de la tabla TBLBUDGETOBRAS actualizando el nuevo importe
    Call actualizarDetallePartidaPresupuestaria(pRsRea("IDOBRA"),pRsRea("IDAREAORIGEN"),pRsRea("IDDETALLEORIGEN"),pRsRea("IDAREADESTINO"),pRsRea("IDDETALLEDESTINO"),pRsRea("IMPORTEPESOS"),pRsRea("IMPORTEDOLARES"),pRsRea("TIPOCAMBIO"))
    
    'Solo si la obra es trimestral se agrega en la tabla de detalles trimestrales
    tipoFormularioObra = getFormTypeByIdObra(pRsRea("IDOBRA"))
    if ( Cdbl(tipoFormularioObra) = OBRA_FORM_TRIM) then Call actualizarDetalleTrimestralPartidaPresupuestaria(pRsRea("IDOBRA"),pRsRea("IDAREAORIGEN"),pRsRea("IDDETALLEORIGEN"),pRsRea("IDAREADESTINO"),pRsRea("IDDETALLEDESTINO"),pRsRea("IMPORTEPESOS"),pRsRea("IMPORTEDOLARES"),pRsRea("TIPOCAMBIO"),pRsRea("PERIODO"))
        

End Function
'---------------------------------------------------------------------------------------------------------------------
Function registrarFirma(pIdReasignacion,pIdObra)
	Dim rs, esAsignacion
    
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACION_UPD_FIRMAR_CBTE", pIdReasignacion &"||"& Session("Usuario"))

    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACIONFIRMAS_INS", pIdReasignacion &"||"& Session("Usuario") &"||"& Session("MmtoDato"))    
    
    
    
    strSQL = "SELECT * FROM TBLBUDGETREASIGNACION WHERE IDREASIGNACION = "& pIdReasignacion 
    Call executeQueryDb(DBSITE_SQL_INTRA, rsRea, "OPEN", strSQL)
    if (not rsRea.Eof) then
        'Verifico si la reasignacion/ajuste tiene un estado finalizado para poder actualizar los importes en las demas tablas de las partidas
        if (Cdbl(rsRea("ESTADO")) = BUDGET_REASIGNACION_FINALIZADA) then 
            Call actualizarImportePartidaPresupuestaria(rsRea)
        else
            esAsignacion = false
            if (Cdbl(rsRea("IDAREAORIGEN")) <> 0) then esAsignacion = true
            Call sendMailNextSignatory(pIdReasignacion, pIdObra, esAsignacion)
        end if
    end if
    registrarFirma = RESPUESTA_OK
End Function
'---------------------------------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds	
	ret = false	
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"		
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true
		else
			gCdUsuario = ""
		end if
	end if		
	leerRegistroFirmas = ret
End Function
'---------------------------------------------------------------------------------------------------------------------
'Se encarga de enviar mail al proximo firmante que tiene el lote, en caso de ser el ultimo envia a los primero autorizantes informando que se aplico
Function sendMailNextSignatory(pIdReasignacion, pIdObra, pEsAsignacion)
    Dim rs, mailMsg, mailOrigen, mailDestino, mailAsunto
    'El sotre procedure devuelve el/los usuarios que deberan ser notificados por la alerta de mail de provisiones
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLBUDGETREASIGNACIONFIRMAS_GET_NEXT_SIGNATORY_BY_IDREASIGNACION", pIdReasignacion)
    if (not rs.Eof) then
        mailOrigen = getTaskMailList(TASK_COM_AUTH_REASSIGNING_BUDGET, MAIL_TASK_SENDER)
        mailAsunto = "Sistema Compras - Alerta de firma"
        mailMsg = "Tiene pendiente para autorizar la siguiente " 
        if (pEsAsignacion) then
            mailMsg = mailMsg  & "Reasignación de Partida Presupuestaria: "& vbcrlf 
        else
            mailMsg = mailMsg  & "Ajuste de Partida Presupuestaria: "& vbcrlf 
        end if
        mailMsg = mailMsg & "Numero: "& pIdReasignacion & vbcrlf
        mailMsg = mailMsg & "Obra: "& getDescripcionObra(pIdObra) & vbcrlf
        while(not rs.Eof)
            mailDestino = getUserMail(Trim(rs("CDUSUARIO")))
            Call GP_ENVIAR_MAIL(mailAsunto, mailMsg, mailOrigen, mailDestino)
            rs.MoveNext()
        wend
    end if
End Function
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim idReasignacion,gCdUsuario,gsecuencia,idObra

idReasignacion = GF_PARAMETROS7("idReasignacion", 0, 6)
idObra = GF_PARAMETROS7("idObra", 0, 6)


Call GP_CONFIGURARMOMENTOS()

respuesta = LLAVE_NO_CORRESPONDE
if (CDbl(idReasignacion) <> 0) then
	if (leerRegistroFirmas()) then respuesta = registrarFirma(idReasignacion,idObra)
else
	respuesta = CODIGO_VACIO
end if	
if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(respuesta)

%>
