<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="geminiIssuePrint.asp"-->
<%
'---------------------------------------------------------------------------------------------
Function registarFirma(pIdTarea)
	Dim strSQL, rs, k, sp_ret 
    registarFirma = "" 
    'Obtengo el estado proximo que tendrá la tarea
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLESTADOSTRANSICION_GET_BY_PARAMETERS", SEC_SYS_GEMINI &"||"& RES_GEM_TAREAS &"||"& getIssueStatus(pIdTarea) &"||"& EVENTO_FIRMA &"||||||$$respuesta")
    if (not rs.Eof) then
        'Actualizo el estado de la Tarea
        g_proximoEstado = rs("ESTADOPROXIMO")
        Call updateIssueStatus(pIdTarea, g_proximoEstado, "")
        'Grabo la firma del usuario
        Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSYSFIRMAS_INS", pIdTarea & "||" & Session("Usuario") & "||" & session("MmtoDato"))	
        registarFirma = RESPUESTA_OK
    end if
End Function
'---------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs,  km, ds
	leerRegistroFirmas = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			if (UCase(rs("CDUSUARIO")) = session("Usuario")) then leerRegistroFirmas = true
		end if
	end if
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Function sendMailNextSignerIssue(p_IdTarea)
	Dim emailTo, emailFrom, asunto, msg, rs
    if ((CInt(g_proximoEstado) = ISSUE_STATUS_CLOSE) or (CInt(g_proximoEstado) = ISSUE_STATUS_HD_CLOSE))then
        'Registro el cierre en Gemini
        Set myFs = Server.CreateObject("Scripting.FileSystemObject")
        emailTo = SENDER_PEDIDOS_IT
        emailFrom = getUserMail(session("Usuario"))        
        asunto = "["& getGeminiTaskCode(p_IdTarea) &"] Tarea Finalizada"
        msg = "La tarea ha sido publicada"
        pathPDF = crearReporte(p_IdTarea)        
        if (myFs.FileExists(pathPDF)) then
            Call GP_ENVIAR_MAIL_ATTACHMENT(asunto, msg, emailFrom, emailTo, pathPDF)                        
		    myFs.DeleteFile(pathPDF)
        end if                    
    else
        'Busco el usuario que firmará el nuevo estado, para enviar el mail       
        Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLESTADOSTRANSICION_GET_BY_PARAMETERS", SEC_SYS_GEMINI &"||"& RES_GEM_TAREAS &"||"& g_proximoEstado &"||"& EVENTO_FIRMA &"||||||")
        if (not rs.Eof) then
            emailTo = getMailNextSignerIssue(CInt(rs("DATOAUXILIAR")), p_IdTarea)
            emailFrom = SENDER_PEDIDOS_IT
            asunto = "["& getGeminiTaskCode(p_IdTarea) &"] - Pedido Realizado a Sistemas - ALERTA FIRMA PENDIENTE "
            msg = "Usted tiene disponible para firmar la Tarea: "& getGeminiTaskCode(p_IdTarea) & "." & vbcrlf & "Por Favor ingrese a la sección de autorizaciones de la Intranet para autorizar la tarea." & vbcrlf & "Muchas Gracias" & vbcrlf & vbcrlf & "El Equipo de Sistemas"            
            For i = 1 To Ubound(emailTo)
                Call GP_ENVIAR_MAIL(asunto, msg , emailFrom, emailTo(i))
            Next            
        end if
    end if
    
End function
'--------------------------------------------------------------------------------------------------------------------------------
'Devuelve un array de mail de un determinado rol o del solicitante de la tarea
Function getMailNextSignerIssue(p_Rol, p_IdTarea)
    Dim rtrn()
    redim rtrn(0)
    if (p_Rol = 0) then
        'El proximo firmante se trata de un Solicitante, se lo busca en la tabla de Tareas del Gemini
        Redim Preserve rtrn(Ubound(rtrn) + 1)
        rtrn(ubound(rtrn)) = getIssueMailApplicant(p_IdTarea)
    else
        'El proximo firmante se trata de un Rol, tester o publicador (puede haber varias personas con ese rol)
        Call executeProcedureDb(DBSITE_SQL_INTRA, rsRol, "TBLROLESUSUARIOS_GET_BY_IDROL_IDSISTEMA", p_Rol &"||"& SEC_SYS_GEMINI )
        while (not rsRol.Eof)
            Redim Preserve rtrn(Ubound(rtrn) + 1)
            rtrn(ubound(rtrn)) = getUserMail(rsRol("CDUSUARIO"))
            rsRol.MoveNext()
        wend
    end if
    getMailNextSignerIssue = rtrn
End Function
'**********************************************************	
'*****************  INICIO DE PAGINA  *********************
'**********************************************************
Dim ret,idTarea,respuesta,g_proximoEstado

idTarea = GF_Parametros7("idTarea",0,6)

Call GP_CONFIGURARMOMENTOS()

ret = LLAVE_NO_CORRESPONDE
if (idTarea <> 0) then
    if (leerRegistroFirmas()) then
    	ret = registarFirma(idTarea)
        Call sendMailNextSignerIssue(idTarea)
    end if
else
    respuesta = CODIGO_VACIO
end if

if (respuesta <> RESPUESTA_OK) then respuesta = respuesta & "-" & errMessage(respuesta)
Call HK_sendResponse(ret)


%>
