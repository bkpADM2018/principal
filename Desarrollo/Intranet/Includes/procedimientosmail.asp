<%
'/* TIPOS DE MAILS */
Const MAIL_CUPO = "C"
Const MAIL_RETENCION = "R"
Const MAIL_BOLETO = "B"
Const MAIL_CONFIRMACION = "T"

'/* Constantes para manejo de listas de mail de tareas */
Const MAIL_TASK_INFO_LIST = "INFO"  'Lista default de informaci�n de la tarea.
Const MAIL_TASK_ERROR_LIST = "ERROR"
Const MAIL_TASK_SENDER = "SENDER"

Const MAIL_TYPE_TEXT = 0
Const MAIL_TYPE_HTML = 1

Dim mail_config_Type

mail_config_Type = MAIL_TYPE_TEXT
'******************************************************************************************
'Autor: Javier Scalisi
'Fecha: 16/09/2016
'Este profcedimiento env�a un mail a la lista especificada de la tarea especificada.
Function SendMail(pIdTask, pCdList, subject, bodyText, PathAttachment)    
    
    if (pCdList = "") then pCdList = MAIL_TASK_INFO_LIST
    
    auxOrigen = getTaskMailList(pIdTask, MAIL_TASK_SENDER)
    auxDestino = getTaskMailList(pIdTask, pCdList)         
    
'   Envia Mail
    Call GP_ENVIAR_MAIL_ATTACHMENT(subject, bodyText, auxOrigen, auxDestino, PathAttachment)
    
End Function    
'******************************************************************************************
'Autor: Javier Scalisi
'Fecha: 12/09/2016
'Este profcedimiento permite obtener la lista de mails para las tarea indicada.
'Levanta todas las listas de manera de tenerlas listas en caso de que se requieran nuevamente.
Function getTaskMailList(pIdTask, pCdList)    
    Dim theList, conn          
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_GET_BY_PARAMETERS", pIdTask & "||" & pCdList)
    while (not rs.eof)     
        theList = theList & rs("EMAIL") &  ";"         
        rs.MoveNext()
    wend
    if (Len(theList) > 1) then theList = Left(theList, Len(theList)-1)    
    getTaskMailList = theList
    
End function
'******************************************************************************************
'Autor: Javier Scalisi
'Fecha: 12/09/2016
'Este profcedimiento permite grabar la lista de mails para las tarea indicada.
'La lista debe contener las direcciones separadas por punto y coma (;)
Function saveTaskMailList(pIdTask, pCdList, theList)        
    Dim arrList, i, strSQL
    
    'Primero borro la lista a grabar
    if (Len(Trim(theList)) > 0) then
        strSQL="Delete from TBLTAREAMAILS where IDTAREA=" & pIdTask & " and CODIGOLISTA='" & pCdList & "'"        
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)                
        arrList = Split(theList, ";")
        For each i in arrList
            Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_INS", pIdTask & "||" & pCdList & "||" & Left(Trim(i), 100))   
        Next        
    end if
End function

'******************************************************************************************
'Autor: Javier a. Scalisi
'Fecha: 27/08/2003
'Este procedimiento envia un mail utilizando CDONTS
Function GP_ENVIAR_MAIL(Asunto,Descripcion,Origen,Destino)
   GP_ENVIAR_MAIL = GP_ENVIAR_MAIL_ATTACHMENT(Asunto, Descripcion, Origen, Destino, "")
End Function
'******************************************************************************************
Function GP_ENVIAR_MAIL_ATTACHMENT(Asunto, Descripcion, Origen, Destino, PathAttachment)
    Dim Mail,i, ret 
    ret = false
    if (Origen <> "") and (Destino <> "") then
        ret = true
      set Mail = Server.CreateObject("CDO.Message")
      Set objConfig = Server.CreateObject("CDO.Configuration")
      Set Fields = objConfig.Fields
      With Fields
		.Item(cdoSmtpUseSSL)			= CDO_SMTP_USE_SSL
    	.Item(cdoSendUsingMethod)       = cdoSendUsingPort
    	.Item(cdoSMTPServer)            = CDO_SMTP_SERVER
    	.Item(cdoSMTPServerPort)        = CDO_SMTP_PORT
    	.Item(cdoSMTPConnectionTimeout) = 10
    	.Item(cdoSMTPAuthenticate)      = CDO_SMTP_AUTHENTICATE
    	if (CDO_SMTP_AUTHENTICATE = cdoAdvance) then
    		.Item(cdoSendUserName)          = CDO_SEND_USER_NAME	    
   			.Item(cdoSendPassword)          = CDO_SEND_PASSWORD
    	end if
    	.Update
    End With
    Set Mail.Configuration = objConfig

      'Preparar el mail.
      Mail.From=Origen
      Mail.To=Destino
      Mail.Subject=Asunto      
      if (CInt(mail_config_Type) = MAIL_TYPE_HTML) then
            Mail.HtmlBody= Descripcion      
      else  
            Mail.TextBody= Descripcion      
      end if                   
      'Mail.Bcc= "scalisij@toepfer.com"
      if PathAttachment <> "" then 
			strAtt = Split(PathAttachment,";")
			For i = LBound(strAtt)  to UBound(strAtt)
				Mail.AddAttachment strAtt(i)
			next
		end if
      'Se envia el mail.
      On Error Resume Next
      Mail.Send
      'response.write "#" & Err.Number & "#"
      if Err.Number <> 0 then 
        response.write "Error al intentar enviar el mail: " & Err.Description & "(Err. Number:" & Err.Number & ")"
        ret = false
      end if        
	  Set Mail = nothing
   end if
   GP_ENVIAR_MAIL_ATTACHMENT = ret
End Function
'******************************************************************************************
Function obtenerMailCuposProveedor(ByVal p_ORKC, ByRef p_vecMail)
    obtenerMailCuposProveedor = obtenerMailTipo(p_ORKC, MAIL_CUPO, p_vecMail)
end function
'*****************************************************************************************
function obtenerMailRetenciones(ByVal p_ORKC, ByRef p_vecMail)
    obtenerMailRetenciones = obtenerMailTipo(p_ORKC, MAIL_RETENCION, p_vecMail)
end function
'*****************************************************************************************
function obtenerMailBoletos(ByVal p_ORKC, ByRef p_vecMail)
    obtenerMailBoletos = obtenerMailTipo(p_ORKC, MAIL_BOLETO, p_vecMail)
end function
'*****************************************************************************************
function obtenerMailConfirmaciones(ByVal p_ORKC, ByRef p_vecMail)
    obtenerMailConfirmaciones = obtenerMailTipo(p_ORKC, MAIL_CONFIRMACION, p_vecMail)
end function
'*****************************************************************************************
function obtenerMailTipo(ByVal p_ORKC, p_tipo, ByRef p_vecMail)
    dim strSQL, conn, rs, i

    strSQL = "Select Mail from toepferdb.tblmailsmercaderias where Tipo='" & p_tipo & "' and idproveedor=" & p_ORKC
    'Response.Write strSQL
    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
    i=0
    while (not rs.eof)
		if len(rs("Mail"))>0 then
			p_vecMail(i) = rs("Mail")
			i=i+1
		end if
        rs.MoveNext
    wend
    call GF_BD_AS400_2(rs, conn, "CLOSE", "")
    obtenerMailTipo = i
end function
'*****************************************************************************************
Sub guardarMailTipo(ByVal p_ORKC, p_tipo, orden, ByRef p_Mail)
    Dim vec(10)
    Dim strSQL, conn, rs

    strSQL = "Select Mail from toepferdb.tblmailsmercaderias where Tipo='" & p_tipo & "' and Orden = " & orden & " and idproveedor=" & p_ORKC
    call executeQuery(rs, "OPEN", strSQL)
    if (not rs.eof) then
        'El orden Existe, se actualiza
        strSQL = "Update toepferdb.tblmailsmercaderias Set Mail='" & p_Mail & "' where Tipo='" & p_tipo & "' and Orden = " & orden & " and idproveedor=" & p_ORKC
        call executeQuery(rs, "EXEC", strSQL)
    else
        if (p_Mail <> "") then
            'No existe se inserta
            strSQL = "Insert into toepferdb.tblmailsmercaderias values(" & p_ORKC & ", '" & p_tipo & "', " & orden & ", '" & p_Mail & "')"
            call executeQuery(rs, "EXEC", strSQL)
        end if
    end if
End Sub
'*****************************************************************************************
function ObtenerMailUsuario(p_user)
    dim my_valor

    my_valor = GF_DT1("READ","SGEMAIL", session("MmtoSistema"), session("MmtoSistema"),"SG",p_user)
    if my_valor <> "?" then
        ObtenerMailUsuario = my_valor
    else
        ObtenerMailUsuario = ""
    end if
end function
'*****************************************************************************************
Function saveMailTipo(pIdProveedor, pTipo, pMail)
	Dim strSQL,lastOrden
	strSQL = "SELECT MAX(ORDEN) AS LAST FROM TOEPFERDB.TBLMAILSMERCADERIAS WHERE IDPROVEEDOR = "&pIdProveedor& " AND TRIM(TIPO) = '"&pTipo&"'"
	Call executeQuery(rs, "OPEN", strSQL)
	lastOrden = 1
	if (not rs.Eof) then
		if (not isNull(rs("LAST"))) then lastOrden = rs("LAST") + 1
	end if
	Call guardarMailTipo(pIdProveedor, pTipo, lastOrden, Trim(pMail))
End Function
'*****************************************************************************************
Function prepareForMailto(pBody)
	Dim ret
	
	ret = Replace(pBody, " ", "%20")
	ret = Replace(ret, chr(10), "")
	ret = Replace(ret, chr(13), "%0D%0A")
	prepareForMailto = ret
	
End Function
%>