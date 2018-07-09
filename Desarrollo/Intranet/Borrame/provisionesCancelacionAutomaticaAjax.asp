<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'------------------------------------------------------------------------------------------------------------------------------
Function borrarCancelacionAutomatica(p_nroLote, p_fechaLote)
    Call executeSP(rsPro, "EJIFL.TBLPROVISIONESCANE_DEL", p_nroLote &"||"& p_fechaLote)
    Call executeSP(rsPro, "EJIFL.TBLPROVISIONESFIRMAS_DEL", p_nroLote &"||"& p_fechaLote)
End Function
'------------------------------------------------------------------------------------------------------------------------------
'Actualiza la marca de inclusion de cancelacion automatica
Function actualizarMarcaInclusionCancelacionAutomatica(p_NroLote, p_FechaLote, p_Moviemiento, p_Secuencia)
    Call executeSP(rs, "EJIFL.TBLPROVISIONESCANE_UPD_MARCAINCLUSION_BY_PARAMETERS", p_NroLote &"||"& p_FechaLote &"||"& p_Secuencia &"||"& Trim(p_Moviemiento))
End Function
'------------------------------------------------------------------------------------------------------------------------------
'Al rechazar el lote vuelve a estado generado (inicial) y se borran las firmas que tenia
Function retrasarEstadoCancelacionAutomatica(p_NroLote, p_FechaLote)
    Call executeSP(rs, "EJIFL.TBLPROVISIONESCANE_UPD_ESTADO_BY_PARAMETERS", p_NroLote &"||"& p_FechaLote &"||"& PROVISCIONES_ESTADO_GENERADO )
    Call sendMailCancelacion(p_NroLote, p_FechaLote)
    Call executeSP(rs, "EJIFL.TBLPROVISIONESFIRMAS_DEL", p_NroLote &"||"& p_FechaLote)
End Function
'------------------------------------------------------------------------------------------------------------------------------
'Envia mail a los autorizantes que participaron en la provision indicando que se cancelo
Function sendMailCancelacion(pNroLote ,pFechaLote)
    Dim rs, mailMsg, mailOrigen, mailDestino, mailAsunto
    'El sotre procedure devuelve los usuarios que deberan ser notificados por la alerta de mail de provisiones
    Call executeSP(rs, "EJIFL.TBLPROVISIONESFIRMAS_GET_NEXT_SIGNATORY_BY_PARAMETERS", pNroLote &"||"& pFechaLote &"||"& Session("Usuario"))
    if (not rs.Eof) then
        mailOrigen = getTaskMailList(TASK_EJE_PROVISIONS, MAIL_TASK_SENDER)
        mailAsunto = "Sistema Provisiones - Alerta de firma"
        mailMsg = "Se modifico la provisión, para su aprobacion deberán autorizar todos los usuarios nuevamente. "& vbcrlf
        mailMsg = mailMsg & "Nro.Lote: "& pNroLote & vbcrlf
        mailMsg = mailMsg & "Fecha Lote: "& GF_FN2DTE(pFechaLote) & vbcrlf
        mailMsg = mailMsg & "Usuario responsable: "& getUserDescription(Session("Usuario")) & vbcrlf
        while(not rs.Eof)
            mailDestino = getUserMail(Trim(rs("CDUSUARIO")))
            Call GP_ENVIAR_MAIL(mailAsunto, mailMsg, mailOrigen, mailDestino)
            rs.MoveNext()
        wend
    end if
End Function
'******************************************************************************************************************************
'***************************************************  INICIO DE LA PAGINA  ****************************************************
'******************************************************************************************************************************
dim nroLote, fechaLote, rsPro,accion, indice,estado

nroLote   = GF_PARAMETROS7("nroLote",0,6)
fechaLote = GF_PARAMETROS7("fechaLote",0,6)
accion    = GF_PARAMETROS7("accion","",6)

SELECT CASE accion
    CASE ACCION_BORRAR
        Call borrarCancelacionAutomatica(nroLote, fechaLote)
    CASE ACCION_GRABAR
       'Primero obtengo el maximo de indices de filas que tiene el lote
        indice = GF_PARAMETROS7("indice",0,6)
        for i = 0 to Cint(indice)-1
            secuencia = GF_PARAMETROS7("secuencia_"& i,0,6)
            marcaInclusion = GF_PARAMETROS7("marcaInclusion_"& i,"",6)
            Call actualizarMarcaInclusionCancelacionAutomatica(nroLote, fechaLote, marcaInclusion, secuencia)
        next
        'Si cuando modifica las marcas de inclucion el estado es diferente a cuando genera el lote (estado generado) se eliminan las firmas y se vuelve para atas el estado
        estado = GF_PARAMETROS7("estado","",6)
        if (estado <> PROVISCIONES_ESTADO_GENERADO) then Call retrasarEstadoCancelacionAutomatica(nroLote, fechaLote)
END SELECT


%>