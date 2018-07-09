<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosAS400.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->
<!-- #include file="Includes/procedimientosRetenciones.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp" -->
<!-- include file="Includes/procedimientosLog.asp" -->
<!-- #include file="Includes/procedimientosTraducir.asp" -->
<!-- #include file="Includes/procedimientosfechas.asp" -->
<!-- #include file="Includes/procedimientosvalidacion.asp" -->
<!-- #include file="Includes/cor-IncludePC.asp"-->
<!-- #Include File="Includes/ExternalFunctions.ASP" -->
<!-- #include file="Includes/procedimientosEmpresas.asp"-->
<%
'***************************************************************************
sub marcarDuplicado(p_oPDF, p_RetNro)
    dim rs, conn, strSQL

    strSQL = "select * from controlEnvioRetencion where RetNro=" & p_RetNro & " and MrcEnvioMail='V'"
    call GF_BD_AS400(rs, conn, "OPEN", strSQl)
    if not rs.eof then
        Call GF_setFont(Gbl_oPDF,"ARIAL", 16,8)
        Call GF_writeTextAlign(p_oPDF,290, 770, GF_Traducir("COPIA DEL DUPLICADO"),300, 2)
        Call GF_writeTextAlign(p_oPDF,290, 790, GF_Traducir("QUE OBRA EN NUESTRO PODER"),300, 2)
    end if
    call GF_BD_AS400(rs, conn, "CLOSE", "")
end sub
'***************************************************************************
function getFileName(p_tipo, p_nro)
'EAB Agrege la X
    getFileName = getTituloTipo(p_tipo) & " Nro " & GF_EDIT_CBTE(p_nro) & "Y.pdf"
end function
'***************************************************************************
Dim i,k,h,p_maxRet, p_vecRetenciones(), inLimite, vecMailsRetenciones(1)', hayError

on Error Resume next
'Tomo los parametros
'Nro de retenciones a generar.
p_maxRet = GF_Parametros7("P_MAXRET",0,6)
'Tomo las retenciones a generar.
Redim Preserve p_vecRetenciones(p_maxret,3)
response.write "TOTAL RETENCIONES A ENVIAR= " & (p_maxRet + 1) & "<br>"
i = 0
For h = 0 to p_maxRet
    if ucase(GF_Parametros7("CHK" & i,"",6))="ON" then
        p_vecRetenciones(i,0) = GF_Parametros7("P_RET" & i & "Nro","",6)
        p_vecRetenciones(i,1) = GF_Parametros7("P_RET" & i & "Tipo","",6)
        p_vecRetenciones(i,2) = GF_Parametros7("P_RET" & i & "Fecha","",6)
        p_vecRetenciones(i,3) = GF_Parametros7("P_RET" & i & "KCPRO","",6)
        i = i + 1
    end if
Next

'Se generan las retenciones.
set oFS = Server.CreateObject("Scripting.FileSystemObject")

%>
<br><br>
<table border=1>
    <tr>
        <td>Nro</td>
        <td>Tipo</td>
        <td>Fecha</td>
        <td>KCPRO</td>
        <td>Resultado</td>
    </tr>
<%For h = 0 to i - 1%>
    <tr>
        <td><%=p_vecRetenciones(h,0)%></td>
        <td><%=p_vecRetenciones(h,1)%></td>
        <td><%=p_vecRetenciones(h,2)%></td>
        <td><%=p_vecRetenciones(h,3)%></td>
    <%
    call obtenerMailRetenciones(p_vecRetenciones(h,3), vecMailsRetenciones)
    vecMailsRetenciones(0) = "BacariniE@toepfer.com"
    'response.write vecMailsRetenciones(0) & "6"
    if ((vecMailsRetenciones(0)<>"") or (vecMailsRetenciones(1)<>"")) then
        strTituloTipo = getTituloTipo(p_vecRetenciones(h,1))
        if oFS.fileExists(Server.MapPath("temp\" & getFileName(p_vecRetenciones(h,1), p_vecRetenciones(h,0)))) then
            call oFS.deleteFile(Server.MapPath("temp\" & getFileName(p_vecRetenciones(h,1), p_vecRetenciones(h,0))), true)
        end if
        if err.Number = 0 then
            'Si no hubo error genero el pdf y lo mando
            Set Gbl_oPDF = GF_createPDF(Server.MapPath("temp\" & getFileName(p_vecRetenciones(h,1), p_vecRetenciones(h,0))))
            Select Case p_vecRetenciones(h,1)
                   Case "C": ret = GF_Retencion_C(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "E": ret = GF_Retencion_E(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "B": ret = GF_Retencion_B(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "H": ret = GF_Retencion_H(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "D": ret = GF_Retencion_D(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "G": ret = GF_Retencion_G(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "J": ret = GF_Retencion_J(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "I": ret = GF_Retencion_I(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "K": ret = GF_Retencion_K(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "L": ret = GF_Retencion_L(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "M": ret = GF_Retencion_M(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
                   Case "P": ret = GF_Retencion_P(p_vecRetenciones(h,0),p_vecRetenciones(h,1),p_vecRetenciones(h,2))
            End Select
            call marcarDuplicado(Gbl_oPDF, p_vecRetenciones(h,0))
            Call GF_closePDF(Gbl_oPDF)
            call enviarMailsRetenciones(p_vecRetenciones(h,3),p_vecRetenciones(h,1),p_vecRetenciones(h,0))%>
        <%else%>
            <td>Error al intentar borrar e archivo de la retencion</td>
            <%'Si hubo error lo logueo
            strSQL = "update controlEnvioRetencion set MrcEnvioMail = 'X' where RetNro=" & p_vecRetenciones(h,0)
            call GF_BD_AS400("", conn, "EXEC", strSQL)
            call GF_LogError(err, "Error al intentar borrar el archivo de las retenciones." & chr(13) & chr(10) & "cor-RetGenerator linea 34")
            call Err.clear()
        end if
    else%>
        <td>No se pudo enviar la retencion por no tener seteadas las direcciones.</td>
        <%strSQL = "update controlEnvioRetencion set MrcEnvioMail = 'N' where RetNro=" & p_vecRetenciones(h,0)
        call GF_BD_AS400("", conn, "EXEC", strSQL)
    end if%>
    </tr>
<%Next
call GF_BD_AS400("", conn, "CLOSE", "")%>
</table>