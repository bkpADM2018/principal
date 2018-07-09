<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosAS400.asp"-->
<!-- #include file="Includes/procedimientosRetenciones.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp" -->
<!-- #include file="Includes/procedimientosfechas.asp" -->
<!-- #include file="Includes/procedimientosvalidacion.asp" -->
<!-- #include file="Includes/cor-IncludePC.asp"-->
<!-- #include file="Includes/procedimientosFormato.asp" -->
<!-- #Include File="Includes/ExternalFunctions.ASP" -->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientos.asp"-->
<%
'Call ProcedimientoControl("RETGEN")
'Constantes de intentos de creacion de archivo.
Const MAX_TRY = 20

'Obtiene el nombre del archivo a generar.
Function getFilename()
         Randomize()
         getFilename = session.SessionId & "-" & Int(100 * Rnd()) & ".pdf"
End Function

%>
<%
Dim i,k,h,p_maxRet, p_vecRetenciones(), hayErr, inLimite

'On Error Resume Next

'Tomo los parametros

'Nro de retenciones a generar.
p_maxRet = GF_Parametros7("P_MAXRET",0,6)
'Tomo las retenciones a generar.
Redim Preserve p_vecRetenciones(p_maxret,2)
k=-1
For i = 0 to p_maxRet
    if (GF_Parametros7("CHK" & i,"",6) = "on") then
       'Si la retencion debe ser generada, guardo sus datos en el vector.
       k = k + 1
       p_vecRetenciones(k,0) = GF_Parametros7("P_RET" & i & "Nro","",6)
       p_vecRetenciones(k,1) = GF_Parametros7("P_RET" & i & "Tipo","",6)
       p_vecRetenciones(k,2) = GF_Parametros7("P_RET" & i & "Fecha","",6)
    end if
Next
if (k > -1) then
'Se generan las retenciones.
'Se abre el archivo PDF.
hayErr = True
intLimite = 0 'Numero maximo de intentos para crear el archivo
while (hayErr) and (intLimite < MAX_TRY) and (k > -1)
      filename = getFilename()
      Set Gbl_oPDF = GF_createPDF(Server.MapPath("temp\" & filename))
      call GF_setPDFMode(PDF_FILE_MODE)
      hayErr = False
      if (Err.Number <> 0) then hayErr = True
      Err.Clear
      intLimite = intLimite + 1
wend
if (intLimite = MAX_TRY) then
   response.redirect("MGMSG.asp?P_MSG=La informacion solicitada no puede se generada en este momento. Intente mas tarde. Gracias")
end if
ret = False

For h = 0 to k
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
    if (h < k) then Call GF_newPage(Gbl_oPDF)
Next
Call GF_closePDF(Gbl_oPDF)
if (not ret) then response.redirect("MGMSG.asp?P_MSG=Esta retencion aun no se encuentra disponible para impresion.")
%>
<HTML>
<HEAD>
<link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</script>
</HEAD>
<BODY>
      <%'response.end
      response.redirect INTRANET_FILE_PATH & filename%>
</BODY>
</HTML>
<% else %>
<script language="javascript">
        window.close();
</script>
<% end if %>
