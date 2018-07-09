<!-- #include file="Includes/procedimientosUnificador.asp"-->
<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosPDF.asp"-->
<!-- #include file="Includes/procedimientosTraducir.asp" -->
<!-- #include file="Includes/procedimientos.asp"-->
<!-- #include file="Includes/procedimientosBoletos.asp" -->

<%
'Call ProcedimientoControl("CTOGEN")

'Constantes de intentos de creacion de archivo.
Const MAX_TRY = 20

'**************************************************************************************************************************
'Obtiene el nombre del archivo a generar.
Function getFilename(cant, pVecBoletos)
		Dim ret 
		
		if (cant > 1) then
			Randomize()
			ret = session.SessionId & "-" & Int(100 * Rnd()) & ".pdf"
		else
			ret = Replace(GF_EDIT_CONTRATO(pVecBoletos(cant,0), pVecBoletos(cant,1), pVecBoletos(cant,2), pVecBoletos(cant,3), pVecBoletos(cant,4)) & ".pdf", "/", "-")
		end if
		getFilename = ret
End Function
'**************************************************************************************************************************
Dim i, k, p_maxBol, p_vecBoletos(), hayErr, inLimite, bol

'On Error Resume Next

'Tomo los parametros

'Nro de contratos a generar.
p_maxBol = GF_Parametros7("P_MAXCTO",0,6)
'Tomo los contratos a generar.
Redim Preserve p_vecBoletos(p_maxBol,10)
k=-1
For i = 0 to p_maxBol
    if (GF_Parametros7("CHK" & i,"",6) = "on") then
       'El contrato debe ser generada.
       k = k + 1
       p_vecBoletos(k,0) = GF_Parametros7("P_CTO" & i & "Prod","",6)
       p_vecBoletos(k,1) = GF_Parametros7("P_CTO" & i & "Suc","",6)
       p_vecBoletos(k,2) = GF_Parametros7("P_CTO" & i & "Oper","",6)
       p_vecBoletos(k,3) = GF_Parametros7("P_CTO" & i & "Sec","",6)
       p_vecBoletos(k,4) = GF_Parametros7("P_CTO" & i & "Cos","",6)       
   end if
Next
if (k > -1) then
    'Se generan los boletos.
    'Se abre el archivo PDF.
    hayErr = True
    intLimite = 0 'Numero maximo de intentos para crear el archivo

    while (hayErr) and (intLimite < MAX_TRY) and (k > -1)
          filename = getFilename(k, p_vecBoletos)
          Set Gbl_oPDF = GF_createPDF(Server.MapPath("temp\" & filename))
          hayErr = False
          if (Err.Number <> 0) then hayErr = True
          Err.Clear
          intLimite = intLimite + 1
    wend
    if (intLimite = MAX_TRY) then
       response.redirect("MGMSG.asp?P_MSG=La informacion solicitada no puede ser generada en este momento. Intente mas tarde. Gracias")
    end if
    bol = False

    For i = 0 to k		
		bol = GF_GenerarBoleto(Gbl_oPDF, p_vecBoletos(i,0), p_vecBoletos(i,1), p_vecBoletos(i,2), p_vecBoletos(i,3), p_vecBoletos(i,4))        
        if i < k then call GF_newPage(Gbl_oPDF)
    Next
    Call GF_closePDF(Gbl_oPDF)
    if (not bol) then response.redirect("MGMSG.asp?P_MSG=Este boleto aun no se encuentra disponible para impresion.")
    %>
    <HTML>
    <HEAD>
    <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
    </script>
    </HEAD>
    <BODY>
       <% response.redirect INTRANET_FILE_PATH & filename %>
    </BODY>
    </HTML>
<% else %>
<script language="javascript">
        window.close();
</script>
<% end if %>