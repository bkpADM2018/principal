<!--#include file="Includes/procedimientosMG.asp"-->

<%
Dim p_tipo,p_controls,p_ret,params, p_fecha

'Bajo los parametros
p_tipo=GF_Parametros7("P_TIPO","",6)
p_ret=GF_Parametros7("P_RET","",6)
p_fecha=GF_Parametros7("P_FECHA","",6)
'Controlo que los parametros sean validos
if (isEmpty(p_tipo)) or (isEmpty(p_ret)) then
   response.redirect("mgmsg.asp?P_MSG=PARAMETROS INCORRECTOS.")
end if
'Se arma la lista de parametros.
params="?P_TIPO=" & p_tipo & "&P_RET=" & p_ret & "&P_FECHA=" & p_fecha
'Se redirije a la pagina indicada segun retencion.
select case p_tipo
       case "C": response.redirect("RET615_99.asp" & params) 'IVA
       case "E": response.redirect("RET1394_02.asp" & params)'IVA
       case "B": response.redirect("RET830_00.asp" & params) 'Ganancias
       case "H": response.redirect("RETIBBA.asp" & params)
       case "D": response.redirect("RETIBSF.asp" & params)
       case "G": response.redirect("RETRIAS.asp" & params)
       case "J": response.redirect("RETRISM.asp" & params)
       case "I": response.redirect("RETIBCBA.asp" & params)
       case "K": response.redirect("RET4052_95.asp" & params) 'Contrib. Patronales 4052/95
       case "L": response.redirect("RET1784_05.asp" & params) 'Contrib. Patronales 1784/05
       case "M": response.redirect("RET1556_03.asp" & params) 'Contrib. Patronales 1556/03
       case "P": response.redirect("RET1769_04.asp" & params) 'Contrib. Patronales 1769/04
end select
response.redirect("MGMSG.asp?P_MSG=Esta retencion aun no se encuentra disponible para impresion.")
%>

