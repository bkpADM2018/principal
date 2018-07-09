<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<% 
'--------------------------------------------------------------------------------------------------
' Autor: Jonathan Costilla
' Fecha: 15/09/2016
'--------------------------------------------------------------------------------------------------
Function enviarMailCancelacion(motivo)
	Dim strMsg
	
	strMsg = "Se ha cancelado el pedido de cotización: " & pct_cdPedido & vbCrLf & vbCrLf	&_
            "Motivo: " & vbCrLf & motivo
    
	emailDestino = getUserMail(pct_cdSolicitante)	
	emailDestino = obtenerMail(CD_TOEPFER) & ";" & emailDestino
	Call enviarMail("Cancelación de Pedido", strMsg, emailDestino)		
	
End Function
'--------------------------------------------------------------------------------------------------
Function enviarMail(asunto, msg, email)
	Dim emailSender, obras, dsTipoCompra,i, usrAdmin, emailSolicitante
	enviarMail = false
	emailSender = getUserMail(session("Usuario"))
	'email = "scalisij@toepfer.com;" & email
	'Response.Write "(" & email & ")<br>"
	if (email <> "") then
		emailSolicitante = getUserMail(pct_cdSolicitante)
		msg = msg & vbCrLf & vbCrLf
		msg = msg & "Datos del Pedido" & vbCrLf
		msg = msg & String(100, "-") & vbCrLf
		msg = msg & "Codigo asignado.....: " & pct_cdPedido & vbCrLf		
		msg = msg & "Titulo..............: " & pct_tituloPedido & vbCrLf
		'Set obras = obtenerListaObras(pct_idObra, "", "","",OBRA_ACTIVA)		
		Set obras = obtenerDescripcionCompletaDetalle(pct_idObra, pct_idArea, pct_idDetalle)
		if ((not obras.eof) and (CLng(pct_idObra) <> 0)) then
			msg = msg & "Ptda. Presupuestaria: " & obras("CDOBRA") & " - " & obras("DSOBRA") & vbCrLf
			if (pct_idArea <> 0) then			
				msg = msg & "                      ----> " & obras("IDAREA") & " - " & obras("DSAREA") & vbCrLf
				if (pct_idDetalle <> 0) then
					msg = msg & "                            ----> " & obras("IDDETALLE") & " - " & obras("DSDETALLE") & vbCrLf
				end if
			end if			
		end if
		msg = msg & "Solicitante.........: " & pct_cdSolicitante & "-" & pct_dsSolicitante & vbCrLf					
		msg = msg & "Tipo de Pedido......: Pedido de Precios" & vbCrLf		
		msg = msg & "Fecha de Emisión....: " & pct_FechaInicio & vbCrLf		
		msg = msg & "Fecha de Limite.....: " & pct_FechaCierre & vbCrLf
		msg = msg & "Administra..........: " & LICITACIONES_ARGENTINA & vbCrLf
		msg = msg & "Descripcion: " & vbCrLf & pct_dsPedido & vbCrLf
		
		'response.write "<pre>" & msg  & "</pre>"
		'response.end
		Call GP_ENVIAR_MAIL(GF_TRADUCIR("Sistema de Compras Web -" & asunto) & ": " & pct_cdPedido, msg, emailSender, email)
		enviarMail = true
		'Response.Write "Mando!<br>"
	end if	
End function
'*********************************************************************************************'
'********************************	INICIO PAGINA  *******************************************'
'*********************************************************************************************'
DIM idPedido,callFlag,accion,description, motivo

callFlag = false
idPedido = GF_PARAMETROS7("idPedido",0,6)
accion = GF_PARAMETROS7("accion","",6)
motivo = GF_PARAMETROS7("motivo","",6)

Call initHeaderDB(idPedido)
IF accion = ACCION_CANCELAR THEN
    IF motivo <> "" THEN
        Call cancelarPedido(idPedido, motivo)
        Call AnularPIC(idPedido)
        Call enviarMailCancelacion(motivo)
        callFlag = true
    ELSE
        call setError(DESCRIPCION_VACIA)
    END IF
ELSE
    motivo = ""
END IF
%>
<!DOCTYPE html>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Cargar Motivo") %></title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">

        <% if (callFlag) then%>
            window.parent.callback_Cancelar();  
        <%end if%>

    function canelarPedido() {
        var accion = confirm("¿Esta seguro que desea cancelar el pedido?");
        if (accion) {
            document.getElementById("frmSel").submit();
        }
    }
</script>	
</HEAD>
<BODY>
<form name="frmSel" id="frmSel" method="post" action="comprasPCTCancelacionPopUp.asp">
    <div id="errDisplay"><% call showErrors() %></div>
    <table width="472px">
        <thead>
            <tr>
                <td style="background:green;">Ingrese el motivo por el cual desea cancelar el pedido.</td>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>
	                <textarea type="text" name="motivo" cols="77" rows="5" maxLength="4000" id="motivo" style="resize:none;" placeholder="Escriba aqui el motivo de la cancelaci&oacute;n..." ><% =motivo %></textarea>
                </td>
            </tr>
        </tbody>
    </table>
    <br>
    <div style="text-align:center;">
        <input type="button" value="Guardar" onclick="canelarPedido()">
    </div>
    <input type="hidden" id="idPedido" name="idPedido" value="<%=idPedido%>">
    <input type="hidden" id="accion" name="accion" value="<%=ACCION_CANCELAR%>">    
</form>
</BODY>
</HTML>