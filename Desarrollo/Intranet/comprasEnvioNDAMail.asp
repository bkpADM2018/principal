<!--#include file="comprasnotadeaceptacionprint.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
	Dim asunto, origen,destino,mensaje,pathPDF ,errores, fs,idNDA
	
	idPedido = GF_PARAMETROS7("idPedido", 0, 6)
	idCotizacion = GF_PARAMETROS7("idCotizacion", 0, 6)
	idNDA = GF_PARAMETROS7("IdNDA", 0, 6)	
	
	errores =false	
	
	pathPDF = CrearPdf(idPedido,idCotizacion,idNDA,PDF_FILE_MODE)
	if(idPedido > 0) then
		'tiene un pedido asociado, por lo tanto toma para referenciar el codigo del Pedido, entra tambien si tiene un pedido y una cotizacion
		Call initHeader(idPedido)
		asunto  = GF_TRADUCIR("NOTA DE ACEPTACION") & " - REF: " & pct_cdPedido		
		origen = obtenerMail(CD_TOEPFER)		
		'Se envia el mail al Proveedor, al administrador del pedido y a quien envia el mail (dado que puede ser un administrativo y necesita copia para su control)
		destino = obtenerMail(pct_idProveedorElegido) & "; " & origen & "; " & getUserMail(session("Usuario"))
	else
		'tiene una cotizacion asociada y no un pedido, se referencia a esta cotizacion 
		Call readCTZ(idCotizacion)	'esta funcion me permite cargar entre otras cosas el proveedor de esa cotizacion, para obtener el mail'
		asunto  = GF_TRADUCIR("NOTA DE ACEPTACION") & " - REF: " & idCotizacion
		if (origen = "") then origen = obtenerMail(CD_TOEPFER)
		destino = obtenerMail(ctz_IdProveedor) & "; " & origen & "; " & getUserMail(session("Usuario"))
	end if
	
	'destino = "scalisij@toepfer.com"	
	
	mensaje = GF_TRADUCIR("Su cotizacion ha sido aceptada por ADM Agro SRL")	
	Call GP_ENVIAR_MAIL_ATTACHMENT(asunto, mensaje, origen, destino, pathPDF)
		
	'una vez enviado el mail borro el pdf	
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	fs.DeleteFile(pathPDF)	
%>

