<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<% 

'******************************************************************************************************************************************
'Autor: Ajaya Nahuel - 23/03/2012
'Modificacion: Ajaya Nahuel - 26/04/2012

'PROPOSITO:
'En esta pagina se va a crear el POP UP para que se pueda ingrasar el Texto del Mensaje de una Nota de Aceptacion, cuya pagina esta compuesta por un TextAerea
'en el cual va a tener un texto ya puesto por defecto, de ahi el usuario podra o no modificarlo en base a lo que le quiera escribir. A esto se le agrega un boton de 
'imprimir y otro de enviar, ambos primero verifican que el texto ingresado sea logico y luego toman acciones dependiendo cada uno.
'Luego se le agrego una tabla en la que puede modificar el mail del proveedor o agregar otros mas, q permite ya guardarlos 
'Se adapto esta pagina primero para crear una NDA solamente cuando tiene un pedido asociado, pero luego se modifico para que tambien lo haga por un PIC

'******************************************************************************************************************************************
Dim pIdPedido, strMsj, rs, admin, accion, dsMensaje,idProveedor, mailsProveedor,cantRs, dsproveedor

pIdPedido = GF_Parametros7("idPedido",0,6)
'el pedido va a variar por que si viene de la pagina de comprasAdministrarCotizaciones.asp solo va a traer el IdCotizacion por que es una compra directa y su pedido
'va a ser 0 si no lo tiene, en cambio si viene de la pagina comprasFichaPCTtab1.asp va a venir con un IdPedido y un IdCotizacion(la cual va atener como prioridad el PCT y despues el PIC)
pIdCotizacion = GF_Parametros7("idcotizacion",0,6)
accion = GF_PARAMETROS7("accion",0,6)
dsMensaje = replace (GF_PARAMETROS7("notaAceptacion","" ,6),chr(13)&CHR(10),"<br>")


if(pIdPedido > 0)then	
	'si viene con un pedido, el texto del mensaje va hecer referencia al Pedido
	set rs = cargaMensajeNDA(pIdPedido,pIdCotizacion)
	Call GF_MGKS ("SG",rs("cdusradmin"), "", admin)
	if (admin = "") then admin = "Departamento de compras"	
	strMsj= GF_Traducir("At: ")&trim(rs("dsempresa"))&"<br><br>"
	strMsj= strMsj & GF_Traducir("Mediante la presente manifestamos nuestra ")
	strMsj= strMsj & GF_Traducir("conformidad a vuestro presupuesto de fecha ")&left(GF_FN2DTE(rs("fechapresentacion")),10)
	strMsj= strMsj & GF_Traducir(" correspondiente al pedido ")&rs("cdpedido")&" "&rs("titulo")&"<br><br>"
	strMsj= strMsj & GF_Traducir(" Los saluda Atte.")
	strMsj= strMsj & GF_TRADUCIR(admin)			
else	
	'si viene sin pedido pero con IdCotizacion, el texto del mensaje va aperecer la cotizacion solamente
	set rs =  cargaPICMensajeNDA(pIdCotizacion)
	Call GF_MGKS ("SG",rs("CDUSUARIO"), "", admin)
	if (admin = "") then admin = "Departamento de compras"	
	strMsj= GF_Traducir("At: ")&trim(rs("dsempresa"))&"<br><br>"
	strMsj= strMsj & GF_Traducir("Mediante la presente manifestamos nuestra ")
	strMsj= strMsj & GF_Traducir("conformidad a vuestro presupuesto de fecha ")&left(GF_FN2DTE(rs("MOMENTO")),10)
	strMsj= strMsj & GF_Traducir(" correspondiente a la cotizacion Nº ")&pIdCotizacion&"<br><br>"
	strMsj= strMsj & GF_Traducir(" Los saluda Atte.")
	strMsj= strMsj & GF_TRADUCIR(admin)	
end if
If NOT IsNull(rs("idproveedor")) then mailsProveedor = obtenerMail(rs("idproveedor"))

cantRs = rs.recordCount

%>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script defer type="text/javascript" src="scripts/pngfix.js"></script>
	<script>
	var ch= new channel();	
	<% if (mailsProveedor <> "") then %>
	var flagHayMail =  true;
	<% else %>
	var flagHayMail =  false;
	<%end if %>;
	var vProveedores = new Array();
	var vEmail = new Array();
	
	function comprobarIngresoNDA(idPed, idCoti, accion){		
		if (flagHayMail) {
			var pMsj = document.getElementById("notaAceptacion").value;			
			if(pMsj == ""){
				alert("Debe ingresar el mensaje de la Nota de Aceptación");
			} 
			else {		
				//Para remmplazar los saltos de linea en javascript se utiliza la siguiente funcion 
				pMsj = pMsj.replace(/\n/gi,"<br>")
				pMsj = pMsj.replace(/&/gi,"|A|")
				//donde '/\n/gi' es el salto de linea capturado por javascript y lo transforma en <br> que son de asp				
				if(accion == <%=ACCION_ENVIAR_NDA%>){
					//primero actualiza luego envia
					ch.bind("comprasNDAAjax.asp?idPedido="+idPed+"&idcotizacion="+idCoti+"&accion="+accion+"&dsMensaje="+pMsj,"callback_enviarNDA("+idPed+","+idCoti+")");
				}
				if(accion == <%=ACCION_IMPRIMIR_NDA%>){
					//primero actualiza luego imprime
					ch.bind("comprasNDAAjax.asp?idPedido="+idPed+"&idcotizacion="+idCoti+"&accion="+accion+"&dsMensaje="+pMsj,"callback_printNDA("+idPed+","+idCoti+")");
				}	
				ch.send();			
			}
		}			
		else {
			alert("El proveedor no posee ninguan dirección de e-mail registrada.!!");
		}
	}	
	function callback_enviarNDA(pIdPed,pIdCoti) {	
		//encargada de llamar ala pagina que envia , luego esa pagina llama a la que crea el PDF				
		var myNda = ch.response();		
		ch.bind("comprasEnvioNDAMail.asp?idPedido="+pIdPed+"&idCotizacion="+pIdCoti+"&IdNDA="+myNda,"callback_NDAenviado()");
		ch.send();		
	}
	function callback_NDAenviado(){
		// encargada de mostrar por pantalla que el mail fue enviado 
		document.getElementById("avisoNDA").innerHTML="<% =GF_TRADUCIR("La Nota de Aceptación se ha enviado.") %>";
		document.getElementById("avisoNDA").className = "TDSUCCESS";
	}		
	function callback_printNDA(pIdPed,pIdCoti){
		// muestra en una ventana el  PDF
		var myNda = ch.response();
		document.getElementById("avisoNDA").className = "TDBAJAS";
		document.getElementById("avisoNDA").innerHTML="<% =GF_TRADUCIR("Generando PDF...") %>";	
		window.open("comprasnotadeaceptacionprint.asp?idPedido="+pIdPed+"&idCotizacion="+pIdCoti+"&IdNDA="+myNda, "_blank", "location=no,menubar=no,scrollbars=yes,scrolling=yes,height=650,width=1000",false);			
		parent.window.cerrarPopUpNDA();	//cierra el pop up
	}	
	//*************** funciones encargadas de administrar las direcciones de mail  ***********************	
	function editMail() {
		document.getElementById("txtmail").style.display = 'none';
		document.getElementById("inputMail").style.display = 'block';
		document.getElementById("imgMail").innerHTML = '<img src="images/compras/save-16x16.png" onClick="saveMail()" title="Guardar">';
		}	
	function saveMail() {
		var actualMail = document.getElementById("actualMail").value;
		var newMail = document.getElementById("inputMail").value;
		var idProv = document.getElementById("idProveedor").value;
		if ((actualMail != newMail) && (newMail != '')) {
			document.getElementById("imgMail").innerHTML = '<img src="images/loading_small_green.gif">';
			ch.bind("comprasActualizarMail.asp?idEmpresa=" + idProv + "&mail=" + newMail, "saveMailCallBack()");
			ch.send();
		} else {
			alert('No se produjo ningun cambio en el mail');
			updateChange(actualMail);
		}
	}
	function saveMailCallBack() {
		var actualMail = document.getElementById("actualMail").value;
		var newMail = ch.response();
		if (newMail != '') {
			flagHayMail = true;
			updateChange(newMail);
		} else {
			flagHayMail = false;
			alert('El mail ingresado es invalido');
			updateChange(actualMail);
		}
	}
	function updateChange(mail) {
		document.getElementById("inputMail").value = mail;
		document.getElementById("inputMail").style.display = 'none';
		document.getElementById("txtmail").innerHTML = mail;
		document.getElementById("txtmail").style.display = 'block';
		document.getElementById("actualMail").value = mail;
		document.getElementById("imgMail").innerHTML = '<img src="images/compras/edit-16x16.png" onClick="editMail()">';
	}
	
	
	function actualizarMail(c){
		var x=document.getElementById("CmbProveedores").selectedIndex;
		var y=document.getElementById("CmbProveedores").options;		
		var indice = y[x].index;		
		document.getElementById("txtmail").innerHTML = vEmail[indice];
		document.getElementById("actualMail").value  = vEmail[indice];
	}
	
		
	
	//**********************************************************************************************
	</script>
</head>
<body>
	<form id="myForm" name="myForm" action="comprasNDAPopUp.asp" method="post">
	<input type='hidden' name='accion' id='accion' value='<%=accion%>'>	
	<input type="hidden" id="idPedido" name="idPedido" value="<% =pIdPedido %>">	
	<table width="100%">		
		<tr>
			<div id="avisoNDA" align="center" class="TDBAJAS"></div>
			<font class="big" color="#517b4a"><% =GF_TRADUCIR("Ingrese el mensaje para la Nota de Aceptacion") %></font>			
		</tr>
		<tr>
			<div id="MensajeEnviado" name="MensajeEnviado"></div>
			<td>			
				<table width="100%" border="0" id="Detalle" name="Detalle" cellpadding="1" cellspacing="2">				
					<tr><td align="center">
						<textarea rows="10" name="notaAceptacion" id="notaAceptacion" style="width:425px"><%=replace(strMsj,"<br>",chr(13)&CHR(10))%></textarea>
					</td></tr>
					<tr>
						<td align="center">
							<table class="reg_Header" width="100%" align="center">
								<tr class="reg_Header_nav" align="center">
									<td><% =GF_TRADUCIR("Proveedor") %></td>
									<td colspan="2" width="45%" align="center"><% =GF_TRADUCIR("Email") %></td>
								</tr>
								<tr class="reg_Header_navdos">
									<td>
										<div id="Prov"><%=rs("dsempresa")%></div>	
									</td>								
									<input type="hidden" id="idProveedor" value="<% =rs("idproveedor")%>">
									<td>
										<div id="txtmail"><% =mailsProveedor %></div>
										<input type="hidden" id="actualMail" style="display:block;" value="<% =mailsProveedor %>">
										<input type="text" id="inputMail" style="display:none;" size="40" value="<% =mailsProveedor %>">
									</td>
									<td width="15px">
										<div id="imgMail" style="cursor:pointer;">
											<img src="images/compras/edit-16x16.png" onClick="editMail()" title="Editar">
										</div>
									</td>									
								</tr>
							</table>	
						</td>
					</tr>	
				</table>
			</td>
		</tr>			
		<tr>
			<td colspan="2" align="left">
				<a href="javascript:comprobarIngresoNDA('<%= pIdPedido%>','<%= pIdCotizacion%>','<%=ACCION_ENVIAR_NDA%>')"><img id="iconNDA" src="images/compras/mail_sent-16x16.png" title="Enviar"> <% =GF_TRADUCIR("Enviar") %></a>
				&nbsp;&nbsp;&nbsp;
				<a href="javascript:comprobarIngresoNDA('<%= pIdPedido%>','<%= pIdCotizacion%>','<%=ACCION_IMPRIMIR_NDA%>')"><img src="images/compras/printer-16x16.png" title="Imprimir"> <% =GF_TRADUCIR("Imprimir") %></a>			
				
			</td>
		</tr>
	</table>
	</form>
</body>
</html>

