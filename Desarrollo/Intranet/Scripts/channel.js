/**
 *                   COMMUNICATION CHANNEL
 *
 *   Canal de comunicaciones entre el servidor y el cliente para
 *   consulta de datos.
 *   El canal opera en forma secuencial, procesando una solicitud a la vez.
 *   Permite realizar multiples consultas, pero internamente las ira resolviendo una a una.
 *   
 *
 *  version 1.2.0 - 17/01/2008 - Javier A. Scalisi
 *  version 2.0.0 - 25/04/2008 - Javier A. Scalisi
 *		Se incorporo la posibilidad de que la respuesta del canal sea codigo HTML, 
 *		haciendo que la misma se parsee para obtener los JS a recuperar del server.
 *	version 2.1.0 - 28/04/2008 - Javier A. Scalisi
 *		Se incorporo la posibilidad de obtenter la respuesta desde el server ya parseada en una matriz de datos.
 *		Se utilizaron los separadores standard (|) para registros y (;) para campos.
 *		Se incorporo la posibilidad de indicar codigo fuente a ejecutar luego de cargar codigo fuente (responseSource)
 *  Version 2.2.0 - 11/01/2012 - Javier A.Scalisi
 *		Se modifico el componente para que cada llamada al channel sea independiente y asi evitar conflictos entre las llamadas
 */
//Constantes
var DEFAULT_LINE_TOKEN = "|";
var DEFAULT_FIELD_TOKEN = ";";

//Variables globales privadas que se necesitan para recivir
//la respuesta del server.
var $commChs = new Array();		//El Canal 
var $callbacks = new Array();	//Funcion de usuario a llamar.
var $servlets = new Array();	//URL del script de servidor a llamar
var $methods = new Array();		//Metodo HTTP de comuniaion (GET/POST)
var $ticket = new Array();		//Tickets entregados y pendientes de responder.
var $lastTicket = 0;			//Ultimo nro de ticket Entregado
var $response = new Array();	//Repuestas obtenidas por el canal.
var $loadedobjects = "";		//Objetos ya cargados para otras secciones, evita repeticiones.
var $objToLoad = new Array();	//Objetos a cargar para el funcionamiento de la seccion cargada.
var $lastResponse = 0;			//Ultimo ticket respondido
function channel() {
 	
 	//Atributos 	
 	this.servlet;				//Servlet a invocar.
 	this.method = "GET";		//Metodo de conexion.
	this.callback;				//Funcion a llamar con el response
	
	//Metodos
 	this.bind = function (servlet, callback) {			
					this.callback = "doNothing()";
					if (callback != "") {
						this.callback = callback;
						var chr = callback.charAt(callback.length-1);
						if ((chr != ")") && (chr != ";")) this.callback = callback.concat("()");
					}								
					var d = new Date();
					var par = "?" 
					if (servlet.indexOf("?") != -1)  par = "&"						
					par += "TS=" + d.getTime();
					servlet = servlet + par;
					this.servlet = servlet;
				}
	/**
	*	Funcion para dar la orden de envio de datos al server.
	*	Pone el pedido en la cola de pedidos hasta que el canal este disponible.
	*	Devuelve un nro de ticket unico para identificar la respuesta.
	*/
	this.send = function () {
					tkt = this.nextTicket();
					$ticket.push(tkt);
					$commChs.push(this.createAjaxChannel());
					$callbacks.push(this.callback);
					$servlets.push(this.servlet);
					$methods.push(this.method);
					//Se fuerza la comuniacion solo si hay un unico pedido pendiente.
					if ($ticket.length == 1) sendNow();					
					return tkt;
				}	
	//Funcion que devuelve el texto de respuesta.
 	this.response = function (ticket) {
 						if (!ticket) ticket = $lastResponse;  
						
						return $response[ticket];
					}		
	//Funcion que obtiene desde el servidor un HTML y lo coloca en el elemento cuyo ID es el indicado por parametro.	
	this.responseSource =	function(ticket) {
								var resp = this.response(ticket);
								loadScripts(resp);
								return resp;
							}
	//Devuelve la respuesta del canal en una matriz, separando los registros y los campos.	
	this.parsedResponse =	function(ticket, lineToken, fieldToken) {
								var ret = new Array();
								var lToken = lineToken || DEFAULT_LINE_TOKEN;
								var rToken = fieldToken || DEFAULT_FIELD_TOKEN;
								var resp = this.response(ticket);
								if (resp) {
									var regs = resp.split(lToken);
									for (var x in regs) ret.push(regs[x].split(rToken));
								}
								return ret;
							}		
	//Devuelve true si ya se encuentra disponible la respuesta del ticket consultado.
	this.readyTicket = 	function (ticket) {
							return ($response[ticket] != undefined);
						}
	
	this.nextTicket = 	function() {
							$lastTicket++;
							if ($lastTicket > 100) $lastTicket=1;
							return $lastTicket;
						}
	this.createAjaxChannel = function() {						
								if (window.XMLHttpRequest) {
									ajaxCh = new XMLHttpRequest();
								} else if (window.ActiveXObject) {
									ajaxCh = new ActiveXObject("Microsoft.XMLHTTP");
								}
								return ajaxCh;
							}
	
}
/***********************************************************************************************************************************************
* FUNCIONES GLOBALES DE OPERACION
*/
/**
* 	 Funcion que envia el pedido al server
*/
function sendNow () {						
	if ($ticket.length > 0) {	
		$currCh = $commChs[0];
		$currCh.open($methods.shift(), $servlets.shift(), true);   				   				   		
		$response[$ticket[0]] = undefined;						
		$currCh.onreadystatechange = responseHandler;	
		$currCh.send(null);
	}
} 	

//Funion que manipula las respuestas de AJAX y 
	//llama a la funcion de Usuario
function responseHandler() {	
	$currCh = $commChs[0];
	if ($currCh.readyState == 4) {	
		if ($currCh.status == 200) {
			$lastResponse = $ticket.shift();
			$commChs.shift();
			$response[$lastResponse] = $currCh.responseText; //eval(command);
			sendNow();
			eval($callbacks.shift());  							
		} else {
			//el status 0 se produce si hay una transaccion en curso y se abandona la pagina, no se muestra, no hace falta.
			if ($currCh.status != 0) {
				alert("CHANNEL - Status: " + $currCh.status);
			}
		}		
	}
}
/* Default calback */
function doNothing() {}


/* Funcion que parsea una pagina para obtener los scripts a traer del servidor */
function parseCode(texto) {
	var len = 0;
	var posIni = texto.indexOf("<script");		
	for (; posIni != -1;) {			
		//Hay Scripts
		var posFin = texto.indexOf("</script");
		var temp = texto.substring(posIni, posFin); 
		var posSrcIni = temp.indexOf("src=");
		if (posSrcIni != -1) {
			//Es un include.
			posSrcIni += 5;
			var posSrcFin = temp.indexOf(".js") + 3;
			//Si el include es del propio canal, no lo incluyo.
			if (temp.indexOf("channel") == -1)	len = $objToLoad.push(temp.substring(posSrcIni, posSrcFin));		
		} else {
			//Es codigo fuente en la propia pagina.
			var temp2 = temp.indexOf(">"); 
			var posCodeIni = temp2 + 1;
			var posCodeFin = posFin - 2;
			len = $objToLoad.push(temp.substring(posCodeIni, posCodeFin));
		}			
		texto = texto.substr(posFin+4);
		posIni = texto.indexOf("<script");		
	}		
	return len;
}

/* Funcion que obtiene los scripts necesarios para la pagina cargada */
function loadScripts(texto) {				
	var i = parseCode(texto);
	for (k = 0; k < i; k++){			
		var file = $objToLoad.shift();
		var fileref="";
		if ($loadedobjects.indexOf(file)==-1){ 
			if (file.indexOf(".js")!=-1){ 
				fileref=document.createElement('script');
				fileref.setAttribute("type","text/javascript");
				fileref.setAttribute("src", file);
				document.getElementsByTagName("head").item(0).appendChild(fileref);
				$loadedobjects+=file+" ";
			} else {
				eval(file);
			}
		}
	}
}	
