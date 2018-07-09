var HK_APPLET_ROOT = "HKeyctl";
var HK_IMAGE_ROOT = "HKeyimg";
var HK_ACTION_ROOT = "HKeyAcc";

var hk_ch = new channel();
var hk_keys = new Array();
var hk_isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	

var hk_divID = new Array();
var hk_id = new Array();
/**
 * Codigos de error reservados por el componente.
 */
var HK_MSG_VALID = "HK0";
var HK_MSG_NOT_VALID = "HK1";
var HK_ERR_NO_KEY = "HK2";
var HK_ERR_UNKNOWN = "HK3";

/**
 * Objeto para administrar las llaves de seguridad en el cliente.
 * Parametros:
 *		
 *		params	: 	Estructura con los parametros de la consulta a ejecutar (Creado por la función HKEY de ASP)
 *		func	:	Funcion que será llamada por el componente cuando obtenga una respuesta de la validación de la llave.
 *					Esta funcion recibira un parametro con el codigo de la respuesta, el mismo sera agregado al final  
 *					de los parametros que ya posea la funcion.
 *					Por ejemplo:
 *						Si se pasa	 			Quedara como
 *						func();					func(<code>);
 *						func 					func(<code>);
 *						func(<parm1>) 			func(<parm1>, <code>);
 */
function Hkey(divId, link ,params, func, isSmall) {
	this.id = hk_keys.push(this) - 1;	
	this.divId = divId;
	this.link = link;
	this.params = params;
	this.keyData = "";
	this.func = func;
	this.keyImg = "images/hardkey/Authorize-48x48.png";
	this.isSmall = isSmall || false;
	if (this.isSmall) this.keyImg = "images/hardkey/Authorize-16x16.png";

	
	//Llamar esta función en el evento onload del body.
	this.start = 	function () {		
						var div = document.getElementById(this.divId);
						if (div) {
							var accImg = document.createElement("img");						
							accImg.value = "Confirmar";
							accImg.src = this.keyImg;
							accImg.style.cursor="pointer";
							accImg.title = "Autorizar";
							accImg.id = HK_ACTION_ROOT + this.id; 
							if (hk_isFirefox){
								accImg.setAttribute('onclick', "hk_check(" + this.id + ");");
							} else{
								accImg['onclick']=new Function("hk_check(" + this.id + ");return false;");
							}
							div.appendChild(accImg);							
							var img = document.createElement("img");
							img.src= "images/hardkey/loading_small_green.gif";
							img.style.visibility="hidden";
							img.id = HK_IMAGE_ROOT + this.id; 
							div.appendChild(img);							
							this.started = true;
						}
						this._connect();
					}		
	this._connect=function () {
						if (!document.getElementById(HK_APPLET_ROOT)) {
							var applet = document.createElement("applet");
							applet.id= HK_APPLET_ROOT;
							applet.code = "HardKeyApplet.ElApplet";
							applet.archive = "AppletMio.jar";
							applet.style.width= "0px"; 
							applet.style.height= "0px"; 
							applet.style.visibility= "hidden";
							document.body.appendChild(applet);
						}
					}
	this.check = 	function() {						
						//this._connect();
						var img = document.getElementById(HK_IMAGE_ROOT+ this.id);
						if (img) {
							img.style.visibility="visible";
							var accImg = document.getElementById(HK_ACTION_ROOT+ this.id);
							accImg.style.visibility="hidden";							
						}
						var hkCtl = document.getElementById(HK_APPLET_ROOT);
						this.keyData = hkCtl.EnviarComando(this.params);
						var dir = this.link;
						dir += (this.link.indexOf('?') == -1)? '?': '&';
						dir += "HK=" + this.keyData;																		
						hk_ch.bind(dir, "hk_check_callback(" + this.id + ")");							
						hk_ch.send();
					}
	//Metodo que ejecuta la funcion dada por el usuario.
	this.execUserFunc =	function (response) {
							var chr = this.func.charAt(this.func.length-1);								
							if (( chr != ")") && (chr != ";")) this.func += "()";
							eval(this._addParametro(this.func, response));
						}
						
	this._addParametro =function (p_func, p_value) {
						    var ret = p_func;
							//Se agregan parametros a la funcion JavaScript.
							var arrAux = p_func.split(")");
							ret = arrAux[0];
							if (ret.charAt(ret.length-1) != "(") ret += ", "; 
							ret += "'" + p_value + "');";							
							return ret;
						}
}
//---------------------------------------------------------------
//	Funciones Globales
//---------------------------------------------------------------
function hk_check_callback(id) {
	var img = document.getElementById(HK_IMAGE_ROOT+ id);
	if (img) img.style.visibility="hidden";		
	var resp = hk_ch.response();	
	var arrResp = resp.split("|");
	if (arrResp[0] == HK_MSG_VALID) {
		hk_keys[id].execUserFunc(arrResp[1]);
	} else {
		alert("HARDKEY ERROR " + arrResp[0] + "-" + arrResp[1]);
	}
}
//---------------------------------------------------------------
function hk_check(id) {
	hk_keys[id].check();		
}
//---------------------------------------------------------------
