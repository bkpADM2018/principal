/*
TOOLBAR 6.0                 
Barra de herramientas Web.             
                                        
	Release :1.0 - 18/01/2007 - Javier A. Scalisi      
	Release: 2.0 - 14/02/2007 - Javier A. Scalisi
	Release: 3.0 - 14/04/2008 - Javier A. Scalisi
	Release: 4.0 - 14/05/2008 - Javier A. Scalisi 	
	Release: 5.0 - 20/11/2008 - Javier A. Scalisi 
	Release: 6.0 - --/10/2013 - Javier A. Scalisi 	
*/

/****************************************************
*	Como Usar:                            			*
*    Crear un objeto Toolbar y definir un 			*
*    DIV con el mismo ID, al final del  			*
*    body llamar a la funcion draw().   			*
*****************************************************/

var TOOL_IMAGE_DIR = "images/";

/* Constantes de navegadores */
var TOOLB_NAV_FF = 1; //Firefox
var TOOLB_NAV_IE = 0; //Internet Explorer

var TOOLB_NAVIGATOR = (navigator.userAgent.indexOf("MSIE") >= 0) ? TOOLB_NAV_IE : TOOLB_NAV_FF;


	//Parametros deprecated
	//pButtons, imgDir --> No utilizarlos!!
function Toolbar(pid, pButtons, imgDir) {
	
/* Atributos */
this.key = pid;
this.imgDir = imgDir || TOOL_IMAGE_DIR;
if (this.imgDir.slice(-1) != "/") this.imgDir = imgDir + "/";

this.buttons = new Array();


/* Funciones */

/* Botones pre-configurados */
	//DEPRECATED - NO UTILIZAR LOS BOTONES PREDETERMINADOS
this.addButtonADD 			=	function(text, func) { return this.addButton("plus-16.png",text,func); } /*Generico, simbolo + */
this.addButtonADDDOC		=	function(text, func) { return this.addButton("add-16.png",text,func); }	/*Añadir document*/
this.addButtonADDPERFIL     =   function(text, func) { return this.addButton("add-user-16.png", text, func); } /*Añadir user*/
this.addButtonEXPORT 		=	function(text, func) { return this.addButton("export-16.png",text,func); }	/*Generico exportar*/
this.addButtonPRINT			=	function(text, func) { return this.addButton("print-16.png",text,func); }
this.addButtonPDF           =   function(text, func) { return this.addButton("pdf-16.png", text, func); } /*Exportar en PDF*/
this.addButtonEXCEL         =   function(text, func) { return this.addButton("excel-16.png", text, func); } /*Exportar en Excel*/
this.addButtonSEE			=	function(text, func) { return this.addButton("see-16.png",text,func); } /*VER - icon ojo*/
this.addButtonSEARCH		=	function(text, func) { return this.addButton("buscar-16.png",text,func); } /*Lupa de buscar*/
this.addButtonHOME 			=	function(text, func) { return this.addButton("home-16.png",text,func); }
this.addButtonREFRESH 		=	function(text, func) { return this.addButton("refresh-16x16.png",text,func); }
this.addButtonCANCEL 		=	function(text, func) { return this.addButton("cross-16.png",text,func); }
this.addButtonCONFIRM 		=	function(text, func) { return this.addButton("checkmark-16.png",text,func); }
this.addButtonSAVE 			=	function(text, func) { return this.addButton("save-16.png",text,func); }
this.addButtonWATCH			=	function(text, func) { return this.addButton("see-16.png",text,func); }
this.addButtonRETURN 		= 	function(text, func) { return this.addButton("back-16.png",text,func); }
this.addButtonMAXIMIZE 		=	function(text, func) { return this.addButton("zoom_in-16x16.png",text,func); }
this.addButtonREDUCE 		=	function(text, func) { return this.addButton("zoom_out-16x16.png",text,func); }
	//DEPRECATED - NO UTILIZAR LOS BOTONES PREDETERMINADOS



/*========================================================================================
	AGREGAR BOTON
 		Esta funcion permite agregar un boton standard a la barra de herramientas.
========================================================================================*/
this.addButton = function(styleClass, text, func) {                    
                    var button = new GenericButton();	
					button.styleClass += styleClass;
					//Si la clase tiene un punto, asumo que es una imagen. (Metodo toolbar vieja)
					var classorimg = styleClass.indexOf(".");					
					if (classorimg != -1) {
						button.isClass = false;
						button.styleClass = this.imgDir + styleClass;
						//Si la URL no existe, se busca la imagen en el directorio default de imagenes.
						if (!ToolUrlExists(button.styleClass)) button.styleClass = TOOL_IMAGE_DIR + styleClass;						
					}
					button.action = func;
					button.text = text;
					//button.width = this.buttonWidth;
					this.buttons.push(button);
					var nbr = this.buttons.length - 1;
					return nbr;
				}

/*========================================================================================
	SWITCHER
		Funcion que se utiliza para insertar una tecla de dos estados en la barra.
		
		img*: Imagen para on/off;
		func*: Funcion a ejecutar para cuando se cambia al estado on/off;
		text*: Texto a mostrar en el estado on/off.
======================================================================================== */
this.addSwitcher = 	function(styleClass, text, actionOn, actionOff) {
						return this.addSwitch(styleClass, text, actionOn, styleClass, text, actionOff);
					}
this.addSwitch = 	function(styleClassOff, textOff, actionOn, styleClassOn, textOn, actionOff) {
						var button = new GenericSwitcher();
						button.actionOff = actionOff;
						button.action = actionOn;		
						//Si la clase tiene un punto, asumo que es una imagen. (Metodo toolbar vieja)
						
						var classorimg = styleClassOn.indexOf(".");					
						if (classorimg != -1) {
							button.isClass = false;
							button.styleClass = this.imgDir + styleClassOff;
							//Si la URL no existe, se busca la imagen en el directorio default de imagenes.
							if (!ToolUrlExists(button.styleClass)) button.styleClass = TOOL_IMAGE_DIR + styleClass;	
							button.styleClassOn = this.imgDir + styleClassOn;
							//Si la URL no existe, se busca la imagen en el directorio default de imagenes.
							if (!ToolUrlExists(button.styleClassOn)) button.styleClassOn = TOOL_IMAGE_DIR + styleClass;	
						}												
						button.text = textOff;
						button.textOn = textOn;						
						this.buttons.push(button);
						var nbr = this.buttons.length - 1;
						return nbr;						
					}
/* Funcion que permite cambiar el estado de un Switch (Solo para Switchs!!!)) */
this.changeState =	function(idButton) {
						var button = this.buttons[idButton];		
						button.changeState();						
					} 	

/*========================================================================================
	CHANGELOOK
		Funcion que permite modificar el aspecto fisico de un boton.
========================================================================================*/
this.changeLook =	function(idButton, img, text) {
						var button = this.buttons[idButton];		
						button.changeLook(img, text);						
					} 

/*========================================================================================
	DRAW
========================================================================================*/
					this.draw = function() {
					    //Se abre la barra		
					    var html = "<ul class=\"toolBarHolder\">";
					    for (var i in this.buttons) {
					        html += this.buttons[i].draw();
					    }
					    html += "</ul>";
					    document.getElementById(this.key).innerHTML = html;
					}
}

/*========================================================================================
	GENERIC BUTTON
		Clase generica para dibujo de botones.
========================================================================================*/
var $toolButtons = new Array();
var $toolImages = new Array();

function GenericButton() {

	//Atributos
	this.xNumber 	= $toolButtons.length; 
	$toolButtons.push(this);
	this.id			= "TOOLB" + this.xNumber;
	this.styleClass = ""
	this.isClass 	= true;	//Sirve para determinar si el parametro this.styleClass es el nombre de una clase o una imagen (imagenes se usaban en toolbar vieja!)
	this.action 	= "";
	this.text 		= "";

	//FUNCIONES
	//Funcion que devuelve el HTML de un boton.
	this.draw = function() {
	    var html = "<li id=\"" + this.id + "\"";

	    if (this.isClass) html += " class=\"" + this.styleClass + "\" ";

	    if (this.action) html += " onclick=\"" + this.action + "\" ";

	    html += "><span>";

	    if (!this.isClass) {
	        html += "<img id=\"btnImg" + this.xNumber + "\" src=\"" + this.styleClass + "\">";
	        //Se guarda la imagen cargada para checkear la carga OK.
	        $toolImages.push("btnImg" + this.xNumber);	        
	    }

	    html += "</span>";

	    html += "<p id=\"btnText" + this.xNumber + "\">" + this.text + "</p></li>";
	    //this.isDrawn = true;
	    return html;
	}		
		
	this.changeLook = 	function(newStyle, newText) {
							if (!this.isClass) {
								//Botones armados con imagenes. Estilo viejo de toolbar.
								var btnImg = document.getElementById("btnImg" + this.xNumber);
								if ((newStyle) && (btnImg)) {																	
									btnImg.src = newStyle;
							} else {
								//Botones armados solo con estilos.
								var btn = document.getElementById(this.id);
								btn.className=newStyle;
							}
							var btnText = document.getElementById("btnText" + this.xNumber);
							if ((newText) && (btnText)) btnText.innerHTML = newText;
						}
					}
	return this;
}

/*========================================================================================
	GENERIC SWITCHER
========================================================================================*/
function GenericSwitcher() {
	this.base= GenericButton;
	this.base();
	//----------------------------------------------------------------------------------------------------------
	//Atributos
	this.status = 0;
	this.actionOff = "";
	this.styleClassOn = "";
	this.textOn = "";
		
	this.changeState = 	function() {
							if (this.status != 0) {
								this.changeLook(this.styleClass, this.text);
								this._setEvent(this.actionOff);
								this.status = 0;
							} else {	
								this.changeLook(this.styleClassOn, this.textOn);
								this._setEvent(this.action);
								this.status = 1;
							}		
						}
	this._setEvent = function(func) {
						//Se cambia el evento del boton.
						var btn = document.getElementById(this.id);							
						if (TOOLB_NAVIGATOR == TOOLB_NAV_FF) { //FF
							btn.setAttribute("onclick", func);
						} else {   //IE
							eval("btn.onclick= function() { eval(" + func + ") }");
						}
					}
	return this;
}
//----------------------------------------------------------------------------------------------------------
//  TOOLBAR GLOBAL FUNCTIONS
//----------------------------------------------------------------------------------------------------------
function ToolUrlExists(url) {
    var http = new XMLHttpRequest();
    http.open('HEAD', url, false);
    http.send();
    return http.status != "404";
}
