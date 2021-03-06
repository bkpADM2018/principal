var popUpNameRegister = Array();	//Registro de nombre del popUp.
var HEIGHT_POPUP = 80;	//Tama�o agregado del popup titulo y bonones
/**
 *FUNCION PARA CREAR EL OBJETO PARA ABRIR UN POP UP
 *pName			--	NOMBRE QUE DEFINE EL ID DEL IFRAME Y EL INDICE PARA OBTENER EL OBJETO
 *pUrl			--	PAGINA A ABRIR EN EL POP UP
 *ptitle		--  TITULO POP UP
 *pWidth		--  WIDTH POP UP
 *pHeight		--  HEIGHT POP UP
 *pActionClose	--	ACCION QUE SE EJECUTA AL CERRAR EL POPUP
 */
 //---------------------------------------------------------------------
function winPopUp(pName, pUrl, pWidth, pHeight, ptitle, pActionClose) {
	//instancia objeto
	popUpNameRegister[pName] = this;
	//objeto jquery
	var $this = $(this);
	
	this.show = function() {
					$this = $("<iframe id='" + pName + "' class='" + pName + "' src='" + pUrl + "' border='1' style='padding:0px;' />").dialog({
						title: (ptitle) ? ptitle : 'Operaci�n Externa',
						autoOpen: true,
						width: pWidth,
						//el valor agregado mejora el centrado en la pagina
						height: parseInt(pHeight) + HEIGHT_POPUP,
						modal: true,
						draggable: false,
						resizable: false,
						autoResize: true,
						close: function(event, ui){
							eval(pActionClose)
						}
					}).width(pWidth).height(pHeight);
				}
	this.hide = function() {
					$this.dialog('close');
				}
	this.resize = function(pWidth, pHeight) {
					$this.dialog('option', 'width', pWidth);
					$this.dialog('option', 'height', pHeight);
					$this.width(pWidth).height(pHeight);
					$this.dialog('option', 'position', 'center');
				}
	this.show();
	return this;
}
//---------------------------------------------------------------------
/* Devuelve una referencia a la instancia de la variable de la ventana padre
 * que controla el popup. 
 * Se tomo desde la pagina Iwin reemplazada
 */
//---------------------------------------------------------------------
function getObjPopUp(name) {
	if (arguments.length==1) {
		//Child side.
		var p = window.parent;
		if (typeof(p.getObjPopUp) == "function") return p.getObjPopUp(name, true);
	} else {
		//Parent side.
		return popUpNameRegister[name];
	}
}
//---------------------------------------------------------------------
