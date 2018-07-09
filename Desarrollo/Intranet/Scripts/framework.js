/************************************************************\
 *                       FRAMEWORK JS                       *
 *                                                          *
 *  Set de funciones y constantes globales utiles para todas*
 *  las aplicaciones de la empresa. El objetivo es brindar  *
 *  respuesta a usos básicos comunes.                       *
 *  Es impoortante que todas las funciones de este archivo  *
 *  puedan utilizarse en los browsers habilitados.          *
 *                                                          *
 *  Javier A. Scalisi - 14/08/2013                          *
 ************************************************************/
/* Constantes de navegadores */
var FWRK_NAV_FF = 1; //Firefox
var FWRK_NAV_IE = 0;  //Internet Explorer

var FWRK_NAVIGATOR = (navigator.userAgent.indexOf("MSIE") >= 0) ? FWRK_NAV_IE : FWRK_NAV_FF;

/* Constantes para eventos de pantalla */
var FWRK_EVT_ON_CLICK = "onclick";
var FWRK_EVT_ON_MOUSE_DOWN = "onmousedown";
var FWRK_EVT_ON_MOUSE_UP = "onmouseup";
var FWRK_EVT_ON_MOUSE_OVER = "onmouseover";
var FWRK_EVT_ON_MOUSE_OUT = "onmouseout";

/** 
 * Funcion responsable de setear un evento en un determinado objeto de pantalla.
 * Autor: Javier A. Scalisi
 * Fecha: 14/08/2013
 */
function setEvent(obj, evt, func) {
    if (FWRK_NAVIGATOR == FWRK_NAV_FF) { //FF
        obj.setAttribute(evt, func);
    } else {   //IE
        eval("obj." + evt + "= function() { eval(" + func + ") }");
    }
}
