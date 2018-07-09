/*
 * Controla el caracter ingresado por teclado de acuerdo al tipo esperado.
 * Parametros:
 *				evento	: Evento
 * 				campo	: Objeto que tiene el foco
 * 				Tipos	: Tipo de dato que se esperar controlar
 *							N: Numerico
 *  	    				I: Importe
 */ 
function controlIngreso (campo, evento, tipo){
 	var ret = true;	
	var auxText = new String();
    var ascii = (document.all) ? evento.keyCode : evento.which;	
	var caracter = String.fromCharCode(ascii);
	auxText = campo.value;
	if (ascii==8 || ascii==0 || ascii==9 || ascii==13 || (ascii>47 && ascii<58)) return true;//Backspace,Delete,Tab,Enter 
	//alert(ascii);
	if (tipo=="E" && ascii==46) {
	    if (document.all){
	    	evento.keyCode=0;
        	evento.cancelBubble = true;
		    evento.returnValue = false;
	    }
		else{
			if (evento.stopPropagation) {
				evento.stopPropagation();
	            evento.preventDefault();
			} 
		}	
	    return true;
	}
   	else{
		if (ascii == 44) { //La coma
			if (auxText.search(caracter) > 0){
				alert("No puede colocar otro separador de decimales!");
				evento.keyCode = 0;
				//evento.cancel;
				}
		}
		//alert(caracter);
		switch (tipo) {
   			case "N":
   				ret = controlNumero(auxText + caracter); 
   				break;
	   		case "I":
   				ret = controlImporte(auxText + caracter);
   				break;
			case "E":    	
   				ret = controlImporte(auxText + caracter);
   				break;
	   	};
	}
    return ret;
}
//-------------------------------------------------------------------------------------------------------------------
/**
 * Controla el caracter ingresado por teclado de acuerdo al tipo esperado.
 * Si todo esta OK y corresponde, salta al siguiente campo de ingreso de datos.
 * Parametros:
 * 				evento  : Evento
 * 				campo   : Objeto que tiene el foco
 * 				Tipos   : Tipo de dato que se esperar controlar
 *							N: Numerico
 *      					I: Importe
 *				
 */
function controlDatos(campo, evento, tipo) {	
	var ret = controlIngreso (campo, evento, tipo);
	if (ret) controlSalto(campo, evento);
	return ret;
}
//-------------------------------------------------------------------------------------------------------------------
/**
 * Controla el campo de entrada y en caso de error lo resalta.
  * Parametros:
  * 			campo   : Objeto que tiene el foco
 * 				Tipos   : Tipo de dato que se esperar controlar
 *							N: Numerico
 *		      				I: Importe
 *							F: Fecha
 *      					D: Dia
 *      					M: Mes
 *      					Y: Anio
 */
function controlCampo(campo, tipo) {		
	var ret = true;	
	if ((campo.value == "") || (campo.value == undefined))  return ret;
	campo.className = "";
	switch (tipo) {
   		case "N":
   			ret = controlNumero(campo.value);    			
   			if (ret) campo.value = editarNumero(campo.value, 0);
   			break;
   		case "I":    	   			
   			ret = controlImporte(campo.value);
			if (ret) campo.value = editarImporte(campo.value);				
   			break;
		case "F":    	
   			ret = controlFecha(campo.value);   			
   			break;		
   		case "D":    	
   			ret = controlNumero(campo.value, 1, 31);
   			if (ret) campo.value = editarNumero(campo.value, 0);
   			break;
   		case "M":    	
   			ret = controlNumero(campo.value, 1, 12);
   			if (ret) campo.value = editarNumero(campo.value, 0);
   			break;
   		case "Y":    	
   			ret = controlNumero(campo.value, 1980, 2079);
   			if (ret) campo.value = editarNumero(campo.value, 0);
   			break;
   	};
	if (!ret) campo.className = "ERRVALOR";
	return ret;
}
//-------------------------------------------------------------------------------------------------------------------
/**
 * Esta funcion permite detectar cuando saltar al siguiente campo
 * de ingreso de datos automaticamente. 
 *
 * Parametros:  input    : Objeto que tiene el foco 
 *				e        : Evento
 */
function controlSalto(input, e) {

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);
	var keyCode = (isNN) ? e.which : e.keyCode;
	var filtro = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];
	var maxLen = input.maxLength;		
	var cant = (isNN) ? 1 : 0;
	var curLen = input.value.length + cant;
	if(curLen == maxLen && !contieneElementos(filtro,keyCode)) {
		input.value = input.value.slice(0, maxLen);
		input.form[(getIndex(input)+1) % input.form.length].focus();
	}
}
//-------------------------------------------------------------------------------------------------------------------
function contieneElementos(arr, elem) {
    var found = false, index = 0;
    while(!found && index < arr.length)
		if(arr[index] == elem)
			found = true;
		else
			index++;
    return found;
  }
//-------------------------------------------------------------------------------------------------------------------
function getIndex(input) {
    var index = -1, i = 0, found = false;
    while (i < input.form.length && index == -1)
    if (input.form[i] == input)index = i;
    else i++;
    return index;
  }

//-------------------------------------------------------------------------------------------------------------------
/**
 * Controla que el valor sea un numero
 */
function controlNumero (valor, min, max){
	if (valor == "-") return true;
	if (isNaN(valor)) return false;
	if (min != undefined) if ((valor < min) || (valor > max)) return false;
	return true;
}
//------------------------------------------------------------------------------------------------------------------- 
/**	
 * Controla que el parametro sea un importe valido. 
 * Numerico con a lo sumo 2 decimales.
 */
function controlImporte (valor){
	//alert("Controlar Importe(" + valor + ")")
 	//Remplazo coma por punto dado que los numeros solo aceptan punto.    	
 	valor = valor.replace(/\,/,".");
    if (controlNumero(valor)) {        	
		var aux = valor.indexOf(".");
    	if (aux != -1) {
    		if (valor.length > (aux + 3)) return false
    	}
    	return true
    }
    return false;
    
} 
//-------------------------------------------------------------------------------------------------------------------
/**
* Controlar fecha
* es llamada con desc="dia", "mes" o "anio" 
* y valor="DD/MM/AAAA
*/
function controlFecha(pfecha) {

var fecha;
var diasFebrero=0;
var diasMes;
var fechaValida=true;
var intAnioCtl=0;

fecha = pfecha.split("/");

if ((fecha[0]=="") || (fecha[1]=="") || (fecha[2]=="")) {
	fechaValida=false;
} else {
	// Determinacion de los dias del mes de febrero
	fecha[2] = Number(fecha[2]);
 	if ((fecha[2] %  4) == 0) {
   		// El año es biciesto.
   		diasFebrero= 29;
	} else {
   		//El año no es biciesto.
   		diasFebrero= 28;
	}
   		   	
   	//Determinacion de los dias del mes.
	switch (fecha[1]) {
  		case "01":
  		case "03":
  		case "05":
		case "07":
  		case "08":
  		case "10":
  		case "12": diasMes = 31; 
			break;
	   	case "02": diasMes = diasFebrero; 
			break;
  		case "04":
  		case "06":
  		case "09":
   		case "11": diasMes = 30; 
			break;
   		default: diasMes=0;
	}
 	// Se controla la fecha
   	if (diasMes == 0) 
   		fechaValida=false;
	else if ((parseInt(fecha[0]) < 1) || (parseInt(fecha[0]) > diasMes)) 
	   		fechaValida=false;
}

return fechaValida;
}
