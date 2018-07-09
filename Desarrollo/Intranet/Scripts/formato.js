/**
 * CONSTANTES UTILIZADAS PARA LAS FUNCIONES DE FORMATO
 */
 var CHR_FWD = 0;	//Indica posicion prefijo
 var CHR_AFT = 1;	//Indica posicion sufijo
//--------------------------------------------------------------------- 
function editarNumero(valor, decimales) {
/**
 * Edita un numero para que tenga la cantidad de decimales indicada.
 */	
	valor = valor.replace(/\,/,".");
	valor = Number(valor);
	valor = valor.toFixed(decimales);
	return valor;	
}
//---------------------------------------------------------------------
function editarCaracteres(str, pChar, size, pos) {
/**
 *	Formatea el String indicado agregandole el caracter
 *  pasado por parametro hasta completar la longitud
 *  correspondiente.
 */
 var ret = "";
 if (str.length < size) {
 	//Asume que posicion es CHR_AFT
 	var prefix = "";
 	var suffix = pChar; 	
 	if (pos == CHR_FWD) {
 		prefix = pChar;
	 	suffix = ""; 	
 	}
	ret = str;
 	while (ret.length < size) {		
 		ret = prefix + ret + suffix;
 	} 	
 } else {
	if (pos == CHR_FWD) {
		ret = str.substr(str.length - size);
	} else {
		ret = str.substr(0, size);
	}
 }
 return ret;
}

//---------------------------------------------------------------------
// Edita un importe para que se muestre el importe en el formato correcto, con coma y dos decimales.
function editarImporte(valor) {
		return editarNumero(valor, 2);			
}
//---------------------------------------------------------------------
// Formatea un importe para que se muestre el importe en el formato correcto, con coma y dos decimales.
function formatearImporte(valor, decimales) {
pot = Math.pow(10,decimales);
num = parseInt(valor * pot) / pot;
nume = num.toString().split('.');

entero = nume[0];
decima = nume[1];

if (decima != undefined) {
	fin = decimales-decima.length; }
else {
	decima = '';
	fin = decimales; }

for(i=0;i<fin;i++)
  decima+=String.fromCharCode(48); 

buffer="";
marca=entero.length-1;
chars=1;
while(marca>=0){
   if((chars%4)==0){
	  buffer="."+buffer;
	  chars=1;
   }
   buffer=entero.charAt(marca)+buffer;
   marca--;
   chars++;
}
if (decimales >= 0) {
    num = buffer; +',' + decima;
}
else {
    num = buffer;
}
return num;
}

//---------------------------------------------------------------------
/**
 *	Formatea el String indicado recortando su longitud para adaptarla al tamaño pedido.
 * 
 *	@return	El string del tamaño pedido. 
 */
function limitarString(str, size) {
 if (str.length > size) {
 	str = str.substr(0, size-3);
 	str = str.concat("..."); 	
 }
 return str;
}