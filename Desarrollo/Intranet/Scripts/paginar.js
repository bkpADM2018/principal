//*****************************************************************************\
//*
//*                        COMPONENTE DE PAGINACION
//*
//* Autor   : Javier A. Scalisi
//* Fecha   : 18/04/2008
//* Version : 3.0
//* Autor   : Javier A. Scalisi
//* Fecha   : 30/10/2013
//* Version : 4.0
//*****************************************************************************/
 /**
  IMPORTANTE!!!!
 
  Para usar este componente se puede definir un DIV con el ID asignado al momento de crear el objeto.
  
 var o = new Paginacion("myPaging");
  ...
 <div id="myPaging"></div>
 
  Para iniciar la paginacion, se debe invoca  el metodo 'pagina'	

  o.paginar(p_paginaActual, p_cantLineas, p_lineasPagina, p_maxLineasPagina, url)
  
  p_paginaActual	: Numero de pagina mostrado.
  p_cantLineas		: Cantidad de lineas totales.
  p_lineasPagina	: Cantidad de lineas por pagina
  p_maxLineasPagina	: Maxima cantidad de lineas por pagina que se pueden mostrar
  url | callback	: URL del la pagina que resuelve la consulta de los datos a paginar o funcion a llamar al cambiar de pagina.(ver notas mas arriba).
  
  El numero de pagina seleccionada y la cantidad de registros por pagina se
  devuelven en los parametros "numeroPagina" y "registrosPorPagina" respectivamente. En caso de necesitar
  submitir la pagina, incluir el DIV de paginacion dentro del form.
 
  Si en lugar de una URL se desea pasar una funcion para que sea llamada cuando el usuario indique cambiar de pagina, se debe generar un String 
  con el formato de la llamda a la funcion, el String solicitado sera el utilizado para armar el link. 
  Puede chequearse el resultado en la barra de status del explorador pasando mause sobre los links de las paginas.
  No olvidar, la funcion indicada debe esperar por lo menos dos pgnParametros: numero de pagina solicitado y cantidad de registros por pagina, 
  siempre seran los dos ultimos pgnParametros respectivamente.
  Si el string pasado no usa pgnParametros, se agregaran por lo menos los dos requeridos por este componente.
  Por ejemplo:    
  		Si se pasa	 			Quedara como  
  		func() ;				func(<pagina>, <regxPag>);
  		func 					func(<pagina>, <regxPag>);
  		func(<parm1>) 			func(<parm1>, <pagina>, <regxPag>);
 
  Si no se crea el div en la poagina, la funcion "paginar" devuelve el HTML del componente.
 */ 
 
 /* Constantes de pgnParametros propios de la paginacion */
	var PGN_LINES_X_PAGE = "registrosPorPagina";
	var PGN_ACTUAL_PAGE = "numeroPagina";
	/**
	* Constantes generales
	*/   	
	var PGN_PARAM_DELIMITER = "?";
	var PGN_PARAM_SEPARATOR = "&";	

	var PGN_MODO_JS = 0;
	var PGN_MODO_HTML = 1;

	var PGN_PAGE_LINKS = 7;
	     
	var $pgnRegistro = new Array();	//Registro de Paginaciones.
 
/* Objeto a crear para manejar la paginacion de datos */ 
function Paginacion(id) {		
	/* Variables de control */
    this._number = $pgnRegistro.push(this) - 1;
    this.id = id + this._number;
    this.div_paginacion = id;
    this.lppCmb = id + "SEL";
	this.paginaActual;			        //Pagina actualmente seleccionada.
	this.lpp;					        //Lineas por paginaseleccionadas.
	this.modoPaginacion;		        //Modo en que opera la paginacion.
	this.url;					        //La URL.
	this.cantidadLineas = 0; 	        //Numero total de lineas resultado de la tabla.
	this.cantidadPaginas = 0; 	        //Numero total de Paginas.
	this.lastIni = 1;                   //Ultima pagina que encabezo el dibujo de links individuales.
	// Ejecutar este metodo para que se iniciar la paginacion.
	// Parametros:
	//      p_paginaActual  : Numero de pagina seleccionada.
	//      p_cantLineas    : Cantidad de lineas total de la tabla a paginar.
	//      p_lineasPagina  : Cantidad de lineas a mostrar por pagina.
	//      p_maxLineasPagina: NO SE USA MAS - Queda por compatibilidad.
	//      url             : Link a llamar/funcion a ejecutar.
	this.paginar = 	function (p_paginaActual, p_cantLineas, p_lineasPagina, p_maxLineasPagina, url){
		var html;	
		// paginar solo si hay registros.
		if (p_cantLineas > 0) {
			//Se detrmina el modo.
			this._setMode(url);
			//Se setea la URL
			this._setUrl(url);
			//Se dibuja
			this.cantidadLineas = p_cantLineas;
			this._draw(p_paginaActual, p_lineasPagina);
		}		
	}
	this._setMode =	function (url) {
		this.modoPaginacion = PGN_MODO_HTML;
		//Si no hay parentesis se asume que se paso una funcion javascript
		if ((url.indexOf(".") == -1) || (url.indexOf(")") != -1)) this.modoPaginacion = PGN_MODO_JS;
	}

	this._setUrl = function(url) {
		var chr = url.charAt(url.length-1);
		this.url = url;
		if (( chr != ")") && (chr != ";") && (this.modoPaginacion == PGN_MODO_JS)) this.url = url.concat("()");
	}
	//Dibuja la paginacion.
	this._draw = function(page, lpp) {
	    this.lpp = lpp;
	    this.paginaActual = page;
	    this.lastIni = page - Math.floor((PGN_PAGE_LINKS-1)/2);
		if (this.lastIni <= 0) this.lastIni = 1;
	    // cantidad de paginas
	    this.cantidadPaginas = Math.floor(this.cantidadLineas / lpp);
	    if (this.cantidadLineas % lpp > 0) this.cantidadPaginas++;
	    var html = this._fullHTML(this.cantidadPaginas);
	    var div = document.getElementById(this.div_paginacion);
	    if (div != undefined) div.innerHTML = html;
	}
	//Genera la estructura de paginacion.
	this._fullHTML = 	function (p_total) {
        var html = "<div>"
        html += this._lppDesign();
        html += this._pslDesign(p_total);
        html += "</div>";
        //Parametros para cuando se submite la pagina.
        html += "<input id=\"" + PGN_ACTUAL_PAGE + "\" name=\"" + PGN_ACTUAL_PAGE + "\" type=\"hidden\" value=\"" + this.paginaActual + "\">"
        html += "<input id=\"" + PGN_LINES_X_PAGE+ "\" name=\"" + PGN_LINES_X_PAGE+ "\" type=\"hidden\" value=\""+ this.lpp + "\">"
        return html;
	}
	//Carga los valores de control y llama a la funcion de usuario.
	this.solveRequest = function(page, lpp) {
	    if (lpp != this.lpp) {
	        //Se cambiaron las lineas por pagina, se redibuja todo.
	        var lppCmb = document.getElementById(this.lppCmb);
	        lpp = lppCmb.options[lppCmb.selectedIndex].value;
	        this._draw(page, lpp);
	    }
	    if (page != this.paginaActual) {
	        //Se cambia la pagina actual a la solicitada y se redibuja el rango.
	        this.paginaActual = page;
	        this.pslChange(this.lastIni, this.cantidadPaginas);
	    }
	    //Se prepara la URL.
	    var aUrl = this._addParametro(this.url, PGN_ACTUAL_PAGE, page);
	    aUrl = this._addParametro(aUrl, PGN_LINES_X_PAGE, lpp);
	    if (this.modoPaginacion == PGN_MODO_JS) {
	        //Se ejecuta la funcion Javascript.
	        eval(aUrl);
	    } else {
	        //Se setea la URL en el navegador.						     
	        document.location.href = aUrl;
	    }

	}	
	//Arma el HTML para las lineas por pagina incluyendo el contenedor.
	this._lppDesign = function() {
	    var html = "<div class=\"resultxpage\">";
	    html += "<span> Resultados por p&aacutegina: </span>";
	    html += "<select id=\"" + this.lppCmb + "\" class=\"idResultados\" onchange=\"javascript:solveRequest(" + this._number + ",1,0)\">";
	    for (var i = 10; i <= 100; i+=10) {
	        html += "<option value=\"" + i + "\"";
	        if (this.lpp == i) html += " selected=\"selected\"";
	        html += ">" + i + " </option>";
	    }
	    html += "</select>";
	    html += "</div>";

	    return html;
	}	
	//Arma el HTML de un grupo
	this._grpHTML = function(p_ini, p_total) {
	    this.lastIni = p_ini;
	    var html = "";
		var max = p_ini + PGN_PAGE_LINKS;	    
	    if (p_total < max) max = p_total + 1;
	    for (var i = p_ini; i < max; i++) {
	        html += "<button ";
	        if (i == this.paginaActual) html += "class=\"actual\""
	        html += "onclick=\"javascript:solveRequest(" + this._number + "," + i + "," + this.lpp + ")\">" + i + "</button>"
	    }
	    return html;
	}
	this._prevLink = function(p_ini, p_total) {
        var html = "";
        var ppPrev = p_ini - 1;                    //Pagina anterior a la primera del rango actual.
        var ppFarPrev = p_ini - PGN_PAGE_LINKS;    //Primera pagina del rango anterior menor.
        //Se dibuja el link para corrimiento individual de paginas. (+1 pagina)
        if (ppPrev > 0) {
            if (ppFarPrev > 0) html += "<button onclick=\"javascript:pslChange(" + this._number + "," + ppFarPrev + "," + p_total + ")\">&lt;&lt;</button>";
            html += "<button onclick=\"javascript:pslChange(" + this._number + "," + ppPrev + "," + p_total + ")\">&lt;</button>";	                            
        }
        return html;							
	}
	this._nextLink =	function (p_ini, p_total) {
        var html = "";
        var ppNext = p_ini + PGN_PAGE_LINKS;        //Primera pagina del siguiente rango.
        var ppFarNext = ppNext + PGN_PAGE_LINKS;    //Primera pagina del rango segundo mayor.
        //Se dibuja el link para corrimiento individual de paginas. (+1 pagina)
        if (p_total >= ppNext) {
            html += "<button onclick=\"javascript:pslChange(" + this._number + "," + (p_ini+1) + "," + p_total + ")\">&gt;</button>";
            if (p_total >= ppFarNext) html += "<button onclick=\"javascript:pslChange(" + this._number + "," + ppNext + "," + p_total + ")\">&gt;&gt;</button>";
        }														
		return html;
	}	
	//Arma los links para las paginas dentro de su contenedor. (Pages Section Links (psl))
	//Se llama solo al cargar la pagina.
	this._pslDesign = function(p_total) {
	    var html = "<div class=\"pageselector\" id=\"" + this.id + "\">"
	    //Dibujo el link a paginas anteriores si corresponde.
	    html += this._prevLink(this.lastIni, p_total);
	    //Dibujo los links de paginas.
	    html += this._grpHTML(this.lastIni, p_total)
	    //Dibujo el link de paginas siguientes si corresponde.
	    html += this._nextLink(this.lastIni, p_total);	    
	    html += "</div>";
	    return html;
	}	
	//Modifica el HTML de las paginas individuales.
	//Se llama al utilizar el componente del lado del cliente.
	this.pslChange = function(p_ini, p_total) {
	    //Dibujo el link a paginas anteriores si corresponde.
	    var html = this._prevLink(p_ini, p_total);
	    //Dibujo los links de paginas.
	    html += this._grpHTML(p_ini, p_total)
	    //Dibujo el link de paginas siguientes si corresponde.
	    html += this._nextLink(p_ini, p_total);

	    document.getElementById(this.id).innerHTML = html;
	}
    //-------------------------------------------------------------------
	this._addParametro =	function (p_url, p_key, p_value) {
	    var ret = p_url;
		if (this.modoPaginacion == PGN_MODO_HTML) {
			//Se agregan parametros a la URL
		    var aux = PGN_PARAM_SEPARATOR;
			if (ret.indexOf(PGN_PARAM_DELIMITER) == -1) aux = PGN_PARAM_DELIMITER;
		    ret += aux + p_key + "=" + p_value;
		} else {
			//Se agregan parametros a la funcion JavaScript.
			var arrAux = p_url.split(")");
			ret = arrAux[0];
			if (ret.charAt(ret.length-1) != "(") ret += ", "; 
			ret += p_value + ");";
		}
	    return ret;
	}
}
/**********************************************************************************
 ***	FUNCIONES GLOBALES 
***********************************************************************************/
function solveRequest(pNbr, page, lpp) {
	$pgnRegistro[pNbr].solveRequest(page, lpp);
}
function pslChange(pNbr, p_ini, p_total) {
	$pgnRegistro[pNbr].pslChange(p_ini, p_total);
}