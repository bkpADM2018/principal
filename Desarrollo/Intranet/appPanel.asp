<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosAplicaciones.asp"-->
<%
'***************************************************************************************************************'
'******************************************* INICIO DE PAGINA **************************************************'
'***************************************************************************************************************'
Dim vTipoDocumento, cAutorizaciones, cPoseidon, cCompliance, cPermisos
if isToepfer(session("KCOrganizacion")) then
	set cAutorizaciones = new clsAutorizaciones
	cAutorizaciones.drawPanel()
	set cPoseidon = new clsPoseidon
	cPoseidon.drawPanel()
	set cCompliance = new clsCompliance
	cCompliance.drawPanel()
	set cPermisos = new clsPermisos
	cPermisos.drawPanel()
else
    if ((UCase(session("Usuario")) = "ADUANASL") or (UCase(session("Usuario")) = "ADUANAROS") or (UCase(session("Usuario")) = "ADUANABBA")) then
        set cAduana = new clsAduana
	    cAduana.drawPanel()
    else
	    set cPoseidon = new clsPoseidon
		cPoseidon.drawPanel()
	    set cPagos = new clsPagos
	    cPagos.drawPanel()
	    set cAFIP = new clsAFIP
	    cAFIP.drawPanel()
	    set cContratos = new clsContratos
	    cContratos.drawPanel()
    end if
end if	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="css/main.css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript">
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
	
	var ch = new channel();	
	var vTipoDocumento = new Array();
	var arrTipoDoc = new Array();
	var indexPendientes = 0;
	var objDoc = "";
	var objDoc2 = "";
	
	function bodyOnLoad(){
		//Verifico que se halla instanciado el quadrante de Autorizaciones, caso contrario no llama al ajax		
		<% if (existQuadrant(QUADRANT_HOME_POSEIDON)) then %>
			  ch.bind("getCamionesEnPlanta_AJAX.asp" ,"CallBack_getCamionesEnPlanta()");
			  ch.send();
		<% end if %>
		<% if (existQuadrant(QUADRANT_HOME_AUTORIZACIONES)) then
			  if (not isFormSubmit()) then
				  vTipoDocumento = getDocumentoFirmar()
				  for i=0 to UBound(vTipoDocumento)-1 %>
		   			  vTipoDocumento.push('<% =vTipoDocumento(i) %>');
			  <%  next
			  end if %>
			  var tipoDoc = vTipoDocumento.pop();
			  ch.bind("comprasAutorizaciones.asp?Tipo=" + tipoDoc + "&accion=<%=ACCION_PROCESAR%>&origen=1","CallBack_getAutorizaciones('"+tipoDoc+"')");
			  ch.send();
		<% end if %>	
					
	}
	function CallBack_getCamionesEnPlanta(){
		var rtrn = ch.response();
		var arr = rtrn.split("|");
		if(rtrn.length > 0){
			objDoc2 = $('#titleQuadrant_<%=QUADRANT_HOME_POSEIDON%>').next();
			for (i in arr) {
				$(objDoc2).children().text(arr[i]);
				objDoc2 = objDoc2.next();
			}
		}	
	}
	function CallBack_getAutorizaciones(pTipoDoc){
		var rtrn = ch.response();
		var arr = rtrn.split(";");
		var auxDsDocumento = "";
		var countDoc = 0;
		var flagNewFormat = false;
		if(rtrn.length > 0){
			for (i in arr) {
				var val = arr[i].split("|");
				if (val[11] == 0){					
					if (existeTipoDocumento(val[0]) == true){
						countDoc++;
					}
					else{
						if (arrTipoDoc.length <= 3){							
							if (arrTipoDoc.length > 0) {								
								showTypeDocument(auxDsDocumento, countDoc, arrTipoDoc.pop()); 
								countDoc = 0;
							}
							countDoc++;
							arrTipoDoc.push(val[0]);
							auxDsDocumento = val[9];
						}
					}
				}
				else{
					flagNewFormat = true;
					showTypeDocument(val[9], val[11], val[0]);
				}
			}
			if ((arrTipoDoc.length <= 3)&&(flagNewFormat == false)) showTypeDocument(auxDsDocumento, countDoc, arrTipoDoc.pop());
		}
		if (vTipoDocumento.length > 0) {
			var tipoDoc = vTipoDocumento.pop();
			ch.bind("comprasAutorizaciones.asp?Tipo=" +  tipoDoc + "&accion=<%=ACCION_PROCESAR%>&origen=1","CallBack_getAutorizaciones('"+tipoDoc+"')");
			ch.send();
		}
	}
	
	//function showTypeDocument : se encarga de mostrar el Tipo de Documento y las cantidades que tiene para firmar
	function showTypeDocument(pDsDocument, pCantidad, pTipoDoc){				
		if (objDoc == "")
			objDoc = $('#titleQuadrant_<%=QUADRANT_HOME_AUTORIZACIONES%>').next();
		else
			objDoc = objDoc.next();
		var descripcionDoc = pDsDocument;
		var pos = pDsDocument.indexOf("("+pTipoDoc+")");
		if (pos != -1) descripcionDoc  = pDsDocument.substring(0,pos);
		$(objDoc).children().text(descripcionDoc + " ("+pCantidad+")");
		indexPendientes = parseInt(indexPendientes) + parseInt(pCantidad);
		document.getElementById('titleQuadrant_<%=QUADRANT_HOME_AUTORIZACIONES%>').innerHTML = "Pendientes ("+indexPendientes+")";	
	}
	
	//function existeTipoDocumento : Determina si un tipo de documento ya se cargo al Array. 
	function existeTipoDocumento(pTipo){
		if(isFirefox == true){
			if( arrTipoDoc.indexOf(pTipo) >= 0) {
				return true
			}
			else{
				return false
			}
		}
		else{
			for(var i=0;i<arrTipoDoc.length;i++){
				if(arrTipoDoc[i] == pTipo) {
					return true
				}
				else{
					return false
				}
			}
		}
	}
	function abrirRegistrosBalanza(p_pto){
		var link1 = "Poseidon/Embarques/registroBalanzaOnLine.asp?pto=" + p_pto + "&bza=<% =EMBARQUE_BALANZA_1 %>"; 
		var link2 = "Poseidon/Embarques/registroBalanzaOnLine.asp?pto=" + p_pto + "&bza=<% =EMBARQUE_BALANZA_2 %>"; 
	    var Ancho = screen.width - 950;
	    window.open(link1, "_blank", 'width=950,height=750,left=10,top=100,toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=no');
	    window.open(link2, "_blank", 'width=950,height=750,left=' + Ancho + ',top=100,toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=no');
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
</body>
</html>
