<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosMap.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 01/12/2008                         -->
'<!--               Pagina     : Map.ASP		                    -->
'<!--               Descripcion: Mapa para destacar zonas de interes-->
'<!------------------------------------------------------------------->
%>
<html>
<head>
<style type="text/css">
	.infowindow { font-size: smaller; font-family:Verdana; }
	.infolink { margin:0 25px; }
</style>
   <script src="http://maps.google.com/maps?file=api&amp;v=2&amp;sensor=true&amp;key=ABQIAAAAt8GkK4aA78xxsJkqioLBAxRo_FjsXXdVepeX-tz_U9bCCpIBsxTJzUPVoarKphXxi-XRqWQA8ZAIvw" type="text/javascript"></script>
</head>
<body onload="buildMap();load();" onunload="GUnload()">
<%
if session("Usuario") = "" then
	Response.Write "La session ha expirado, por favor haga click <a href='http://bai-sys-1/actisaintra/'>aqui</a> para actualizar los datos de usuario."
	Response.End
end if	
%>
<Link REL=stylesheet href="CSS/Iwin.css" type="text/css">
<script language="JavaScript" src="Scripts/channel.js"></script>
<script language="JavaScript" src="Scripts/iwin.js"></script>
<script type="text/javascript">
//<![CDATA[
var ch = new channel();
var map, actual, actualPolygon, actualPolyline;
var gmarkers = [];
var gpolygons = [];
var gpolylines = [];
var count = 0;
var countPolygon = 0;
var countPolyline = 0;
var userText;
var actualTDColor;
var actualTDIcon;
var actualTDShape;
var gColor, gIcon, gEvent, gMarker;
var isEditing = false;
var overlayOld;
var currentUser;
currentUser = '<%=session("Usuario")%>';

gColor = "ff4500";
gIcon = "images/Map Icons/K08.gif";

	if (window.navigator.userAgent.indexOf('MSIE')< 0){           HTMLElement.prototype.__defineGetter__('innerText',function () { return(this.textContent); });           HTMLElement.prototype.__defineSetter__('innerText',function (txt) { this.textContent = txt; });      }

//----------------------------------------------------------------------------------------------
function createPopUpWindow(){
	pp = new PopUpWindow('Nueva Zona', 'MapNewZone.asp', '330', '80');
}
//----------------------------------------------------------------------------------------------
function addToCombo(pId, pCd){
var pObj;
	pObj = document.getElementById('userGroup');
	pObj.options[pObj.length]=new Option(pCd, pId);	
	pObj.options[pObj.length - 1].selected = true;
	pObj = document.getElementById('shapeByGroup');
	pObj.options[pObj.length]=new Option(pCd, pId);	

}
//----------------------------------------------------------------------------------------------
function createPolygon(points) {
var letter;
	countPolygon++;
	// Set draggable markers
	var	polygon = new GPolygon(points	, "#" + gColor, 2, 1, "#" + gColor, 0.2);	
	polygon.content = countPolygon;
	gpolygons.push(polygon);


		/*
		GEvent.addListener(polygon, "dragstart", function() {
			// Close infowindow when dragging a marker
			map.closeInfoWindow();
		});

		GEvent.addListener(polygon, "dragend", function() {
			// Update gmarkers array to get the right points
			for(var i = 0; i < gpolygons.length; i++) {
				if(gpolygons[i] == polygon) {
					gpolygons.splice(i, 1, polygon);
				}
			}
		});*/
	return polygon;
}
//----------------------------------------------------------------------------------------------
function createPolyline(points) {
var letter;
	countPolyline++;
	var	polyline = new GPolyline(points	, "#" + gColor, 4);	
	polyline.content = countPolyline;
	gpolylines.push(polyline);

		GEvent.addListener(polyline, "dragstart", function() {
			map.closeInfoWindow();
		});

		GEvent.addListener(polyline, "dragend", function() {
			for(var i = 0; i < gpolylines.length; i++) {
				if(gpolylines[i] == polyline) {
					gpolylines.splice(i, 1, polyline);
				}
			}
		});
	return polyline;
}
//----------------------------------------------------------------------------------------------
function createMarker(point, icon, name) {
var letter;
	count++;
	// Set draggable markers
	var marker = new GMarker(point, {icon:icon, draggable:false, bouncy:false, dragCrossMove:false, title:name});
	marker.content = count;
	gmarkers.push(marker);

		
		GEvent.addListener(marker, "dragstart", function() {
			// Close infowindow when dragging a marker
			map.closeInfoWindow();
		});

		GEvent.addListener(marker, "dragend", function() {
			// Update gmarkers array to get the right points
			for(var i = 0; i < gmarkers.length; i++) {
				if(gmarkers[i] == marker) {
					gmarkers.splice(i, 1, marker);
				}
			}
		});
		GEvent.addListener(marker, "infowindowclose", function() {
			isEditing = false;
		});		
	return marker;
}
//----------------------------------------------------------------------------------------------
function removeFromInput(pName){
var auxName = new String();
var auxString = new String();
var pos, posFirst, posLast;
auxName = pName;
auxString = document.getElementById("pointsUsersToSave").value; 
//Buscar la posicion en la que aparece el nombre de la figura a borrar
pos = auxString.indexOf('||' + auxName + '||');
//Buscar el primer corchete que delimita las figuras
posFirst = auxString.lastIndexOf('[',pos);
//Buscar el ultimo corchete que delimita las figuras
posLast = auxString.indexOf(']',pos);
auxString = auxString.replace(auxString.substring(posFirst,posLast+1),''); 
document.getElementById("pointsUsersToSave").value = auxString;  	
}
//----------------------------------------------------------------------------------------------
function removePolyline(pName) {
	removeFromInput(pName);
	for(var i = 0; i < gpolylines.length; i++) {
		if(gpolylines[i] == actualPolyline) {
			map.removeOverlay(actualPolyline);
			gpolylines.splice(i, 1); break;
		}
	}
	if(gpolylines.length == 0){ 
		countPolyline = 0; 
	}
	else{
		countPolyline = gpolylines[gpolylines.length-1].content;
	}
	map.closeInfoWindow();
	return false;
}
//----------------------------------------------------------------------------------------------
function removePolygon(pName) {
	removeFromInput(pName);
	for(var i = 0; i < gpolygons.length; i++) {
		if(gpolygons[i] == actualPolygon) {
			map.removeOverlay(actualPolygon);
			gpolygons.splice(i, 1); break;
		}
	}
	if(gpolygons.length == 0){ 
		countPolygon = 0; 
	}
	else{
		countPolygon = gpolygons[gpolygons.length-1].content;
	}
	map.closeInfoWindow();
	return false;
}
//----------------------------------------------------------------------------------------------
function removeMarker() {
	for(var i = 0; i < gmarkers.length; i++) {
		if(gmarkers[i] == actual) {
			map.removeOverlay(actual);
			removeShape(actual.getLatLng());
			gmarkers.splice(i, 1); break;
		}
	}
	if(gmarkers.length == 0){ 
		count = 0; 
	}
	else{
		count = gmarkers[gmarkers.length-1].content;
	}
return false;
}
//----------------------------------------------------------------------------------------------
function removeShape(pCoords){
var pointsToSave = new String();
var aux = new String();
var auxPosIni, auxPosEnd, auxSubString, auxShape;
pointsToSave = document.getElementById("pointsUsersToSave").value;
//Buscar posicion de coordenadas
auxPosIni = pointsToSave.search(pCoords);
//Resta por caracteres separadores
auxPosIni = auxPosIni - 5;
//Separar el string desde la posicion indicada hasta el final
auxSubString = pointsToSave.substring(auxPosIni,pointsToSave.length);
//Buscar el separador de fin ']'
auxPosEnd = auxSubString.search(']') + 1;
//Eliminar la figura completa
auxShape = auxSubString.substring(0,auxPosEnd);
pointsToSave = pointsToSave.replace(auxShape,"");
document.getElementById("pointsUsersToSave").value = pointsToSave;
}
//----------------------------------------------------------------------------------------------
function makeHTML(marker) {
	for(var j = 0; j < gmarkers.length; j++) {
		if(gmarkers[j] == marker) {
			var point= gmarkers[j].getLatLng();
		}
	}
	var html = "<div class='infowindow'>" +
			"<b> " + userText +"<\/b>" +
			"<p><a href='#' onclick='return removeMarker()'>Quitar Marca<\/a>" +
			"<\/p><\/div>";
	return html;
}
//----------------------------------------------------------------------------------------------
function mapClick(overlay, point) {
if (overlay) {
	try {
		map.setCenter(overlay.getLatLng(), map.getZoom(), map.getCurrentMapType());
	}
	catch(e){
	}
}	
if (document.getElementById("TBL_DRAW").style.visibility != 'visible') return 0;
var typeOfDrawing, myMarker, number1, number2, mySplit, myAuxPoint, myAuxCoords;
typeOfDrawing = document.getElementById("typeOfDrawing").value;
 if(point) {
	var letter;
	letter = document.getElementById("userElement").value;
	if (letter==''){
		alert("Debe asignarle un nombre a la figura!");
		return 0;
	}	
	var icon = new GIcon();
	var iconPoint = new GIcon();
	addIcon(icon, gIcon);
	switch (typeOfDrawing){
		case 'P':
				
				var letter = "<div class='infowindow'>" +
						"<table width='60%'><tr><td align=center><b> " + document.getElementById("userElement").value + "<\/b><\/td><\/tr>" + 
						"<tr><td align=left class='reg_header8'>" + document.getElementById("userText").value + "<\/td><\/tr>" + 
						"<tr><td align=left><p><a href='#' onclick='return removeMarker()'>Quitar Marca<\/a><\/p><\/td><\/tr><\/table>" +
						"<\/div>";	
				myMarker = createMarker(point, icon, document.getElementById("userElement").value);
				map.addOverlay(myMarker);
			
				GEvent.addListener(myMarker, "click", function() {
					actual = myMarker;
					myMarker.openInfoWindowHtml(letter);
				});
				
				break;
		case 'A':
				//myMarker = createMarker(point, iconPoint);
				//map.addOverlay(myMarker);		
				//actualPolygon =	overlay;
				//break;
		case 'T':
				//myMarker = createMarker(point, iconPoint);
				//map.addOverlay(myMarker);		
				//actualPolygon = overlay;
				//break;
	}
	//Redondear punto
	myAuxPoint = point.toString();
	myAuxPoint = myAuxPoint.replace('(','');
	myAuxPoint = myAuxPoint.replace(')','');
	mySplit = myAuxPoint.split(',');
	number1 = mySplit[0].substr(0,6);
	number2 = mySplit[1].substr(0,7);
	point = new GPoint(number1,number2);
	//Controlar maxima cantidad de caracteres
	myAuxCoords = document.getElementById("pointsUsers").value;
	if ((myAuxCoords.length + myAuxPoint.length) > 600){  
		alert("No es posible almacenar más puntos!");
	}
	else{
		document.getElementById("pointsUsers").value = document.getElementById("pointsUsers").value + "$(" + point + ")" ;	
		if (typeOfDrawing=='P'){
			document.getElementById("pointsUsersToSave").value = "[" + typeOfDrawing + document.getElementById("pointsUsers").value + "||" + document.getElementById("userElement").value + "||" + document.getElementById("userGroup").value + "||" + gColor + "||" + gIcon + "||" + document.getElementById("userText").value + "||]" + document.getElementById("pointsUsersToSave").value;	
			document.getElementById("userText").value = "";
			document.getElementById("pointsUsers").value = "";
			document.getElementById("userElement").value = '';
			document.getElementById("userGroup").value = '';			
		}	
	}	

 }
 else
 {
 	switch (typeOfDrawing){
		case 'P':
				actual = overlay;
				break;
		case 'A':
				actualPolygon  = overlay;
				break;
		case 'T':
				actualPolyline  = overlay;
				break;
	}			
 }
} 
//----------------------------------------------------------------------------------------------
function addIcon(icon, pImage) {
	icon.image = pImage;
	icon.iconSize = new GSize(22, 22);
	icon.shadowSize = new GSize(10, 10);
	icon.iconAnchor = new GPoint(10, 10);
	icon.infoWindowAnchor = new GPoint(5, 10);
}

//----------------------------------------------------------------------------------------------
function clearPoints(){
	document.getElementById("pointsUsers").value = "";
}
//----------------------------------------------------------------------------------------------
function draw(){
	var myPoints  = new String();
	var mySplit, mySplit2;
	var point1 = new Number();
	var point2 = new Number();
	var myOnClick;
	var myPolygon, myPolyline, point, typeOfDrawing, myMarker;
	var GLatLngs = new Array();
	myPoints = document.getElementById("pointsUsers").value;
	if (myPoints!=""){
		typeOfDrawing = document.getElementById("typeOfDrawing").value;
		myPoints = myPoints.replace(/\(\(/g,"");
		myPoints = myPoints.replace(/\)\)/g,"");
		myPoints = myPoints.substr(1,myPoints.length);
		mySplit = myPoints.split('$');

		//Cargar puntos en el array		for (i=0; i < mySplit.length;i++){
				mySplit2 = mySplit[i].split(",");
				point1 = mySplit2[0];
				point2 = mySplit2[1];				GLatLngs[i] = new GLatLng(point1, point2);
		}

		switch (typeOfDrawing){
			case 'A': //Es un area
				GLatLngs[i] = GLatLngs[0];
				myPolygon = createPolygon(GLatLngs)
				map.addOverlay(myPolygon);
	
				var icon = new GIcon();
				var letter;
				icon.image = gIcon;
				addIcon(icon);
				var auxDel = "[" + typeOfDrawing + document.getElementById("pointsUsers").value + "||" + document.getElementById("userElement").value + "||" + document.getElementById("userGroup").value + "||" + gColor + "||" + gIcon + "||" + document.getElementById("userText").value + "||]";
				var letter = "<div class='infowindow'>" +
						"<table width='60%'><tr><td align=center><b> " + document.getElementById("userElement").value + "<\/b><\/td><\/tr>" + 
						"<tr><td align=left class='reg_header8'>" + document.getElementById("userText").value + "<\/td><\/tr>" + 
						"<tr><td align=left><p><a href='#' onclick=removePolygon('" + document.getElementById("userElement").value + "')>Quitar Zona<\/a><\/p><\/td><\/tr><\/table>" +
						"<\/div>";	
										
				myMarker = createMarker(GLatLngs[0], icon, document.getElementById("userElement").value);
				map.addOverlay(myMarker);

					GEvent.addListener (myPolygon, "mouseover", function(){
						actualPolygon = myPolygon;
						myMarker.openInfoWindowHtml(letter);
					});
				break;	
			case 'T': //Es un trayecto
				myPolyline = createPolyline(GLatLngs)
				map.addOverlay(myPolyline);
				
				var icon = new GIcon();
				icon.image = gIcon;
				addIcon(icon);
			
				var letter = "<div class='infowindow'>" +
						"<table width='80%'><tr><td align=center><b> " + document.getElementById("userElement").value + "<\/b><\/td><\/tr>" + 
						"<tr><td align=left class='reg_header8'>" + document.getElementById("userText").value + "<\/td><\/tr>" + 
						"<tr><td align=left><p><a href='#' onclick=removePolyline('" + document.getElementById("userElement").value + "')>Quitar Trayecto<\/a><\/p><\/td><\/tr><\/table>" +
						"<\/div>";	
				
				myMarker = createMarker(GLatLngs[0], icon, document.getElementById("userElement").value);
				map.addOverlay(myMarker);

					GEvent.addListener(myPolyline, "mouseover", function(){
						actualPolyline = myPolyline;
						myMarker.openInfoWindowHtml(letter);
					});			

				break;	
		}
		document.getElementById("pointsUsersToSave").value = "[" + typeOfDrawing + document.getElementById("pointsUsers").value + "||" + document.getElementById("userElement").value + "||" + document.getElementById("userGroup").value + "||" + gColor + "||" + gIcon + "||" + document.getElementById("userText").value + "||]" + document.getElementById("pointsUsersToSave").value;	
		document.getElementById("userText").value = "";		
		document.getElementById("pointsUsers").value = "";
		document.getElementById("userElement").value = '';
		document.getElementById("userGroup").value = '';			
	}	
}
//----------------------------------------------------------------------------------------------
function setColor(pObj, pColorHEX){
	gColor = pColorHEX; 
	if (actualTDColor){
		//actualTDColor.className = '';
		//actualTDColor.innerHTML = "&nbsp;&nbsp;";
	}
	else{
		actualTDColor = document.getElementById("firstColor");
	}

	actualTDColor.className = '';
	actualTDColor.innerHTML = "&nbsp;&nbsp;";
	actualTDColor = pObj;
	pObj.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
	//document.getElementById("drawColor").style.backgroundColor  = pColorHEX;	
	actualTDColor.className = 'recuadro';
}
//----------------------------------------------------------------------------------------------
function setIcon(pImg){
	gIcon = pImg.src; 
	if (actualTDIcon){
		//actualTDIcon.className = '';
	}
	else{
		actualTDIcon = document.getElementById('firstIcon');
	}
	actualTDIcon.className = '';
	actualTDIcon = pImg;	
	actualTDIcon.className = 'recuadro';
}
//----------------------------------------------------------------------------------------------
function setGroup(pObj){
	if (pObj.value == 'NEW'){
		createPopUpWindow();
	}
}	
//----------------------------------------------------------------------------------------------
function setElement(pObj){
//Controlar que el nombre no exista
var myObj = document.getElementById("shapeByName");
var textAux1 = new String();
var textAux2 = new String();
textAux1 = pObj.value;
	for (i=0;i < myObj.options.length;i++){
		textAux2 = myObj[i].value;
		if (textAux1.toUpperCase()==textAux2.toUpperCase()){
			alert("Ya existe otra figura con ese nombre!");
			document.getElementById("userElement").focus();
			//pObj.value = '';
			
			//pObj.focus();
			//return 0;
        }
	} 
}
//----------------------------------------------------------------------------------------------
function buildMap() {
	if(GBrowserIsCompatible()) {
		map=new GMap2(document.getElementById("map"),{draggableCursor:'auto',draggingCursor:'move'});
		var terr =new GMapType(G_PHYSICAL_MAP.getTileLayers(),G_PHYSICAL_MAP.getProjection(),"Relief");
		map.addMapType(terr);
		map.setCenter(new GLatLng(-35.53, -59.88), 7, G_NORMAL_MAP);
		map.addControl(new GLargeMapControl());
		map.addControl(new GMenuMapTypeControl());
		//map.enableScrollWheelZoom();
		GEvent.addListener(map, "click", mapClick);
	}
}
//----------------------------------------------------------------------------------------------
function define(pType){
	document.getElementById("typeOfDrawing").value = pType;
}
//----------------------------------------------------------------------------------------------
function save(){
var pointsToSave = new String();
pointsToSave = document.getElementById("pointsUsersToSave").value;
if (pointsToSave!=""){
		useAjax ("AjaxPlace", pointsToSave);
	}		
}
//----------------------------------------------------------------------------------------------
function useAjax_callback(pDIV){		
	document.getElementById(pDIV).innerHTML = ch.response();
	document.getElementById("pointsUsersToSave").value = "";
	document.location.reload();
}
//----------------------------------------------------------------------------------------------
function useAjax(pDIV, pParam){
	document.getElementById(pDIV).innerHTML = "";
	var link = "MapAjax.asp?param=" + pParam;
	var param = "useAjax_callback('" + pDIV + "')";
	ch.bind(link, param);
	ch.send();
}   
//----------------------------------------------------------------------------------------------
function load_callback(pDIV){		
	//recibir del response ch.response();
var myMarker, letterAux;
var myAjaxPoints = new String();
var typeOfDrawing = new String();
var mySplit0, mySplit1 ,mySplit2, mySplit3, mySplit4;
var point1, point2;
var myPoint;
var point1 = new Number();
var point2 = new Number();
myAjaxPoints = ch.response();
buildMap()
mySplit0 = myAjaxPoints.split("[");
for (i=1; i < mySplit0.length;i++){	mySplit1 = mySplit0[i].split("||");			typeOfDrawing = mySplit1[0];
			switch (typeOfDrawing){
					case 'P':
							var icon= new GIcon();
							addIcon(icon, mySplit1[6]);
							letterAux = '';
							letterAux = "<tr><td align=left><font class='SMALL' id='commentPlace'>" + mySplit1[8].substr(0,mySplit1[8].length-1) + "<\/font><\/td><\/tr>";
							if (mySplit1[8].length > 1) letterAux = letterAux + "<tr><td align=right><hr><font onclick=prepareComment(" + mySplit1[2] + ") class='SMALLLNK'>[Editar]<\/font>&nbsp;&nbsp;<font onclick=saveComment(2," + mySplit1[2] + "); class='SMALLLNK'>[Eliminar]<\/font><\/td><\/tr>";
							var letter = "<div class='infowindow' style='width:400px;height:200px;'id='showCommentDiv'>" +
								"<table width='100%' cellspacing=0 cellpadding=0>" + 
									"<tr><td align=center><font size='+1'><b>" + mySplit1[3] + "<\/b><\/font><\/td><\/tr>" + 
									"<tr><td align=left class='reg_header8'>" + mySplit1[7] + "<\/td><\/tr>" + 
									"<tr><td align=right><font onclick='saveComment(3, " + mySplit1[2] + ");' class='SMALLLNK'><hr>[Quitar Ciudad]</font><\/td><\/tr>" + 
									"<tr><td align=center><b>Comentarios<\/b><\/td><\/tr>" + 
									"<tr><td align=left>" + 
									"<div style='overflow:auto;'><table width='100%'>" + letterAux + "<\/table><\/div>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=center valign='bottom' nowrap><br><textarea rows='3' cols='50' id='newComment'></textarea>" + 
									"&nbsp;<img id='IMG_P_OK' style='cursor:pointer;' title='Agregar comentario' onclick='saveComment(0, " + mySplit1[2] + ");' src='images/ok2.gif'>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=left><div id='comments'></div><\/td><\/tr>" +
								"<\/table>" +
								"<\/div>";	

							mySplit2 = mySplit1[1].split(",");
							point1 = mySplit2[0].replace("(","");
							point2 = mySplit2[1].replace(")","");
							myPoint = new GPoint(point2,point1);				
							myMarker = createMarker(myPoint, icon, mySplit1[3]);
							map.addOverlay(myMarker);
							myMarker.bindInfoWindowHtml(letter);
							break;
					case 'A':
							var GLatLngs = new Array();
							mySplit = mySplit1[1].split('(');
							for (var j=1; j < mySplit.length;j++){
								mySplit2 = mySplit[j].split(",");
								point1 = mySplit2[0].replace("(","");
								point2 = mySplit2[1].replace(")","");
								GLatLngs[j] = new GLatLng(point1, point2);
							}					
							GLatLngs[j] = GLatLngs[1];
							gColor = mySplit1[5];
							myPolygon = createPolygon(GLatLngs)
							map.addOverlay(myPolygon);
							var icon = new GIcon();
							var letter;
							addIcon(icon, mySplit1[6]);
							
							
							letterAux = '';
							letterAux = "<tr><td align=left><font class='SMALL' id='commentPlace'>" + mySplit1[8].substr(0,mySplit1[8].length-1) + "<\/font><\/td><\/tr>";
							if (mySplit1[8].length > 1) letterAux = letterAux + "<tr><td align=right><hr><font onclick=prepareComment(" + mySplit1[2] + ") class='SMALLLNK'>[Editar]<\/font>&nbsp;&nbsp;<font onclick=saveComment(2," + mySplit1[2] + "); class='SMALLLNK'>[Eliminar]<\/font><\/td><\/tr>";
							var letter = "<div class='infowindow' style='width:400px;height:200px;'id='showCommentDiv'>" +
								"<table width='100%' cellspacing=0 cellpadding=0>" + 
									"<tr><td align=center><font size='+1'><b>" + mySplit1[3] + "<\/b><\/font><\/td><\/tr>" + 
									"<tr><td align=left class='reg_header8'>" + mySplit1[7] + "<\/td><\/tr>" + 
									"<tr><td align=right><font onclick='saveComment(3, " + mySplit1[2] + ");' class='SMALLLNK'><hr>[Quitar Area]</font><\/td><\/tr>" + 
									"<tr><td align=center><b>Comentarios<\/b><\/td><\/tr>" + 
									"<tr><td align=left>" + 
									"<div style='overflow:auto;'><table width='100%'>" + letterAux + "<\/table><\/div>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=center valign='bottom' nowrap><br><textarea rows='3' cols='50' id='newComment'></textarea>" + 
									"&nbsp;<img id='IMG_P_OK' style='cursor:pointer;' title='Agregar comentario' onclick='saveComment(0, " + mySplit1[2] + ");' src='images/ok2.gif'>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=left><div id='comments'></div><\/td><\/tr>" +
								"<\/table>" +
								"<\/div>";	

							myMarker = createMarker(GLatLngs[1], icon, mySplit1[3]);
							map.addOverlay(myMarker);
							myMarker.bindInfoWindowHtml(letter);
							break;
				
					case 'T':
							var GLatLngs = new Array();
							mySplit = mySplit1[1].split('(');
							for (var j=1; j < mySplit.length;j++){
								mySplit2 = mySplit[j].split(",");
								point1 = mySplit2[0].replace("(","");
								point2 = mySplit2[1].replace(")","");
								GLatLngs[j] = new GLatLng(point1, point2);
							}					
							gColor = mySplit1[5];
							myPolyline = createPolyline(GLatLngs)
							map.addOverlay(myPolyline);

							var icon = new GIcon();
							var letter;
							addIcon(icon, mySplit1[6]);
							
							
							letterAux = '';
							letterAux = "<tr><td align=left><font class='SMALL' id='commentPlace'>" + mySplit1[8].substr(0,mySplit1[8].length-1) + "<\/font><\/td><\/tr>";
							if (mySplit1[8].length > 1) letterAux = letterAux + "<tr><td align=right><font onclick=prepareComment(" + mySplit1[2] + ") class='SMALLLNK'>[Editar]<\/font>&nbsp;&nbsp;<font onclick=saveComment(2," + mySplit1[2] + "); class='SMALLLNK'>[Eliminar]<\/font><\/td><\/tr>";
							var letter = "<div class='infowindow' style='width:400px;height:200px;'id='showCommentDiv'>" +
								"<table width='100%' cellspacing=0 cellpadding=0>" + 
									"<tr><td align=center><font size='+1'><b>" + mySplit1[3] + "<\/b><\/font><\/td><\/tr>" + 
									"<tr><td align=left class='reg_header8'>" + mySplit1[7] + "<\/td><\/tr>" + 
									"<tr><td align=right><font onclick='saveComment(3, " + mySplit1[2] + ");' class='SMALLLNK'><hr>[Quitar Trayecto]</font><\/td><\/tr>" + 
									"<tr><td align=center><b>Comentarios<\/b><\/td><\/tr>" + 
									"<tr><td align=left>" + 
									"<div style='overflow:auto;'><table width='100%'>" + letterAux + "<\/table><\/div>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=center valign='bottom' nowrap><br><textarea rows='3' cols='50' id='newComment'></textarea>" + 
									"&nbsp;<img id='IMG_P_OK' style='cursor:pointer;' title='Agregar comentario' onclick='saveComment(0, " + mySplit1[2] + ");' src='images/ok2.gif'>" + 
									"<\/td><\/tr>" + 
									"<tr><td align=left><div id='comments'></div><\/td><\/tr>" +
								"<\/table>" +
								"<\/div>";	

							myMarker = createMarker(GLatLngs[1], icon, mySplit1[3]);
							map.addOverlay(myMarker);
							myMarker.bindInfoWindowHtml(letter);
							break;
					}
}
}
//----------------------------------------------------------------------------------------------
function prepareComment(pIdShape){
var myElement, myElementImg;
if (isEditing == false){
	document.getElementById("newComment").innerText = document.getElementById("commentPlace").innerText;
	myElement = document.getElementById("IMG_P_OK");
	myElement.setAttribute('src','images/modificar.gif');
	myElement.setAttribute('title','Modificar Comentario');
	myElement['onclick']=new Function("saveComment(1," + pIdShape + ");return false;");
	isEditing = true;
}
else{
	alert("No se pueden editar dos comentarios al mismo tiempo!");
}
}
//----------------------------------------------------------------------------------------------
function load(){
var i;
var myOptionsGroups = new String();
var myOptionsUsers = new String();
var myOptionsNames = new String();
var myOptionsCommentsByUser = new String();
var myObjGroup = document.getElementById("shapeByGroup");
var myObjUser = document.getElementById("shapeByUser");
var myObjName = document.getElementById("shapeByName");
var myObjComment = document.getElementById("shapeByComment");
//var myObjCommentsByUser = document.getElementById("commentByUser");
               for (i=0;i < myObjGroup.options.length;i++)
               {
                  if (myObjGroup.options[i].selected)
                  {
                     myOptionsGroups  = myOptionsGroups + "," + myObjGroup.options[i].value;
                  }
               } 
               for (i=0;i < myObjUser.options.length;i++)
               {
                  if (myObjUser.options[i].selected)
                  {
                     myOptionsUsers  = myOptionsUsers + ",'" + myObjUser.options[i].value + "'";
                  }
               }
               for (i=0;i < myObjName.options.length;i++)
               {
                  if (myObjName.options[i].selected)
                  {
                     myOptionsNames  = myOptionsNames + ",'" + myObjName.options[i].value + "'";
                  }
               }
               /*
               for (i=0;i < myObjCommentsByUser.options.length;i++)
               {
                  if (myObjCommentsByUser.options[i].selected)
                  {
                     myOptionsCommentsByUser  = myOptionsCommentsByUser + ",'" + myObjCommentsByUser.options[i].value + "'";
                  }
               }
               */
		var link = "MapLoadAjax.asp?byGroup=" + myOptionsGroups.substr(1,myOptionsGroups.length-1) + "&byName=" + myOptionsNames.substr(1,myOptionsNames.length-1) + "&byUser=" + myOptionsUsers.substr(1,myOptionsUsers.length-1) + "&byComment=" + myObjComment.value;
		var param = "load_callback()";
		ch.bind(link, param);
		ch.send();
}
//----------------------------------------------------------------------------------------------
function showComments(pImg, pId){
	if (document.getElementById("comments").innerHTML==''){
		document.getElementById("comments").innerHTML = "<textarea size='30' id='newComment'></textarea><img style='cursor:pointer;' onclick='saveComment(0, " + pId + ");' src='images/aceptar.gif'>";
		pImg.src = "images/arrow_up.gif";
	}
	else{
		document.getElementById("comments").innerHTML = "";
		pImg.src = "images/arrow_down.gif";
	}
}
//----------------------------------------------------------------------------------------------
function saveComment_callback(){
	//document.getElementById("IMG_P_OK").src = 'images/CheckOK.gif';
	load();
	//document.getElementById("commentPlace").innerText = document.getElementById("newComment").value; 
}
//----------------------------------------------------------------------------------------------
function saveComment(pAction, pIdShape){
	var myComment = new String();
	var myFinalComment = new String();
	var myRegExp, i;
	var mySplit;
	isEditing = false;
	/*
	Options:
		0.Save
		1.Update
		2.Delete
	*/
	switch(pAction){
		case 0:
			myComment = document.getElementById("newComment").value;
			if (myComment.length  > 0){
				document.getElementById("IMG_P_OK").src = 'images/Loading6.gif';
				myRegExp = String.fromCharCode(10);
				mySplit = myComment.split(myRegExp);
				for (i=0;i<mySplit.length;i++){
					myFinalComment = myFinalComment + '<br>' + mySplit[i];
				}
				myFinalComment = myFinalComment.substring(4,myFinalComment.length);		
				var previousComment = new String();
				previousComment = document.getElementById("commentPlace").innerHTML; 
				if (previousComment.length > 0) pAction = 1
				myFinalComment = myFinalComment + '<br>' +  previousComment;
				myFinalComment = getFinalComment(myFinalComment);
			}	
			else{
				alert('El comentario esta vacio!');
				return 0;
			}
			break;
		case 1:
			//alert('Modificar');
			myComment = document.getElementById("newComment").value;
			if (myComment.length  > 0){
				document.getElementById("IMG_P_OK").src = 'images/Loading6.gif';
				myRegExp = String.fromCharCode(10);
				mySplit = myComment.split(myRegExp);
				for (i=0;i<mySplit.length;i++){
					myFinalComment = myFinalComment + '<br>' + mySplit[i];
				}
				myFinalComment = myFinalComment.substring(4,myFinalComment.length);		
				myFinalComment = getFinalComment(myFinalComment);
			}
			else{
				alert('El comentario esta vacio!');
				return 0;
			}				
			break;	
		case 2:
			if (!confirm("Esta seguro que desea eliminar este comentario?")){
				return 0;
			}
			break;
		case 3:
			if (!confirm("Esta seguro que desea eliminar esta ciudad?")){
				return 0;
			}
			break;
	}	
	//return 0;
	//alert("MapSaveComment.asp?pAction=" + pAction + "&pIdShape=" + pIdShape + "&pComment=" + myFinalComment + "&pUser=" + currentUser);
	var link = "MapSaveComment.asp?pAction=" + pAction + "&pIdShape=" + pIdShape + "&pComment=" + myFinalComment + "&pUser=" + currentUser;
	var param = "saveComment_callback()";
	ch.bind(link, param);
	ch.send();
	
}
//----------------------------------------------------------------------------------------------
function getFinalComment(pComment){
var myReturn = new String();
myReturn = pComment;
				myReturn = myReturn.replace(/á/g,'a');
				myReturn = myReturn.replace(/à/g,'a');
				myReturn = myReturn.replace(/é/g,'e');
				myReturn = myReturn.replace(/è/g,'e');
				myReturn = myReturn.replace(/í/g,'i');
				myReturn = myReturn.replace(/ì/g,'i');
				myReturn = myReturn.replace(/ó/g,'o');
				myReturn = myReturn.replace(/ò/g,'o');
				myReturn = myReturn.replace(/ú/g,'u');
				myReturn = myReturn.replace(/ù/g,'u');
				myReturn = myReturn.replace(/Á/g,'A');
				myReturn = myReturn.replace(/À/g,'A');
				myReturn = myReturn.replace(/É/g,'E');
				myReturn = myReturn.replace(/È/g,'E');
				myReturn = myReturn.replace(/Í/g,'I');
				myReturn = myReturn.replace(/Ì/g,'I');
				myReturn = myReturn.replace(/Ó/g,'O');
				myReturn = myReturn.replace(/Ò/g,'O');
				myReturn = myReturn.replace(/Ú/g,'U');
				myReturn = myReturn.replace(/Ù/g,'U');
				myReturn = myReturn.replace(/Ñ/g,'N');
				myReturn = myReturn.replace(/ñ/g,'n');
				myReturn = myReturn.replace(/'/g,'');
				myReturn = myReturn.replace(/String.fromCharCode(34)/g,'');
return myReturn; 				
}
//----------------------------------------------------------------------------------------------
function controlKey(obj, e){
var keyPressed;
keyPressed = (document.all) ? e.keyCode : e.which;
if (((keyPressed<65 || keyPressed>122) && (keyPressed<48 || keyPressed>57)) && (keyPressed!=0 && keyPressed!=9 && keyPressed!=8 && keyPressed!=32 && keyPressed!=13)){
	alert("Caracter no valido");
	if (document.all)		e.keyCode = 0;	else		e.which = 0;	return 0;
	}
}
//----------------------------------------------------------------------------------------------
function showHide(pIdImage, pIdTable){
	if (document.getElementById(pIdTable).style.visibility == 'visible'){
		document.getElementById(pIdImage).src = 'images/flechitaderecha.gif';
		document.getElementById(pIdTable).style.visibility = 'hidden';
		document.getElementById(pIdTable).style.position = 'absolute';
	}
	else{
		document.getElementById(pIdImage).src = 'images/flechitaabajo.gif';
		document.getElementById(pIdTable).style.visibility = 'visible';
		document.getElementById(pIdTable).style.position = 'relative';
	}	
}
//----------------------------------------------------------------------------------------------
function recuadro(pObj){
	if (actualTDShape){
		//actualTDShape.className = '';
	}
	else{
		actualTDShape = document.getElementById('firstShape');
	}
	actualTDShape.className = '';
	actualTDShape = pObj;	
	actualTDShape.className = 'recuadro';
}
//----------------------------------------------------------------------------------------------
</script>

<Link REL="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
<title>Intranet ActiSA - Toepfer Maps</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<form name="frmMain" action="" method="post">
<% 'GF_TITULO("globe-128x128.png","Toepfer Maps") %>
<table width="100%" align="center" border=0 >
	<tr>
		<td valign="top">
			<div id="map" style="width: 900px; height: 850px"></div>
		</td>	
			
		<td align="left" valign=top>
			<table width="100%" style="width:220px;" cellpadding=1 cellspacing=0>
				<tr style="cursor:pointer;" class="reg_header_navK" onclick="showHide('IMG_LOAD', 'TBL_LOAD');">
					<td valign="center" width="10%">
						<img id="IMG_LOAD" src="images/flechitaderecha.gif">
					</td>
					<td valign="center">
						<font><b>Filtrar</b></font>	
					</td>
				</tr>
			</table>		
			<!--<font><b><%=GF_Traducir("Cargar")%></b></font>-->
			<table align="center" id="TBL_LOAD" style="visibility:hidden;position:absolute;">
				<tr>
					<td>
						<!--Cargar-->			
						<table align="center" border=0 width="70%" cellpadding=0 cellspacing=0>
							<tr>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r1_c2.gif"></td>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c3.gif"></td>
							</tr>				
							<tr>
								<td align="center"><img height="320" width="8" src="images/marco_r2_c1.gif"></td>				
								<td align="center" colspan="3">
									<b><%=GF_Traducir("Por Zona")%></b><br>
									<select style="width:130px;" name="shapeByGroup[]" id="shapeByGroup" multiple size=5>
										<option value=""></option>	
										<option value="ALL" selected>Todas</option>					
											<%
											sql = "Select g.idgroup, g.dsgroup  from Drawings d inner join groups g on d.drawinggroup=g.idgroup group by g.idgroup, g.dsgroup"
											call GF_BD_Control_Map (rs, cn, "OPEN", sql)
											while not rs.eof	
												%>
												<option title="<%=trim(rs("dsGroup"))%>" value="<%=trim(rs("idGroup"))%>"><%=left(trim(rs("dsGroup")),20)%></option>
												<%
												rs.movenext
											wend	
											call GF_BD_Control_Map (rs, cn, "CLOSE", sql)
											%>
									</select>	
									<br><br>
									<b><%=GF_Traducir("Por Nombre")%></b><br>
									<select style="width:130px;" name="shapeByName" id="shapeByName" multiple size=5>
										<option value=""></option>						
											<%
											sql = "Select * from Drawings order by dsShape asc"
											call GF_BD_Control_Map (rs, cn, "OPEN", sql)
											while not rs.eof	
												%>
												<option title="<%=trim(rs("dsShape"))%>" value="<%=trim(rs("dsShape"))%>"><%=left(trim(rs("dsShape")),20)%></option>
												<%
												rs.movenext
											wend	
											call GF_BD_Control_Map (rs, cn, "CLOSE", sql)
											%>
									</select>		
									<br><br>
									<b><%=GF_Traducir("Por Autor")%></b><br>
									<select style="width:130px;" name="shapeByUser" id="shapeByUser" multiple size=5>
										<option value=""></option>						
											<%
											sql = "Select distinct(owner) as shapeByUser from Drawings order by owner asc"
											call GF_BD_Control_Map (rs, cn, "OPEN", sql)
											while not rs.eof
												%>
												<option title="<%=trim(rs("shapeByUser"))%>" value="<%=trim(rs("shapeByUser"))%>"><%=left(trim(rs("shapeByUser")),20)%></option>
												<%
												rs.movenext
											wend
											call GF_BD_Control_Map (rs, cn, "CLOSE", sql)
											%>
									</select>
									<br><br>
									<b><%=GF_Traducir("Por Comentario")%></b><br>
									<input type="text" style="width:130px;" name="shapeByComment" id="shapeByComment">
								</td>
								<td align="center"><img height="320" width="8" src="images/marco_r2_c3.gif"></td>
							</tr>
							<tr>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c1.gif"></td>
								<td align="center" colspan="3" valign=bottom>
									<input type="button" onclick="load()" value= "Cargar" id=LoadPoints>
								</td>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c3.gif"></td>
							</tr>
							<tr>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r3_c2.gif"></td>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c3.gif"></td>
							</tr>
						</table>
						<!--Fin Cargar-->
					</td>
				</tr>
			</table>
			<table width="100%" style="width:220px;" cellpadding=1 cellspacing=0>
				<tr class="reg_header_navK" style="cursor:pointer;" onclick="showHide('IMG_DRAW', 'TBL_DRAW');">
					<td valign="center" width="10%">
						<img id="IMG_DRAW" src="images/flechitaderecha.gif">
					</td>
					<td valign="center">
						<font><b>Crear</b></font>
					</td>
				</tr>
			</table>				
			<table align="center" id="TBL_DRAW" style="visibility:hidden;position:absolute;">
				<tr>
					<td>
						<!--Figura a Dibujar-->
						<font><b><%=GF_Traducir("Figura a dibujar")%></b></font>
						<table border=0 width="70%" cellpadding=0 cellspacing=0>
							<tr>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r1_c2.gif"></td>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c3.gif"></td>
							</tr>				
							<tr align="center">
								<td align="center"><img height="40" width="8" src="images/marco_r2_c1.gif"></td>				
								<td id="firstShape" class="recuadro" width="33%" align="center" onclick="recuadro(this);">
									<img style="cursor:pointer;" onclick="define('P');" src="images/mapa/Punto32.png">
								</td>	
								<td align="center" valign=middle width="33%" onclick="recuadro(this);">
									<img style="cursor:pointer;" onclick="define('A');" src="images/mapa/Area32.png">
								</td>
								<td align="center" width="33%" onclick="recuadro(this);"	>
									<img style="cursor:pointer;" onclick="define('T');" src="images/mapa/Trayecto32.png">
								</td>
								<td align="center"><img height="40" width="8" src="images/marco_r2_c3.gif"></td>										
							</tr>
							<tr>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c1.gif"></td>				
								<td align="center">Punto</td>
								<td align="center">Area</td>
								<td align="center">Trayecto</td>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c3.gif"></td>										
							</tr>	
							<tr>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r3_c2.gif"></td>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c3.gif"></td>
							</tr>					
						</table>
						<!--Fin Figura a Dibujar-->
						<!--Atributos-->
						<font><b><%=GF_Traducir("Atributos")%></b></font>
						<table border=0 cellpadding=0 cellspacing=0>
							<tr>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c1.gif"></td>
								<td align="center"><img height="8" width="160" src="images/marco_r1_c2.gif"></td>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c3.gif"></td>
							</tr>

							<tr>
								<td align="center"><img height="43" width="8" height="100%" src="images/marco_r2_c1.gif"></td>
								<td align="left">
									<b><%=GF_Traducir("Color")%></b><br>
									<% call ShowColors %>
								</td>
								<td align="center"><img height="43" width="8" src="images/marco_r2_c3.gif"></td>
							</tr>
							<tr>
								<td align="center"><img height="43" width="8" src="images/marco_r2_c1.gif"></td>
								<td align="left">
									<b><%=GF_Traducir("Marca")%></b><br>
									<% call ShowIcons %>
								</td>
								<td align="center"><img height="43" width="8" src="images/marco_r2_c3.gif"></td>
							</tr>		
							<tr>
								<td align="center"><img height="38" width="8" src="images/marco_r2_c1.gif"></td>
								<td align="left">
									<b><%=GF_Traducir("Nombre")%></b><br>
									<input type="text" id=userElement onchange="setElement(this)" size=20 onkeypress="return controlKey(this, event);">
								</td>
								<td align="center"><img height="38" width="8" src="images/marco_r2_c3.gif"></td>					
							</tr>	
							<tr>
								<td align="center"><img height="38" width="8" src="images/marco_r2_c1.gif"></td>
								<td align="left">
									<b><%=GF_Traducir("Zona")%></b><br>
									<select name="userGroup" id="userGroup" onchange="setGroup(this)">
										<option value=""></option>
										<option value="NEW">Nueva...</option>
											<%
											sql = "Select * from Groups order by dsGroup asc"
											call GF_BD_Control_Map (rs, cn, "OPEN", sql)
											while not rs.eof	
												%>
												<option title="<%=trim(rs("dsGroup"))%>" value="<%=trim(rs("idGroup"))%>"><%=left(trim(rs("dsGroup")),20)%></option>
												<%
												rs.movenext
											wend	
											call GF_BD_Control_Map (rs, cn, "CLOSE", sql)
											%>
									</select>	
								</td>
								<td align="center"><img height="38" width="8" src="images/marco_r2_c3.gif"></td>					
							</tr>		
							<tr>
								<td align="center"><img height="120" width="8" height="100%" src="images/marco_r2_c1.gif"></td>
								<td align="left">
									<b><%=GF_Traducir("Descripcion")%></b><br>
									<textarea cols="20" rows="5" id="userText"></textArea>
								</td>
								<td align="center"><img height="120" width="8"src="images/marco_r2_c3.gif"></td>					
							</tr>		
							<tr>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c1.gif"></td>				
								<td align="center">
									<input type="button" onclick="draw()" value= "&nbsp;&nbsp;Hecho&nbsp;&nbsp;" id=drawPoints>
									<input type="button" onclick="clearPoints()" value="Deshacer" id=ClearPoints>
								</td>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c3.gif"></td>					
							</tr>		
							<tr>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c1.gif"></td>
								<td align="center"><img height="8" width="160" src="images/marco_r3_c2.gif"></td>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c3.gif"></td>
							</tr>		
						</table>	
						<!--Fin Atributos-->
						<!--Guardar puntos-->
						<font><b><%=GF_Traducir("Guardar cambios")%></b></font>
						<table border=0 width="100%" cellpadding=0 cellspacing=0>
							<tr>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r1_c2.gif"></td>
								<td align="center"><img height="8" width="8"  src="images/marco_r1_c3.gif"></td>
							</tr>				
							<tr>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c1.gif"></td>								
								<td align="center" colspan="3">
									<input type="button" onclick="save()" value= "&nbsp;Guardar&nbsp;" id=SavePoints>
									<input type="button" onclick="clear()" value="Cancelar" id=Cancel>
								</td>
								<td align="center"><img height="20" width="8" src="images/marco_r2_c3.gif"></td>				
							</tr>
							<tr>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c1.gif"></td>
								<td align="center" colspan="3"><img height="8" width="160" src="images/marco_r3_c2.gif"></td>
								<td align="center"><img height="8" width="8" src="images/marco_r3_c3.gif"></td>
							</tr>					
						</table>			
						<!--Fin Guardar puntos-->
					</td>
				</tr>
			</table>			
			
		</td>
	</tr>
</table>	
<div id="AjaxPlace"></div>
<input type="hidden" id="typeOfDrawing" size="5" value="P">
<input type="hidden" id="pointsUsers" size="120">
<input type="hidden" id="pointsUsersSaved" size="180">
<input type="hidden" id="pointsUsersToSave" size="180">
</center>
</form>
</body>
</html>
<%
sub ShowColors()
%>
<table cellpadding=2 cellspacing=2>
	<tr>
		<td id="firstColor" style="cursor:pointer;" bgcolor=#000000 onclick="setColor(this, '000000')">&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td style="cursor:pointer;" bgcolor=#1e90ff onclick="setColor(this, '1e90ff')">&nbsp;&nbsp;</td>
		<!--<td style="cursor:pointer;" bgcolor=#ff00ff onclick="setColor(this, 'ff00ff')">&nbsp;&nbsp;</td>-->
		<td style="cursor:pointer;" bgcolor=#ff4500 onclick="setColor(this, 'ff4500')">&nbsp;&nbsp;</td>
		<td style="cursor:pointer;" bgcolor=#3cb371 onclick="setColor(this, '3cb371')">&nbsp;&nbsp;</td>
		<td style="cursor:pointer;" bgcolor=#cc00ff onclick="setColor(this, 'cc00ff')">&nbsp;&nbsp;</td>
		<td style="cursor:pointer;" bgcolor=#330099 onclick="setColor(this, '330099')">&nbsp;&nbsp;</td>
		<td style="cursor:pointer;" bgcolor=#778899 onclick="setColor(this, '778899')">&nbsp;&nbsp;</td>
		<!--<td style="cursor:pointer;" bgcolor=#00ff7f onclick="setColor(this, '00ff7f')">&nbsp;&nbsp;</td>-->
	</tr>
</table>
<%
end sub
sub ShowIcons()
%>
<table cellpadding=2 cellspacing=2>
	<tr>
		<td><img id="firstIcon" class="recuadro" style="cursor:pointer;" height="18" width="18" src="images/Map Icons/K08.gif"	onclick="setIcon(this)"></td>
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/11.gif"	onclick="setIcon(this)"></td>
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/10.gif"	onclick="setIcon(this)"></td>
		<!--
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/4.gif"	onclick="setIcon(this)"></td>
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/32.gif"	onclick="setIcon(this)"></td>
		-->
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/1.gif"	onclick="setIcon(this)"></td>
		<td><img style="cursor:pointer;" height="18" width="18" src="images/Map Icons/toepfer.gif"	onclick="setIcon(this)"></td>
	</tr>
</table>
<%
end sub
%>