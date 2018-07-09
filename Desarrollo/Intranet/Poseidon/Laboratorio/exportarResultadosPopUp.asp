<!--#include file="../../Includes/procedimientosCompras.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosuser.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosLaboratorio.asp"-->
<!--#include file="../../Includes/procedimientosSeguridad.asp"-->
<%
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
'-----      COMIENZO DE LA PAGINA
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Dim pto,arrFile,valParameterPath,accion,chkDetalle,valParameterFile,g_Camara,valMax

Call initTaskAccessInfo(TASK_POS_INFO_ANALISIS, session("DIVISION_PUERTO"))

pto = GF_Parametros7("pto","",6)
myHoy = Left(session("MmtoDato"), 8)
myHasta = myHoy

%>
<HTML>
<HEAD>
	<meta http-equiv="X-UA-Compatible" content="IE=9">
	
	<TITLE>Laboratorio - Exportar archivo camara - <% =pto %> </TITLE>
    
	<link rel="stylesheet" href="../../css/ActisaIntra-1.css" type="text/css" />
    <link rel="stylesheet" href="../../css/main.css" type="text/css"/>
    <link rel="stylesheet" href="../../css/calendar-win2k-2.css" type="text/css">	
    <style type="text/css">
        
        .color_Black {        	
	        color: #333;
	        font-size: 18px;
        }
        .color_Golden {        	
	        color: GoldenRod;
	        font-size: 18px;
	    }
    </style>
	<script type="text/javascript" src="../../scripts/channel.js"></script>
	<script type="text/javascript" src="../../scripts/controles.js"></script>	
    <script type="text/javascript" src="../../scripts/calendar.js"></script>
	<script type="text/javascript" src="../../scripts/calendar-1.js"></script>
	<script type="text/javascript">
		var MODO_FECHA = "F";
		var ch = new channel();	
		var segments;		
        var ch = new channel();	
	    var maxSegments;
	    var currSegment=0;
	    var MS_X_DAY = 86400000 //Milisegundos por día.	
	    var d = new Date();
	    var changeFilters = false;
        var lst;
		var errFlag = false;
		
        function procesarExportacion(p_Pto){
            var modo = document.querySelector('input[name="modo"]:checked').value;
            //Controlo las fechas
			if (modo == MODO_FECHA) {
				//solicitud por fecha
				var fd = document.getElementById("fdymd").value;
				var fh = document.getElementById("fhymd").value;
				var mc = document.getElementById("muestraComercial").checked;
				var mb = document.getElementById("muestraBiotecnologia").checked;
				if (fd <= fh) {
					if (mc || mb) {
						generateFile(modo);
					} else {
						alert("Error: Debe seleccionar al menos un tipo de muestra para la solicitud.");
					}
				} else {
					alert("Error: El período ingresado no es correcto!");
				}
			} else {
				//solicitud por muestras
				var txt = document.getElementById("lstMuestras").value;
				txt = txt.replace(';', ',');
				txt = txt.replace('-', ',');
				if (txt.length > 0) {
					lst = txt.split(',');
					generateFile(modo);
				} else {
					alert("Error: Debe incluir al menos un numero de muestra.");
				}
			}
        }
        
		function generateFile(pModo) {
			document.getElementById("actionLabel").style.textAlign = 'center';		    
		    document.getElementById("actionLabel").className = "confirmsj";	
			document.getElementById("actionLabel").style.visibility = 'visible';			
		    document.getElementById("actionLabel").innerHTML = "Inicializando... ";
		    calculateSegments(pModo);
		    document.getElementById("accion").value = "<%=ACCION_PROCESAR %>";   
		    generateSegment(currSegment, pModo)
	    }	
        function generateSegment(currSegment, pModo) {
		    document.getElementById("actionLabel").innerHTML = "Recopilando datos...  ( " + (currSegment+1) + " / " + (maxSegments+1) + " )";
			if (pModo == MODO_FECHA) {
				var strFecha = document.getElementById("fdymd").value;
				var d = strFecha.substr(6,2);
				var m = strFecha.substr(4,2)-1;
				var y = strFecha.substr(0,4);
				var fd = new Date(y, m, d, 0, 0, 0, 0);
				var d = new Date(fd.getTime() + (MS_X_DAY*currSegment));
				document.getElementById("fecContableDS").value = d.getDate();
				document.getElementById("fecContableMS").value = d.getMonth()+1;
				document.getElementById("fecContableAS").value = d.getFullYear();
			} else {
				document.getElementById("muestra").value = lst[currSegment];
			}			
		    document.getElementById("frmSel").action="exportarResultadosE1.asp";
		    document.getElementById("frmSel").target = "ifrmXLS";		    
            document.getElementById("frmSel").submit();
	    }
       	function calculateSegments(pModo) {
			if (pModo == MODO_FECHA) {
				var strFechaD = document.getElementById("fdymd").value;
				var dd = strFechaD.substr(6,2);
				var md = strFechaD.substr(4,2)-1;
				var yd = strFechaD.substr(0,4);
				var fd = new Date(yd, md, dd, 0, 0, 0, 0);
				var strFechaH = document.getElementById("fhymd").value;
				var dh = strFechaH.substr(6,2);
				var mh = strFechaH.substr(4,2)-1;
				var yh = strFechaH.substr(0,4);
				var fh = new Date(yh, mh, dh, 0, 0, 0, 0);
				maxSegments = Math.round((fh.getTime() - fd.getTime())/MS_X_DAY)
			} else {
				maxSegments = lst.length-1;
			}
	    }
       function generateSegment_callback(pflag, pModo) {
            //Si algun segmento viene con error, salva la marca y sigue con el siguiente segmento para ver si ha más errores.			
			if (pflag == "True") errFlag = true;
			if (currSegment < maxSegments) {
				currSegment += 1;
				document.getElementById("accion").value = "";
				generateSegment(currSegment, pModo);
			} else {
				//Finalizo la generacion de datos, si hubo error mando el archivo de errores por mail, sino genero el reporte de datos final.
				if (errFlag) {					
					document.getElementById("actionLabel").innerHTML = 'Se han detectado errores!, se ha enviado un mail con la lista de items a corregir.';
					document.getElementById("actionLabel").className = "errormsj";					
					ch.bind("generarSolicitudesEnvioMailErroresAjax.asp?pto=<%=pto%>","CallBack_getMail()");
					ch.send();
				} else {
					document.getElementById("maxSegment").value = currSegment;
					generateReport();
				}
			}
	    }
        function generateReport(){            		    
		    document.getElementById("actionLabel").innerHTML = "Generando Reporte...";
            document.getElementById("accion").value = "";
            document.getElementById("frmSel").action="exportarResultadosE2.asp";            
            document.getElementById("frmSel").target="ifrmXLS";
            document.getElementById("frmSel").submit();
        }
        function generateReport_callback() {
        	sendMailFileExportacion();
        }
        function sendMailFileExportacion() {		    
		    document.getElementById("actionLabel").innerHTML = "Enviando mail de los archivos generados";
		    ch.bind("generarSolicitudesCamaraEnvioMailAjax.asp?pto=<%=pto%>","CallBack_getMail()");
		    ch.send();
	    }
        function CallBack_getMail(){
		    var rtrn = ch.response();
		    if (rtrn != '<%=FILE_MISSING%>'){							    
			    if (!errFlag) document.getElementById("actionLabel").innerHTML = "Los archivos generados se enviaron a " + rtrn;			    
		    }
		    else{
				document.getElementById("actionLabel").className = "errormsj";
			    document.getElementById("actionLabel").innerHTML = "Se produjo un error al intentar mandar el mail";			    
		    }
			restartAttForm();
	    }
        function restartAttForm(){
		    document.getElementById("accion").value = '<%=ACCION_SUBMITIR%>';
		    var obj = document.getElementById("frmSel");
		    obj.action = "exportarResultadosPopUp.asp";
		    obj.removeAttribute('target');
			errFlag = false;
	    }	
		
        function lightOn(tr) {
		    tr.className = "color_Golden";
	    }
	
    	function lightOff(tr) {		    
            tr.className = "color_Black";
	    }
        function MostrarCalendarioDesde(p_objID, funcSel) {
				
					var dte= new Date();		    	    
					var elem= document.getElementById(p_objID);
					if (calendar != null) calendar.hide();		
					var cal = new Calendar(false, dte, funcSel, CerrarCal);
					cal.weekNumbers = false;
					cal.setRange(1993, 2045);
					cal.create();
					calendar = cal;		
					calendar.setDateFormat("dd/mm/y");
					calendar.showAtElement(elem);
			}	

			function CerrarCal(cal) {
				cal.hide();
			}

			function SeleccionarCalDesde(cal, date) {
				var str= new String(date);
				document.getElementById("fd").value = str;
				document.getElementById("fdymd").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
			    if (cal) cal.hide();
			}	
								
			function MostrarCalendarioHasta(p_objID, funcSel) {				
				var dte= new Date();		    	    
				var elem= document.getElementById(p_objID);
				if (calendar != null) calendar.hide();		
				var cal = new Calendar(false, dte, funcSel, CerrarCal);
				cal.weekNumbers = false;
				cal.setRange(1993, 2045);
				cal.create();
				calendar = cal;		
				calendar.setDateFormat("dd/mm/y");
				calendar.showAtElement(elem);
			}	

			function SeleccionarCalHasta(cal, date) {
				var str= new String(date);
				document.getElementById("fh").value = str;
				document.getElementById("fhymd").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
			    if (cal) cal.hide();
			}	
			function QuitarFechaHasta(){
				document.getElementById("fh").value = "";
			    
			}	
            function QuitarFechaDesde(){
				document.getElementById("fd").value = "";			    
			}	
    
			function switchGenerador(pType) {
				if (pType == 'M') {
					document.getElementById('divQFecha').style.visibility= 'hidden';
					document.getElementById('divQFecha').style.position= 'absolute';
					document.getElementById('divQMuestra').style.visibility= 'visible';
					document.getElementById('divQMuestra').style.position= 'relative';
				} else {
					document.getElementById('divQMuestra').style.visibility= 'hidden';
					document.getElementById('divQMuestra').style.position= 'absolute';
					document.getElementById('divQFecha').style.visibility= 'visible';
					document.getElementById('divQFecha').style.position= 'relative';
				}
			}
	</script>
	
</HEAD>

<BODY>

	<form id="frmSel" name="frmSel" method="post" action="exportarResultadosPopUp.asp">

	<div class="col66"></div>

	<div ><% Call showMessages() %></div>	
	
    <div class="free10"></div>	
	<div class="col06"></div>
	<div >
		Seleccione la modalidad de Generaci&oacute;n: 
		<input type="radio" name="modo" value="F" onclick="javascript:switchGenerador('F')" checked> por fecha. 
		<input type="radio" name="modo" value="M" onclick="javascript:switchGenerador('M')"> por muestra.
	</div>	
	<div id="divQFecha">
		<div class="free10"></div>	
		<div class="col06"></div>
		<div >Seleccione el periodo a incluir en las solicitudes:</div>
		<div class="free10"></div>	
		<div class="col06"></div>
		<div>
			<div class="col36">
				Fecha Desde:
				<input type="text" name="fd" id="fd" onClick="javascript:MostrarCalendarioDesde('fd', SeleccionarCalDesde)" value="<% =GF_FN2DTE(myHoy) %>" size="10">
				<input type="hidden" name="fdymd" id="fdymd" value="<% =myHoy %>">
			</div>
			<div class="col36">
				Fecha Hasta:
				<input type="text" name="fh" id="fh" onClick="javascript:MostrarCalendarioHasta('fh', SeleccionarCalHasta)" value="<% =GF_FN2DTE(myHasta) %>" size="10">
				<input type="hidden" name="fhymd" id="fhymd" value="<% =myHasta %>">
			</div>
		</div>
		<div class="free10"></div>	
		<div class="col06"></div>
		<div>
			<div class="col36">
				Muestra comercial:&nbsp&nbsp
				<INPUT type="checkbox" name="muestraComercial" id="muestraComercial" style="cursor:pointer;" checked />
			</div>
			<div class="col36">
				Muestra Biotecnologia:&nbsp&nbsp
				<INPUT type="checkbox" name="muestraBiotecnologia" id="muestraBiotecnologia" style="cursor:pointer;" checked />
			</div>					
		</div>
	</div>
	<div id="divQMuestra" style="visibility:hidden; position:absolute;">
		<div class="free10"></div>	
		<div class="col06"></div>
		<div >Indique las meustras a incluir (lista separada por comas):</div>
		<div class="free10"></div>
		<div class="col06"></div>
		<div>
			<textarea id="lstMuestras" name="lstMuestras" cols="70" rows="5"></textarea>
		</div>
	</div>
	<div class="free20"></div>
	<div class="col35"></div>		
	<div>
		<input type="button" onclick="procesarExportacion('<%=pto%>')" value="Exportar"/>
	</div>	
	<div class="free20"></div>
	<div id="actionLabel" class="confirmsj" style="visibility:hidden; "></div>
	<input type="hidden" id="accion" name="accion" value="<% =accion%>" />	
	<input type="hidden" id="pto" name="pto" value="<% =pto %>" />	
	<input type="hidden" id="muestra" name="muestra" />
	<input type="hidden" id="fecContableDS" name="fecContableDS" />
	<input type="hidden" id="fecContableMS" name="fecContableMS" />
	<input type="hidden" id="fecContableAS" name="fecContableAS" />
	<input type="hidden" id="maxSegment" name="maxSegment" />
	<input type="hidden" id="usr" name="usr" value="<% =session("Usuario") %>" />        
    
    
    <div id="respuestaImportacion" style="width:100%;visibility:hidden;"></div>
    </form>
    <iframe name="ifrmXLS" id="ifrmXLS" width="0px" height="0px" style="visibility:hidden"></iframe>
</BODY>
</HTML>
