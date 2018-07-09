<!--#include file="../Includes/procedimientosCompras.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<%
    Dim FechaNombreDesde, FechaNombreHasta
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" type="text/css" href="../css/main.css" />
<style type="text/css">
.table, th, td {
	vertical-align:top;
	}

/*----------->Box ADMININSTRACION Control Panel*/
#cell {
	/*max-height: 100%;*/
	background-color: #fff;
	}
	#cell .boxround {
		padding: 7px;
/*		border: solid 2px rgba(120, 180, 40, 1);
		border-radius: 12px;
		background: rgba(255, 255, 255, 1);*/
	}
	#cell .boxround:hover {
		/*border: solid 2px rgba(46, 107, 77, 1);*/
		border-radius: 12px;
		background: rgba(230, 250, 200, 1);
	}
/*Box ADMININSTRACION Control Panel<-----------*/

.title_sec_section {
	text-align:left;
	color: #000;
	font-size: 12px;
	font-weight: bold;
	font-family: sans-serif;
}
.textoSeccion {
	text-align:left;
	vertical-align:text-top;
	color: #2e6b4d;
	font-size: 12px;
	font-family: sans-serif;
}
	.textoSeccion:hover {
	color: #78b428;
	}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.title_sec_section {
	font-size: 12px;
	font-weight: bold;
}
.textoSeccion {
	font-size: 12px;
}
</style>
<title>Reporte De Facturacion</title>
    <script type="text/javascript" src="../scripts/calendar.js"></script>
	<script type="text/javascript" src="../scripts/calendar-1.js"></script>
    <script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="../scripts/jQueryObject.js"></script>
	<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
    <script type="text/javascript">
        function bodyOnLoad() {
            //tb = new Toolbar('toolbar', 6,'images/');		
            //tb.addButton("Home-16x16.png", "Home", "cerrar()");
            //tb.draw();	
        }
        function MostrarCalendario(p_objID, funcSel) {
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

        function SeleccionarCalDesde_VCer(cal, date) {
            var str= new String(date);
            document.getElementById("dtFechaDesde_VCer").value = str;
            document.getElementById("f").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
            if (cal) cal.hide();
        }
        function SeleccionarCalHasta_VCer(cal, date) {
            var str = new String(date);
            document.getElementById("dtFechaHasta_VCer").value = str;
            document.getElementById("ff").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
            if (cal) cal.hide();
        }
        function SeleccionarCalDesde_VNeg(cal, date) {
            var str= new String(date);
            document.getElementById("dtFechaDesde_VNeg").value = str;
            document.getElementById("f").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
            if (cal) cal.hide();
        }
        function SeleccionarCalHasta_VNeg(cal, date) {
            var str = new String(date);
            document.getElementById("dtFechaHasta_VNeg").value = str;
            document.getElementById("ff").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
            if (cal) cal.hide();
        }
        function VerFechaAnalisisVN() {
            
            if (document.getElementById('target_VNeg').style.display == "none") {
                $('#target_VNeg').slideDown('slow');
                document.getElementById('target_VNeg').style.display = 'block';
            }
            else {
                document.getElementById('target_VNeg').style.display = 'none';
            }
        }
        function VerFechaAnalisisVC() {
            if (document.getElementById('target_VCer').style.display == "none") {
                $('#target_VCer').slideDown('slow');
                document.getElementById('target_VCer').style.display = 'block';
            }
            else {
                document.getElementById('target_VCer').style.display = 'none';
            }
        }        
        //------------------------------------------------
        function armarXLSNegCer(p_IsNeg){
            var auxParametro;
            if (p_IsNeg) { 
                auxParametro = -1;
            }
            else {
                auxParametro = 0;
            }
            window.open("interfacturasReportePrintXLS.asp?fechaDesde=" + document.getElementById("f").value + "&fechaHasta=" + document.getElementById("ff").value + "&modoImporte=" + auxParametro, "_blank", 'width=1000,height=500,top=10, left=10, scrollbars=YES,resizable=YES,titlebar=NO');
        }
        function generateFile() {
            window.open("interfacturasGenerarArchivo.asp", "_blank", "toolbar=no, scrollbars=yes, resizable=yes, width=750, height=650");
        }
    </script>
</head>
<body OnLoad ='bodyOnLoad();' style="overflow:hidden">
<table width="500px" border="0" align="center" cellpadding="6" cellspacing="0">
    <tr>
    <td width="50%" valign="top">
    <section id="cell">
        <!--------------------------------------------CAJA 1 (VALORES EN CERO)---------------------------------------------->
        <div class="boxround">                  
	        <table width="480" class="boxround">
                <tbody onMouseOver="" onMouseOut="" onmouseover="" onClick="VerFechaAnalisisVC();" style="cursor: pointer;">
	                <tr>	                        
	                    <td class="title_sec_section"><% =GF_TRADUCIR("Facturas con valor 0 (cero)")%></td>
	                </tr>
	                <tr>
	                    <td class="textoSeccion">
                            <%=GF_TRADUCIR("Identificar facturas  con importe cero, ya sea por errores en su emisión o para evitar la generación de notas de crédito con su correspondiente autorización.")%>
	                    </td>
	                </tr>
                </tbody>
                <tr id="target_VCer" style="display:none;"> <!--FECHA OCULTA(tr)-->
                    <td>
                        <div>
                            <div class="col36">
                                <%Response.Write GF_TRADUCIR("Fecha Desde: ") %>
                                <input type="text" name="dtFechaDesde_VCer" id="dtFechaDesde_VCer" onClick="javascript: MostrarCalendario('dtFechaDesde_VCer', SeleccionarCalDesde_VCer)" value="" size="10">
                            </div>
                            <input type="hidden" id="f" name="f" value="" />
                            <div class="col36">
                                <% Response.Write GF_TRADUCIR("Fecha Hasta: ") %>
                                <input type="text" name="dtFechaHasta_VCer" id="dtFechaHasta_VCer" onClick="javascript: MostrarCalendario('dtFechaHasta_VCer', SeleccionarCalHasta_VCer)" value="" size="10">
                            </div>
                            <input type="hidden" id="ff" name="ff" value="" />
                            <div class="col36"> <!--BOTON GENERAR-->
                                  <input type="button" value="Generar" onclick="armarXLSNegCer(false)"/> 
                            </div>  
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <!--------------------------------------------CAJA 2 (VALORES EN NEGATIVOS)----------------------------------------------->
        <div class="boxround">                  
	        <table width="480" class="boxround">
                <tbody onMouseOver="" onMouseOut="" onClick="VerFechaAnalisisVN();" style="cursor: pointer;" >
	                <tr>
	                    <td class="title_sec_section"><% =GF_TRADUCIR("Facturas con valores negativos")%></td>
	                </tr>
	                <tr>
	                    <td class="textoSeccion"><% =GF_TRADUCIR("Detectar facturas con ítems de valores negativos que puedan corresponder a descuentos otorgados en exceso, o a errores en el proceso de  facturación.") %></td>
	                </tr>
                </tbody>
                <tr id="target_VNeg" style="display:none;"> <!--FECHA OCULTA(tr)-->
                    <td>
                        <div >
                            <div class="col36">
                                <% Response.Write GF_TRADUCIR("Fecha Desde: ") %>
                                <input type="text" name="dtFechaDesde_VNeg" id="dtFechaDesde_VNeg" onClick="javascript: MostrarCalendario('dtFechaDesde_VNeg', SeleccionarCalDesde_VNeg)" value="" size="10">
                            </div>
                            <input type="hidden" id="f" name="f" value="" />
                            <div class="col36">
                                <% Response.Write GF_TRADUCIR("Fecha Hasta: ") %>
                                <input type="text" name="dtFechaHasta_VNeg" id="dtFechaHasta_VNeg" onClick="javascript: MostrarCalendario('dtFechaHasta_VNeg', SeleccionarCalHasta_VNeg)" value="" size="10">
                            </div>
                            <input type="hidden" id="ff" name="ff" value="">
                            <div class="col36"> <!--BOTON GENERAR-->
                                <input type="button" value="Generar" onclick="armarXLSNegCer(true)"/> 
                            </div>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <!--------------------------------CAJA 5 (EL UNICO REPORTE QUE PUEDE VER EL PROVEEDOR)------------------------------------>
        <div class="boxround">                  
	        <table width="480" class="boxround" onclick="generateFile()">
                <tbody onMouseOver="" onMouseOut="" style="cursor: pointer;">
	                <tr>	                 
	                    <td class="title_sec_section"><% =GF_TRADUCIR("Reporte Para Proveedores")%></td>
	                </tr>
	                <tr>
	                    <td class="textoSeccion"><% =GF_TRADUCIR("Resumen de las facturas emitidas a un determinado proveedor en un período.") %></td>
	                </tr>
                </tbody>
            </table>
        </div>
    </section>
    </td>
    </tr>
</table>
</body>
</html>