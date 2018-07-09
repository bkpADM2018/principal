<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<%
Call comprasControlAccesoCM(RES_CC)
'--------------------------------------------------------------------------
'controla los unicos datos que ingresa el usuario a la hora de crear el contrato
'los demas datos son asignados desde el pedido
Function controlCTC()
	Dim rtrn, rs, strSQL
	rtrn = false
	if (CTC_cdResponsable <> "") then
		if (ControlCTCProveedor()) then		    
			if (CTC_TotalImporte > 0)then					
				if (CTC_fechaVto <> "") then
					if (Cdbl(Left(session("MmtoDato"), 8)) <= Cdbl(CTC_fechaVto)) then		
						'Response.Write "<BR></BR> CTC_idObra:"& CTC_idObra & " CTC_areaObra:" & CTC_areaObra & " CTC_detalleObra:" &CTC_detalleObra & " CTC_tipo:" & CTC_tipo
						'JASif (CTC_tipo = CTC_TIPO_GENERAL) then
						'	rtrn = true
						'else									
                        if (CTC_idDivision > 0) then							
							rtrn = true
						else
                            setError(DIVISION_NO_EXISTE)
                        end if		
					else
						setError(PERIODO_ERRONEO)
					end if
				else
					setError(FECHA_ENTREGA_INCORRECTA)
				end if
			else
				setError(IMPORTE_NO_EXISTE)
			end if
		else
			setError(PROVEEDOR_NO_HAB_CTC)
		end if					
	else
		setError(FALTA_RESPONSABLE)
	end if		
	controlCTC = rtrn
End Function
'---------------------------------------------------------------------------------------------------
'Determina si los importes de pantalla son totales o representa valores de cuota o unitarios.
Function prepararImportesContrato(pCdMoneda, importePantalla, unidades, pTipoCambio)
    Dim importe
                    
    importe = importePantalla                    
    if (tieneValorUnitario(CTC_tipo)) then
        importe = CLng(unidades) * CDbl(importePantalla)
        if (pCdMoneda = MONEDA_PESO) then
            CTC_ContratoPesos = importe
            CTC_ContratoDolares = Round(CDbl(CTC_ContratoPesos) / CDbl(pTipoCambio), 0)
            CTC_valorUnitarioPesos = importePantalla
            CTC_valorUnitarioDolares = Round(CDbl(importePantalla) / CDbl(pTipoCambio), 0)
        else
            CTC_ContratoDolares = importe
            CTC_ContratoPesos = Round(Cdbl(CTC_ContratoDolares) * CDbl(pTipoCambio), 0)   
            CTC_valorUnitarioDolares = importePantalla
            CTC_valorUnitarioPesos = Round(CDbl(importePantalla) * CDbl(pTipoCambio), 0)
        end if
    else
        CTC_valorUnitarioPesos = 0
        CTC_valorUnitarioDolares = 0
        if (pCdMoneda = MONEDA_PESO) then
            CTC_ContratoPesos = importe
            CTC_ContratoDolares = Round(CDbl(CTC_ContratoPesos) / CDbl(pTipoCambio), 0)
        else
            CTC_ContratoDolares = importe
            CTC_ContratoPesos = Round(CDbl(CTC_ContratoDolares) * CDbl(pTipoCambio), 0)
        end if
    end if
    
End Function
'--------------------------------------------------------------------------
Function ControlCTCProveedor()
    ControlCTCProveedor = true
    if (UCase(CTC_contratoFisico) = "ON") then 
        if (not habilitadoParaContratos(CTC_idProveedor)) then ControlCTCProveedor = false
    end if
End Function 
'--------------------------------------------------------------------------
'GENERAR NUEVOS CONTRATOS, SE TOMAN LA MAYORIA DE LOS DATOS DEL PEDIDO, EL-
'RESTO SE PIDEN AQUI, SI EL CONTRATO ES MANTENIMIENTO SE PEDIRA EL AREA Y -
'EL DETALLE DE LA OBRA SELECCIONADA EN CASO DE SER UINA INVERSION SE ------
'ASIGNAN LOS VALORES 0 EN AREA Y DETALLE, A LA HORA DE REALIZAR EL PAGO SE-
'DARA LA OPCION DE ELEGIR EL BUDGET DONDE IMPACTARA EL MISMO --------------
'UNA VEZ CARGADO EL PAGO SE ENVIA UN MAIL A LEGALES PARA QUE AUTORIZEN EL -
'CONTRATO, HASTA NO ESTAR AUTORIZADO EL CONTRATO PERMANECE INAVILITADO ----
'PARA RECIBIR PAGOS -------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'***********************************************
'*************  COMIENZO DE PAGINA  ************
'***********************************************
Dim rsCTC, txtTitulo, importeOriginal
Dim CTC_TotalPesos, CTC_TotalDolar ,ctcEstado
Dim controlOK, flagGrabar, aux, aux2, CTC_Unidades


Call GP_CONFIGURARMOMENTOS

accion = GF_PARAMETROS7("accion","",6)
CTC_idContrato = GF_PARAMETROS7("CTC_idContrato",0,6)
CTC_idPedido = GF_PARAMETROS7("idPedido",0,6)
if (CTC_idPedido <> 0) then 
	Call initHeaderDB(CTC_idPedido)
else
	pct_cdPedido = "Sin Pedido"
end if
flagGrabar = false
'debe tener un pedido asignado, en caso contrario se niega el acceso
CTC_tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
if (isFormSubmit() or (CTC_idContrato <> 0)) then
	if (isFormSubmit()) then
		'se toman los paramtros desde la pagina
		CTC_tipo = GF_PARAMETROS7("CTC_tipo","",6)
		CTC_FReparo = GF_PARAMETROS7("CTC_FReparo",0,6)
		CTC_Titulo = GF_PARAMETROS7("CTC_Titulo", "",6)
		CTC_idProveedor = GF_PARAMETROS7("CTC_idProveedor",0,6)
		CTC_dsProveedor = GF_PARAMETROS7("CTC_dsProveedor","",6)
		CTC_cdResponsable = GF_PARAMETROS7("CTC_cdResponsable","",6)
		CTC_dsResponsable = getUserDescription(CTC_cdResponsable)
		CTC_cdMoneda = GF_PARAMETROS7("CTC_cdMoneda","",6)		
		if (CTC_cdMoneda ="") then CTC_cdMoneda = MONEDA_PESO		
		CTC_TotalImporte = GF_PARAMETROS7("CTC_TotalImporte",0,6)
		importeOriginal = GF_PARAMETROS7("importeOriginal",0,6)
        CTC_contratoFisico = GF_PARAMETROS7("CTC_contratoFisico","",6)
		CTC_cdContrato = GF_PARAMETROS7("CTC_cdContrato","",6)	
		if (CTC_cdContrato = "") then CTC_cdContrato = CONTRATO_A_CONFIRMAR
		CTC_fechaVto = GF_PARAMETROS7("issuedate","",6)	        		
        CTC_idDivision = GF_PARAMETROS7("idDivision",0,6)
        CTC_Unidades = GF_PARAMETROS7("unidades",0,6)   
        if (CTC_Unidades = 0) then CTC_Unidades = 1     
	else
		'Se esta modificando un contrato existente.			
		Set rsCTC = readCTC(CTC_idContrato)
		if (not rsCTC.eof) then			
			if (CDbl(rsCTC("ESTADO")) = ESTADO_CTC_PENDIENTE) then						
				CTC_idPedido = CDbl(rsCTC("IDPEDIDO"))
				CTC_cdMoneda = rsCTC("CDMONEDA")	
				CTC_Titulo = rsCTC("TITULO")
				CTC_tipo = rsCTC("TIPO")
				'En caso de que este vacio los tomo como Obra, ya que son CTC antes de aplicar este cambio
				if(CTC_tipo = "") then CTC_tipo = CTC_TIPO_OBRA
				if (tieneValorUnitario(CTC_tipo)) then				    
				    CTC_TotalImporte = CDbl(rsCTC("IMPORTEUNITARIOPESOS")) 				    
				    if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_TotalImporte = CDbl(rsCTC("IMPORTEUNITARIODOLARES"))
				else
				    CTC_TotalImporte = CDbl(rsCTC("IMPORTEPESOS"))
				    if (CTC_cdMoneda = MONEDA_DOLAR) then CTC_TotalImporte = CDbl(rsCTC("IMPORTEDOLARES"))
				end if
				importeOriginal = CDbl(rsCTC("IMPORTEPESOS"))
				if (CTC_cdMoneda = MONEDA_DOLAR) then importeOriginal = CDbl(rsCTC("IMPORTEDOLARES"))
				CTC_FReparo = CDbl(rsCTC("FONDOREPARO"))
				CTC_idProveedor = CDbl(rsCTC("IDPROVEEDOR"))				
				CTC_dsProveedor =getDescripcionProveedor(CTC_idProveedor)
				CTC_cdResponsable = rsCTC("CDRESPONSABLE")
				CTC_dsResponsable = getUserDescription(CTC_cdResponsable)
				CTC_fechaVto = rsCTC("FECHAVTO")	
				'En caso de que este vacio los tomo como Fecha Actual, ya que son CTC antes de aplicar este cambio
				if (Cdbl(rsCTC("FECHAVTO")) = 0) then CTC_fechaVto = Left(session("MmtoDato"), 8)
                CTC_idDivision = rsCTC("IDDIVISION")
                CTC_Unidades = 1
			else
				Response.Redirect "comprasAccesoDenegado.asp"
			end if
		else
			Response.Redirect "comprasAccesoDenegado.asp"
		end if
	end if	
	'se controlan los datos del contrato
	controlOK = controlCTC()
	if ((accion = ACCION_GRABAR) and (controlOK)) then
        If (UCASE(CTC_contratoFisico)  <> "ON") then 
            ctcEstado = ESTADO_CTC_AUTORIZADO
            CTC_cdContrato = CONTRATO_TIPO_SERVICIO
        else
            ctcEstado = ESTADO_CTC_PENDIENTE
        End If
		'-JAS if(CTC_tipo = CTC_TIPO_GENERAL)then
		'	CTC_areaObra = 0 
		'	CTC_detalleObra = 0
		'end if
		'Se preparan los importes para grabar el contrato.
		Call prepararImportesContrato(CTC_cdMoneda, CTC_TotalImporte, CTC_Unidades, CTC_tipoCambio)
		flagGrabar = grabarCTC(CTC_idContrato, CTC_Titulo, CTC_idPedido, CTC_idProveedor, CTC_ContratoPesos, CTC_ContratoDolares, CTC_valorUnitarioPesos, CTC_valorUnitarioDolares, CTC_FReparo, CTC_cdResponsable, CTC_cdContrato, CTC_cdMoneda, CTC_fechaVto, CTC_tipo, CTC_idDivision, ctcEstado)		
		if ((flagGuardar) and (CTC_idPedido > 0)) then
		    if ((pct_idObra > 0) and (pct_idArea > 0) and (pct_idDetalle > 0)) then
		        auxImportePartida = CTC_ContratoPesos 'VARIABLE QUE SE USA PARA CARGAR COLUMNA SALDO SEGUN MONEDA DEL CONTRATO
                if (pCdMoneda = MONEDA_DOLAR) then auxImportePartida = CTC_ContratoDolares
		        Call grabarPartidaCTC(CTC_idContrato, pct_idObra, pct_idArea, pct_idDetalle, CTC_fechaVto, Left(session("MmtoDato"),8), CTC_cdMoneda, auxImportePartida, session("Usuario"))
            end if		        
		end if
	end if
else	
	'Nuevo Contrato		
	CTC_idProveedor = 0
	CTC_dsProveedor = ""	
	CTC_TotalPesos = 0
	CTC_TotalDolar = 0		
	importeOriginal	= 0
	CTC_tipo = CTC_TIPO_OBRA
	if (CTC_idPedido <> 0) then		
		CTC_idProveedor = pct_idProveedorElegido
		CTC_dsProveedor = pct_dsProveedorElegido
		Call obtenerGanadorPlanilla(CTC_cdMoneda, aux, aux2)
		CTC_TotalImporte = aux		
		importeOriginal = aux
	end if 	
	CTC_FReparo = F_REPARO_INICIAL			
    CTC_contratoFisico = "ON"		
    CTC_Unidades = 1	
end if
%>
<html>
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=Edge">
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title><% =GF_TRADUCIR("Sistema de Compras - Contratos") %></title>
	<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
	<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
    <link rel="stylesheet" href="css/main.css" type="text/css">
	<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
	<style type="text/css">
	.labelStyle {
		font-weight: bold;	
	}
    input {
        height: 18px;
        background: #D8D8D8;
        border: 1px solid #bbb;
        padding: 2px;
        margin: 0;
    }

        input:focus {
            background: #D8D8D8;
        }

    select {
        background:#D8D8D8;
        border: 1px solid #ccc;
        padding: 2px;
        margin: 0;
    }

        select:focus {
            background:#D8D8D8;
        }

    textarea {
        background: #dddddd;
        border: 1px solid #ccc;
        padding: 2px;
        margin: 0;
    }

        textarea:focus {
            background: #D8D8D8;
        }
.popstyle {
background-color:white!important;
font-family:Arial;
font-style:italic;
}
.inputs{margin-left:10px;margin-top:5px;margin-bottom:5px;}
input[type="radio"] {margin:5px;}
.fontstile{text-align:center;font-weight:700;font-size:12px;font-style:initial;display:block;}
.datas tbody td:first-child {border-radius:0px;}
.datas tbody tr:nth-child(odd) {background-color: #eeeeee;}/* Color for td alternative */
.datagridlv1 tbody tr:nth-child(odd) {background-color: #fff !important;}/* Color for td alternative IN TABLE FATHER-SON*/
.datagridlv1 tbody tr:nth-child(4n+1) {background-color: #eeeeee !important;}/* Color for td alternative IN TABLE FATHER-SON */
.datagridlv1 tbody tr:nth-child(4n+2) {background-color: #eeeeee !important;}/* Color for td alternative IN TABLE FATHER-SON*/
.datagrid tbody td {border-bottom: 1px solid #CCCCFF;}
table {border:1px solid #CCCCFF;}
.tr-titulo{padding-right:10px;}
td{font-weight:bold;}
	    tr { border-bottom:1px solid lightgray;
	    }
	</style>	
	<script type="text/javascript" src="scripts/Toolbar.js"></script>
	<script type="text/javascript" src="scripts/formato.js"></script>
	<script type="text/javascript" src="scripts/calendar.js"></script>
	<script type="text/javascript" src="scripts/calendar-1.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/controles.js"></script>
	
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>	
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	<script type="text/javascript" src="scripts/jQueryAutocomplete.js"></script>
		
	<script type="text/javascript">
		var ch = new channel;
		function bodyOnLoad(){
			<% if (flagGrabar) then 
		        if (ctcEstado = ESTADO_CTC_PENDIENTE) then %>
				    alert('El contrato debe ser confirmado por el Departamento de Legales.');
                <%End if%>
		        cerrar();
			<% end if%>
			var	tb = new Toolbar('toolbar', 4, "images/compras/");
			idBtnGuardar = tb.addButtonSAVE("Guardar", "submitInfo('<% =ACCION_GRABAR %>')");
			idBtnControl = tb.addButtonCONFIRM("Controlar",  "submitInfo('<% =ACCION_CONTROLAR %>')");			
			tb.addButton("Close-16x16.png", "Cerrar", "cerrar()");
			tb.draw();			
			cambiarTipoCTC('<% =CTC_tipo %>');			
		}
		
		function submitInfo(acc){
			document.getElementById("accion").value = acc;
			document.getElementById("frmSel").submit();
		}

		function cerrar() {
			window.close();
		}

		function SeleccionarResponsable(ms){
			var desc = ms.getSelectedItem();
			if (desc.indexOf('-') != -1) {
				var arr = desc.split('-');
				document.getElementById("CTC_cdResponsable").value = arr[0];
				ms.setValue(arr[1]);
			} else {
				if (desc == "") document.getElementById("CTC_cdResponsable").value = "";
			}				
		}
	
		function keyPressEvent(obj, evt) {		
			return controlIngreso(obj, evt, 'I');
		}

		function controlPercent(pObj){	
			if (pObj.value > 100){
				alert("El porcentaje no puede ser mayor a 100!");
				pObj.value = 0; 
			}
		}

		
		function MostrarCalendario(p_objID, funcSel){
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
		
		function CerrarCal(cal){
			cal.hide();
		}
		
		function SeleccionarCalEmision(cal, date){
			var str= new String(date);
			document.getElementById("issuedateDiv").value = str;
		    document.getElementById("issuedate").value = str.substr(6,4) + str.substr(3,2) + str.substr(0,2);
			if (cal) cal.hide();
		}
		function QuitarFechaHasta(){
			document.getElementById("issuedateDiv").value = "";
			document.getElementById("issuedate").value = "";
		}	
		
		function cambiarTipoCTC(pTipo){			
			/* Default: Pago Contra Certificados de Obra */						
			document.getElementById("CTC_FReparo").disabled = false;
			document.getElementById("lblImportes").innerHTML = "Importe Total";			
			document.getElementById("CTC_FReparo").value = 5;			
			document.getElementById("cantUnidades").style.display = "none";
			if (pTipo == '<%=CTC_TIPO_OBRA%>') document.getElementById("CTC_contratoFisico").checked = true;
			/* Servicio Repetitivo - Valor Unidad Fijo */
			if (pTipo == '<%=CTC_TIPO_UNITARIO %>'){				                
				document.getElementById("CTC_FReparo").disabled = true;
				document.getElementById("CTC_FReparo").value = 0;
				document.getElementById("lblImportes").innerHTML = "Valor Unitario";
                document.getElementById("cantUnidades").style.display = "block"; 
			}
			/* Contrato General */
			if (pTipo == '<%=CTC_TIPO_GENERAL %>') {
				document.getElementById("CTC_FReparo").disabled = true;
				document.getElementById("CTC_FReparo").value = 0;
			}
		}			
		
		$(function() {
			
				$( "#CTC_Responsable" ).autocomplete({
				minLength: 2,
				source: "comprasStreamElementos.asp?tipo=JQPersonas",
				focus: function( event, ui ) {
					$( "#CTC_Responsable" ).val(ui.item.nombre);
					return false;
				},
				select: function( event, ui ) {
					$( "#CTC_Responsable"    ).val (ui.item.nombre);
					$( "#CTC_cdResponsable"  ).val (ui.item.cdusuario );
					
					return false;
				},
				change: function( event, ui ) {
					if (!ui.item)
					{
						$( "#articulo" ).val("");
						$( "#CTC_cdResponsable"  ).val ("");
						
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.cdusuario + " - <font style='font-size:10;'>" + item.nombre + "</font></a>" )
					.appendTo( ul );
			};			
		});
		
		<% if (CTC_idPedido = 0) then %>
		
		$(function() {			
				$( "#CTC_dsProveedor" ).autocomplete({
				minLength: 2,
				source: "comprasStreamElementos.asp?tipo=JQEmpresas&linea=0",
				focus: function( event, ui ) {
					$( "#CTC_dsProveedor").val(ui.item.dsempresa);
					return false;
				},
				select: function( event, ui ) {
					$( "#CTC_dsProveedor").val (ui.item.dsempresa);
					$( "#CTC_idProveedor").val (ui.item.idempresa );
					
					return false;
				},
				change: function( event, ui ) {
					if (!ui.item)
					{
						$( "#CTC_idEmpresa"  ).val ("0");
						$( "#CTC_dsEmpresa"  ).val ("");
						
					}
				}
			})
			.data( "autocomplete" )._renderItem = function( ul, item ) {
				return $( "<li></li>" )
					.data( "item.autocomplete", item )
					.append( "<a>" + item.idempresa + " - <font style='font-size:10;'>" + item.dsempresa + "</font></a>" )
					.appendTo( ul );
			};
		});
		<% end if %>
		function calcularVU() {				
		<% if (CTC_idPedido <> 0) then %>
			var u = editarNumero(document.getElementById("unidades").value, 0); //Si no es del tipo unitario las unidades siempre sera 1.			
			if (u <= 0) u = 1;			
			var vjoI = document.getElementById("importeOriginal").value;			
			var nvoI = editarNumero(vjoI/u, 0);				
			document.getElementById("CTC_TotalImporte").value=nvoI;
			document.getElementById("unidades").value = u;			
			var sm = document.getElementById("hdnSM").value;
			document.getElementById("lblImpPedido").innerHTML=sm + ' ' + editarImporte(nvoI/100);			
		<% end if %>
		}
		function copiarImporte() {
			document.getElementById("CTC_TotalImporte").value= editarNumero(document.getElementById("totalImporte").value, 2)*100;
		}
	</script>
</head>
<body onLoad="bodyOnLoad()">
	<form method="post" id="frmSel">
		<div id="toolbar"></div>
        <br>
        <b class="fontstile" style="margin-bottom:-10px;padding-top: 13px;"><% =GF_TRADUCIR("DATOS DEL CONTRATO") %></b>
		<table class="popstyle datas" align="center" border="0" style="border-color: lightgray;width:95%; border-radius:5px;">
			<tr><td colspan="4" style="background:white;"><% call showErrors() %></td></tr>
		<%	if (CTC_idContrato <> 0) then %>
        <tr>
            <td class="" colspan="4" style="background-color:white!important;color:black;">
				<b class="fontstile">
					<% =GF_TRADUCIR("Id Contrato:") %>&nbsp;<% =CTC_idContrato %>
				</b>
			</td>
			<input type="hidden" name="CTC_idContrato" id="CTC_idContrato" value="<%=CTC_idContrato%>">
        </tr>
		<%	end if	%>
                <tbody>
                    <tr>
                        <td class="tr-titulo" style="text-align: right; background-color: white;" rowspan=2"><% =GF_TRADUCIR("Tipo") %></td>
                        <td style="background-color: white; padding-left: 7px;">
                            <input style='cursor:pointer;border:none;' id='CTC_tipo' type='radio' value="<% =CTC_TIPO_OBRA %>" name='CTC_tipo' <% if (CTC_tipo = CTC_TIPO_OBRA) then Response.write "CHECKED" %> onclick="cambiarTipoCTC('<%=CTC_TIPO_OBRA%>')">
                            <label for="tipoReporte">Pago Contra Certificados de Obra</label><br />
                            
                            <input style='cursor:pointer;border:none;' id='CTC_tipo' type='radio' value="<% =CTC_TIPO_GENERAL %>" name='CTC_tipo' <% if (CTC_tipo = CTC_TIPO_GENERAL) then Response.write "CHECKED" %> onclick="cambiarTipoCTC('<%=CTC_TIPO_GENERAL%>')">
                            <label for="tipoReporte">General</label>

							<input style='cursor:pointer;border:none;' id='CTC_tipo' type='radio' value="<% =CTC_TIPO_UNITARIO %>" name='CTC_tipo' <% if (CTC_tipo = CTC_TIPO_UNITARIO) then Response.write "CHECKED" %> onclick="cambiarTipoCTC('<% =CTC_TIPO_UNITARIO %>')">
                            <label for="tipoReporte">Servicio Repetitivo x Valor Unitario</label>
                        </td>
					</tr>

                   
                    <tr>&nbsp;</tr>
                    <tr>
                        <td class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Con Contrato Fisico")%></td>
                        <td colspan="3" >
                            <input type="checkbox" id="CTC_contratoFisico" name="CTC_contratoFisico" style="margin-left:10px;" <% if (ucase(CTC_contratoFisico) = "ON") then response.write "CHECKED" end if %> />
                        </td>
                    </tr>                    
                    <tr>
                        <td class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Division") %></td>
                        <td colspan="3">
                <%  Call executeProcedureDb(DBSITE_SQL_INTRA, rsDivision, "TBLDIVISIONES_GET","") %>
                    <select id="idDivision" name="idDivision" class="inputs">
					    <option value="0" <%if (Cint(CTC_idDivision) = 0) then %> selected='true' <%end if%>><% =GF_TRADUCIR("-Seleccione-") %></option>
						<% while (not rsDivision.eof) %>
						     <option value="<% =rsDivision("IDDIVISION") %>" <% if (Cint(CTC_idDivision) = Cint(rsDivision("IDDIVISION"))) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %></option>
						<%    rsDivision.MoveNext()
						  wend %>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td class="tr-titulo inputs " width=20% style="text-align: right; height:15px;" width="20%"><% =GF_TRADUCIR("Pedido") %></td>
                        <td colspan="3" style="padding-left: 10px; padding-top: 10px; padding-bottom: 10px;"><% =pct_cdPedido %></td>
                        <input  type="hidden" name="idPedido" id="idPedido" value="<%=CTC_idPedido%>">
                    </tr>
                    <tr>
                        <td width=20% class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Titulo") %></td>
                        <td colspan="3"><input class="inputs" type="text" name="CTC_Titulo" id="CTC_Titulo" size="80" value="<%=CTC_Titulo%>" ></td>
                    </tr>                    
                    <tr>
                        <td class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Proveedor") %></td>
                        <td>
                            <input type="hidden" id="CTC_idProveedor" name="CTC_idProveedor" value="<%=CTC_idProveedor%>">
					<%	if (CTC_idPedido = 0) then	%>
							<input id="CTC_dsProveedor" name="CTC_dsProveedor" class="inputs" size="80px" value="<%=CTC_dsProveedor%>">
					<%	else	%>
                            <input type="hidden" id="CTC_dsProveedor" name="CTC_dsProveedor" size="30" class="inputs" value="<%=CTC_dsProveedor%>">
					<%		Response.Write CTC_dsProveedor 
						end if			%>
                        </td>
                    </tr>
                    <tr>
                        <td class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Fondo de Reparo") %></td>
                        <td style="font-weight: bold;">
                            <input class="inputs" type="text" id="CTC_FReparo" name="CTC_FReparo" value="<% =CTC_FReparo %>" size="8" maxlength="6" style="text-align: right;" onkeypress="controlPercent(this); return keyPressEvent(this, event)" onblur="controlPercent(this)" <% if (CTC_tipo = CTC_TIPO_GENERAL) then response.write "disabled=true" %>>
                            %
                        </td>
                    </tr>
                    <tr>
                        <td class="tr-titulo" style="text-align:right;"><% =GF_TRADUCIR("Fecha Vto") %></td>
                        <td>
                            <input type="text" name="issuedateDiv" id="issuedateDiv" readonly onclick="javascript:MostrarCalendario('issuedateDiv', SeleccionarCalEmision)" value="<% =GF_FN2DTE(CTC_fechaVto) %>" class="inputs" size="30">
                            <input type="hidden" id="issuedate" name="issuedate" value="<%=CTC_fechaVto%>">
                        </td>
                    </tr>
                    <tr>
                        <td class="tr-titulo" style="text-align: right;"><% =GF_TRADUCIR("Moneda") %></td>
                        <td style="padding-left:7px;">
                            <input type="radio" id="optPesos" name="CTC_cdMoneda" value="<% =MONEDA_PESO %>" <% if (CTC_cdMoneda = MONEDA_PESO) then Response.write "CHECKED" %>> Pesos
                            <input type="radio" id="optDolares" name="CTC_cdMoneda" value="<% =MONEDA_DOLAR %>" <% if (CTC_cdMoneda = MONEDA_DOLAR) then Response.write "CHECKED" %>> Dolares
                        </td>

                    </tr>
                    <tr>
                        <td class="tr-titulo" id="lblImportes" style="text-align:right;"><% =GF_TRADUCIR("Importe Total") %></td>
                        <td style="font-weight: bold;">						
					<% if (CTC_idPedido = 0) then %>
                            <input type="text" id="totalImporte" name="totalImporte" style="text-align: right;" size="30" onblur="copiarImporte()" class="inputs" value="<%=CTC_TotalImporte/100 %>">                            
					<% else %>						
						<input type="hidden" id="hdnSM" value="<% =getSimboloMoneda(CTC_cdMoneda) %>">
						<span id="lblImpPedido"><% =getSimboloMoneda(CTC_cdMoneda) & " " & GF_EDIT_DECIMALS(CTC_TotalImporte, 2) %></span>
						<input type="hidden" id="importeOriginal"  name="importeOriginal" value="<% =importeOriginal %>">
					<% end if %>
						<input type="hidden" id="CTC_TotalImporte" name="CTC_TotalImporte" value="<% =CTC_TotalImporte %>">						
                        </td>
                    </tr>	
					<tr id="cantUnidades" style="display:none;">
                        <td class="tr-titulo" style="text-align:right;"><% =GF_TRADUCIR("Unidades estimadas:") %></td>
                    	<td><input type="text" size="5" name="unidades" id="unidades" onBlur="calcularVU()" value="<% =CTC_Unidades %>"/></span></td>
                    </tr>						
                    <tr>
                        <td class="tr-titulo" style="text-align:right;"><% =GF_TRADUCIR("Responsable") %></td>
                        <td>
                            <span class="ui-widget  ">
                                <input class="inputs" id="CTC_Responsable" name="CTC_Responsable" style="margin-left: 10px;" size="30" value="<%=CTC_dsResponsable%>">
                                <input type="hidden" name="CTC_cdResponsable" id="CTC_cdResponsable" value="<%=CTC_cdResponsable%>">
                            </span>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
		<input type="hidden" id="accion" name="accion" value="">
		<input type="hidden" id="CTC_idContrato" name="CTC_idContrato" value="<% =CTC_idContrato %>">		
	</form>
</body>
</html>