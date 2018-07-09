<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosFormato.asp"-->		
<!--#include file="Includes/procedimientosCompras.asp"-->	
<!--#include file="Includes/procedimientosSql.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<% 
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 22/04/2013
' Nombre: 	loadPolizaCaucion(pIdPoliza, pIdPedido)
' Objetivo:
'			Carga los valores de una Poliza de Caucion, ademas devuelve el importe ganador de la Planilla Comparativa
' Parametros:
'			[int]	  pIdPoliza			
'			[int]	  pIdPedido
' Devuelve:  -
'--------------------------------------------------------------------------------------------------
Function loadPolizaCaucion(pIdPoliza, pIdPedido)
	Dim strSQL
	strSQL = "			SELECT PCT.CDPEDIDO, POL.IMPORTE, POL.CDMONEDA,DET.IMPORTE AS IMP "					
	strSQL = strSQL & "	FROM ( SELECT IMPORTE, CDMONEDA, IDPEDIDO						  "			
	strSQL = strSQL & "		   FROM TBLPOLIZASCAUCION						              "
	strSQL = strSQL & "		   WHERE IDPDC = " & pIdPoliza & " ) POL					  "
	strSQL = strSQL & "		INNER JOIN (SELECT IDPROVEEDOR,								  "
	strSQL = strSQL & "						   CDPEDIDO,								  "
	strSQL = strSQL & "						   IDPEDIDO									  "
	strSQL = strSQL & "					FROM TBLPCTCABECERA					              "
	strSQL = strSQL & "					WHERE IDPEDIDO = " & pIdPedido
	strSQL = strSQL & "					) AS  PCT										  "
	strSQL = strSQL & "			ON PCT.IDPEDIDO = POL.IDPEDIDO							  "	
	strSQL = strSQL & "		INNER JOIN TBLPCPDETALLE DET						          "	
	strSQL = strSQL & "			ON DET.IDPROVEEDOR = PCT.IDPROVEEDOR					  "	
	strSQL = strSQL & "			AND DET.IDPEDIDO = " & pIdPedido
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
	if not rs.EoF then
		cdPedido	= rs("CDPEDIDO")
		auxCdMoneda = rs("CDMONEDA")
		monto	= cdbl(rs("IMPORTE"))
		monto_old = cdbl(rs("IMPORTE")) 
		auxImporte	= cdbl(rs("IMP"))
	end if	
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 22/04/2013
' Nombre: 	getSaldoAcumulado(pIdPedido)
' Objetivo:
'			Devuelve el saldo de las PDC que se utilizo hasta el momento para el pedido
'			(con estado Recibida,Vencida,Devuelta)
' Parametros:
'			[int]	  pIdPedido
' Devuelve:  
'			[int]	  pSaldo
'----------------------------------------------------------------------------------------------
Function getSaldoAcumulado(pIdPedido)
	Dim strSQL, auxSaldo
	strSQL = "			 SELECT SUM(IMPORTE) AS SALDO				  "				
	strSQL = strSQL & "	 FROM TBLPOLIZASCAUCION		                  "
	strSQL = strSQL & "	 WHERE IDPEDIDO = " & pIdPedido 
	strSQL = strSQL & "		AND ESTADO IN (" & ESTADO_PDC_RECIBIDA & "," & ESTADO_PDC_VENCIDA & "," & ESTADO_PDC_DEVUELTA & ")"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if not IsNull(rs("SALDO")) then	auxSaldo  = cdbl(rs("SALDO"))		
	getSaldoAcumulado = auxSaldo
End Function
'---------------------------------------------------------------------------------------------
Function controlPolizaCaucion(pAseguradora, pTomador, pMonto, pFecha, pNroPoliza, pMonto_old, pIdPedido)
	Dim rtrn
	rtrn = false	
	if (GF_CONTROL_PERIODO_2(GF_FN2DTE(Left(session("MmtoDato"),8)), pFecha) = 0) then
		if((Cdbl(pMonto) <= Cdbl(auxImporte))and(Cdbl(pMonto) > 0))then			
			if((Cdbl(pMonto) + getSaldoAcumulado(pIdPedido)) <= Cdbl(auxImporte))then
				if(pAseguradora > 0)then
					if(pNroPoliza <> "")then
						if(pTomador > 0)then
							rtrn = true
						else		
							Call setError(PROVEEDOR_NO_EXISTE)
						end if
					else
						Call setError(PDC_NRO_ASEGURADORA_NO_EXISTE)
					end if	
				else		
					Call setError(PDC_ASEGURADORA_NO_EXISTE)
				end if
			else		
				Call setError(PDC_MONTO_INCORRECTO)
			end if
		else		
			Call setError(PDC_MONTO_INCORRECTO)
		end if
	else		
		Call setError(PERIODO_ERRONEO)
	end if	
	controlPolizaCaucion = rtrn
End Function
'--------------------------------------------------------------------------------------------
Function generarPolizaCaucionParcial(pIdPoliza, pMonto_old, pMonto, pMoneda, pIdPedido)
	Dim strSQL, auxMonto
	auxMonto = pMonto_old - pMonto
	Call updateSaldoPolizaCaucion(pIdPoliza, auxMonto, pMoneda)	
	generarPolizaCaucionParcial = addPolizaCaucion(pIdPedido,TIPO_PDC_POR_ADELANTO, pMonto, pMoneda, session("MmtoSistema"), session("Usuario"),ESTADO_PDC_PENDIENTE)
End Function
'*********************************************************************************************'
'********************************	INICIO PAGINA  *******************************************'
'*********************************************************************************************'
Dim idPoliza, auxCdPedido, auxImporte, auxCdMoneda,dsProveedor, idProveedor, idAseguradora, dsAseguradora, fechaVenc
Dim flagControlar, monto, monto_old, idPdcNew

idPoliza	  = GF_Parametros7("idPoliza",0,6)
cdPedido	  = GF_Parametros7("cdPedido","",6)
idPedido	  = GF_Parametros7("idPedido",0,6)
accion		  = GF_PARAMETROS7("accion","",6)
idAseguradora = GF_PARAMETROS7("idAseguradora", 0, 6)
dsAseguradora = Trim(Ucase(GF_PARAMETROS7("dsAseguradora", "", 6)))
idProveedor   = GF_PARAMETROS7("idProveedor", 0, 6)
dsProveedor   = GF_PARAMETROS7("dsProveedor", "", 6)
fechaVenc	  = GF_PARAMETROS7("fechaVenc","",6)
monto		  = GF_PARAMETROS7("monto", 0,6)
monto_old	  = GF_PARAMETROS7("monto_old", 0,6)
auxImporte	  = GF_PARAMETROS7("auxImporte", "",6)
nroPoliza	  = GF_PARAMETROS7("nroPoliza", "",6)
auxCdMoneda	  = GF_PARAMETROS7("auxCdMoneda", "",6)

if fechaVenc = "" then fechaVenc = GF_FN2DTE(Left(session("MmtoDato"),8))
if(accion = ACCION_GRABAR)then
	flagControlar = controlPolizaCaucion(idAseguradora, idProveedor, monto, fechaVenc, nroPoliza, monto_old, idPedido)
	if(flagControlar)then
		idPdcNew = idPoliza
		if(monto < monto_old)then idPdcNew = generarPolizaCaucionParcial(idPoliza, monto_old, monto, auxCdMoneda, idPedido)
		Call updatePolizaCaucion(idPdcNew, nroPoliza, idAseguradora, idProveedor, monto, GF_DTE2FN(fechaVenc), ESTADO_PDC_RECIBIDA)
	end if
else
	Call loadPolizaCaucion(idPoliza, idPedido)	
end if	
input_Imp = monto/100


%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Administrar PDC") %></title>
<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<link rel="stylesheet" href="css/calendar-win2k-2.css" type="text/css">
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="Scripts/botoneraPopUp.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="Scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryAutocomplete.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-1.js"></script>
<script type="text/javascript">
	var calendar;
	var botones = new botonera("botones");
	
	function onLoadPage(){
		autocompleteAseguradora();
		var msProveedor = new MagicSearch("", "companyName0", 30, 2, "comprasStreamElementos.asp?tipo=empresas");
		msProveedor.setMinChar(3);
		msProveedor.setToken(";");
		msProveedor.onBlur = SeleccionarProveedor;
		msProveedor.setValue('<% =dsProveedor %>');	
		botones.addbutton('<%=GF_Traducir("Guardar")%>','guardar()');
		botones.show();
		<% if(flagControlar)then %>
			parent.window.submitInfo();
		<% end if %>
	}
	
	function autocompleteAseguradora()
	{	
		$(function() {
		$( "#dsAseguradora" ).autocomplete({
		minLength: 2,
		source: function(request,response){
			$.ajax({
				url: "comprasStreamElementos.asp",
				dataType: "json",
			data: {				
				term : request.term,
				Tipo : "JQAseguradoras",
				DsLista : document.getElementById("dsAseguradora").value
				 },
		    success: function(data) {				
				response(data);
				}
			});	
		},		
		focus: function( event, ui ) {
				$( "#dsAseguradora").val(ui.item.descr);
				$( "#idAseguradora").val(ui.item.id);
				return false;
			},
		select: function( event, ui ) {
		        $( "#dsAseguradora").val (ui.item.descr);				
				$( "#idAseguradora").val (ui.item.id);				
				return false;
			}		
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
			.data( "item.autocomplete", item )
			.append( "<a><font style='font-size:10;'>" + item.descr + "</font></a>" )
			.appendTo( ul );
			};
		});
	}
	
	function SeleccionarProveedor(ms){
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("idProveedor").value = arr[0];
			document.getElementById("dsProveedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("idProveedor").value = 0;
				document.getElementById("dsProveedor").value = "";
				ms.setValue("");
			}	
		}				
	}
	
	function SeleccionarCalFin(cal, date) {
		var str= new String(date);		
		document.getElementById("fdateDiv").innerHTML = str;
		document.getElementById("fechaVenc").value = str;
		if (cal) cal.hide();
	}

	function CerrarCal(cal) {
		cal.hide();
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
	
	function guardar()
	{	
		document.getElementById("accion").value = '<%=ACCION_GRABAR%>';		
		document.getElementById("myForm").submit();
	}
	
	function asignarMonto(){
		var tipoCambio = document.getElementById("input_Imp").value;
		tipoCambio = tipoCambio.replace(/,/,".");
		document.getElementById("monto").value = tipoCambio * 100;		
	}
	
	function keyPressEvent(obj, evt) {
		return controlIngreso(obj, evt, 'I');
	}
	
	</SCRIPT>
</HEAD>
<BODY onload="onLoadPage();">
	<FORM id="myForm" name="myForm" action="comprasPDCPopUp.asp" method="post">
	<INPUT type='hidden' name='accion' id='accion'>	
	<INPUT type="hidden" id="idPoliza" name="idPoliza" value="<% =idPoliza %>">
	<INPUT type="hidden" id="cdPedido" name="cdPedido" value="<% =cdPedido %>">
	<INPUT type="hidden" id="idPedido" name="idPedido" value="<% =idPedido %>">
	<INPUT type="hidden" id="auxCdMoneda" name="auxCdMoneda" value="<% =auxCdMoneda %>">	
	<INPUT type="hidden" id="auxImporte" name="auxImporte" value="<% =auxImporte %>">	
	<% call showErrors() %>
	<TABLE  width="100%">			
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Pedido") %></TD>
			<TD colspan=2 ><B><% =cdPedido %></B></TD>
		</TR>										
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Vencimiento") %></TD>			
			<TD >
				<a href="javascript:MostrarCalendario('imgFin', SeleccionarCalFin)">
					<img id="imgFin" src="images/compras/calendar-16x16.png">
				</a>
				<input type="hidden" id="fechaVenc" name="fechaVenc" value="<%=fechaVenc%>">				
			</TD>
			<TD >
				<div id="fdateDiv" class="labelStyle"><% =fechaVenc %></div>
			</TD>	
		</TR>			
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Monto") %></TD>
			<TD colspan=2>
				<%= getSimboloMoneda(auxCdMoneda)%>&nbsp				
				<input type="text" id="input_Imp" name="input_Imp" value="<% =input_Imp %>" style="text-align:right;" onKeyPress="return keyPressEvent(this, event)" onBlur="asignarMonto();">
				<input type="hidden" id="monto" name="monto" value="<% =monto %>">
				<input type="hidden" id="monto_old" name="monto_old" value="<% =monto_old %>">
			</TD>
		</TR>			
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Aseguradora") %></TD>
			<TD colspan=2>				
				<INPUT id="dsAseguradora" name="dsAseguradora" size=40 value="<%=dsAseguradora%>">									
				<INPUT type="hidden" id="idAseguradora" name="idAseguradora" value="<%=idAseguradora%>">
			</TD>
		</TR>
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Nro Poliza") %></TD>
			<TD colspan=2>				
				<input type="text" id="nroPoliza" name="nroPoliza" size=30 value="<%=nroPoliza%>">				
			</TD>
		</TR>			
		<TR>			
			<TD width="30%" class="reg_header"><% =GF_TRADUCIR("Tomador") %></TD>
			<TD colspan=2>
				<div id="companyName0"></div>
				<input type="hidden" id="idProveedor" name="idProveedor" value="<%=idProveedor%>">
				<input type="hidden" id="dsProveedor" name="dsProveedor" value="<%=dsProveedor%>">
			</TD>
		</TR>		
	</TABLE>
	<div id="botones"></div>
	</FORM>
</BODY>
</HTML>

