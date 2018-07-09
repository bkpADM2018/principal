<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosTitulos.asp"-->
<!--#include file="../../includes/procedimientosSQL.asp"-->
<!--#include file="Include/procedimientoProcedencias.asp"-->
<%

'----------------------------------------------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function

'-----------------------------------------------------------------------------------------------------------------------
Function readProcedenciaDB(pCdProcedencia)
	Dim rs
	call getProcedenciaByCdProcedencia(pCdProcedencia,g_strPuerto)	
End Function
'-----------------------------------------------------------------------------------------------------------------------
Function readProcedenciaParams()
	call getProcedenciaByParams()
End Function
'---------------------------------------------------------------------------------------------------------
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto,params,g_cdProducto,g_TipoEnvio,g_CodigoCamara,g_UltimoTurno,accion,g_UltimaBoleta,g_IsEdit,flagGrabar,flagAdd, g_Humedimetro
Dim g_DescripcionAbr,g_Descripcion,g_HumedadRecep,g_HumedadBase,g_Coeficiente2,g_Coeficiente1,g_BoletaCamara,g_BaseTrigo, g_TipoProducto
Dim flagPermiso 

g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)
flagPermiso = true
'if (leerPermisos(g_strPuerto, TASK_PRODUCT_USER) = NO_TIENE_PERMISO) then flagPermiso = false

accion = GF_Parametros7("accion","",6)
gCdProcedencia = GF_PARAMETROS7("cdProcedencia",0,6)
call addParam("cdProcedencia", gCdProcedencia, params)

g_IsEdit = GF_Parametros7("isEdit","",6)
call addParam("isEdit", g_IsEdit, params)

if (not isFormSubmit()) then	
	g_IsEdit = false	
	if Cdbl(gCdProcedencia) <> 0 then Call readProcedenciaDB(gCdProcedencia)	
else
	call readProcedenciaParams 	
	if (accion = ACCION_GRABAR) then	
		if checkProcedencia then		
			if Cdbl(gCdProcedencia) <> 0 then
				Call updateProcedencia()
			else
				Call addProcedencia()
			end if	
			flagGrabar = true
			g_IsEdit = true
		end if			
	end if
end if	

%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administración de Procedencias </TITLE>
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
				
	<style type="text/css">
		
	</style>
</HEAD>
<script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script language="javascript">
	var refPopUpAdd;
	var ch = new channel();
	var up1;	
	function onLoadPage(){				
		refPopUpAdd = getObjPopUp('popupProcedencia');
		tb = new Toolbar('toolbar', 6,'../../images/');				
		tb.addButton("back-16.png","Volver", "volver()");
		<% if (flagPermiso) then %>
        tb.addButton("save-16.png", "Grabar", "submitInfo('<%=ACCION_GRABAR%>')");		
        <% end if %>
		tb.draw();
		document.getElementById("msjGrabar").innerHTML  = "";		
		<% if(flagGrabar) then %>
				refPopUpAdd.hide();
				document.getElementById("msjGrabar").className  = "reg_Header_success";
				document.getElementById("msjGrabar").innerHTML  = "Se grabó correctamente."
				document.getElementById("accion").value = "";
				document.getElementById("isEdit").value = "<%=g_IsEdit%>";
		<% end if %>		
		autoCompleteLocalidadOncca();
		autoCompleteLocalidadCamara();
	}
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function volver(){	
		refPopUpAdd.hide();
	}
	function reloadPage(prod,edit){
		document.getElementById("cdProducto").value = prod;
		document.getElementById("isEdit").value = edit;
		submitInfo("");
	}
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function submitInfo(acc){		
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();		
	}	
	
    function convertToBoolean(pVal){
        if(pVal == true)
            return "1";
        else
            return "0";

    }

    function autoCompleteLocalidadOncca(){
        $( "#gDsProcedenciaOncca").autocomplete({
			minLength: 1,
			source: "../puertosStreamElementos.asp?tipo=JQProcedenciasONCCA&pto=<%=g_strPuerto%>&pcia=<%=gCdProvincia%>",
			focus: function( event, ui ) {
				$( "#gDsProcedenciaOncca").val(ui.item.dsloc);
			return false;
			},
			select: function( event, ui ) {
				$( "#gDsProcedenciaOncca"   ).val (ui.item.dsloc);
				if (document.getElementById("gDsProcedencia").value == '') {
					$( "#gDsProcedencia"		).val (ui.item.dsloc);
					}
				$( "#gCdProcedenciaOncca"   ).val (ui.item.idloc);		
				$( "#gCdProvincia"			).val (ui.item.idprov);
				
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#gDsProcedenciaOncca").val ("");
					$( "#gCdProcedenciaOncca").val ("");					
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a><font style='font-size:10;'>" + item.idloc + " - " + item.dsloc + "</font><br> <font style='font-size:8;'>" + item.dsprov + " - " + item.dspart + "</font></a>" )
				.appendTo( ul );
		};
	}
    function autoCompleteLocalidadCamara(){
		var auxStr = new String();
        $( "#gDsProcedenciaCamara").autocomplete({
			minLength: 1,
			source: "../puertosStreamElementos.asp?tipo=JQProcedenciasCamara&pto=<%=g_strPuerto%>&pcia=<%=gCdProvincia%>",
			focus: function( event, ui ) {
				$( "#gDsProcedenciaCamara").val(ui.item.dsloc);
			return false;
			},
			select: function( event, ui ) {
				$( "#gDsProcedenciaCamara"   ).val (ui.item.dslocalidad);
				//$( "#gCdLocalidadCamara"   ).val (ui.item.cdlocalidadcamara + ui.item.cdlocalidadsubcamara);
				auxStr = ui.item.cdlocalidadsubcamara;
				if (auxStr.length==1){
					auxStr = "00" + auxStr;
				}else if (auxStr.length==2){
					auxStr = "0" + auxStr;
				}	
				$( "#gCdProcedenciaCamaraAsoc"   ).val (ui.item.cdlocalidadcamara + auxStr);
				$( "#gCdLocalidadCamara"   ).val (ui.item.cdlocalidadcamara);
				$( "#gDsLocalidadCamara"   ).val (ui.item.dslocalidad);
				$( "#gDsProvinciaCamara"   ).val (ui.item.dsprov);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#gDsProcedenciaCamara").val ("");
					$( "#gCdLocalidadCamara").val ("");
					$( "#gDsLocalidadCamara").val ("");
					$( "#gDsProvinciaCamara").val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a><font style='font-size:10;'>" + item.cdlocalidadcamara + " - " + item.dslocalidad + " - " + item.dsprov + "</font></a>" )
				.appendTo( ul );
		};
	}
	
</script>
<BODY onload="onLoadPage()">
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post>					
<div class="tableaside"> 
	<div class="tableasidecontent"><% call showErrors() %></div>
	<div id="msjGrabar"></div>
	<h3><%=GF_Traducir("Datos de la Procedencia")%></h3>
    <div class="tableasidecontent">
        <h3><%=GF_Traducir("Búsqueda a partir de Oncca")%></h3>
	    <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Provincia") %> </div>
	    <div class="col46"> 
			<select id="gCdProvincia" name="gCdProvincia" onchange="submitInfo('<%=ACCION_CONTROLAR%>');">
				<option value="0"><%= GF_TRADUCIR("Seleccione...")%></option>
				<%
				strSQL = "SELECT CDPROVINCIA, DSPROVINCIA FROM dbo.PROVINCIAS ORDER BY DSPROVINCIA"
				call GF_BD_Puertos (g_strPuerto, rsProvincia, "OPEN",strSQL)										
				while not rsProvincia.eof 
					mySelected = ""
					if cint(gCdProvincia) = cint(rsProvincia("CDPROVINCIA")) then mySelected = "SELECTED"
				%>
						<option value="<%=rsProvincia("CDPROVINCIA")%>" <%=mySelected%>><%=UCASE(rsProvincia("DSPROVINCIA"))%></option>
				<%	rsProvincia.movenext
				wend
				%>							
			</select>									
		</div>	        
        <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Descripción") %> </div>
        <div class="col46">
			<input size=50 type="text" id="gDsProcedenciaOncca" name="gDsProcedenciaOncca"  value="<%= gDsProcedenciaOncca %>">
			<input size=50 type="hidden" id="gCdProcedenciaOncca" name="gCdProcedenciaOncca"  value="<%= gCdProcedenciaOncca %>">			
		</div>
		<div class="col56"><h3><%=GF_Traducir("Datos Internos")%></h3></div>
		<div class="col16"></div>
        <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Código") %> </div>        
        <div class="col46">
			<% If (gCdProcedencia<>0) Then 
				Response.Write gCdProcedencia %>
				<input type="hidden" id="gCdProcedencia" name="gCdProcedencia" value="<%=gCdProcedencia%>">
			<% else 
				Response.Write "Automático"
				%>
					
				<input type="hidden" size=10 id="gCdProcedencia" name="gCdProcedencia" value="<%=gCdProcedencia%>" readonly>
			<% end if %>			
			<input type="hidden" id="gCdSubProcedencia" name="gCdSubProcedencia" value="<%=gCdSubProcedencia%>">
		</div>
	    <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Descripción") %> </div>
        <div class="col46">
			<input size=50 type="text" id="gDsProcedencia" name="gDsProcedencia" value="<%= gDsProcedencia %>">
		</div>

	    <br>
        <div class="col56"><h3><%=GF_Traducir("Datos Cámara (Obligatorio)")%></h3></div>
		<div class="col16"></div>
        <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Búsqueda") %> </div>
        <div class="col46"><input size=50 type="text" id="gDsProcedenciaCamara" name="gDsProcedenciaCamara" value="<%= gDsProcedenciaCamara %>"></div>
        
        <div class="col16 reg_header_navdos"> <% =GF_TRADUCIR("Código Cámara") %> </div>
        <div class="col46">
			<input size=10 type="text" id="gCdProcedenciaCamaraAsoc" name="gCdProcedenciaCamaraAsoc" value="<%= gCdProcedenciaCamaraAsoc %>" readonly>
		</div>

        <!--
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripción Cámara") %> </div>-->
        <div class="col66">
			<FONT size=6>ATENCIÓN: Las modificaciones realizadas impactarán en todos los puertos.</font>
			<input size=50 type="hidden" id="gCdLocalidadCamara" name="gCdLocalidadCamara"  value="<%= gCdLocalidadCamara %>" readonly>
			<input size=50 type="hidden" id="gDsLocalidadCamara" name="gDsLocalidadCamara"  value="<%= gDsLocalidadCamara %>" readonly>
		</div>
        
        
	</div>	
</div>
<input type="hidden" name="accion" id="accion" <%=accion%>>
<input type="hidden" name="isEdit" id="isEdit" value="<%=g_IsEdit%>">
</form>
</BODY>
</HTML>