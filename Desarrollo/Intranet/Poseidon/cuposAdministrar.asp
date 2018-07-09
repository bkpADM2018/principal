<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<%
Const PERIODO_CUPOS = 9 
'-------------------------------------------------------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------------------------------------------------------
Function dibujaLineaTitulo(ByRef p_Rs, pFechaDesde, pDftMax)
    Dim   i, auxCantidad
    i = 0  %>
    <tr id="trDetalle_0">
            <td align="center">Destinatario</td>         
            <td align="center">Corredor</td>
            <td align="center">Vendedor</td>
            <td align="center">Maximo</td>
        <% while (i <= diasCupos)
                fechaActual = GF_DTEADD(pFechaDesde, i, "D")          
                auxCantidad = pDftMax
                if (not p_Rs.eof) then
                    if (CDbl(p_Rs("FECHACUPO")) = CDbl(fechaActual)) then
                        auxCantidad = p_Rs("CANTIDAD")                             
                        p_Rs.MoveNext()
                    end if                        
               end if      %>
               <td align="center">
                    <% =auxCantidad %>  
                    <input type="hidden" id="cantidad_<%=i%>_0" value="<%=auxCantidad %>"/>                    
                </td>
<%              i = i + 1
           wend   %>        
           <td align="center"></td>
           <%  if (puedeAgregar) then %>
           <td align="center"></td>
           <%  end if %>
    </tr>
<%
End Function
'-------------------------------------------------------------------------------------------------------------------------------
Function dibujaLineaCupos(ByRef p_Rs, p_Cliente, p_Corredor, p_Vendedor, p_index, ByRef totCantidad)
    Dim   i
    i = 0  %>
    <tr id="trDetalle_<%=p_index %>">
            <td align="left"><%= getDsClienteByCUIT(p_Cliente) %></td>
            <td align="left"><% if (CLng(p_Corredor) > 0) then response.Write p_Corredor &"-"& getDsCorredor(p_Corredor) %></td>
            <td align="left"><% if (CLng(p_Vendedor) > 0) then response.Write p_Vendedor &"-"& getDsVendedor(p_Vendedor) %></td>
            <td align="center"><div id="totalFila_<%=p_index%>"></div></td>
        <% while (i <= diasCupos)
             fechaActual = GF_DTEADD(fechaDesde, i, "D")
             if (validarFechaCupos(p_Rs, fechaActual, p_Cliente, p_Corredor, p_Vendedor)) then %>
                <td align="center" id="td_<%=i%>_<%=p_index%>"  style="cursor:pointer;" onclick="abrirNominacion(<%=p_Cliente %>,<%=p_Corredor %>,<%=p_Vendedor %>,<%=fechaActual %>,<%=i%>,<%=p_index%>)">
                    <%=p_Rs("CANTIDAD") %>
                    <input type="hidden" id="cantidad_<%=i%>_<%=p_index%>" value="<%=p_Rs("CANTIDAD") %>"/>
                </td>
            <%  totCantidad = Cdbl(totCantidad) + CLng(p_Rs("CANTIDAD"))
                p_Rs.MoveNext()
             else  %>
                <td align="center" id="td_<%=i%>_<%=p_index%>"></td>
        <%   end if                             
             i = i + 1
           wend   %>
        <td align="center">
            <img src="../images/compras/close-16x16.png" title="Eliminar fila" style="cursor:pointer;" onclick="javascript:eliminarNominacion(<%=p_Cliente %>,<%=p_Corredor %>,<%=p_Vendedor %>,'<%=fechaDesde %>','<%=fechaHasta %>',<%=p_index %>,'')" />
        </td>
        <%  if ((puedeAgregar) or (CDbl(p_Cliente) <> CDbl(CUIT_TOEPFER)) or (CLng(p_Vendedor) > 0))then 	%>
        <td align="center">            
            <img src="../images/mail-16.png" title="Enviar Cupos Por Mail" style="cursor:pointer;" onclick="javascript:enviarMail(<%=p_Cliente %>,<%=p_Corredor %>,<%=p_Vendedor %>,'<%=fechaDesde %>','<%=fechaHasta %>')" />            
        </td>        
        <%  end if %>
    </tr>
<%
End Function
'-------------------------------------------------------------------------------------------------------------------------------
Function validarFechaCupos(p_Rs, p_Fecha, p_Cliente, p_Corredor, p_Vendedor)
    validarFechaCupos = false
    if (not p_Rs.Eof) then
        if ((Cdbl(p_Rs("CUITCLIENTE")) = Cdbl(p_Cliente)) and (Cdbl(p_Rs("CDCORREDOR")) = Cdbl(p_Corredor))and(Cdbl(p_Rs("CDVENDEDOR")) = Cdbl(p_Vendedor))) then
            if(Cdbl(p_Rs("FECHACUPO")) = Cdbl(p_Fecha)) then validarFechaCupos = true
        end if
    end if
End function
'********************************************************************************************************************************
'********************************************************* INICIO DE PAGINA *****************************************************
'********************************************************************************************************************************
Dim  rs,g_strPuerto,rsCup,rsDet,fechaDesde,fechaHasta,diasCupos,strParam, cdProducto
Dim strSQL,  cuitCupeador, myWhere, puedeAgregar, maxCuposDisponibles
Dim myLckUsr, myLckKey, chkDetallar, fc
Dim myDsCorredor, flagEsCorredor, myDsCliente, myCuitCliente

Call initTaskAccessInfo(TASK_POS_ADMIN_CUPOS, session("DIVISION_PUERTO"))
myLckUsr = getLckUser(session("Usuario"))
'*** DEFINICIONES ***
'Cupeador: es quien ingresa a la pagina para cargar/nominar cupos.
'Cliente : Destinatario de los cupos. PAra Toepfer, puede elegir cliente, para terceros, solo trabajaran con sus cupos y podran nominar corredor y vendedor.
'********************

cuitCupeador = GF_PARAMETROS7("cuitCupeador",0,6)
Call addParam("cuitCupeador",cuitCupeador,strParam)

'Se controla el acceso - Solo se permite elegir el proveedor por parametro si el usuario de la session es TOEPFER
if (CDbl(cuitCupeador) <> CDbl(session("CuitOrganizacion"))) then
    response.Redirect "../comprasAccesoDenegado.asp"
end if


cdProducto = GF_PARAMETROS7("cdProducto",0,6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
Call addParam("pto",g_strPuerto,strParam)

Call GP_CONFIGURARMOMENTOS

lckMsg = false
'if (CDbl(cuitCupeador) = CDbl(CUIT_TOEPFER)) then
'    if (cdProducto > 0) then	
'	    myLckKey = LCK_LOGISTICA & "_" & cuitCupeador & "_" & cdProducto
'	    if ((myLckKey <> session("LastLCK_" & LCK_LOGISTICA)) and (session("LastLCK_" & LCK_LOGISTICA) <> "")) then Call releaseLckKey(g_strPuerto, myLckUsr, session("LastLCK_" & LCK_LOGISTICA))
'	    if (not checkLckKey(g_strPuerto, myLckKey, myLckUsr)) then
'		    cdProducto = 0
'		    lckMsg = true		
'		    session("LastLCK_" & LCK_LOGISTICA) = ""		
'	    else
'		    'Si anters estaba en otro producto debo liberarlo.		
'		    session("LastLCK_" & LCK_LOGISTICA) = myLckKey
'	    end if
 '   else
'	    Call releaseLckKey(g_strPuerto, myLckUsr, session("LastLCK_" & LCK_LOGISTICA))
'	    session("LastLCK_" & LCK_LOGISTICA) = ""
 '   end if
'end if

fechaDesde = GF_PARAMETROS7("fd", "", 6)
if (fechaDesde = "") then fechaDesde = GF_DTEADD(Left(Session("MmtoDato"),8), 0, "D")
fechaHasta = GF_DTEADD(fechaDesde, PERIODO_CUPOS, "D")
diasCupos = GF_DTEDIFF(fechaDesde ,fechaHasta ,"D")
chkDetallar = GF_PARAMETROS7("chkDetallar", "", 6)
if (chkDetallar <> "") then chkDetallar = "checked"

fc = GF_PARAMETROS7("fc","",6)
flagEsCorredor = false
if (fc="C") then flagEsCorredor = true                    
if (fc = "") then
    'No se indico el rol, se determina uno para sugerir.
    myDsCliente=getDsClienteByCUIT(cuitCupeador)
    'Si no se encuentra la descripcion del cliente, implica que quien entro a cupear es un corredor por ende los cupos son para Destinatario = ADM Agro.
    if (Trim(myDsCliente) = "") then  
        flagEsCorredor = true					                          
        fc="C"
    end if
end if					    
if (not flagEsCorredor) then
    myDsCliente=getDsClienteByCUIT(cuitCupeador)
    myCuitCliente = cuitCupeador
else
	myDsCliente=Trim(getDsClienteByCUIT(CUIT_TOEPFER))
	myCuitCliente = CUIT_TOEPFER										
end if
		                    
puedeAgregar = (CDbl(cuitCupeador) = CDbl(CUIT_TOEPFER))

'Leo los productos
if (not puedeAgregar) then    
    strSQL= "Select DISTINCT C.CDPRODUCTO, DSPRODUCTO, count(*) CANTIDAD from CODIGOSCUPO C " &_
            "   inner join PRODUCTOS P on C.CDPRODUCTO=P.CDPRODUCTO where "  &_ 
            " ((C.CUITCLIENTE = " &  cuitCupeador & " and C.ESTADO >= " & CUPO_OTORGADO & ")" &_    
	        " or (C.CUITCLIENTE = '" & CUIT_TOEPFER & "' and C.CDCORREDOR = " & session("KCOrganizacion") & " and C.ESTADO >= " & CUPO_PROVISORIO & "))"	&_
            "      and C.FECHACUPO <= " & fechaHasta &_
            "       and C.FECHACUPO >= " & fechaDesde &_
            "   group by C.CDPRODUCTO, DSPRODUCTO " &_
            "   order by DSPRODUCTO"            
    Call executeQueryDb(g_strPuerto, rsProductos, "OPEN", strSQL)    
else    
    strSQL= "Select DISTINCT P.CDPRODUCTO, DSPRODUCTO, CANTIDAD " &_
            " from PRODUCTOS P " &_
            " left join (Select CDPRODUCTO, count(*) CANTIDAD " &_
			"               from CODIGOSCUPO  " &_
            "               where FECHACUPO <= " & fechaHasta &_
            "               and FECHACUPO >= " & fechaDesde &_          
            "               and ESTADO > " & CUPO_CANCELADO &_
            "               group by CDPRODUCTO ) C on C.CDPRODUCTO=P.CDPRODUCTO " &_
            "   order by CANTIDAD desc, DSPRODUCTO "            
    Call executeQueryDb(g_strPuerto, rsProductos, "OPEN", strSQL)        
end if
%>
<html>
<head>
<title>Sistema de Cupos</title>

<meta http-equiv="x-ua-compatible" content="IE=11">

<link rel="stylesheet" href="../css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="../css/tabs-print.css" TYPE="text/css" MEDIA="print">
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
    .selectorColumn {
        background: #FF6666;
        color: #FFFFFF;
    }
    .inputImgNominacion {
        background:none;
        border:none;
    }
</style>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>

<script type="text/javascript">
    var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");
    //Cargamos la funcion inicio de la pagina
    var ch = new channel();        
    
    var arrTotales = new Array();
    var arrPendientes = new Array();
    function bodyOnLoad() {
        <% if (Cdbl(cdProducto) <> 0) then 
            if (puedeAgregar) then
        %>
        document.getElementById("dsCliente").focus();
        <%  else %>            
        document.getElementById("dsCorredor").focus();
        <%  end if %>
        var maxIndexV = document.getElementById("maxIndexVertical").value;
        //Dibujo los valores totales por fila en el div de la celda Total
        for (var i = 1; i < maxIndexV; i++) {
            if (document.getElementById("totalFila_" + i)) document.getElementById("totalFila_" + i).innerHTML = document.getElementById("vlTotal_" + i).value;
        }
        autoCompleteCorredor();
        autoCompleteVendedor();
        <%  if (puedeAgregar) then %>
        autoCompleteCliente();
        <%  end if %>
        cargarNominacionesParciales();            
        <% else %>
            document.getElementById("cdProducto").focus();
        <% end if %>
        }
    function cargarNominacionesParciales() {
        var maxIndexV = document.getElementById("maxIndexVertical").value;
        //Primero inicializo el array con valores por defecto
        iniciarArrayTotal();
        //Luego calculo las nominaciones realizadas
        //i: indice vertical (variable dependiendo las lineas que tenga)
        for (var i = 0; i < maxIndexV; i++) {
            //y: indice horizontal
            for (var y = 0; y < arrPendientes.length; y++) {
                if (document.getElementById("cantidad_" + y + "_" + i)) {
                    if (i == 0)
                        arrPendientes[y] = document.getElementById("cantidad_" + y + "_" + i).value;
                    else {
                        arrPendientes[y] = arrPendientes[y] - document.getElementById("cantidad_" + y + "_" + i).value;
                        arrTotales[y] = parseInt(arrTotales[y]) + parseInt(document.getElementById("cantidad_" + y + "_" + i).value);
                    }                        
                }
            }
        }
        //Por ultimo imprimo el resultado de los totales en la tabla
        dibujarTotalesColumnas();
    }
    function iniciarArrayTotal() {
        for (var i = 0; i <= "<%= diasCupos%>"; i++) {
            arrPendientes[i] = 0;
            arrTotales[i] = 0;
        }
       
    }
    //Creo los elementos totales para la tabla
    function dibujarTotalesColumnas() {
        for (var i = 0; i < arrPendientes.length; i++) {
            var td = document.getElementById("tdPendiente_"+i);            
            td.innerHTML = arrPendientes[i];
            if (parseInt(arrPendientes[i]) < 0)td.style.color = "red";
            var td2 = document.getElementById("tdTotalProd_"+i);
            td2.innerHTML = arrTotales[i];
        }
    }
    //Permite reacalcular los totales de la columna del cupo que se edito, se ejecuta con el evento onblur
    function actualizarTotales(e, p_Index) {            
        var parcial = arrPendientes[p_Index];
        var subTotal = arrTotales[p_Index];
        var cant = e.value;       
        if (isNaN(cant) || cant == '') cant = 0;
        parcial -= cant;
        subTotal += parseInt(cant);
        document.getElementById("tdPendiente_" + p_Index).innerHTML = parcial;
        document.getElementById("tdTotalProd_" + p_Index).innerHTML = subTotal;
        if (parseInt(parcial) < 0)
            document.getElementById("tdPendiente_" + p_Index).style.color = "red";
        else
            document.getElementById("tdPendiente_" + p_Index).style.color  = "#666";
    }

    function comprobarDescripcion(elementDs, elementCd, elementCuit) {
        if (trimStr(elementDs.value) == "") {
            elementCd.value = "";
            elementCuit.value = "";
        }
    }

    //Esta funcion quita los espacios en blanco de un string
    //NOTA: no se utiliza la funcion trim() de javascript debido a que en Internet Explorer 8 no es compatible
    function trimStr(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }
    function autoCompleteCorredor() {
        $("#dsCorredor").autocomplete({
            minLength: 4,
            source: "puertosStreamElementos.asp?tipo=JQCorredores&pto=<%=g_strPuerto%>",
            focus: function (event, ui) {
                $("#dsCorredor").val(ui.item.dscorredor);
                return false;
            },
            select: function (event, ui) {
                $("#dsCorredor").val(ui.item.dscorredor);
                $("#cdCorredor").val(ui.item.cdcorredor);
                $("#cuitCorredor").val(ui.item.nucuit);
                if (document.getElementById("cuitCorredor").style.display == "block") {
                    document.getElementById("cuitCorredor").style.display = "none";
                    document.getElementById("divCuitCorredor").style.display = "none";
                }
            }
        })
		.data("autocomplete")._renderItem = function (ul, item) {
		    return $("<li></li>")
				.data("item.autocomplete", item)
				.append("<a>" + item.cdcorredor + " - <font style='font-size:10;'>" + item.dscorredor + "</font></a>")
				.appendTo(ul);
		};
    }
    function autoCompleteVendedor() {
        $("#dsVendedor").autocomplete({
            minLength: 4,
            source: "puertosStreamElementos.asp?tipo=JQVendedores&pto=<%=g_strPuerto%>",
            focus: function (event, ui) {
                $("#dsVendedor").val(ui.item.dsvendedor);
                return false;
            },
            select: function (event, ui) {
                $("#dsVendedor").val(ui.item.dsvendedor);
                $("#cdVendedor").val(ui.item.cdvendedor);
                $("#cuitVendedor").val(ui.item.nudocumento);
                if (document.getElementById("cuitVendedor").style.display == "block"){
                    document.getElementById("cuitVendedor").style.display = "none";
                    document.getElementById("divCuitVendedor").style.display = "none";
                }
            }
        })
		.data("autocomplete")._renderItem = function (ul, item) {
		    return $("<li></li>")
				.data("item.autocomplete", item)
				.append("<a>" + item.cdvendedor + " - <font style='font-size:10;'>" + item.dsvendedor + "</font></a>")
				.appendTo(ul);
		};
    }
    
    function autoCompleteCliente() {
        $("#dsCliente").autocomplete({
            minLength: 4,
            source: "puertosStreamElementos.asp?tipo=JQClientes&pto=<%=g_strPuerto%>",
            focus: function (event, ui) {
                $("#dsCliente").val(ui.item.dscliente);
                return false;
            },
            select: function (event, ui) {
                $("#dsCliente").val(ui.item.dsvendedor);
                $("#cdCliente").val(ui.item.cdCliente);
                $("#cuitCliente").val(ui.item.nucuit);                
            }
        })
		.data("autocomplete")._renderItem = function (ul, item) {
		    return $("<li></li>")
				.data("item.autocomplete", item)
				.append("<a>" + item.cdcliente + " - <font style='font-size:10;'>" + item.dscliente + "</font></a>")
				.appendTo(ul);
		};
    }
    
    function controlarNominacion() {
        document.getElementById("btnSaveNominacion").style.display = "none";
        document.getElementById("loadingNominacion").style.display = "block";
        var msgError = "";
        manejoErrorNominacion(msgError);

        //Controlo Cliente        
        if (document.getElementById("cuitCliente").value == "") {
            document.getElementById("cuitCliente").value='<% =myCuitCliente %>';
            document.getElementById("dsCliente").value = '<% =myDsCliente %>';
        }
        var dsCliente   = document.getElementById("dsCliente").value;
        var cuitCliente = document.getElementById("cuitCliente").value;
        if(!controlarCliente(cuitCliente)) msgError = msgError + "<li>No se encuentra registrado el Destinatario.</li>";
        
        //Controlo el corredor
        var dsCorredor   = document.getElementById("dsCorredor").value;
        var cdCorredor   = document.getElementById("cdCorredor").value;
        var cuitCorredor = document.getElementById("cuitCorredor").value;        
        if(!controlarCorredor(cdCorredor, dsCorredor, cuitCorredor)) msgError = msgError + "<li>No se encuentra registrado el corredor, ingrese el CUIT para completar el registro.</li>";

        //Controlo el vendedor
        var dsVendedor = document.getElementById("dsVendedor").value;
        var cdVendedor = document.getElementById("cdVendedor").value;
        var cuitVendedor = document.getElementById("cuitVendedor").value;        
        if(!controlarVendedor(cuitCliente, cdVendedor, dsVendedor, cuitVendedor)) msgError = msgError + "<li>No se encuentra registrado el vendedor, ingrese el CUIT para completar el registro.</li>";

        //Controlo la cantidad
        if (!controlarCantidadCupo()) msgError = msgError + "<li>Se encontro un error en la cantidad de cupo ingresada.</li>";

        //Controlo los totales
        if(!controlarTotales()) msgError = msgError + "<li>Se encontraron totales negativos.</li>";

        if (msgError == "") {
            var strParameter = "accion=<%=ACCION_GRABAR %>&cuitCupeador=<% =cuitCupeador %>&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&fechaDesde=<%=fechaDesde%>&fechaHasta=<%=fechaHasta%>&fc=<% =fc %>";
            for (var i = 0; i <= "<%=diasCupos%>"; i++) {
                strParameter = strParameter + "&cupo_" + i + "=" + document.getElementById("cupo_" + i).value;
            }
            strParameter = strParameter + "&cuitCliente=" + cuitCliente;
            strParameter = strParameter + "&cdVendedor=" + cdVendedor + "&dsVendedor=" + dsVendedor + "&cuitVendedor=" + cuitVendedor;
            strParameter = strParameter + "&cdCorredor=" + cdCorredor + "&dsCorredor=" + dsCorredor + "&cuitCorredor=" + cuitCorredor;
            
            ch.bind("cuposAdministrarAjax.asp?" + strParameter, "nominarCupoControl_Callback()");
            ch.send();
        }
        else {
            manejoErrorNominacion(msgError);
            document.getElementById("btnSaveNominacion").style.display = "block";
            document.getElementById("loadingNominacion").style.display = "none";
        }

    }

    function nominarCupoControl_Callback() {
        var respuesta = ch.response();
        if (respuesta == "LCK") cambiarProducto();
        document.getElementById("btnSaveNominacion").style.display = "block";
        document.getElementById("loadingNominacion").style.display = "none";
        respuesta = respuesta.split("|");
        if (respuesta[0] == "<%=RESPUESTA_OK%>") {
            restaurarNominacion(respuesta[1],respuesta[2],respuesta[3]);
            //verifico si esta seleccionada una columna para eliminar, en ese caso se quita la seleccion
            if (parseInt(document.getElementById("indiceCol_Old").value) >= 0){
                seleccionarColumnaFecha(document.getElementById("indiceCol_Old").value, "");
                document.getElementById("trEliminar").style.display = "none";
                document.getElementById("trMail").style.display = "none";                
            }
<%      if (puedeAgregar) then %>
            document.getElementById("dsCliente").focus();
<%      else             
            if (flagEsCorredor) then %>             
            document.getElementById("dsVendedor").focus();
<%          else %>            
            document.getElementById("dsCorredor").focus();
<%          end if
        end if           %>            
        }
        else {
            manejoErrorNominacion(respuesta[0]);
        }
    }
    
    function controlarVendedor(p_cuitCliente, p_CdVendedor, p_DsVendedor, p_CuitVendedor) {
        var ret = true;
        if (p_CdVendedor == "") {
            if ((trimStr(p_DsVendedor) != "") && (trimStr(p_CuitVendedor) != "")) {
                ret = true;
            } else {
                ret = false;
            <% if (puedeAgregar) then %>
				if ((trimStr(p_DsVendedor) == "") && (trimStr(p_CuitVendedor) == "")) ret = true;
               //if ((trimStr(p_DsVendedor) == "") && (trimStr(p_CuitVendedor) == "") && (p_cuitCliente != '<% =CUIT_TOEPFER %>')) ret = true;                
            <% end if %>                                    
                if ((document.getElementById("cuitVendedor").style.display == "none") && (!ret)){
                    document.getElementById("divCuitVendedor").style.display = "block";
                    document.getElementById("cuitVendedor").style.display = "block";                    
                }                     
            }
        } else {
            if (document.getElementById("cuitVendedor").style.display == "block"){
                document.getElementById("divCuitVendedor").style.display = "none";
                document.getElementById("cuitVendedor").style.display = "none";
            } 
        }
        return ret;
    }
    function controlarCorredor(p_CdCorredor, p_DsCorredor, p_CuitCorredor) {
        var ret = true;
        if (p_CdCorredor == "") {
            if (((trimStr(p_DsCorredor) != "") && (trimStr(p_CuitCorredor) != "")) || ((trimStr(p_DsCorredor) == "") && (trimStr(p_CuitCorredor) == ""))) {
                ret = true;
            }
            else {                
                ret = false;
                if (document.getElementById("cuitCorredor").style.display == "none") {
                    document.getElementById("cuitCorredor").style.display = "block";
                    document.getElementById("divCuitCorredor").style.display = "block";
                }                                   
            }
        }
        else {
            if (document.getElementById("cuitCorredor").style.display == "block"){
                document.getElementById("divCuitCorredor").style.display = "none";
                document.getElementById("cuitCorredor").style.display = "none";
            } 
        }
        return ret;
    }
    function controlarCliente(p_CuitCliente) {
        var ret = true;
        if (trimStr(p_CuitCliente) == "") ret = false;                
        return ret;
    }
    
    //Es el control que se realiza del lado del cliente para ver las cantidades ingresadas, para eso se toman en cuenta los 
    //datos de la cabecera (por columna), las nominaciones realizadas y los totales
    function controlarCantidadCupo() {
        var ret = true;
        //este flag permite saber si se ingreso algun cupo en los textbox de cada columna, de esta manera valido que no llame al ajax sin haber completado algo
        var flagCompletoCupos = false;
        var i = 0;
        while ((i < arrPendientes.length) && (ret)) {
            if (document.getElementById("cantidad_" + i + "_0")) { 
                var cupo = document.getElementById("cupo_" + i).value;
                if (cupo > 0) flagCompletoCupos = true
                if (parseInt(cupo) < 0) ret = false;
            }
            else {
                if ((parseInt(document.getElementById("cupo_" + i).value) > 0)||(parseInt(document.getElementById("cupo_" + i).value) < 0)) ret = false;
            }
            i ++;
        }
        if (!flagCompletoCupos) ret = false;

        return ret;
    }
    function controlarTotales(){
        var ret = true;
        var i = 0;
        while ((i < arrPendientes.length) && (ret)) {
            var cupo = 0;
            if(document.getElementById("cupo_" + i).value > 0) cupo = document.getElementById("cupo_" + i).value;
            var totales = parseInt(arrPendientes[i]) - parseInt(cupo);
            if(parseInt(totales) < 0) ret = false;
            i++;
        }
        return ret;
    }
    function manejoErrorNominacion(p_Error) {
        if (trimStr(p_Error) != "") {
            document.getElementById("tdError").className = "TDERROR";
            document.getElementById("dsError").style.display = "block";
        }
        else {
            document.getElementById("tdError").className = "";
            document.getElementById("dsError").style.display = "none";
        }
        document.getElementById("dsError").innerHTML = p_Error;
    }

    function eliminarNominacion(p_cuitCliente, p_CdCorredor, p_CdVendedor, p_FechaDesde, p_FechaHasta,p_IndiceFila,p_IndiceColumna) {
        if (confirm("Desea eliminar las nominaciones?")) {
            var strParameter = "accion=<%=ACCION_BORRAR %>&cuitCupeador=<% =cuitCupeador %>&cuitCliente=" + p_cuitCliente + "&fechaDesde=" + p_FechaDesde + "&fechaHasta=" + p_FechaHasta + "&cdVendedor=" + p_CdVendedor + "&cdCorredor=" + p_CdCorredor + "&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&fc=<% =fc %>";
            ch.bind("cuposAdministrarAjax.asp?" + strParameter, "eliminarNominacion_Callback('"+p_IndiceFila+"','"+p_IndiceColumna+"')");
            ch.send();
        }
    }
        
    //Elimino la fila o columna fisicamente de la pagina
    function eliminarNominacion_Callback(p_IndiceFila, p_IndiceColumna) {
        var resp = ch.response();
        if (resp == "LCK") cambiarProducto();
        if (p_IndiceFila != "") {
            //ELIMNA FILA
            $("#trDetalle_"+p_IndiceFila).remove();
            //Elimino los elementos hidden relacionados con la fila
            $("#vlTotal_"+p_IndiceFila).remove();
            $("#cdCorredor_"+p_IndiceFila).remove();
            $("#cdVendedor_"+p_IndiceFila).remove();
            //verifico si esta seleccionada una columna para eliminar, en ese caso se quita la seleccion
            //verifico si esta seleccionada una columna para eliminar, en ese caso se quita la seleccion
            if (parseInt(document.getElementById("indiceCol_Old").value) >= 0){
                seleccionarColumnaFecha(document.getElementById("indiceCol_Old").value, "");
                document.getElementById("trEliminar").style.display = "none";
                document.getElementById("trMail").style.display = "none";                
            }
        }
        else{
            //ELIMNA COLUMNA
            var arrFilasBorradas = new Array();
            //Obtenemos el maximo indice de filas
            var indexMaxFilas = document.getElementById("maxIndexVertical").value;
            for (var i = 1; i < indexMaxFilas; i++) {
                if (document.getElementById("td_" + p_IndiceColumna + "_" + i)){
                    if(document.getElementById("cantidad_" + p_IndiceColumna + "_" + i)){
                        //actualizo la columna de total por fila restandole el valor que tenia la celda
                        document.getElementById("vlTotal_"+i).value = parseInt(document.getElementById("vlTotal_"+i).value) - parseInt(document.getElementById("cantidad_" + p_IndiceColumna + "_" + i).value);
                        document.getElementById("totalFila_"+i).innerHTML = document.getElementById("vlTotal_"+i).value;
                        //Si el total por fila es 0 nominaciones significa que la fila debe eliminarse, para eso se guarda en un array las filas que no tienen datos y deben borrarse
                        if (parseInt(document.getElementById("vlTotal_"+i).value) == 0) arrFilasBorradas.push(i);
                    }
                    //limpio la celda
                    document.getElementById("td_" + p_IndiceColumna + "_" + i).innerHTML = "";
                    $("#td_" + p_IndiceColumna + "_" + i).removeAttr("onclick");
                    $("#td_" + p_IndiceColumna + "_" + i).removeAttr("style");
                }
            }
            //quito el efecto de eliminacion en la columna y oculto la fila de eliminar
            seleccionarColumnaFecha(document.getElementById("indiceCol_Old").value, "");
            document.getElementById("trEliminar").style.display = "none";
            document.getElementById("trMail").style.display = "none";            
            //Recorro las filas que no tienen nominaciones para eliminarlas
            for (var i = 0; i <= arrFilasBorradas.length; i++) {
                var fila = arrFilasBorradas.pop();
                if (existeFilaNominacion(fila)) {
                    //Elimino la fila
                    $("#trDetalle_" + fila).remove();
                    //Elimino los elementos hidden relacionados con la fila
                    $("#vlTotal_"+fila).remove();
                    $("#cdCorredor_"+fila).remove();
                    $("#cdVendedor_"+fila).remove();
                }
            }
        }
        //Recalculo todos los totales nuevos y los imprimo
        cargarNominacionesParciales();
        
    }

    function enviarMail(p_cuitCliente, p_CdCorredor, p_CdVendedor, p_FechaDesde, p_FechaHasta) {        
        var strParameter = "accion=<%=ACCION_EMAIL %>&cuitCupeador=<% =cuitCupeador %>&cuitCliente=" + p_cuitCliente + "&fechaDesde=" + p_FechaDesde + "&fechaHasta=" + p_FechaHasta + "&cdVendedor=" + p_CdVendedor + "&cdCorredor=" + p_CdCorredor + "&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&fc=<% =fc %>";
        var puw = new winPopUp('popupMail',"cuposMailPopUp.asp?" + strParameter, 450, 480, 'Envio de Mail', "");                
        //ch.bind("cuposAdministrarAjax.asp?" + strParameter, "enviarMail_callback()");
        //ch.send();
    }
    
    function verCupos(p_cuitCliente, p_CdCorredor, p_CdVendedor, p_FechaDesde, p_FechaHasta) {                
        var strParameter = "accion=<%=ACCION_VISUALIZAR %>&cuitCupeador=<% =cuitCupeador %>&cuitCliente=" + p_cuitCliente + "&fechaDesde=" + p_FechaDesde + "&fechaHasta=" + p_FechaHasta + "&cdVendedor=" + p_CdVendedor + "&cdCorredor=" + p_CdCorredor + "&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&fc=<% =fc %>";
        window.open("cuposAdministrarAjax.asp?" + strParameter);
    }
    
    function enviarMail_callback(p_IndiceFila, p_IndiceColumna) {
        var resp = ch.response();
        if (resp == "LCK") cambiarProducto();
        if (p_IndiceFila == "") {        
            //quito el efecto de eliminacion en la columna y oculto la fila de eliminar
            seleccionarColumnaFecha(document.getElementById("indiceCol_Old").value, "");            
            document.getElementById("trEliminar").style.display = "none";
            document.getElementById("trMail").style.display = "none";   
        }        
        document.getElementById("dsError").className = "TDSUCCESS";        
        document.getElementById("dsError").innerHTML = resp;
        $("#dsError").removeAttr("style");
        var om = document.getElementById("dsError");
        setTimeout(function() { om.style.display = "none"; } , 5000)
    }
    
    //Esta funcion permite seleccionar una columna para poder eliminarla, crea un estilo que la diferencia de otras columnas
    function seleccionarColumnaFecha(p_IndiceColumna, p_Fecha) {
        var maxIndex = document.getElementById("maxIndexVertical").value;
        var indexColumnOld = document.getElementById("indiceCol_Old").value;
        
        if (document.getElementById("imgEliminar")) {
            $("#imgEliminar").remove();            
            $("#imgMail").remove();            
            for (var i = 0; i <= maxIndex; i++) {
                if (document.getElementById("td_" + indexColumnOld + "_" + i)) document.getElementById("td_" + indexColumnOld + "_" + i).className = "";
            }
            if (document.getElementById("tdPendiente_" + indexColumnOld)) document.getElementById("tdPendiente_" + indexColumnOld).className = "";
        }
        document.getElementById("indiceCol_Old").value = -1;
        document.getElementById("trEliminar").style.display = "none";
        //document.getElementById("trMail").style.display = "none";        
        if (indexColumnOld != p_IndiceColumna) {
            //creo el efecto de la seleccion para la cabecera y las nominaciones ya echas
            for (var i = 0; i <= maxIndex; i++) {
                if (document.getElementById("td_" + p_IndiceColumna + "_" + i)) document.getElementById("td_" + p_IndiceColumna + "_" + i).className = "selectorColumn";
            }
            if (document.getElementById("tdPendiente_" + p_IndiceColumna)) document.getElementById("tdPendiente_" + p_IndiceColumna).className = "selectorColumn";
            //dibujo la imagen para eliminar
            var td = document.getElementById("tdEliminar_" + p_IndiceColumna)
            td.align = "center";
            var img = document.createElement("img");
            img.id = "imgEliminar"
            img.src = "../images/compras/close-16x16.png";
            img.title = "Quitar asignacion a cupos del dia";
            img.style.cursor = "pointer"
            if (isFirefox) {
                img.setAttribute('onclick', "eliminarNominacion('','','',"+p_Fecha+","+p_Fecha+",'',"+p_IndiceColumna+")");
            } else {
                img['onclick'] = new Function("eliminarNominacion('','','',"+p_Fecha+","+p_Fecha+",'',"+p_IndiceColumna+")");
            }
            td.appendChild(img);
            $("#trEliminar").removeAttr("style");
            <%  'if (puedeAgregar) then %>
            //dibujo la imagen para mail
            //var td = document.getElementById("tdMail_" + p_IndiceColumna)
            //td.align = "center";
            //var img = document.createElement("img");
            //img.id = "imgMail"
            //img.src = "../images/mail-16.png";
            //img.title = "Otorgar por mail cupos del dia";
            //img.style.cursor = "pointer"
            //if (isFirefox) {
            //    img.setAttribute('onclick', "enviarMail('','','',"+p_Fecha+","+p_Fecha+",'',"+p_IndiceColumna+")");
            //} else {
            //    img['onclick'] = new Function("enviarMail('','','',"+p_Fecha+","+p_Fecha+",'',"+p_IndiceColumna+")");
            //}
            //td.appendChild(img);            
            //$("#trMail").removeAttr("style");
            <%  'end if %>
            document.getElementById("indiceCol_Old").value = p_IndiceColumna;
        }            
    }
    
    function keyPressedEnter(evento) {
        key=(document.all) ? evento.keyCode : evento.which;
        if(key==13){        
            controlarNominacion();
            return false;
        }
        else{
            return true;
        }
    }
    function controlNewkeyPressed(elemento,evento,tipo){
        if (keyPressedEnter(evento)){
            return controlIngreso(elemento,evento,tipo);
        }
    }
    function cambiarProducto(){
        document.getElementById("frmSel").submit();
    }
    function restaurarNominacion(p_cuitCliente, p_CdVendedor, p_CdCorredor){
        
        //Obtengo el numero de indices de las columnas
        var maxIndex = document.getElementById("maxIndexVertical").value;

        //Primero debo saber si el codigo de corredor y de vendedor grabados se encuentran ya publicados en las nominaciones
        //Para eso debo recorrer la columna de corredor/vendedor para actualizar sus valores, si hay coincidencia guardo el indice encontrado
        var indexLinea = buscarClienteCorredorVendedor(p_cuitCliente, p_CdVendedor, p_CdCorredor, maxIndex);
        
        if (indexLinea == 0){
            //El cliente/corredor/vendedor son nuevos ,creo una nueva fila donde fijara los datos grabados
            var tr = document.createElement("tr");
            tr.id = "trDetalle_"+maxIndex;
            
            var tdCliente = document.createElement("td");
            tdCliente.align = "left";
            tdCliente.innerHTML = document.getElementById("dsCliente").value;
            tr.appendChild(tdCliente);
    
            var tdCorredor = document.createElement("td");
            tdCorredor.align = "left";
            var auxCorredor = "";
            if (p_CdCorredor > 0) auxCorredor = p_CdCorredor +"-"+ document.getElementById("dsCorredor").value;
            tdCorredor.innerHTML = auxCorredor;
            tr.appendChild(tdCorredor);

            var tdVendedor = document.createElement("td");
            tdVendedor.align = "left";
            var auxVendedor = "";
            if (p_CdVendedor > 0) auxVendedor = p_CdVendedor +"-"+ document.getElementById("dsVendedor").value;
            tdVendedor.innerHTML = auxVendedor;
            tr.appendChild(tdVendedor);

            var tdTotal = document.createElement("td");
            tdTotal.align = "center";
            var divTotal = document.createElement("div");
            divTotal.id = "totalFila_" + maxIndex;
            tdTotal.appendChild(divTotal);
            tr.appendChild(tdTotal);

            var totalCupos = 0;
            for (var i = 0; i <= "<%= diasCupos%>"; i++){
                var td = document.createElement("td");
                td.id = "td_"+ i +"_" + maxIndex;
                td.align = "center";
                //obtengo los valores de cupos ingresados end cada TextBox
                var valorCupo = document.getElementById("cupo_" + i).value;
                if (parseInt(valorCupo) > 0){
                    td.style.cursor = "pointer";
                    if (isFirefox)
                        td.setAttribute('onclick', "javascript:abrirNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+","+ document.getElementById("colFecha_"+i).value +","+ i +","+maxIndex+");");
                    else 
                        td['onclick'] = new Function("javascript:abrirNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+","+ document.getElementById("colFecha_"+i).value +","+ i +","+maxIndex+");");
                    td.innerHTML = valorCupo; 
                    var vlCupo = document.createElement("input");
                    vlCupo.type = "hidden";
                    vlCupo.id = "cantidad_"+ i +"_" +maxIndex;
                    vlCupo.value = valorCupo;
                    td.appendChild(vlCupo);
                    totalCupos = parseInt(totalCupos) + parseInt(valorCupo);
                }
                if (parseInt(document.getElementById("indiceCol_Old").value) == i) td.className = "selectorColumn";
                tr.appendChild(td);
            }
            divTotal.innerHTML = totalCupos;

            var tdDelete  = document.createElement("td");
            tdDelete.align = "center";
            var imgDelete = document.createElement("img");
            imgDelete.src = "../images/compras/close-16x16.png";
            imgDelete.title = "Eliminar fila";
            imgDelete.style.cursor = "pointer";
            if (isFirefox)
                imgDelete.setAttribute('onclick', "javascript:eliminarNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+",<%=fechaDesde %>,<%=fechaHasta %>,"+maxIndex+",'');");
            else 
                imgDelete['onclick'] = new Function("javascript:eliminarNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+",<%=fechaDesde %>,<%=fechaHasta %>,"+maxIndex+",'');");
            tdDelete.appendChild(imgDelete);
            tr.appendChild(tdDelete);            
            
			var mailEnvio = false
            <%  
				if (puedeAgregar) then 
			%>
				mailEnvio = true
			<%
				end if
			%>
			if ((mailEnvio) || (p_cuitCliente != <% =CUIT_TOEPFER %>) || (p_CdVendedor > 0)) {
				var tdMail  = document.createElement("td");            
				tdMail.align = "center";
				var imgMail = document.createElement("img");
				imgMail.src = "../images/mail-16.png";
				imgMail.title = "Eliminar fila";
				imgMail.style.cursor = "pointer";
				if (isFirefox)
					imgMail.setAttribute('onclick', "javascript:enviarMail("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+",<%=fechaDesde %>,<%=fechaHasta %>);");
				else 
					imgMail['onclick'] = new Function("javascript:enviarMail("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+",<%=fechaDesde %>,<%=fechaHasta %>);");
				tdMail.appendChild(imgMail);            
				tr.appendChild(tdMail);     
			}
            var parentTR = document.getElementById("trCupos").parentNode;
            parentTR.insertBefore(tr, document.getElementById("trCupos"));

            //Agrego los input hidden necesarios para cada Fila (se los agrega al final)
            var vlTotal = document.createElement("input");
            vlTotal.type = "hidden";
            vlTotal.id = "vlTotal_"+ maxIndex;
            vlTotal.name = "vlTotal_"+ maxIndex;
            vlTotal.value = totalCupos;
            parentTR.insertBefore(vlTotal, document.getElementById("trCupos"));

            var cuitCliente = document.createElement("input");
            cuitCliente.type = "hidden";
            cuitCliente.id = "cuitCliente_"+ maxIndex;
            cuitCliente.name = "cuitCliente_"+ maxIndex;
            cuitCliente.value = p_cuitCliente;
            parentTR.insertBefore(cuitCliente, document.getElementById("trCupos"));
            
            var cdCorredor = document.createElement("input");
            cdCorredor.type = "hidden";
            cdCorredor.id = "cdCorredor_"+ maxIndex;
            cdCorredor.name = "cdCorredor_"+ maxIndex;
            cdCorredor.value = p_CdCorredor;
            parentTR.insertBefore(cdCorredor, document.getElementById("trCupos"));

            var cdVendedor = document.createElement("input");
            cdVendedor.type = "hidden";
            cdVendedor.id = "cdVendedor_"+ maxIndex;
            cdVendedor.name = "cdVendedor_"+ maxIndex;
            cdVendedor.value = p_CdVendedor;
            parentTR.insertBefore(cdVendedor, document.getElementById("trCupos"));

            //incremento en uno el indice de las filas de la tabla
            document.getElementById("maxIndexVertical").value = parseInt(maxIndex) + 1;
        }
        else {
            //ya tiene cargado nominaciones el corredor/vendedor, actualizo la fila con los nuevos valores 
            var totalCupos = 0;
            for (var i = 0; i <= "<%= diasCupos%>"; i++){
                //obtengo los valores de cupos ingresados end cada TextBox
                var valorCupo = document.getElementById("cupo_" + i).value;
                if (parseInt(valorCupo) > 0){
                    if (!document.getElementById("cantidad_"+ i +"_" + indexLinea)){
                        document.getElementById("td_"+ i +"_" + indexLinea).innerHTML = valorCupo;
                        document.getElementById("td_"+ i +"_" + indexLinea).style.cursor = "pointer";
                        if (isFirefox)
                            document.getElementById("td_"+ i +"_" + indexLinea).setAttribute('onclick', "javascript:abrirNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+","+ document.getElementById("colFecha_"+i).value +","+i +"," + indexLinea+");");
                        else 
                            document.getElementById("td_"+ i +"_" + indexLinea)['onclick'] = new Function("javascript:abrirNominacion("+p_cuitCliente+","+p_CdCorredor+","+p_CdVendedor+","+ document.getElementById("colFecha_"+i).value +","+i +"," + indexLinea+");");
                        var auxCantidad = document.createElement("input")
                        auxCantidad.type = "hidden";
                        auxCantidad.id = "cantidad_"+ i +"_" + indexLinea;
                        auxCantidad.name = "cantidad_"+ i +"_" + indexLinea;
                        auxCantidad.value = valorCupo;
                        var cantidadCupo = valorCupo;
                        document.getElementById("td_"+ i +"_" + indexLinea).appendChild(auxCantidad);
                    }
                    else{
                        //calculo el nuevo valor de la nominacion que tendra la celda
                        var cantidadCupo = parseInt(document.getElementById("cantidad_"+ i +"_" + indexLinea).value) + parseInt(valorCupo);
                        // asigno el nuevo valor al input de la celda y guardo el objeto para limpiar el html de la celda
                        document.getElementById("cantidad_"+ i +"_" + indexLinea).value = cantidadCupo;
                        var obj = document.getElementById("cantidad_"+ i +"_" + indexLinea);
                        document.getElementById("td_"+ i +"_" + indexLinea).innerHTML = cantidadCupo;
                        document.getElementById("td_"+ i +"_" + indexLinea).appendChild(obj);
                    }
                    var totalCupos = parseInt(totalCupos) + parseInt(valorCupo);
                }
            }
            document.getElementById("totalFila_" + indexLinea).innerHTML =  parseInt(document.getElementById("vlTotal_" + indexLinea).value) + parseInt(totalCupos);
            document.getElementById("vlTotal_" + indexLinea).value = parseInt(document.getElementById("vlTotal_" + indexLinea).value) + parseInt(totalCupos);
        }

        //Finalmente actualizo y visualizo los nuevos totales por dia
        cargarNominacionesParciales();
        //Limpio la fila ingreso de datos
        limpiarIngresoCupos()    
        
    }

    //Obtiene el indice de columna en caso de coincidir el corredor/vendedor
    //retorna 0 en caso de no encontrar nada, retorna el numero de indice en caso de encontrar coincidencia
    function buscarClienteCorredorVendedor(p_cuitCliente, p_CdVendedor, p_CdCorredor, p_Indice){
        var indexLinea = 0;
        for (var i = 1; i < p_Indice; i++) {
            if (existeFilaNominacion(i)){
                var auxCuitCliente = document.getElementById("cuitCliente_" + i).value;
                var auxCdCorredor = document.getElementById("cdCorredor_" + i).value;
                var auxCdVendedor = document.getElementById("cdVendedor_" + i).value;
                if ((auxCuitCliente == p_cuitCliente)&&(parseInt(p_CdVendedor) == parseInt(auxCdVendedor))&&(parseInt(p_CdCorredor) == parseInt(auxCdCorredor))) indexLinea = i;
            }
        }
        return indexLinea;
    }
    
    //controla si existe la fila, si existe retorna true, caso contrario false
    function existeFilaNominacion(p_IndexFila){
        var ret = false;
        if (document.getElementById("trDetalle_"+p_IndexFila)) ret = true;
        return ret;
    }
    function limpiarIngresoCupos(){
        if(document.getElementById("dsCliente")) document.getElementById("dsCliente").value = "";
        if(document.getElementById("cdCliente")) document.getElementById("cdCliente").value = "";
        if(document.getElementById("cuitCliente")) document.getElementById("cuitCliente").value = "";
<%      if (not flagEsCorredor) then %>        
        if(document.getElementById("dsCorredor")) document.getElementById("dsCorredor").value = "";
        if(document.getElementById("cdCorredor")) document.getElementById("cdCorredor").value = "";
<%      end if %>        
        if(document.getElementById("divCuitCorredor")) document.getElementById("divCuitCorredor").style.display ="none";
        if(document.getElementById("cuitCorredor")){
            document.getElementById("cuitCorredor").style.display ="none";
            document.getElementById("cuitCorredor").value = "";
        } 
        if(document.getElementById("dsVendedor")) document.getElementById("dsVendedor").value = "";
        if(document.getElementById("cdVendedor")) document.getElementById("cdVendedor").value = "";
        if(document.getElementById("divCuitVendedor")) document.getElementById("divCuitVendedor").style.display ="none";
        if(document.getElementById("cuitVendedor")){
            document.getElementById("cuitVendedor").style.display ="none";
            document.getElementById("cuitVendedor").value = "";
        } 
        for (var i = 0; i <= "<%= diasCupos%>"; i++){
            if (document.getElementById("cupo_"+i)) document.getElementById("cupo_"+i).value = "";
        }
        if(document.getElementById("btnSaveNominacion")) document.getElementById("btnSaveNominacion").style.display = "block";
        if(document.getElementById("loadingNominacion")) document.getElementById("loadingNominacion").style.display = "none";
    
    }
    function abrirNominacion(p_cuitCliente, p_Corredor, p_Vendedor, p_Fecha, p_IndexCol, p_IndexRow){
        var puw = new winPopUp('popupNominaciones',"cuposNominacionesPopUp.asp?cuitCupeador=<% =cuitCupeador %>&cuitCliente="+p_cuitCliente+"&cdProducto=<%=cdProducto%>&pto=<%=g_strPuerto%>&fecha="+p_Fecha+"&idCorredor="+p_Corredor+"&idVendedor="+p_Vendedor + "&fc=<% =fc %>", 650, 550,'Nominaciones', "recalcularNominaciones("+p_IndexCol +","+ p_IndexRow+")");
    }
        
    //Esta variable setea la cantidad de cupos que fueron eliminados desde el pop up, de esta manera actualizamos la grilla con los 
    //nuevos valores de los cupos sin necesidad de recargar la pagina
    var cuposEliminadosPopUp = 0;
    
    //Esta funcion recalcula la celda de nominaciones modificada en el popUp, por lo que tambien afectara al total de filas y columnas.
    function recalcularNominaciones(p_IndexCol, p_IndexRow) {
        if (parseInt(cuposEliminadosPopUp) > 0){
            //Se eliminaron nominaciones
            var obj = document.getElementById("cantidad_"+ p_IndexCol +"_"+ p_IndexRow);
            var cupos = parseInt(obj.value) - parseInt(cuposEliminadosPopUp);
            //actualizo la columna de total por fila restandole el valor que tenia la celda
            document.getElementById("vlTotal_"+ p_IndexRow).value = parseInt(document.getElementById("vlTotal_"+ p_IndexRow).value) - parseInt(cuposEliminadosPopUp);
            document.getElementById("totalFila_"+ p_IndexRow).innerHTML = document.getElementById("vlTotal_"+ p_IndexRow).value;
            if (parseInt(cupos) > 0) {
                //El nuevo calculo sigue teniendo nominados, actualizo la celda
                document.getElementById("td_"+ p_IndexCol +"_" + p_IndexRow).innerHTML = cupos;
                document.getElementById("td_"+ p_IndexCol +"_" + p_IndexRow).appendChild(obj);
                obj.value = cupos;
            }
            else {
                //No se tiene mas nominados en la celda, se limpia
                document.getElementById("td_"+ p_IndexCol +"_" + p_IndexRow).innerHTML = "";
                $("#td_" + p_IndexCol + "_" + p_IndexRow).removeAttr("onclick");
                $("#td_" + p_IndexCol + "_" + p_IndexRow).removeAttr("style");
                //Si el total por fila es 0 nominaciones significa que la fila debe eliminarse
                if (parseInt(document.getElementById("vlTotal_"+ p_IndexRow).value) == 0) {
                    $("#trDetalle_" + p_IndexRow).remove();
                    $("#vlTotal_"+ p_IndexRow).remove();
                    $("#cdCorredor_"+ p_IndexRow).remove();
                    $("#cdVendedor_"+ p_IndexRow).remove();
                }
            }
            cuposEliminadosPopUp = 0;
            //Por ultimo actualizamos los totales por fechas (por columna)
            cargarNominacionesParciales();
        }
    }
    function controlarDatosGrabados(){
        var ret = false;        
        if ((document.getElementById("cdCorredor").value == "")&&(document.getElementById("dsCorredor").value == "")&&(document.getElementById("cuitCorredor").value == "")) {
            if ((document.getElementById("cdVendedor").value == "")&&(document.getElementById("dsVendedor").value == "")&&(document.getElementById("cuitVendedor").value == "")) ret = true;
        }
        return ret;
    }
    function carta() {
        window.open("cuposEmitirCarta.asp");        
    }
    function MostrarCalendario(p_objID, funcSel) {
		var dte= new Date();		    	    
		var elem= document.getElementById(p_objID);
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, SeleccionarCal, CerrarCal);
	    cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	function SeleccionarCal(cal, date) {
		var str= new String(date);
		document.getElementById("fdVisible").value = str;
        document.getElementById("fd").value = str.substring(6, 10).concat(str.substring(3, 5).concat(str.substring(0, 2)));
		if (cal) cal.hide();	
		cambiarProducto();
	}
	function CerrarCal(cal) {
		cal.hide();
	}
</script>
</head>
<body onload="bodyOnLoad()">
    <div class="tableaside size100">
	    <h3> ADMINISTRACI&Oacute;N DE CUPOS - <% =g_strPuerto %> </h3>
        <br>
        <form id="frmSel" name="frmSel" action="cuposAdministrar.asp<%=strParam %>" method="POST">
             <input type="hidden" id="cuitCupeador" name="cuitCupeador" value="<%=cuitCupeador %>"/>   
             <input type="hidden" id="pto" name="pto" value="<%=pto %>"/>   
        
                <div id="searchfilter" class="tableasidecontent">	
   	            <div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Producto:") %> </div>
	            <div class="col26"> 
		            <select name="cdProducto" id="cdProducto" onchange="cambiarProducto()">
			            <option value=""> <%=GF_Traducir("- Selecciones -")%></option>
				            <%  if (not rsProductos.eof) then
				                    'Primero listo los productos con cupo.
				                    if (not isNull(rsProductos("CANTIDAD"))) then    	%>
				                        <optgroup style="font-weight: bold" label="Con Cupo Asignado">
                            <%          prodSinCupo = false
                                        while ((not rsProductos.eof) and (not prodSinCupo))
                                            if (not isNull(rsProductos("CANTIDAD"))) then
					                            mySelected = ""					                                                    				                
					                            if trim(rsProductos("CDPRODUCTO")) = trim(cdProducto) then mySelected = "SELECTED"%>
					                            <option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>> 
					                                <% response.write rsProductos("DSPRODUCTO") 
					                                    if (not isNull(rsProductos("CANTIDAD"))) then response.write " (" & rsProductos("CANTIDAD") & ")" %>
					                             </option>
					                            <%rsProductos.MoveNext()
                                            else
                                                prodSinCupo = true
                                            end if					                            
				                        wend        %>
				                        </optgroup>
                            <%       end if				         
                                    'Si hay productos, listo los productos sin cupo.               
                                    if (not rsProductos.eof) then   %>
				                        <optgroup style="font-weight: bold" label="Otros Productos">
                            <%          while (not rsProductos.eof)                                            
					                            mySelected = ""					                                                    				                
					                            if trim(rsProductos("CDPRODUCTO")) = trim(cdProducto) then mySelected = "SELECTED"%>
					                            <option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>> 
					                                <% =rsProductos("DSPRODUCTO")  %>
					                             </option>
					                            <%rsProductos.MoveNext()
				                        wend       %>
				                        </optgroup>
				            <%      end if
				                 end if
				            %>
		            </select>
                </div>
	        </div>
	        <div id="searchfilter" class="tableasidecontent">	
	            <div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Desde:") %> </div>
                <div class="col26"> 
                    <input type="text" id="fdVisible" onclick="javascript:MostrarCalendario('fdVisible')" value="<% =GF_FN2DTE(fechaDesde) %>" />
                    <input type="hidden" id="fd" name="fd" value="<% =fechaDesde %>" />                    
                </div>                
            </div>
            <div id="searchfilter" class="tableasidecontent">
<%				if (puedeAgregar) then				%>
				    <div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Detallar Cupos de Terceros:") %> </div>
                    <div class="col16"> 
                        <input type="checkbox" id="chkDetallar" name="chkDetallar" onclick="javascript:cambiarProducto()" <% =chkDetallar %>/>
                    </div>
                <% else %>
                    <div class="col26 reg_header_navdos"> <% = GF_TRADUCIR("Ver cupos donde act&uacuteo como:") %> </div>
                    <div class="col26">                  
                        <input type="radio" name="fc" id="fcD" value="D" <% if (not flagEsCorredor) then response.write "checked" %> onclick="cambiarProducto()"> Destinatario
                        <input type="radio" name="fc" id="fcC" value="C" <% if (flagEsCorredor) then response.write "checked" %> onclick="cambiarProducto()"> Corredor
                    </div>
<%				end if								%>				
            </div>
<%      if (lckMsg) then %>	        
            <div>El acceso a la carga de cupos de este puerto se encuentra bloqueado por otro usuario.</div>
<%      end if %>
        </form>	        
    <% if (Cdbl(cdProducto) <> 0) then        
            if (not puedeAgregar) then            
                maxCuposDisponibles = 0
                'Si no es Toepfer entonces Cliente=Cupeador
                if (not flagEsCorredor) then
                    myWhere= myWhere & " where C.CUITCLIENTE = " &  cuitCupeador & " and C.ESTADO >= " & CUPO_OTORGADO
                else
	                myWhere= myWhere & "where  C.CUITCLIENTE = '" & CUIT_TOEPFER & "' and C.CDCORREDOR = " & session("KCOrganizacion") & " and C.ESTADO >= " & CUPO_PROVISORIO
	            end if
                myWhere = myWhere & "       and C.FECHACUPO <= " & fechaHasta &_
                                    "       and C.FECHACUPO >= " & fechaDesde &_                        
                                    "       and C.CDPRODUCTO = " & cdProducto
                strSQL= " Select FECHACUPO, count(*) CANTIDAD from CODIGOSCUPO C " & myWhere &_                        
                        "   group by FECHACUPO " &_
                        "   order by FECHACUPO"            
                Call executeQueryDb(g_strPuerto, rsCup, "OPEN", strSQL)                
                strSQL= " Select FECHACUPO, CUITCLIENTE, case when CDCORREDOR = " & SIN_CORREDOR & " then 0 else CDCORREDOR end CDCORREDOR, CDVENDEDOR, count(*) CANTIDAD from CODIGOSCUPO C " & myWhere &_                        
                        "   and CDVENDEDOR <> 0" &_                        
                        "   group by FECHACUPO, CUITCLIENTE, CDCORREDOR, CDVENDEDOR " &_
                        "   order by CUITCLIENTE, CDCORREDOR, CDVENDEDOR, FECHACUPO"                                    
                Call executeQueryDb(g_strPuerto, rsDet, "OPEN", strSQL)            
            else                    
                maxCuposDisponibles = getValueParametro(CUPOS_MAX_DISPONIBLES, g_strPuerto)                
                strSQL= " Select FECHACUPO, (" & maxCuposDisponibles & "-T.CANTIDAD) CANTIDAD from (" &_
                        " Select FECHACUPO, count(*) CANTIDAD from CODIGOSCUPO C " &_
                        "   where C.FECHACUPO <= " & fechaHasta &_
                        "       and C.FECHACUPO >= " & fechaDesde &_            
                        "       and C.CDPRODUCTO <> " & cdProducto &_
                        "       and ESTADO <> " & CUPO_CANCELADO &_
                        "   group by FECHACUPO) T " &_
                        "   order by T.FECHACUPO"                                    
                Call executeQueryDb(g_strPuerto, rsCup, "OPEN", strSQL)                
                strSQL= " Select FECHACUPO, CUITCLIENTE, CDCORREDOR, CDVENDEDOR, count(*) CANTIDAD from "
				if (chkDetallar <> "") then
					strSQL= strSQL & " (Select FECHACUPO, CUITCLIENTE, case when CDCORREDOR = " & SIN_CORREDOR & " then 0 else CDCORREDOR end CDCORREDOR, CDVENDEDOR from CODIGOSCUPO C "
				else	
					strSQL= strSQL & " (Select FECHACUPO, CUITCLIENTE, case when (CUITCLIENTE = '" & CUIT_TOEPFER & "') and (CDCORREDOR <> " & SIN_CORREDOR & ") then CDCORREDOR else 0 end CDCORREDOR,  case when (CUITCLIENTE = '" & CUIT_TOEPFER & "') then CDVENDEDOR else 0 end CDVENDEDOR  from CODIGOSCUPO C " 
				end if
                strSQL= strSQL & " where C.FECHACUPO <= " & fechaHasta &_
								"       and C.FECHACUPO >= " & fechaDesde &_          
								"       and ESTADO <> " & CUPO_CANCELADO &_
								"       and C.CDPRODUCTO = " & cdProducto & ") T " &_
								" group by FECHACUPO, CUITCLIENTE, CDCORREDOR, CDVENDEDOR " &_
								" order by CUITCLIENTE, CDCORREDOR, CDVENDEDOR, FECHACUPO"    
                Call executeQueryDb(g_strPuerto, rsDet, "OPEN", strSQL)
            end if
			if (puedeAgregar) then
     %>     	
				<input type="button" onclick="carta()" value="Emitir Carta Cupos" />       
	<%		end if	%>
            <table align="center" width="95%">
                <tr>
                    <td class="reg_Header_Warning" style="padding:4px;"> 
                        <img src="../images/info-16.png" alt="info-16" />
                        &nbsp; <% =GF_TRADUCIR("Para eliminar asiganciones de una fecha en particular debe seleccionar el t&iacute;tulo de la columna y aparecer&aacute; la opci&oacute;n eliminar") %>
                    </td>
                </tr>
                <tr><td style="height:5px;"></td></tr>
                <tr>
                    <td class="reg_Header_Warning" style="padding:4px;"> 
                        <img src="../images/info-16.png" alt="info-16" />
                        &nbsp; <% =GF_TRADUCIR("Para eliminar una asignaci&oacute;n de una fecha y corredor/vendedor en particular debe hacer click en la celda correspondiente") %>
                    </td>
                </tr>
            </table>
            <table class="datagrid" align="center" width="95%">
                <thead>
                    <tr>
                       <th style="width:15%;"></th>
                       <th style="width:15%;"></th>
                       <th style="width:15%;"></th>
                       <th style="width:7%;"></th>
                     <% i = 0
                        auxDesde = fechaDesde 
                        while (auxDesde < fechaHasta)
                          auxDesde = GF_DTEADD(fechaDesde, i, "D")                           %>
                          <th align="center" width="6%" id="th_<%=i %>" style="cursor:pointer;" onclick="javascript:seleccionarColumnaFecha(<%=i %>, <%=auxDesde %>)">
                              <%=getDayName(auxDesde) & "<br>" & LEFT(GF_FN2DTE(auxDesde), 5) %>
                              <input type="hidden" id="colFecha_<%=i %>" name="colFecha_<%=i %>"  value="<%=auxDesde %>">
                          </th>
                    <%     i = i + 1
                        wend %>
                        <th style="width:3%;" align="center">.</th>
                        <%  if (puedeAgregar) then %>
                        <th style="width:3%;" align="center">.</th>
                        <%  end if %>
                    </tr>
                </thead>
                <tbody>
                    <% Call dibujaLineaTitulo(rsCup, fechaDesde, maxCuposDisponibles) 
                       index = 1
                       while (not rsDet.Eof)
                            cantidad = 0
                            'Se -|uncion por que dentro de ella tira error
                            auxCuitCliente = rsDet("CUITCLIENTE")
                            auxIdCorredor = rsDet("CDCORREDOR")
                            auxIdVendedor = rsDet("CDVENDEDOR")
                            Call dibujaLineaCupos(rsDet, auxCuitCliente, auxIdCorredor, auxIdVendedor, index, cantidad) %>
                            <input id="vlTotal_<%=index%>" name="vlTotal_<%=index%>" value="<%=cantidad%>" type="hidden" />
                            <input id="cuitCliente_<%=index%>" name="cuitCliente_<%=index%>" value="<%=auxCuitCliente%>" type="hidden" />
                            <input id="cdCorredor_<%=index%>" name="cdCorredor_<%=index%>" value="<%=auxIdCorredor%>" type="hidden" />
                            <input id="cdVendedor_<%=index%>" name="cdVendedor_<%=index%>" value="<%=auxIdVendedor%>" type="hidden" />
                    <%      index = index + 1
                        wend
                         %>                        
                        <!-- ******************************* NUEVA CARGA ******************************* -->
                        <tr id="trCupos">
                            <td>
                            <%  if (puedeAgregar) then %>
                                <input id="dsCliente" name="dsCliente" type="text" onblur="comprobarDescripcion(this,document.getElementById('cdCliente'),document.getElementById('cuitCliente'))" style="width:98%;" onkeypress="return keyPressedEnter(event);"/>
                                <input type="hidden" id="cdCliente" name="cdCliente" />
                                <input type="hidden" id="cuitCliente" name="cuitCliente" />
                            <%  else									
									response.Write myDsCliente %>   
									<input type="hidden" id="cdCliente" name="cdCliente" />
									<input type="hidden" id="dsCliente" name="dCsliente" value="<% =myDsCliente %>" />
									<input type="hidden" id="cuitCliente" name="cuitCliente" value="<% =myCuitCliente %>"/>
                            <%  end if %>
                                
                            </td>
                            <td>
							<%	if (not flagEsCorredor) then	%>
                                <input id="dsCorredor" name="dsCorredor" type="text" onblur="comprobarDescripcion(this,document.getElementById('cdCorredor'),document.getElementById('cuitCorredor'))" style="width:98%;" onkeypress="return keyPressedEnter(event);"/>
                                <input type="hidden" id="cdCorredor" name="cdCorredor" />
                                <div id="divCuitCorredor" style="float:left;display:none;width:15%;margin-top:5px;">Cuit:</div>
                                <input type="text" id="cuitCorredor" name="cuitCorredor" maxlength="11" onkeypress="return controlNewkeyPressed(this,event,'N');" style="display:none;width:83%;margin-top:5px;"/>
							<%	else	
								myDsCorredor = Trim(getDsCorredor(session("KCOrganizacion")))
								response.write myDsCorredor
							%>
								<input type="hidden" id="dsCorredor" name="dsCorredor" value="<% =myDsCorredor %>"/>
                                <input type="hidden" id="cdCorredor" name="cdCorredor" value="<% =session("KCOrganizacion") %>" />								
                                <input type="hidden" id="cuitCorredor" name="cuitCorredor" value="<% =session("CuitOrganizacion")%>" />
							<% end if	%>
                            </td>
                            <td>
                                <input id="dsVendedor" name="dsVendedor" type="text" onblur="comprobarDescripcion(this,document.getElementById('cdVendedor'),document.getElementById('cuitVendedor'))" style="width:98%;" onkeypress="return keyPressedEnter(event);"/>
                                <input type="hidden" id="cdVendedor" name="cdVendedor" />
                                <div id="divCuitVendedor" style="float:left;display:none;width:15%;margin-top:5px;">Cuit:</div>
                                <input type="text" id="cuitVendedor" name="cuitVendedor" maxlength="11" onkeypress="return controlNewkeyPressed(this,event,'N');" style="display:none;width:83%;margin-top:5px;"/>
                            </td>
                            <td></td>
                            <% i = 0
                                while (i <= diasCupos) %>
                                    <td align="center">
                                        <input type="text" id="cupo_<%=i %>" name="cupo_<%=i %>" onblur="actualizarTotales(this,<%=i %>);" onkeypress="return controlNewkeyPressed(this,event,'N');" size="3"/>

                                    </td>
                            <%     i = i + 1
                                wend %>

                            <td colspan="2" align="center">
                                <input type="image" id="btnSaveNominacion" class="inputImgNominacion" onclick="javascript:controlarNominacion();" name="btnSaveNominacion" src="../images/save-16.png"/>
                                <img id="loadingNominacion" class="inputImgNominacion" src="../images/loading_small_green.gif" title="Guardando" alt="Guardando" style="display:none;"/>
                            </td>
                        </tr>
                        <thead>
                    <tr>
                       <th style="width:15%;"></th>
                       <th style="width:15%;"></th>
                       <th style="width:15%;"></th>
                       <th style="width:7%;"></th>
                     <% i = 0
                        auxDesde = fechaDesde 
                        while (auxDesde < fechaHasta)
                          auxDesde = GF_DTEADD(fechaDesde, i, "D")                           %>
                          <th align="center" width="6%" id="th1" style="cursor:pointer;" onclick="javascript:seleccionarColumnaFecha(<%=i %>, <%=auxDesde %>)">
                              <%=getDayName(auxDesde) & "<br>" & LEFT(GF_FN2DTE(auxDesde), 5) %>
                              <input type="hidden" id="Hidden1" name="colFecha_<%=i %>"  value="<%=auxDesde %>">
                          </th>
                    <%     i = i + 1
                        wend %>
                        <th style="width:3%;" align="center">.</th>
                        <%  if (puedeAgregar) then %>
                        <th style="width:3%;" align="center">.</th>
                        <%  end if %>
                    </tr>
                </thead>
                        <!-- ******************************* TOTALES DIARIOS ******************************* -->
                        <tr id="tr1">
                            <td colspan="3"></td>
                            <td align="center">TOTAL</td>
                            <% i = 0
                                while (i <= diasCupos) %>
                                    <td align="center" id="tdTotalProd_<% =i %>"></td>
                            <%     i = i + 1
                                wend %>
                            <td colspan="2"></td>
                        </tr>
                      <!-- ******************************* PENDIENTES  ******************************* -->
                        <tr id="trTotal">
                            <td colspan="3"></td>
                            <td align="center"><%=GF_TRADUCIR("Disponibles") %></td>
                            <% i = 0
                                while (i <= diasCupos) %>
                                    <td align="center" id="tdPendiente_<%=i %>"></td>
                            <%     i = i + 1
                                wend %>
                            <td colspan="2"></td>
                        </tr>
                        <!-- ******************************* ELIMINAR COLUMNA ******************************* -->
                        <tr style="display:none;" id="trEliminar">
                            <td colspan="3"></td>
                            <td align="center"><%=GF_TRADUCIR("Eliminar") %></td>
                            <% i = 0
                                while (i <= diasCupos) %>
                                    <td align="center" id="tdEliminar_<%=i %>"></td>
                            <%     i = i + 1
                                wend %>
                            <td colspan="2"></td>                            
                        </tr>
                        <!-- ******************************* MAIL COLUMNA ******************************* -->                                                                        
                        <!--
                        <tr style="display:none;" id="trMail">
                        <%  if (puedeAgregar) then %>
                            <td colspan="3"></td>
                            <td align="center"><%=GF_TRADUCIR("Enivar x Mail") %></td>
                            
                            <%  i = 0                                
                                while (i <= diasCupos) %>                                                          
                                    <td align="center" id="tdMail_<%=i %>"></td>
                        <%          i = i + 1
                                wend %>
                            <td colspan="2"></td>    
                            <%  end if %>                        
                        </tr>                                                                        
                        -->                        
                        <input type="hidden" id="indiceCol_Old" name="indiceCol_Old" value="-1"/>
                        <!-- ******************************* MSJ ERROR  ******************************* -->
                        <tr>
                            <td colspan="16" id="tdError" >
                                <div id="dsError" style="display:none;"></div>
                            </td>
                        </tr>
                        <input id="maxIndexVertical" name="maxIndexVertical" value="<%=index%>" type="hidden" />
                </tbody>
            </table>                     
<%          if (not puedeAgregar) then 
                myCuitCliente = cuitCupeador
                myCdCorredor = ""
                if (flagEsCorredor) then
                    myCuitCliente = 0
                    myCdCorredor = session("KCOrganizacion")
                end if
            
%>
            <div class="col26"></div>
            <span class="btnaction">
                <input type="button" value="Finalizar Nominaci�n y Obtener c&oacute;digos de cupo" onclick="javascript:verCupos(<%=myCuitCliente %>,'<% =myCdCorredor %>','','<%=fechaDesde %>','<%=fechaHasta %>')" ></input>
            </span>                   
<%          end if %>                 
        <% end if %>            
    </div>
    
</body>
</html>