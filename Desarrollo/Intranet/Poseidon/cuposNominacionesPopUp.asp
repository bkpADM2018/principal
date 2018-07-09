<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientosAS400.asp"-->
<%

'********************************************************************************************************************************
'********************************************************* INICIO DE PAGINA *****************************************************
'********************************************************************************************************************************
Dim cuitCupeador, cdProducto, nroPuerto, fecha, idCorredor , idVendedor, rsNom, g_strPuerto,fechaPermitida, totalNominados, cuitCliente
Dim strSQL, rs, fc

cuitCupeador = GF_PARAMETROS7("cuitCupeador",0,6)
cuitCliente = GF_PARAMETROS7("cuitCliente",0,6)
cdProducto  = GF_PARAMETROS7("cdProducto",0,6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
fecha       = GF_PARAMETROS7("fecha","",6)
idCorredor  = GF_PARAMETROS7("idCorredor","",6)
idVendedor  = GF_PARAMETROS7("idVendedor","",6)
'Esta fecha permitida valida que solo se puedan eliminar nominaciones de ma�ana en adelante
fechaPermitida = GF_DTEADD(Left(Session("MmtoDato"),8), 1, "D")
fc = GF_PARAMETROS7("fc","",6)

'Obtengo las nominaciones
strSQL ="Select C.*, E.DSESTADO ESTADO_CAMION,  DSCORREDOR, DSVENDEDOR from CODIGOSCUPO C " &_
        " left join CAMIONES D on D.NUCUPO=C.CODIGOCUPO" &_
        " left join ESTADOS E on D.CDESTADO=E.CDESTADO" &_
		" left join CORREDORES COR on COR.CDCORREDOR=C.CDCORREDOR " &_
		" left join VENDEDORES VEN on VEN.CDVENDEDOR=C.CDVENDEDOR " &_
        " where FECHACUPO=" & fecha &_
        "       and CUITCLIENTE='" & cuitCliente & "'" &_
        "       and C.ESTADO <> " & CUPO_CANCELADO 
if (CDbl(idCorredor) <> 0) then strSQL = strSQL & " and C.CDCORREDOR = " & idCorredor 
if (CDbl(idVendedor) <> 0) then	strSQL = strSQL & " and C.CDVENDEDOR = " & idVendedor 
strSQL = strSQL & "       and C.CDPRODUCTO=" & cdProducto        		
strSQL = strSQL & " order by DSCORREDOR, DSVENDEDOR, C.CODIGOCUPO"
Call executeQueryDb(g_strPuerto, rsNom, "OPEN", strSQL)
%>
<html>
<head>
<title>Sistema de Cupos</title>
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<style type="text/css">
    .tdTotalNominacion {
        font-weight:bold;
        font-size:11px;
        background-color:#2e6b4d;
        color:#fff;
    }
</style>
    
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript">
    var ch = new channel();
    var cuposEliminados = 0;

    function BodyOnLoad() {
        var tb = new Toolbar('toolbar', 3, "..images/almacenes/");
        <% if (fecha >= fechaPermitida) then %>
        tb.addButton("toolbar-cancel", "Cancelar Todos", "eliminarNominacion('',-1)");
        <%  end if %>
        tb.draw();
    }
    
    function eliminarNominacion(p_CdCupo, p_Index) {
        if (confirm("Desea eliminar la nominacion?")) {
            var strParameter = "accion=<%=ACCION_BORRAR %>&cuitCupeador=<% =cuitCupeador %>&cuitCliente=<% =cuitCliente %>&fechaDesde=<%=fecha%>&fechaHasta=<%=fecha%>&cdVendedor=<%=idVendedor%>&cdCorredor=<%=idCorredor%>&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&cdCupo=" + p_CdCupo + "&fc=<% =fc %>";
            ch.bind("cuposAdministrarAjax.asp?" + strParameter, "eliminarNominacion_Callback(" + p_Index + ")");
            ch.send();
        }
    }
    function eliminarNominacion_Callback(p_Index) {
        if (p_Index >= 0) {
            eliminarRegistroTabla(p_Index);
        } else {
            var idxs = [];
            var tabla = document.getElementsByTagName("tbody")[0];
            for (var i = 0, row; row = tabla.rows[i]; i++) {
                var res = row.id.split("_");                
                idxs.push(res[1]); 
            }
            for (var i = 0, idr; idr = idxs[i]; i++) {
                eliminarRegistroTabla(idr);
            }
        }
    }
    function eliminarRegistroTabla(p_Index) {
        $("#tr_" + p_Index).remove();
        cuposEliminados++;
        parent.window.cuposEliminadosPopUp = cuposEliminados;
        var totalNominados = document.getElementById("totalNominados").value;        
        var total = parseInt(totalNominados) - parseInt(cuposEliminados);
        if (total == 0) {
            var tabla = document.getElementsByTagName("tbody")[0];
            var tr = document.createElement("tr");
            var td = document.createElement("td");
            td.colSpan = "3";
            td.align = "center";
            td.innerHTML = "No se encontraron nominaciones";
            tr.appendChild(td);
            tabla.appendChild(tr);
            document.getElementById("trTotal").style.display = "none";
            document.getElementById("delAll").style.display = "none";
        }
        else {            
            document.getElementById("tdTotal").innerHTML = "Total: " + total;
        }
    }

    function agregar(pIdx) {
        var fgo = true;
        while ((pIdx <= 5) && (fgo)) {     
            var cant = document.getElementById("txtCantidad_" + pIdx).value;
            var cantH = document.getElementById("txtCantidadH_" + pIdx).value;            
            if (cant != cantH) {
                document.getElementById("tdError").className = "";
                document.getElementById("dsError").style.display = "none";
                var cond = document.getElementById("txtCondicion_" + pIdx).value;
                var strParameter = "accion=<%=ACCION_GRABAR %>&especial=1&cuitCupeador=<% =cuitCupeador %>&cuitCliente=<% =cuitCliente %>&fechaDesde=<%=fecha%>&fechaHasta=<%=fecha%>&cdVendedor=<%=idVendedor%>&cdCorredor=<%=idCorredor%>&pto=<%=g_strPuerto%>&cdProducto=<% =cdProducto %>&cond=" + cond + "&cant=" + cant + "&fc=<% =fc %>";                
                ch.bind("cuposAdministrarAjax.asp?" + strParameter, "agregar_Callback(" + pIdx + ")");
                ch.send();
                fgo = false;
            } else {
                pIdx++;
            }
        }                                
    }   
    
    //Esta funcion quita los espacios en blanco de un string
    //NOTA: no se utiliza la funcion trim() de javascript debido a que en Internet Explorer 8 no es compatible
    function trimStr(str) {
        return str.replace(/^\s+|\s+$/g, '');
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
    
    function agregar_Callback(pIdx) {
        var resp = ch.response();
        if (resp != '') {
            manejoErrorNominacion(resp);
        } else {
            document.getElementById("txtCantidadH_" + pIdx).value=document.getElementById("txtCantidad_" + pIdx).value;
            pIdx++;
            if (pIdx <= 5) agregar(pIdx);
        }
        
                
    } 
</script>
</head>
<body onload="BodyOnLoad()" >
    <div id="toolbar"></div>
    <div class="tableasidecontent">
    
        <div class="col26 reg_header_navdos"> Fecha </div>
        <div class="col46"><% =GF_FN2DTE(fecha) %> </div>
        
        <div class="col26 reg_header_navdos"> Producto </div>
        <div class="col46"><%= cdProducto &"-"& Trim(getDsProducto(cdProducto)) %></div>
                
        <div class="col26 reg_header_navdos"> Destinatario </div>
        <div class="col46"><% =getDsClienteByCUIT(cuitCliente) %></div>
        
    </div>       
     <div class="col66" />
    <table class="datagrid" align="center" width="95%">
        <thead>
            <tr>
                <th align="center"><%=GF_TRADUCIR("Cupo") %></th>				
				<th style="width:30%;" align="center"><%=GF_TRADUCIR("Corredor") %></th>
				<th style="width:30%;" align="center"><%=GF_TRADUCIR("Vendedor") %></th>				
                <th style="width:30%;" align="center"><%=GF_TRADUCIR("Estado") %></th>
                <th style="width:16px"  align="center">.</th>
            </tr>
        </thead>

        <tbody>        
      <%    index = 0
            while (not rsNom.Eof) 
                'Determino que estado tendra el cupo
                auxEstado = "PROVISORIO"
                flagCupoRecibido = false
                if (CInt(rsNom("ESTADO")) = CUPO_OTORGADO) then auxEstado = "OTORGADO"                
                if (CInt(rsNom("ESTADO")) = CUPO_NOMINADO) then auxEstado = "NOMINADO"
                if (CInt(rsNom("ESTADO")) = CUPO_PUBLICADO_AFIP) then auxEstado = "DISPONIBLE SIN CTG"
                if (CInt(rsNom("ESTADO")) > CUPO_PUBLICADO_AFIP) then 
                    flagCupoRecibido = true                    
                    if (CInt(rsNom("ESTADO")) = CUPO_ACTIVADO_AFIP) then auxEstado = "CUPO ACTIVADO"
                    'Una vez que aparece en nuestro sistema se informa el estado dentro de planta.
                    if (not isNull(rsNom("ESTADO_CAMION"))) then
                        auxEstado = UCase(rsNom("ESTADO_CAMION"))
                    end if
                    if (CInt(rsNom("ESTADO")) = CUPO_DESCARGADO_AFIP) then auxEstado = "CUPO DESCARGADO"
                end if                    
%>
                <tr id="tr_<%=index %>">
                    <td align="center"><% if (CInt(rsNom("ESTADO")) = CUPO_PROVISORIO) then response.write Left(rsNom("CODIGOCUPO"), 6) & "????" else response.write rsNom("CODIGOCUPO") end if %></td>					
					<td align="center"><%= rsNom("DSCORREDOR") %></td>
					<td align="center"><%= rsNom("DSVENDEDOR") %></td>					
                    <td align="center"><%= auxEstado %></td>
                    <td align="center">
                    <% if ((fecha >= fechaPermitida)and(not flagCupoRecibido)) then %>
                        <img src="../images/compras/cancel-16x16.png" style="cursor:pointer;" title="Eliminar" onclick="eliminarNominacion('<%=rsNom("CODIGOCUPO") %>',<%=index %>)"/>
                    <% end if %>
                    </td>
                </tr>
          <%    index = index + 1
                rsNom.MoveNext() %>
          <% wend %>
            </tbody>
            <tfoot>
            	<tr id="trTotal">
					<td align="center" colspan="5" id="tdTotal"><%=GF_TRADUCIR("Total: ") & index %></td>					
				</tr>
			</tfoot>
    </table>
<%  if ((CDbl(cuitCupeador) = CDbl(CUIT_TOEPFER)) and (CDbl(cuitCliente) = CDbl(CUIT_TOEPFER)) and ((CDbl(idCorredor) <> 0) or (CDbl(idVendedor) <> 0))) then   %>
    <br />
    <h1>CUPOS ESPECIALES</h1>
    <table class="datagrid" align="center" width="95%">
        <thead>
            <tr>
                <th style="width:45%;" align="center"><%=GF_TRADUCIR("Condici&oacute;n") %></th>
                <th style="width:50%;" align="center"><%=GF_TRADUCIR("Asignados") %></th>
                <th style="width:50%;" align="center"><%=GF_TRADUCIR("Ingresados") %></th>                
            </tr>            
        </thead>        
        <tbody> 
<%      strSQL="Select * from CODIGOSCUPOESPECIALES where CUITCLIENTE='" & cuitCliente & "' " &_
                    " and FECHACUPO=" & fecha &_
                    " and CDPRODUCTO=" & cdProducto  &_
					" and CDCORREDOR=" & defineCdCorredor(cuitCliente, idCorredor) &_
                    " and CDVENDEDOR=" & idVendedor        
        Call executeQueryDb(g_strPuerto, rs, "OPEN", strSQL)        
        idx = 0
        while (idx <= 5)      
            idx = idx + 1
            myCondicion= ""
            myCantidad = 0
            myIngresados = 0   
            if (not rs.eof) then
                myCondicion= rs("CONDICION")
                myCantidad = rs("QTASIGNADOS")
                myIngresados = rs("QTINGRESADOS")
                rs.MoveNext()
            end if
%>
             <tr>
                <td>
<%                  if (myCondicion="") then %>                    
                    <input type="text" name="txtCondicion_<% =idx %>" id="txtCondicion_<% =idx %>" maxlength="50" size="20"/></td>
<%                  else %>                                    
                    <% =myCondicion %>
                    <input type="hidden" name="txtCondicion_<% =idx %>" id="txtCondicion_<% =idx %>" maxlength="50" size="20" value="<% =myCondicion %>"/></td>
<%                  end if %>
                <td>
                    <input type="text" name="txtCantidad_<% =idx %>" id="txtCantidad_<% =idx %>" size="10" value="<% =myCantidad %>"/>
                    <input type="hidden" id="txtCantidadH_<% =idx %>" value="<% =myCantidad %>"/>
                </td>
                <td><% =myIngresados %></td>                
             </tr>      
<%      wend %>             
             <tr>
                <td colspan="3" align="center">
                    <input type="button" onclick="javascript:agregar(1);" value="Grabar"/>
                </td>
             </tr>
             <tr>
                <td colspan="16" id="tdError" >
                    <div id="dsError" style="display:none;"></div>
                </td>
            </tr>
        </tbody>
    </table>
<%  end if %>    
    <input type="hidden" id="totalNominados" value="<% =index %>" />
</body>
</html>