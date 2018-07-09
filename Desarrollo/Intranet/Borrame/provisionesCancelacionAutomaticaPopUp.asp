<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->

<%
'--------------------------------------------------------------------------------------------------------------------------------
Dim nroLote,fechaLote, auxTotalAnuladoPesos, auxTotalPesos, indice, accion, secuencia, i, marcaInclusion, estado
Dim auxTotalAnuladoDolares, auxTotalDolares

nroLote   = GF_PARAMETROS7("nroLote",0,6)
fechaLote = GF_PARAMETROS7("fechaLote","",6)
indice    = GF_PARAMETROS7("indice",0,6)
auxTotalAnulado = 0
auxTotal = 0

Set sp_ret = executeSP(rsPro, "EJIFL.TBLPROVISIONESCANE_GET_BY_PARAMETERS", nroLote &"||"& fechaLote &"||||1||0$$totalRegistros")

%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>SISTEMA DE PROVISIONES - Detalle</title>
    <link rel="stylesheet" type="text/css" href="css/main.css" />	
    
	<script type="text/javascript" src="scripts/channel.js"></script>
    <script type="text/javascript" src="scripts/formato.js"></script>
    <script type="text/javascript" src="scripts/controles.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript">
	    var ch= new channel();
	    var arrDetalle = new Array();

	    function bodyOnload() {
	        var totalPesos = document.getElementById("totalPesos").value;
	        var totalDolares = document.getElementById("totalDolares").value;
	        var totalAcumuladoPesos = document.getElementById("totalAnuladoPesos").value;
	        var totalAcumuladoDolares = document.getElementById("totalAnuladoDolares").value;	        
	        document.getElementById("divTotalPesos").innerHTML = "<%=TIPO_MONEDA_PESO%> "+ editarNumero(String(totalPesos),2).replace(".", ",");
	        document.getElementById("divTotalAnuladoPesos").innerHTML = "<%=TIPO_MONEDA_PESO%> "+ editarNumero(String(totalAcumuladoPesos),2).replace(".", ",");
            document.getElementById("divTotalDolares").innerHTML = "<%=TIPO_MONEDA_DOLAR%> "+ editarNumero(String(totalDolares),2).replace(".", ",");
	        document.getElementById("divTotalAnuladoDolares").innerHTML = "<%=TIPO_MONEDA_DOLAR%> "+ editarNumero(String(totalAcumuladoDolares),2).replace(".", ",");
	        
	    }
	    function asignarMarcaInclusion(p_Index){
	        var totalAnuladoPesos = document.getElementById("totalAnuladoPesos").value;
	        var totalAnuladoDolares = document.getElementById("totalAnuladoDolares").value;
	        var importeCancelacionPesos = document.getElementById("importeCancelacionPesos_"+p_Index).value;
	        var importeCancelacionDolares = document.getElementById("importeCancelacionDolares_"+p_Index).value;
	        var totalPesos = 0;
	        var totalDolares = 0;
	        if (document.getElementById("chkMovimiento_"+p_Index).checked){
	            document.getElementById("marcaInclusion_"+p_Index).value = "S";
	            var totalPesos = parseFloat(totalAnuladoPesos) + parseFloat(importeCancelacionPesos);
	            var totalDolares = parseFloat(totalAnuladoDolares) + parseFloat(importeCancelacionDolares);
	        }
	        else{
	            document.getElementById("marcaInclusion_"+p_Index).value = "N";
	            var totalPesos = parseFloat(totalAnuladoPesos) - parseFloat(importeCancelacionPesos);
	            var totalDolares = parseFloat(totalAnuladoDolares) - parseFloat(importeCancelacionDolares);
	        }
	        document.getElementById("totalAnuladoPesos").value = editarNumero(String(totalPesos),2);
	        document.getElementById("totalAnuladoDolares").value = editarNumero(String(totalDolares),2);
	        document.getElementById("divTotalAnuladoPesos").innerHTML = "<%=TIPO_MONEDA_PESO%> "+  editarNumero(String(totalPesos),2).replace(".", ",");
	        document.getElementById("divTotalAnuladoDolares").innerHTML = "<%=TIPO_MONEDA_DOLAR%> "+  editarNumero(String(totalDolares),2).replace(".", ",");
	    }
	    function grabarDetalle(){
	        var count = 0;
	        var strParameter = "";
	        //Tomo los valores que necesito para enviar por ajax (recordar que los valores poseen indices)
            //Para eso se controla que la secuencia del lote se halla cambiado la marca de inclusion
	        var indice = document.getElementById("indice").value;
	        for (var i = 0; i < indice; i++) {
	            var secuencia = document.getElementById("secuencia_"+ i).value;
	            if (parseInt(secuencia) != 0) {
	                var marcaInclusion = document.getElementById("marcaInclusion_"+ i).value;
	                var marcaInclusionOld = document.getElementById("marcaInclusionOld_"+ i).value;
	                if (marcaInclusion != marcaInclusionOld){
	                    strParameter = strParameter + "&marcaInclusion_"+ count +"="+marcaInclusion+"&secuencia_"+ count +"="+secuencia;
	                    arrDetalle.push(i)
	                    count++;
	                }
	            }
	        }
	        if (strParameter != "") {
	            document.getElementById("actionLabel").style.visibility = 'visible';
	            document.getElementById("actionLabel").style.textAlign = 'center';
	            document.getElementById("actionLabel").style.fontSize = "16";
	            document.getElementById("actionLabel").innerHTML = "Guardando marca de inclusión...";
	            var nroLote = document.getElementById("nroLote").value;
	            var fechaLote = document.getElementById("fechaLote").value;
	            var estado = document.getElementById("estado").value;
	            ch.bind("provisionesCancelacionAutomaticaAjax.asp?accion=<%=ACCION_GRABAR%>&nroLote="+ nroLote +"&fechaLote="+ fechaLote +"&estado="+ estado +"&indice="+ count + strParameter, "grabarDetalle_Callback()");
	            ch.send();
	        }
	    }
	    function grabarDetalle_Callback(){
	        for (var i = 0; i < arrDetalle.length; i++) {
	            document.getElementById("marcaInclusionOld_"+ arrDetalle[i]).value = document.getElementById("marcaInclusion_"+ arrDetalle[i]).value;
	        }
	        document.getElementById("actionLabel").style.visibility = 'hidden';
	        document.getElementById("actionLabel").innerHTML = "";
	    }
	</script>
</head>
<BODY onload="bodyOnload()">
	<div class="col66"></div>
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Nro.Lote")%></div>
        <div class="col26"> <%= nroLote %>  </div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Fecha Lote")%></div>
        <div class="col26"> <%= GF_FN2DTE(fechaLote) %>  </div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Total pesos")%></div>
        <div class="col26"><div id="divTotalPesos"></div></div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Total incluidos pesos")%></div>
        <div class="col26"><div id="divTotalAnuladoPesos"></div></div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Total dolares")%></div>
        <div class="col26"><div id="divTotalDolares"></div></div>
        <div class="col26 reg_header_navdos"> <%=GF_Traducir("Total incluidos dolares")%></div>
        <div class="col26"><div id="divTotalAnuladoDolares"></div></div>
    </div>

    <div class="col66"></div>
  <% if (not rsPro.Eof) then
         estado = Trim(rsPro("ESTADO")) %>
         <form id="frmSel" name="frmSel" action="provisionesCancelacionAutomaticaPopUp.asp" method="post">
            <table class="datagrid" width="98%" align="center">
                <thead>
                    <tr>
                        <th align="center" rowspan="2" width="7%"><%=GF_Traducir("Buque")%></th>
                        <th align="center" rowspan="2" width="7%"><%=GF_Traducir("Nominaci&oacuten")%></th>
                        <th align="center" rowspan="2" width="15%"><%=GF_Traducir("Concepto")%></th>
				        <th align="center" colspan="3" width="33%"><%=GF_Traducir("Dolares")%></th>
                        <th align="center" colspan="3" width="33%"><%=GF_Traducir("Pesos")%></th>
				        <th align="center" rowspan="2" width="5%" ><%=GF_Traducir("Inclusi&oacuten")%></th>
                    </tr>
                    <tr>
                        <td align="center" width="11%"><%=GF_Traducir("Provisi&oacuten")%></td>
                        <td align="center" width="11%"><%=GF_Traducir("Gasto")%></td>
                        <td align="center" width="11%"><%=GF_Traducir("Cancelaci&oacuten")%></td>
                        <td align="center" width="11%"><%=GF_Traducir("Provisi&oacuten")%></td>
                        <td align="center" width="11%"><%=GF_Traducir("Gasto")%></td>
                        <td align="center" width="11%"><%=GF_Traducir("Cancelaci&oacuten")%></td>
                    </tr>
                </thead>
                <tbody>
              <% indice = 0
                 while (not rsPro.Eof) 
                       auxTotalPesos = auxTotalPesos + Cdbl(rsPro("IMPORTEPESOS"))
                       auxTotalDolares = auxTotalDolares + Cdbl(rsPro("IMPORTEDOLAR"))
                       if (Cstr(rsPro("MARCAINCLUSION")) = "S") then 
                            auxTotalAnuladoPesos = auxTotalAnuladoPesos + Cdbl(rsPro("IMPORTEPESOS")) 
                            auxTotalAnuladoDolares = auxTotalAnuladoDolares + Cdbl(rsPro("IMPORTEDOLAR")) 
                        end if
                    %>
                       <tr>
                            <td align="center"><%= rsPro("BUQUE") %></td>
                            <td align="center"><%= rsPro("NOMINACION") %></td>
                            <td align="left"><%= rsPro("CONCEPTO") &"-"& Trim(rsPro("MGDES")) %></td>
                            <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(rsPro("PROVISIONDOLARES"))*100,2) %></td>
                            <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(rsPro("GASTODOLARES"))*100,2) %></td>
                            <td align="right">
                                <%= GF_EDIT_DECIMALS(Cdbl(rsPro("IMPORTEDOLAR"))*100,2) %>
                                <input type="hidden" name="importeCancelacionDolares_<%= indice %>" id="importeCancelacionDolares_<%= indice %>" value="<%=rsPro("IMPORTEDOLAR") %>"  />
                            </td>
                            <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(rsPro("PROVISIONPESOS"))*100,2) %></td>
                            <td align="right"><%= GF_EDIT_DECIMALS(Cdbl(rsPro("GASTOPESOS"))*100,2) %></td>
                            <td align="right">
                                <%= GF_EDIT_DECIMALS(Cdbl(rsPro("IMPORTEPESOS"))*100,2) %>
                                <input type="hidden" name="importeCancelacionPesos_<%= indice %>" id="importeCancelacionPesos_<%= indice %>" value="<%=rsPro("IMPORTEPESOS") %>"  />
                            </td>
                            <td align="center">
                                <% if (CStr(estado) <> PROVISCIONES_ESTADO_APLICADO) then %>
                                    <input type="checkbox" id="chkMovimiento_<%= indice %>" name="chkMovimiento_<%= indice %>" title="Aplicar inclusión" <% if (Cstr(rsPro("MARCAINCLUSION")) = "S") then %> checked <% end if %> onclick="asignarMarcaInclusion(<%= indice %>)"/>
                                    <input type="hidden" id="marcaInclusion_<%= indice %>" name="marcaInclusion_<%= indice %>" value="<%= Cstr(rsPro("MARCAINCLUSION")) %>" />
                                    <input type="hidden" id="marcaInclusionOld_<%= indice %>" name="marcaInclusionOld_<%= indice %>" value="<%= Cstr(rsPro("MARCAINCLUSION")) %>" />
                                    <input type="hidden" id="secuencia_<%= indice %>" name="secuencia_<%= indice %>" value="<%=rsPro("SECUENCIA") %>" />
                                <% else %>
                                    <%= rsPro("MARCAINCLUSION") %>
                                <% end if %>
                            </td>
                       </tr>
                <%     indice = indice + 1
                       rsPro.MoveNext()
                 wend %>
                </tbody>
            </table>
            <input type="hidden" id="totalPesos" name="totalPesos"  value="<%=auxTotalPesos %>"  />
            <input type="hidden" id="totalDolares" name="totalDolares"  value="<%=auxTotalDolares %>"  />
            <input type="hidden" id="totalAnuladoPesos" name="totalAnuladoPesos"  value="<%=auxTotalAnuladoPesos %>"  />
             <input type="hidden" id="totalAnuladoDolares" name="totalAnuladoDolares"  value="<%=auxTotalAnuladoDolares %>"  />
            <input type="hidden" id="indice" name="indice"  value="<%=indice %>" />
            <input type="hidden" id="nroLote" name="nroLote"  value="<%=nroLote %>" />
            <input type="hidden" id="fechaLote" name="fechaLote"  value="<%=fechaLote %>" />
            <input type="hidden" id="estado" name="estado"  value="<%=estado %>" /> 
            <br />
            <% if (CStr(estado) <> PROVISCIONES_ESTADO_APLICADO) then %>
             <span class="btnaction">
                <input type="button" value="Guardar" id="btnGuardar" onclick="grabarDetalle()"/>
            </span>
             <% end if %>
            <div class="col66">&nbsp</div>
            <div id="actionLabel" class="confirmsj" style="width:100%;visibility:hidden;margin-top:10px;"></div>
        </form>
    <% end if  %>
</BODY>
</html>