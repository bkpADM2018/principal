<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosFechas.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosTitulos.asp"-->
<!--#include file="../includes/procedimientosSQL.asp"-->
<!--#include file="../includes/procedimientosLog.asp"-->
<%
Function getMermaVolatil(p_fecha ,p_cdProducto, p_cdCliente, p_cdSilo,p_Pto)
    Dim strSQL
    strSQL = "SELECT  (YEAR(MV.DTCONTABLE)*10000 + Month(MV.DTCONTABLE)*100 + DAY(MV.DTCONTABLE)) AS FECHA, "&_
			 "	       MV.CDPRODUCTO, "&_
			 "	       PRO.DSPRODUCTO, "&_
			 "	       MV.CDCLIENTE, "&_
			 "	       CLI.DSCLIENTE, "&_
			 "	       MV.CDSILO, "&_
			 "	       MV.RATIO, "&_
			 "	       MV.CDUSUARIO, "&_
			 "	       MV.MMTO "&_
		     "FROM DBO.TBLREGLASMERMAVOLATIL MV "&_
			 "   LEFT JOIN PRODUCTOS PRO ON MV.CDPRODUCTO = PRO.CDPRODUCTO "&_
			 "   LEFT JOIN CLIENTES CLI ON MV.CDCLIENTE = CLI.CDCLIENTE "&_
		     "WHERE 1 = 1 "
	    	'Si viene una fecha asignada significa que es una modificacion, no se consulta los demas campos por que una merma vlatil no necesariamente tiene producto o cliente o silo	 
            if (p_fecha <> "") then
                'Modificacion 
                strSQL = strSQL & " AND MV.DTCONTABLE = '"& GF_FN2DTCONTABLE(p_fecha) &"' "&_
                                  " AND MV.CDPRODUCTO = "& p_cdProducto &_
                                  " AND MV.CDCLIENTE = "& p_cdCliente &_
	                              " AND MV.CDSILO = '"& p_cdSilo &"'"
             else
                'Nueva merma
                strSQL = strSQL & " AND DTCONTABLE = (SELECT MAX(DTCONTABLE) FROM TBLREGLASMERMAVOLATIL ) "
             end if
             
             strSQL = strSQL & " ORDER BY DTCONTABLE DESC,CDPRODUCTO,CDCLIENTE,CDSILO "
    Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSQL) 
    
    Set getMermaVolatil = rs
End Function
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
'----------------------------------------------------------------------------------------------------------------------
Function agregarMermaVolatil(pPto,pFecha,pCdProducto,pCdCliente,pCdSilo,pRatio)
    Dim strSQL
    strSQL = "INSERT INTO DBO.TBLREGLASMERMAVOLATIL (DTCONTABLE,CDPRODUCTO,CDCLIENTE,CDSILO,RATIO,CDUSUARIO,MMTO) "&_
             " VALUES('"& GF_FN2DTCONTABLE(pFecha) &"',"& pCdProducto &","& pCdCliente &",'"& pCdSilo &"',"& pRatio &",'"& Session("Usuario") &"',"& Session("MmtoDato") &")"
    Call GF_BD_Puertos(pPto, rs, "EXEC", strSQL)
    Call logMig.info("NUEVA MERMA VOLATIL:")
    Call logMig.info("--> FECHA: "& GF_FN2DTE(pFecha))
    Call logMig.info("--> PRODUCTO: "& getDsProducto(pCdProducto))
    Call logMig.info("--> CLIENTE: "& getDsCliente(pCdCliente))
    Call logMig.info("--> SILO: "& pCdSilo)
    Call logMig.info("--> RATIO: "& pRatio)
End Function
'----------------------------------------------------------------------------------------------------------------------
Function inicializarLogMermaVolatil(p_pto)
    Set logMig = new classLog
    Call startLog(HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
    logMig.fileName = "MERMA_VOLATIL_"& p_pto & "_"& left(session("MmtoDato"),8)
    Call logMig.info("-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* INICIA TRANSACCION -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*")
End Function 
'---------------------------------------------------------------------------------------------------------------------- 
Function eliminarMermaVolatilPorFecha(pPto,pFecha)
    Dim strSQL
    strSQL = "DELETE FROM DBO.TBLREGLASMERMAVOLATIL "&_
             "WHERE DTCONTABLE = '"& GF_FN2DTCONTABLE(pFecha) &"'"
    Call GF_BD_Puertos(pPto, rs, "EXEC", strSQL) 
    Call logMig.info("ELIMINO MERMA VOLATIL:")
    Call logMig.info("--> FECHA: "& GF_FN2DTE(pFecha))
End Function
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto, params,lineasTotales, mostrar,paginaActual,g_cdProducto,accion,logMig


accion = GF_Parametros7("accion","",6)
call addParam("accion", accion, params)
g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)
g_cdProducto = GF_PARAMETROS7("cdProducto",0,6)
call addParam("cdProducto", g_cdProducto, params)
g_fecha = GF_PARAMETROS7("fecha", "", 6)
call addParam("fecha", g_fechaD, params)
g_cdCliente = GF_PARAMETROS7("cdCliente", 0, 6)
call addParam("cdCliente", g_cdCliente, params)
g_cdSilo = GF_PARAMETROS7("cdSilo", "", 6)
call addParam("cdSilo", g_cdSilo, params)



if (accion = ACCION_GRABAR) then
    index = GF_PARAMETROS7("index", 0, 6)
    if (g_fecha <> "") then
        Call inicializarLogMermaVolatil(g_strPuerto)
        Call eliminarMermaVolatilPorFecha(g_strPuerto,g_fecha)
        for i = 0 to index - 1
            cdProducto = GF_PARAMETROS7("cdProducto_"& i, 0, 6)
            cdCliente = GF_PARAMETROS7("cdCliente_"& i, 0, 6)
            cdSilo = GF_PARAMETROS7("cdSilo_"& i, "", 6)
            ratio = GF_PARAMETROS7("ratio_"& i, 2, 6)
            estado = GF_PARAMETROS7("estado_"& i, 0, 6)
            if (Cint(estado) = ESTADO_ACTIVO) then Call agregarMermaVolatil(g_strPuerto,g_fecha,cdProducto,cdCliente,Trim(cdSilo),ratio)
        next
        flagGrabar = true 
    else
        Call setError(FECHA_INICIO_INCORRECTA)
    end if
end if

flagEdit = false
if (g_fecha <> "") then flagEdit = true
Set rs = getMermaVolatil(g_fecha ,g_cdProducto, g_cdCliente, g_cdSilo, g_strPuerto)


%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Merma Volatil </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../css/main.css" type="text/css">	
    <link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">	
	<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
	<style type="text/css">
		.reg_header_total {			
			BACKGROUND-COLOR: #BDBDBD;			
			FONT-FAMILY: verdana, arial, san-serif;			
		}	
	</style>

<script type="text/javascript" src="../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script language="javascript">
	var ch = new channel();
	
	<% if (flagGrabar) then %>
        parent.window.submitInfo('<%=ACCION_SUBMITIR%>');
	<% end if %>
	function onLoadPage(){
	    
	}	
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();
	}
	
	function CerrarCal(cal) {
	    cal.hide();
	}		
	function MostrarCalendario(funcSel, pObj) {
	    var dte= new Date();		    	    
	    var elem= document.getElementById(pObj);
	    if (calendar != null) calendar.hide();		
	    var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    cal.weekNumbers = false;
	    cal.setRange(1993, 2045);
	    cal.create();
	    calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	function SeleccionarCal(cal, date) {
	    var str = new String(date);
	    if (validarFechaMermaVolatil(str)){
	        document.getElementById("divFecha").innerHTML = str;
	        document.getElementById("fecha").value = str.substr(6, 4) + str.substr(3, 2) + str.substr(0, 2);
	        if (cal) cal.hide();
	    }
	    else{
	        alert("La fecha seleccionada no debe ser menor a la fecha actual");
	    }
	}	
	function validarFechaMermaVolatil(pFecha){
	   var x=new Date();
	   var fecha = pFecha.split("/");
	   x.setFullYear(fecha[2],fecha[1]-1,fecha[0]);
	   var today = new Date();
	   if (x >= today)
	        return true;
	    else
	        return false;
	}
	function eliminarMermaVolatil(pIndex,pFecha,pCdProducto,pCdCliente,pCdSilo) {
	    document.getElementById("tr_"+pIndex).style.display = "none";
	    document.getElementById("estado_"+pIndex).value = "<%=ESTADO_BAJA%>";
    }
	function agregarMermaVolatil(){
	    var tblDetalle = document.getElementsByTagName("tbody")[0];
	    var index = document.getElementById("index").value;
	    //index = parseInt(index) + 1;
	    //var rDetalle   = tblDetalle.insertRow(parseInt(index));
	    var rDetalle = document.createElement("tr");
	    rDetalle.id= "tr_" + index;
	    rDetalle.className = "reg_Header_navdos";
	    
	    var cProducto = rDetalle.insertCell(0);
	    var cCliente = rDetalle.insertCell(1);
	    var cSilo = rDetalle.insertCell(2);
	    var cRatio = rDetalle.insertCell(3);
	    var cGuardar = rDetalle.insertCell(4);


        //Producto
	    cProducto.id = "tdProdcuto_"+index;
	    ch.bind("mermaVolatilAjax.asp?pto=<%=g_strPuerto %>&accion=<%=ACCION_VISUALIZAR %>", "obtenerProducto_Callback("+ index +")");
	    ch.send();
	    cProducto.align = "center";
	    //Cliente
	    var iCliente = document.createElement('input');
	    iCliente.type = "text";
	    iCliente.id = "dsCliente_"+index;
	    iCliente.name = "dsCliente_"+index;
	    iCliente.setAttribute('style',"width:90%");
	    cCliente.appendChild(iCliente);
	    cCliente.align = "left";
	    var hCliente = document.createElement('input');
	    hCliente.type = "hidden";
	    hCliente.id = "cdCliente_"+index;
	    hCliente.name = "cdCliente_"+index;
	    cCliente.appendChild(hCliente);
        //Silo
	    var iSilo = document.createElement('input');
	    iSilo.type = "text";
	    iSilo.id = "cdSilo_"+index;
	    iSilo.name = "cdSilo_"+index;
	    iSilo.maxLength = 10;
	    iSilo.size = 10;
	    cSilo.appendChild(iSilo);
	    cSilo.align = "left";
	    //ratio
	    var iRatio = document.createElement('input');
	    iRatio.type = "text";
	    iRatio.id = "ratio_"+index;
	    iRatio.name = "ratio_"+index;
	    iRatio.setAttribute('style',"text-align:right;");
	    iRatio.size = 6;
	    iRatio.setAttribute('onkeypress',"return controlIngreso(this,event,'N')");
	    cRatio.appendChild(iRatio);
	    cRatio.align = "left";
	    var hEstado = document.createElement('input');
	    hEstado.type = "hidden";
	    hEstado.id = "estado_"+index;
	    hEstado.name = "estado_"+index;
	    hEstado.value = "<%=ESTADO_ACTIVO%>"
	    tblDetalle.appendChild(hEstado);
	    tblDetalle.appendChild(rDetalle);
	    autoCompleteCorredor(index);

	    document.getElementById("index").value = parseInt(index) + 1 ;
	}
	
	function obtenerProducto_Callback(pIndex){
	    var respuesta = ch.response();
	    document.getElementById("tdProdcuto_"+pIndex).innerHTML = respuesta;
	    renameComboBoxProducto(pIndex)
	}
	function renameComboBoxProducto(pIndex){
	    $("#tdProdcuto_"+ pIndex).each(function(){
	        var id = "#" + this.id;
	        $(id + " select").attr("id", "cdProducto_" + pIndex);
	        $(id + " select").attr("name", "cdProducto_" + pIndex);
	    })
	}
	function verificarCliente(pIndex){
	    var dsCliente = document.getElementById("dsCliente_"+pIndex).value.toString();
	    if (dsCliente.trim() == "") {
	        document.getElementById("dsCliente_"+pIndex).value = "";
	        document.getElementById("cdCliente_"+pIndex).value = 0;
	    }

	}
	function autoCompleteCorredor(pIndex){
	    $( "#dsCliente_"+pIndex).autocomplete({
	        minLength: 2,
	        source: "puertosStreamElementos.asp?tipo=JQClientes&pto=<%=g_strPuerto%>",
	        focus: function( event, ui ) {
	            $( "#dsCliente_"+pIndex).val(ui.item.dscliente);
	            return false;
	        },
	        select: function( event, ui ) {
	            $( "#dsCliente_"+pIndex).val (ui.item.dscliente);
	            $( "#cdCliente_"+pIndex).val (ui.item.cdcliente);
	            return false;
	        },
	        change: function( event, ui ) {				
	            if (!ui.item) {					
	                $( "#dsCliente_"+pIndex).val ("");
	                $( "#cdCliente_"+pIndex).val ("");
	            }
	        }
	    })
		.data( "autocomplete" )._renderItem = function( ul, item ) {
		    return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdcliente + " - <font style='font-size:10;'>" + item.dscliente + "</font></a>" )
				.appendTo( ul );
		};
	}		
</script>
    </HEAD>
<BODY onload="onLoadPage()">	
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post action="mermaVolatilPopUp.asp">
    <%=showMessages()%>
    <div class="col66"></div>
    
    <div class="col26 reg_header_navdos"> Valido desde: </div>
    <div class="col26">
        <% if (not rs.eof) then g_fecha = rs("FECHA")  %>
        <div id="divFecha" style="float:left" ><% =GF_FN2DTE(g_fecha) %></div>&nbsp&nbsp
        <input type="hidden" id="fecha" name="fecha" value="<%=g_fecha %>">
        <% if (not flagEdit) then %>
        <a href="javascript:MostrarCalendario(SeleccionarCal, 'fecha')"  >
    	    <img id="img_fecha" src="../images/calendar-16.png">
		</a>
        <% end if %>
    </div>
    
    <div class="col66"></div>	

	<table width="95%" class="datagrid" align="center" id="tblResult" name="tblResult">           
        <thead>
            <tr>
			    <th width="30%" align="center"><% =GF_TRADUCIR("Producto") %></th>
				<th width="35%" align="center"><% =GF_TRADUCIR("Cliente") %></th>
                <th width="20%" align="center"><% =GF_TRADUCIR("Silo") %></th>
				<th width="5%" align="center"><% =GF_TRADUCIR("Ratio") %></th>
				<th width="5%"  align="center">.</th>
            </tr>
        </thead>
        <tbody>
     <% index = 0
     	if (not rs.eof) then 
        while not rs.EOF  %>
            <tr class="reg_Header_navdos" id="tr_<%=index %>">
               <% if (not flagEdit) then %>
                    <td align="center">
                        <select id="cdProducto_<%=index %>" name="cdProducto_<%=index %>" style="width:95%;" >
		                    <option value="0"><%= GF_TRADUCIR("Selccione...")%></option>
			                <%  strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
			                    call GF_BD_Puertos (g_strPuerto, rsProductos, "OPEN",strSQL)
				                while not rsProductos.eof
				                    if cint(rs("CDPRODUCTO")) = cint(rsProductos("CDPRODUCTO")) then
					                    mySelected = "SELECTED"
					                else
					                    mySelected = ""
					                end if	%>
					                <option value="<%=rsProductos("CDPRODUCTO")%>" <%=mySelected%>><%=rsProductos("DSPRODUCTO")%></option>
				                <%	rsProductos.movenext
				                wend %>
                        </select>
                    </td>
               <% else %>
                    <td align="left">
                        <div><%=rs("DSPRODUCTO") %></div>
                        <input type="hidden" id="cdProducto_<%=index %>" name="cdProducto_<%=index %>" value="<%=rs("CDPRODUCTO") %>" />
                    </td>
               <% end if %>
                <td align="left">
               <% if (not flagEdit) then %>
                    <input type="text" id="dsCliente_<%=index %>" name="dsCliente_<%=index %>" value="<% =rs("DSCLIENTE")%>" style="width:90%" onblur="verificarCliente(<%=index %>)"/>
               <% else %>
                    <div><%=rs("DSCLIENTE") %></div>
               <% end if %>
                    <input type="hidden" id="cdCliente_<%=index %>" name="cdCliente_<%=index %>" value="<% =rs("CDCLIENTE")%>"/>
                </td>
                <td align="left">
               <% if (not flagEdit) then %>
                    <input type="text" id="cdSilo_<%=index %>" name="cdSilo_<%=index %>" value="<% =Trim(rs("CDSILO")) %>" size="10"/>
               <% else %>
                    <div><% =Trim(rs("CDSILO")) %></div>
                    <input type="hidden" id="cdSilo_<%=index %>" name="cdSilo_<%=index %>" value="<% =Trim(rs("CDSILO")) %>"/>
               <% end if %>
                </td>
                <td align="center">
                    <input type="text" id="ratio_<%=index %>" name="ratio_<%=index %>" value="<% =GF_EDIT_DECIMALS(Cdbl(rs("RATIO"))*100,2) %>" size="6" style="text-align:right;" onkeypress="return controlIngreso(this,event,'N');"/>
                </td>
                <% if (not flagEdit) then %>
					<td align="center"><img src="../images/cross-16.png" title="Eliminar" style="cursor:pointer;" onclick="eliminarMermaVolatil(<%=index%>,'<%=rs("FECHA")%>',<%=rs("CDPRODUCTO")%>,<%=rs("CDCLIENTE")%>,'<%=rs("CDSILO")%>')"></td>
                <% else %>
                    <td align="center"></td>
                <% end if %>
                <input type="hidden" value="<%=ESTADO_ACTIVO %>" id="estado_<%=index %>" name="estado_<%=index %>" />
            </tr>
        <%  index = index + 1 
            rs.movenext()
		wend %>
<%	end if %>
    </tbody>
<tfoot>
    <% if (not flagEdit) then %>
    <tr>
        <td colspan="6" align="right"><img src="../images/add.gif" title="Agregar" style="cursor:pointer;" onclick="javascript:agregarMermaVolatil()"/></td>
    </tr>
    <% end if %>
</tfoot>
</TABLE>
    <div class="col56"> </div>
    <span class="btnaction"><input type="submit" id="btnAccion" value="Guardar" /></span>
<input type="hidden" name="index" id="index" value="<%=index %>">
<input type="hidden" name="accion" id="accion" value="<%= ACCION_GRABAR %>">
<input type="hidden" name="Pto" id="Pto" value="<%= g_strPuerto %>">


<input type="hidden" name="cdProducto" id="cdProducto" value="<%= g_cdProducto %>">
<input type="hidden" name="cdCliente" id="cdCliente" value="<%= g_cdCliente %>">
<input type="hidden" name="cdSilo" id="cdSilo" value="<%= g_cdSilo %>">

</form>
</BODY>
</HTML>