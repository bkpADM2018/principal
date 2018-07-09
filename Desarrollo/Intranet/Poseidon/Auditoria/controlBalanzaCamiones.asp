<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientostitulos.asp"-->
<!--#include file="../../Includes/procedimientospuertos.asp"-->
<!--#include file="../../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../../Includes/procedimientosUser.asp"-->
<%
Const FUNCION_BZA_BRUTO = "BRUTO"
Const FUNCION_BZA_TARA = "TARA"
Const FUNCION_BZA_FUERA_DE_LINEA = "FUERA DE LINEA"
'-- Parametro generico para identificar a la balanza correspondiente. Reemplazar la X por el nro que corresponde a la balanza.
Const PARAM_FUNCION_BZA_CAMIONES_X = "CTRLBZAXFUNCION"

'---------------------------------------------------------------------------------------------------------------
Function drawEstadoBzaCam(pIdEstado)
	Dim rtrn 	
	'UTILIZO ESTA COPROBACION EN CASO DE QUE EN EL RM&D ACTUALICE EL ESTADO AUTOMATICAMENTE CADA VEZ QUE 
	'CARGA UNA NUEVA BALANZA		
	select case pIdEstado					
		case BZA_CAM_ESTADO_FINALIZADO
			Response.Write "<IMG src='../../images/finalizado.png' title='Finalizado'>"
		case BZA_CAM_ESTADO_CANCELADO
			Response.Write "<IMG src='../../images/cancelado.png' title='Cancelado'>"	
		case else 			
			Response.Write "<IMG src='../../images/activo.png' title='Activo'>"			
	end select 	
End Function
'---------------------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------------------
' Función:	getCantidadBZAPuerto
' Autor: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	15/02/13
' Objetivo:	
'			Identificar cuantas balanzas utiliza el puerto, y su orden
' Parametros:
'			fechaDesde 	[date] 	fecha Inicio
'			fechaHasta 	[date] 	fecha Final
'			patente 	[string]
'			acoplado 	[string]
'			estado 		[int] 	
'Devuelve : 
'			Carga en un array las posiciones de las balanzas
'---------------------------------------------------------------------------------------------------------------
Function getCantidadBZAPuerto(fechaDesde,fechaHasta,patente,acoplado,estado, tControl)
	Dim strSQL, myWhere, cantBZA 
	Call buscarFiltrosControlBalanza(myWhere,fechaDesde,fechaHasta,patente,acoplado,estado, tControl)
	strSQL = " SELECT COUNT(BRUTO1) CANT1, COUNT(BRUTO2) CANT2, COUNT(BRUTO3) CANT3, COUNT(BRUTO4) CANT4, COUNT(TARA) CANT5 FROM CTRLBZACAMIONES " & myWhere	
    Call executeQueryDb(pto, rs, "OPEN", strSQL)
	cantBZA = 0
	if not rs.Eof then		
		if(rs("CANT1") > 0)then v_haveBZA(0) = true		
		if(rs("CANT2") > 0)then v_haveBZA(1) = true
		if(rs("CANT3") > 0)then v_haveBZA(2) = true
		if(rs("CANT4") > 0)then v_haveBZA(3) = true
		if(rs("CANT5") > 0)then v_haveBZA(4) = true
	end if		
End Function
'---------------------------------------------------------------------------------------------------------------
'Permite saber si el usuario tiene el rol de cancelar un Control, en caso de que pueda se verifica que no este Finalizado
Function canDeleteControl(pMod, pEstado)
	Dim rtrn
	rtrn = false
	if(pMod)then			
		if((pEstado <> BZA_CAM_ESTADO_FINALIZADO)and(pEstado <> BZA_CAM_ESTADO_CANCELADO))then rtrn = true
	end if	
	canDeleteControl = rtrn
End Function
'---------------------------------------------------------------------------------------------------------------
'/***********************************************************************************************************************/
'/******************************************* INICIA PAGINA *************************************************************/
'/***********************************************************************************************************************/


Dim pto, tList, diaDesde, mesDesde, anioDesde, diaHasta, mesHasta, anioHasta, estado, acoplado, patente
Dim paginaActual,mostrar, valueParameter,v_haveBZA(5),flagCall, tControl, bzaName

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)

patente = GF_PARAMETROS7("patente", "", 6)
if (patente <> "") then patente = replace(patente, "-", "")
call addParam("patente", patente, params)

acoplado = GF_PARAMETROS7("acoplado", "", 6)
if (acoplado <> "") then acoplado = replace(acoplado, "-", "")
call addParam("acoplado", acoplado, params)

accion = GF_PARAMETROS7("accion", "", 6)

estado = GF_PARAMETROS7("estado", 0, 6)
if((estado = 0)and(accion <> ACCION_SUBMITIR))then estado = BZA_CAM_ESTADO_TODOS
Call addParam("estado", estado, params)

tControl = GF_PARAMETROS7("tControl", "", 6)
if((tControl = "")and(accion <> ACCION_SUBMITIR))then tControl = BZA_CAM_TIPO_CTRL_TODOS
Call addParam("tControl", tControl, params)

diaDesde = GF_PARAMETROS7("diaDesde", "", 6)
'if (diaDesde = "") then diaDesde = GF_nDigits(Day(Now()),2)
'POR DEFAULT SIEMPRE MUESTRA TODOS LOS CONTROLES DEL MES EN CURSO (EL DIA DEL MES EMPIEZA EL 01)
if (diaDesde = "") then diaDesde = 1
diaDesde = GF_nDigits(diaDesde,2)
call addParam("diaDesde", diaDesde, params)

mesDesde = GF_PARAMETROS7("mesDesde", "", 6)
if (mesDesde = "") then mesDesde= Month(Now())
mesDesde = GF_nDigits(mesDesde,2)
Call addParam("mesDesde", mesDesde, params)

anioDesde = GF_PARAMETROS7("anioDesde", "", 6)
if (anioDesde = "") then anioDesde= Year(Now())
anioDesde = GF_nDigits(anioDesde,4)
Call addParam("anioDesde", anioDesde, params)

diaHasta = GF_PARAMETROS7("diaHasta", "", 6)
if (diaHasta = "") then diaHasta= Day(Now())
diaHasta = GF_nDigits(diaHasta,2)
'POR DEFAULT SIEMPRE MUESTRA TODOS LOS CONTROLES DEL MES EN CURSO (EL DIA DEL MES PUEDE FINALIZAR EL 28,29,30,31)
'if (diaHasta = "") then diaHasta= LastDayOfMonth(GF_nDigits(Year(Now()),4), GF_nDigits(Month(Now()),2))
call addParam("diaHasta", diaHasta, params)

mesHasta = GF_PARAMETROS7("mesHasta", "", 6)
if (mesHasta = "") then mesHasta= Month(Now())
mesHasta = GF_nDigits(mesHasta,2)
call addParam("mesHasta", mesHasta, params)

anioHasta = GF_PARAMETROS7("anioHasta", "", 6)
if (anioHasta = "") then anioHasta= Year(Now())
anioHasta = GF_nDigits(anioHasta,4)
call addParam("anioHasta", anioHasta, params)


ret = GF_CONTROL_PERIODO(diaDesde, diaHasta, mesDesde, mesHasta, anioDesde, anioHasta)
flagCall = false
Select case (ret)
	case 0	
		flagCall=true
	case 1
		Call setError(FECHA_INICIO_INCORRECTA)
	case 2
		Call setError(FECHA_FIN_INCORRECTA)
	case 3
		Call setError(PERIODO_ERRONEO)
end select

mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 50
if (paginaActual = 0) then paginaActual = 1

myFechaDesde = anioDesde & mesDesde & diaDesde
myFechaHasta = anioHasta & mesHasta & diaHasta

Call initTaskAccessInfo(TASK_BZA_CAM_CTRL_PESO, session("DIVISION_PUERTO"))
esModificable = hasWriteAcess(TASK_BZA_CAM_CTRL_PESO, session("DIVISION_PUERTO"))


Call getCantidadBZAPuerto(myFechaDesde,myFechaHasta,patente,acoplado,estado, tControl)

Set rs = leerControlBalanza(pto,myFechaDesde,myFechaHasta,patente,acoplado,estado, tControl)

Call setupPaginacion(rs, paginaActual, mostrar)
lineasTotales = rs.recordcount


%>

<HTML>
<HEAD>

<meta http-equiv="X-UA-Compatible" content="IE=Edge">

<link rel="stylesheet" href="../../css/ActisaIntra-1.css"  type="text/css" />
<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css"  type="text/css" />
<link rel="stylesheet" href="../../css/main.css" type="text/css"> 


<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/formato.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/paginar.js"></script>
<script type="text/javascript" src="../../scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../../Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="../../Scripts/botoneraPopUp.js"></script>
<script type="text/javascript" src="../../Scripts/jQueryPopUp.js"></script>

<script type="text/javascript">	
	var ch = new channel();	
	var popUpPar;
	
	<% if((accion = ACCION_PROCESAR)and(flagCall))then %>
		window.open("controlBalanzaCamionesPrintXLS.asp<% =params %>");
	<% end if %>
	
	function bodyOnLoad() {
		tb = new Toolbar('toolbar', 6,'images/');				
		tb.addButton("refresh-16x16.png", "Recargar", "submitInfo('<%=ACCION_SUBMITIR%>')");
		tb.addButton("excel3.gif", "Imprimir XLS", "submitInfo('<%=ACCION_PROCESAR%>')");
		tb.addButton("Previous-16x16.png", "Volver", "volver()");
		tb.draw();
	<% 	if (not rs.eof) then %>
			var pgn = new Paginacion("paginacion");
		pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 200, "controlBalanzaCamiones.asp<% =params %>");						 
	<%	end if 	%>
	    controlar();
		pngfix();	
	}
	
	function recargar(){	
		window.location.reload();
	}
		
	
	function volver(){
		location.href = "seccionAuditoria.asp?pto=<%=pto%>";
	}
	
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
		
	}
	
	function cancelarControl(pIdControl,pFecha){
		if (confirm("Esta seguro que desea cancelar el control de la fecha " + pFecha + " ? ")) {
			ch.bind("controlBalanzaCamionesAjax.asp?idControl="+pIdControl+"&estado=<%=BZA_CAM_ESTADO_CANCELADO%>&accion=<%=ACCION_CANCELAR%>&pto=<%=pto%>", "cancelarControl_CallBack()");
			ch.send();
		}		
	}
		
	function cancelarControl_CallBack(){			
		submitInfo('<%=ACCION_SUBMITIR%>');
	}
	
	function modificarParametro(pCdParametro){
		var puw = new winPopUp("Iframe","../parametrosPopUp.asp?codParametro=" + pCdParametro + "&pto=<%=pto%>&visitante=<%=esmodificable%>&accion=<%=ACCION_MODIFICAR_PARAMETRO%>","450", "290","Modificar Parametro", "submitInfo('<%=ACCION_SUBMITIR%>');");
	}
	
	function controlar() {
	    //Averiguo si hay exactamente una balanza de tara.
	    var tara = 0
	    var radios = document.getElementsByTagName('input');
	    for (var i = 0; i < radios.length; i++) {
	        if (radios[i].type === 'radio' && radios[i].checked) {	       
	            if (radios[i].value === "<% =FUNCION_BZA_TARA %>") tara++;
	        }
	    }
	    if (tara == 1) {
	        document.getElementById("trWarning").style.visibility = 'hidden';
	        document.getElementById("trWarning").style.position = 'absolute';
	        document.getElementById("trOK").style.visibility = 'visible';
	        document.getElementById("trOK").style.position = 'relative';
	        return true;
	    } else {
	        document.getElementById("trOK").style.visibility = 'hidden';
            document.getElementById("trOK").style.position = 'absolute';
            document.getElementById("trWarning").style.visibility = 'visible';
            document.getElementById("trWarning").style.position = 'relative';
	        return false;
	    }
	}
	function seleccionarFuncion_cb(pParam, pVal) {
	        document.getElementById(pParam + "_valold").value = pVal;
	}
	function seleccionarFuncion() {
	    var ch = new channel();
	    if (controlar()) {
	        var cant = 0;
	        var radios = document.getElementsByTagName('input');
	        for (var i = 0; i < radios.length; i++) {	            
	            if (radios[i].type === 'radio' && radios[i].checked) {
	                if ((radios[i].value === "<% =FUNCION_BZA_TARA %>") || (radios[i].value === "<% =FUNCION_BZA_BRUTO %>")) cant++;
	                var param = document.getElementById(radios[i].name + "_param").value;
                    var desc = document.getElementById(radios[i].name + "_desc").value;
                    var valold = document.getElementById(radios[i].name + "_valold").value;
                    if (radios[i].value != valold) {
                        ch.bind("../ParametrosAjax.asp?cdParam="+param+"&valParam="+radios[i].value+"&valParam_old="+valold+"&nomParam="+desc+"&nomParam_old="+desc+"&accion=<% =ACCION_MODIFICAR_PARAMETRO %>&pto=<% =pto %>","seleccionarFuncion_cb('" + param + "', '" + radios[i].value + "')");	        
			            ch.send();
	                }
	            }
	        }
	        //Se actualiza el nro de balanzas activas.
	        var param = document.getElementById("<% =PARAM_CANT_BZA_CAMIONES %>_param").value;
            var desc = document.getElementById("<% =PARAM_CANT_BZA_CAMIONES %>_desc").value;
            var valold = document.getElementById("<% =PARAM_CANT_BZA_CAMIONES %>_valold").value;	        
            if (cant != valold) {
	            ch.bind("../ParametrosAjax.asp?cdParam="+param+"&valParam="+cant+"&valParam_old="+valold+"&nomParam="+desc+"&nomParam_old="+desc+"&accion=<% =ACCION_MODIFICAR_PARAMETRO %>&pto=<% =pto %>","seleccionarFuncion_cb('" + param + "', '" + cant + "')");
			    ch.send();
			    document.getElementById("<% =PARAM_CANT_BZA_CAMIONES %>_val").innerHTML = cant;
			}
	    }	    
	}
</script>
</HEAD>
<BODY onLoad="bodyOnLoad()">	
<% call GF_TITULO2("kogge64.gif","Control de Balanza de Camiones") %>
<DIV id="toolbar"></DIV>
<FORM id="frmSel" name="frmSel" method="POST" action="controlBalanzaCamiones.asp">
		<div class="tableaside size100"> <!-- BUSCAR -->
            <h3> filtro - Control Balanzas - <% =pto %></h3>
            
            <div id="searchfilter" class="tableasidecontent">
			
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Desde") %> </div>
                <div class="col16"> 
		            <INPUT type="text" id="diaDesde" name="diaDesde" value="<%=diaDesde%>" size=2 maxlength=2>-<INPUT type="text"  id="mesDesde" name="mesDesde" size=2 maxlength=2 value="<%=mesDesde%>">-<INPUT type="text"  id="anioDesde" name="anioDesde" size=4 maxlength=4 value="<%=anioDesde%>">
                </div>			
				
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Hasta") %> </div>
                <div class="col16"> 
		            <INPUT type="text"  id="diaHasta" name="diaHasta" size=2 maxlength=2 value="<%=diaHasta%>">-<INPUT type="text"  id="Text5" name="mesHasta" size=2 maxlength=2 value="<%=mesHasta%>">-<INPUT type="text"  id="Text6" name="anioHasta" size=4 maxlength=4 value="<%=anioHasta%>">
                </div>			
				
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Patente") %> </div>
                <div class="col16"> 
		            <INPUT type="text"  id="patente" name="patente"  size="10" maxlength="10" value="<%=patente%>">
                </div>
				
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Acoplado") %> </div>
                <div class="col16"> 
		            <INPUT type="text"  id="acoplado" name="acoplado"  size="10" maxlength="10" value="<%=acoplado%>">
                </div>			
				
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Estado") %> </div>
                <div class="col16"> 
		            <SELECT id="SELECT1" name="estado">
						<OPTION value="<% =BZA_CAM_ESTADO_TODOS		%>" <%if (estado = BZA_CAM_ESTADO_TODOS)     then %> selected='true' <%end if%>><% =GF_TRADUCIR("-Todos-")   %></OPTION>
						<OPTION value="<% =BZA_CAM_ESTADO_EN_CURSO  %>" <%if (estado = BZA_CAM_ESTADO_EN_CURSO)  then %> selected='true' <%end if%>><% =GF_TRADUCIR("Curso")     %></OPTION>
						<OPTION value="<% =BZA_CAM_ESTADO_FINALIZADO%>" <%if (estado = BZA_CAM_ESTADO_FINALIZADO)then %> selected='true' <%end if%>><% =GF_TRADUCIR("Finalizado")%></OPTION>
						<OPTION value="<% =BZA_CAM_ESTADO_CANCELADO %>" <%if (estado = BZA_CAM_ESTADO_CANCELADO) then %> selected='true' <%end if%>><% =GF_TRADUCIR("Canceldo")  %></OPTION>
					</SELECT>										
                </div>
				
				<div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Tipo") %> </div>
                <div class="col16"> 
		            <SELECT id="SELECT2" name="tControl">
						<OPTION value="<% =BZA_CAM_TIPO_CTRL_TODOS  %>" <%if (tControl = BZA_CAM_TIPO_CTRL_TODOS)     then %> selected='true' <%end if%>><% =GF_TRADUCIR("-Todos-")   %></OPTION>
						<OPTION value="<% =BZA_CAM_TIPO_CTRL_MANUAL %>" <%if (tControl = BZA_CAM_TIPO_CTRL_MANUAL)  then %> selected='true' <%end if%>><% =GF_TRADUCIR("MANUAL")     %></OPTION>
						<OPTION value="<% =BZA_CAM_TIPO_CTRL_AUTOM  %>" <%if (tControl = BZA_CAM_TIPO_CTRL_AUTOM)then %> selected='true' <%end if%>><% =GF_TRADUCIR("AUTOMATICO")%></OPTION>
					</SELECT>	
                </div>			
				
				<span class="btnaction"><input type="submit" value="Buscar"></span>
				
			</div>						
				
        </div><!-- END BUSCAR -->        
    	
<div class="col66"></div>

<%	Call showErrors() 	%>
		
	<table width="100%" align="center">
		<tr>
			<td width="50%" valign="top">
				<table class="datagrid" align="center" width="450px">
					<thead>
						<tr>
							<th align=center width="378px" ><%=GF_TRADUCIR("Controles")%></th>
							<th align=center width="40px" ><%=GF_TRADUCIR("Valor")%></th>
							<th align=center width="16px" ><%=GF_TRADUCIR(".")%></th>
						</tr>
					</thead>
					<tbody>
						<% 	strSQL = "SELECT * FROM PARAMETROS WHERE CDPARAMETRO IN ('" & PARAM_CANT_BZA_CONTROLES & "','" & PARAM_CANT_BZA_CAMIONES & "','" & PARAM_TIPO_CTRL_BZA_CAMIONES & "')"
							Call executeQueryDb(pto, rsParametrosBZA, "OPEN", strSQL)
							if(not rsParametrosBZA.EoF)then
								while(not rsParametrosBZA.EoF)  %>
									<tr>
										<td><%if(Len(CStr(rsParametrosBZA("DSPARAMETRO"))) > 50)then
												  Response.Write left(CStr(rsParametrosBZA("DSPARAMETRO")),50)  & "..."
											  else
												  Response.Write Cstr(rsParametrosBZA("DSPARAMETRO"))
											  end if %>
												<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_param" value="<% =rsParametrosBZA("CDPARAMETRO")%>" />
												<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_valold" value="<% =rsParametrosBZA("VLPARAMETRO")%>" />
												<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_desc" value="<% =rsParametrosBZA("DSPARAMETRO") %>" />
										</td>
										<td align=center ><span id="<% =rsParametrosBZA("CDPARAMETRO") %>_val"><%=rsParametrosBZA("VLPARAMETRO")%></span></td>
									<%  'La cantidad de balanzas no debe ser editable, se modifica automáticamente al cambiar las funciones de las balanzas.
										if((esModificable) and (rsParametrosBZA("CDPARAMETRO") <> PARAM_CANT_BZA_CAMIONES))then %>
											<td onclick="modificarParametro('<%=rsParametrosBZA("CDPARAMETRO")%>')" ><IMG src="images/edit-16x16.png" title="Editar" style="cursor:pointer"></td>
									<%  else  %>	
											<td align="center" ></td>	
									<%  end if  %>
									</tr>
							<%		rsParametrosBZA.MoveNext
								wend	
							else  %>
								<tr>
									<td colspan="4" align="center"><%=GF_TRADUCIR("No se encontraron los parametros de la balanza")%></td>
								</tr>
						<%	end if	%>
					</tbody>
				</table>				
			</td>
			<td>
				<table class="datagrid" align="center" width="500px">								
					<thead>
						<tr id="trWarning">
							<td colspan="3" style="background-color:#ffff99">
								<img src="../../images/compras/warning-16x16.png" align="absMiddle">&nbsp;<b><% =GF_TRADUCIR("Sin una única balanza de TARA los cambios no tendran efecto.") %></b>
							</td>
						</tr>
						<tr id="trOK">
							<td colspan="3" class="reg_Header_success">
								<img src="../../images/compras/accept-16x16.png" align="absMiddle">&nbsp;<b><% =GF_TRADUCIR("La configuracion es correcta.") %></b>
							</td>
						</tr>
						<tr>
							<th align=center width="250px" ><%=GF_TRADUCIR("Balanzas")%></th>
							<th align=center colspan="6" ><%=GF_TRADUCIR("Valor")%></th>							
						</tr>
					</thead>
					<tbody>
						<% 	strSQL = "SELECT * FROM PARAMETROS WHERE CDPARAMETRO IN ("
							for i = 1 to CInt(BZA_CAM_MAX)
								strSQL = strSQL & "'" & replace(PARAM_FUNCION_BZA_CAMIONES_X, "X", i) & "',"
							next
							strSQL = Left(strSQL, len(strSQL)-1) & ") order by CDPARAMETRO"
							Call executeQueryDb(pto, rsParametrosBZA, "OPEN", strSQL)
							if(not rsParametrosBZA.EoF)then
								while(not rsParametrosBZA.EoF)  %>
									<tr>
										<td align="center"><%if(Len(CStr(rsParametrosBZA("DSPARAMETRO"))) > 50)then
												  Response.Write left(CStr(rsParametrosBZA("DSPARAMETRO")),50)  & "..."
											  else
												  Response.Write Cstr(rsParametrosBZA("DSPARAMETRO"))
											  end if %>
										</td>
									<%  if(esModificable)then %>
										<td align=center >												    
											<input type="radio" name="<% =rsParametrosBZA("CDPARAMETRO") %>" onclick="seleccionarFuncion()" value="<% =FUNCION_BZA_BRUTO %>" <% if (rsParametrosBZA("VLPARAMETRO") = FUNCION_BZA_BRUTO) then response.write "checked" %> /> Bruto
											<input type="radio" name="<% =rsParametrosBZA("CDPARAMETRO") %>" onclick="seleccionarFuncion()" value="<% =FUNCION_BZA_TARA %>" <% if (rsParametrosBZA("VLPARAMETRO") = FUNCION_BZA_TARA) then response.write "checked" %> /> Tara
											<input type="radio" name="<% =rsParametrosBZA("CDPARAMETRO") %>" onclick="seleccionarFuncion()" value="<% =FUNCION_BZA_FUERA_DE_LINEA %>" <% if (rsParametrosBZA("VLPARAMETRO") = FUNCION_BZA_FUERA_DE_LINEA) then response.write "checked" %> /> Fuera de Linea
											<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_param" value="<% =rsParametrosBZA("CDPARAMETRO")%>" />
											<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_valold" value="<% =rsParametrosBZA("VLPARAMETRO")%>" />
											<input type="hidden" id="<% =rsParametrosBZA("CDPARAMETRO") %>_desc" value="<% =rsParametrosBZA("DSPARAMETRO") %>" />
										</td>
									<%  else  %>	
										<td align="center" ><%=rsParametrosBZA("VLPARAMETRO")%></td>	
									<%  end if  %>
									</TR>
						<%		rsParametrosBZA.MoveNext
								wend	
							else  %>
								<tr>
									<td colspan="4" align="center"><%=GF_TRADUCIR("No se encontraron los parametros de la balanza")%></td>
								</tr>
						<%	end if	%>	
					</tbody>
				</table>
			</td>
		</tr>
	</table>
	<input type="hidden" id ="pto" 		name="pto" value=<%=pto%>>
	<input type="hidden" id ="accion" 	name="accion" value=<%=accion%>>
</FORM>

<div class="col66"></div>

    <table class="datagrid" width="100%" align="center">
		<thead>		
			<tr>
				<th width="5%" style="text-align: center"><%=GF_Traducir("Fecha")%></th>
				<th width="2%" style="text-align: center"><%=GF_Traducir("Tipo")%></th>
				<th width="6%" align="center"><%=GF_Traducir("Patente")%></th>			
				<th width="6%" align="center"><%=GF_Traducir("Acoplado")%></th>
				<% For i = 0 To UBound(v_haveBZA) - 1 
					  if(v_haveBZA(i))then  
						if(i <> 4)then
							bzaName = "Bruto " & i + 1
						else
							bzaName = "Tara"
						end if
					  
				%>
						<th width="8%" align="center"><% = "Peso " & bzaName %></th>
						<th width="8%" align="center"><% = "Responsable " & bzaName%></th>
						<th width="8%" align="center"><% = "Fecha Control " & bzaName%></th>
				<%    end if	
				   Next %>			
				<th width="2%" align="center"> . </th>
				<th width="2%"align="center"> . </th>
			</tr>			
		</thead>		
<% 	if (not rs.eof) then %>				
		<tbody>
		<% 	while ((not rs.eof) and (CInt(reg) < CInt(mostrar)))
				reg = reg + 1	%>
				<tr>
					<td align="center"><%=GF_FN2DTE(Cdbl(rs("FECHA")))%></td>
					<td align="center"><%=rs("TIPOCONTROL") %></td>
					<td align="center"><%= Left(rs("CDCHAPACAMION"),3) & "-" & Right(rs("CDCHAPACAMION"),3)%></td>		
					<td align="center"><%= Left(rs("CDCHAPAACOPLADO"),3) & "-" & Right(rs("CDCHAPAACOPLADO"),3)%></td>						
				<%  For i = 0 To UBound(v_haveBZA) - 1 
				        if(v_haveBZA(i))then  %>
						    <td align="center">
						<%		if(i <> 4)then
									bzaName = "BRUTO" & i + 1
									auxBZA = rs(bzaName)
								else
									bzaName = "TARA"
								    auxBZA = rs(bzaName)								    
								end if	
								if Not Isnull(auxBZA)then 
								    Response.Write GF_EDIT_DECIMALS(Cdbl(auxBZA)*100,2) & " Kg"
								    if Not Isnull(rs("TARA"))then 
								    	diff = CDbl(auxBZA) - CDbl(rs("TARA"))
								    	if (diff <> 0) then Response.Write "(" & GF_EDIT_DECIMALS(diff*100,2) & " Kg" & ")"
								    end if
								end if
					    %>				
							</td>
							<td align="center"><% =getUserDescription(rs(bzaName & "USR")) %></td>
							<td align="center"><% =GF_FN2DTE(rs(bzaName & "MMTO")) %></td>
				<%     end if	
					Next %>					
					<td align="center"><% Call drawEstadoBzaCam(Cdbl(rs("ESTADO"))) %></td>
		<%			if(canDeleteControl(esModificable, Cdbl(rs("ESTADO"))))then				%>
						<td align="center" onclick="javascript:cancelarControl(<%=rs("IDCONTROL")%>,'<%=GF_FN2DTE(Cdbl(rs("FECHA")))%>')">
							<IMG title="Cancelar" src="images/cancel-16x16.png" style="cursor:pointer">
						</td>
		<%			else		%>		
						<td align="center"></td>
		<%			end if		
				rs.movenext
				wend 	%>
				</tr>
		</tbody>
		<tfoot>
			<TR><td colspan="15"><DIV id="paginacion"></DIV></td></TR>	
		</tfoot>
<%	else	%>
		<tbody>
			<TR class="TDNOHAY"><td colSpan="4"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></TR>
		</tbody>
<%	end if	%>		
	</table>

</BODY>
</HTML>