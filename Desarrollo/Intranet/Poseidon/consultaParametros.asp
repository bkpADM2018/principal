<!--#include file="../Includes/procedimientosMG.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<%
Const SIN_PUESTO_OCUPADO = 0
'----------------------------------------------------------------------------------
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
'-----------------------------------------------------------------------------------
Function validarEditables(modif,ppto,pCdParam)
	dim rtrn, rs, myEditable
	rtrn = false
	if((modif = TASK_PARAM_ADMIN))then
		'muestra los comandos de edicion por que es un administrador
		rtrn = true				
	else	
		if(tieneParametrosExtra(pCdParam, ppto))then 
			'Si ya tiene parametrosExtra entonces traigo el registro EDITABLE'
			Set rs = traerParametrosEditables(pCdParam,ppto)
			myEditable = CStr(rs("EDITABLE"))			
		end if
		if((modif = TASK_PARAM_USER) and (myEditable = PARAMETRO_EDITABLE))then 
			'muestra los comandos de edicion por que el usuario tiene permitido editar ese parametro
			rtrn = true		
		end if
	end if
    validarEditables = rtrn
end function

'-----------------------------------------------------------------------------------

Dim rsPar,NombreParametro,paginaActual,mostrar,lineasTotales,rsLista,esmodificable,pto,v_mostrar
Dim accion, cdParam, nomParam, rsPuesto, v_Idpuesto,rsPue,idpuesto,params,mySelected
Dim setOrder

pto = GF_PARAMETROS7("pto", "", 6)
call addParam("pto", pto, params)
cdParam = Trim(UCASE(GF_PARAMETROS7("cdParam", "", 6)))
call addParam("cdParam", cdParam, params)
nomParam = Trim(UCASE(GF_PARAMETROS7("nomParam", "", 6)))
call addParam("nomParam", nomParam, params)
idpuesto = GF_PARAMETROS7("idpuesto", 0, 6)
call addParam("idpuesto", idpuesto, params)
accion = GF_PARAMETROS7("accion", 0, 6)
setOrder = UCASE(GF_PARAMETROS7("setOrder", "", 6))
call addParam("setOrder", setOrder, params)
if(len(setOrder) <= 0)then setOrder = " ORDER BY A.CDPARAMETRO ASC "
'---------------------------------------------'
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1
'--------------------------------------------------------------------------'	
tList = TASK_PARAM_ADMIN & ", " & TASK_PARAM_USER & ", " & TASK_PARAM_AUDIT
esmodificable = leerPermisos(pto, tList)
if (session("Usuario") = "JAS") then esmodificable = TASK_PARAM_ADMIN

if(esmodificable = NO_TIENE_PERMISO)then response.redirect "../comprasAccesoDenegado.asp"

Set rsPar = leerParametros(pto,cdParam,nomParam,idpuesto,setOrder,false)	
Call setupPaginacion(rsPar, paginaActual, mostrar)
lineasTotales = rsPar.recordcount

select case UCASE(pto)
	case TERMINAL_TRANSITO
		rootLog = NOMBRE_RUTA_PARAMETRO_TRA
	case TERMINAL_ARROYO		
		rootLog = NOMBRE_RUTA_PARAMETRO_ARR
	case TERMINAL_PIEDRABUENA
		rootLog = NOMBRE_RUTA_PARAMETRO_LPB
End select	
%>
<HTML>
<HEAD>

<link rel="stylesheet" href="../css/MagicSearch.css" type="text/css">
<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="../css/JQueryUpload2.css"	 type="text/css">
<link rel="stylesheet" href="../css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="../css/tabs-print.css" TYPE="text/css" MEDIA="print">
<link rel="stylesheet" href="../css/Toolbar.css" type="text/css">
<link href="../css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
</style>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script type="text/javascript" src="../Scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="../Scripts/botoneraPopUp.js"></script>
<script type="text/javascript" src="../Scripts/jQueryPopUp.js"></script>
      
<script type="text/javascript">

	var ch = new channel();	
	var popUpPar;
	
	function onLoadPage(){
		tb = new Toolbar('toolbar', 6,'images/');
		tb.addButtonREFRESH("Recargar", "submitInfo()");		
			tb.addButton("log_16x16.png", "Ver Log", "mostrarLog()");
		<%  if(esmodificable = TASK_PARAM_ADMIN)then%>
			tb.addButton("add.gif", "Agregar", "agregarParametro()");
		<%end if%>	
		tb.draw();			
		<% 	if (not rsPar.eof) then %>					
				var pgn = new Paginacion("paginacion");						
				pgn.paginar(<% =paginaActual %>, <% =lineasTotales %>, <% =mostrar %>, 50, "consultaParametros.asp<% =params %>");					
		<%	end if 	%>			
	}
	
	function mostrarLog()
	{
		winPopUp('Iframe', 'verlogs.asp?root=<% =rootLog %>', "900", "500", 'Logs Parametros', '');
	}			
	
	function irA(pLink) {
		location.href = pLink;
	}
	
	function modificarParametro(pCodP){
		popUpPar = winPopUp('Iframe', 'parametrosPopUp.asp?codParametro='+ pCodP +'&pto=<%=pto%>&visitante=<%=esmodificable%>&accion=<%=ACCION_MODIFICAR_PARAMETRO%>', "450", "290", 'Modificar Parametro', 'submitInfo()');			 				
	}			
	
	function agregarParametro(){
		popUpPar = winPopUp('Iframe', 'parametrosPopUp.asp?pto=<%=pto%>&visitante=<%=esmodificable%>&accion=<%=ACCION_AGREGAR_PARAMETRO%>', "450", "290", 'Agregar Parametro', 'submitInfo()');			 				
	}
	function eliminarParametro(pcod){
		ch.bind("ParametrosAjax.asp?cdParam="+pcod+"&accion=<% =ACCION_ELIMINAR_PARAMETRO%>&pto=<%=pto%>","eliminarParametro_CallBack('"+pcod+"')");
		ch.send();				
	}
	function eliminarParametro_CallBack(pcod){
		alert("Se elimino el Parametro: '" + pcod + "'");	
		submitInfo();	
	}	
	function submitInfo() {				
		document.getElementById("frmSel").submit();
	}	
	
	function cerrarPopUpPar(pCod)
	{		
		document.getElementById("cdParam").value= pCod;		
	}
	function setOrder(p_campo,p_orden){
		document.getElementById("setOrder").value = ' ORDER BY '+p_campo+' '+p_orden;
		document.getElementById("frmSel").submit();
	}
		
</script>

</HEAD>
<BODY onload="onLoadPage();">
<div id="toolbar"></div>
<form name="frmSel" id="frmSel" method="get" action="consultaParametros.asp">
	
	<div>
	<br><br>
	<table id="tblBusqueda" width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="images/marcos/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marcos/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marcos/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
           <td width="8"><img src="images/marcos/marco_r2_c3.gif"></td>
           <td align="right"></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="images/marcos/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="images/marcos/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marcos/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marcos/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="95%" align="center" border="0">
                            <tr>
								<input type="hidden" name="setOrder" id="setOrder" value="<% =setOrder %>">
								<td width="15%" align="right"><% = GF_TRADUCIR("Cod. Parametro") %>:</td>
								<td width="20%">
									<input type="text"  id="cdParam" name="cdParam" value="<%=cdParam%>">
								</td>
								<td width="13%" align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
								<td width="20%">
									<input type="text"  id="nomParam" name="nomParam" value="<%=nomParam%>">
								</td>
                            </tr>      
							<tr>
								<td width="15%" align="right">
									<% = GF_TRADUCIR("Puesto") %>
								</td>
								<td>
									<select style="z-index:-1;" name="idPuesto" id="idPuesto">
										<option SELECTED value="0">-<% =GF_TRADUCIR("Seleccione")%>-
										<%Set rsPue = leerPuestos(pto)										
										while (not rsPue.eof)
										if rsPue("IDPUESTO") = idpuesto then
											mySelected = "SELECTED"
										else
											if idpuesto = 0 then	idpuesto = rsPue("IDPUESTO")
											mySelected = ""
										end if
										%>										
											<option value="<% =rsPue("IDPUESTO")%>" <% =mySelected %>><% =rsPue("DSPUESTO") %></option>
											
											<%rsPue.MoveNext()
										wend%>
									</select>
								</td>
							</tr>
							<tr>															
								<td colspan="4"  align="center"><input type="submit" value="Buscar" id="submit1" name="submit1" onclick='submitInfo();'></td>								
							</tr>		
								
                     </table>
	           </td>
	           <td height="100%"><img src="images/marcos/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="images/marcos/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="images/marcos/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="images/marcos/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	</div> 
		
	<br>
	
	<br>
	<!--*********************************************************************************-->
	<table class="reg_header" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">	
	<!--**************************PAGINACION*****************************************-->
	<% 	if (not rsPar.eof) then %>
		<tr><td colspan="10"><div id="paginacion"></div></td></tr>				
	
				<!--*********************** CABECERA ******************************-->
		<tr>
			<td class="reg_header_nav" width="20%" style="text-align: center">
				<img src="images\arrow_up_12x12.gif" onclick='setOrder("A.CDPARAMETRO","asc")' style="cursor:pointer" title="Ascendente">
					&nbsp <%=GF_Traducir("Cod. Parametro")%>&nbsp 
				<img src="images\arrow_down_12x12.gif" onclick='setOrder("A.CDPARAMETRO","desc")' style="cursor:pointer" title="Descendente">
			</td>
			<td class="reg_header_nav" align="center">
				<img src="images\arrow_up_12x12.gif" onclick='setOrder("A.DSPARAMETRO","asc")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Descripcion")%>
				<img src="images\arrow_down_12x12.gif" onclick='setOrder("A.DSPARAMETRO","desc")' style="cursor:pointer" title="Descendente">
			</td>
			
			<td class="reg_header_nav" align="center">
				<img src="images\arrow_up_12x12.gif" onclick='setOrder("A.VLPARAMETRO","asc")' style="cursor:pointer" title="Ascendente">
					<%=GF_Traducir("Valor")%>
				<img src="images\arrow_down_12x12.gif" onclick='setOrder("A.VLPARAMETRO","desc")' style="cursor:pointer" title="Descendente">
			</td>
			<td class="reg_header_nav" align="center">				
				<%=GF_Traducir("Puesto")%>			
			</td>					
			<%if(esmodificable <> TASK_PARAM_AUDIT)then%>
				<%if(esmodificable <> TASK_PARAM_USER)then%>
					<td class="reg_header_nav" align="center"><%=GF_Traducir(".")%></td>		
					<td class="reg_header_nav" align="center"><%=GF_Traducir(".")%></td>
				<%else%>				
					<td class="reg_header_nav" align="center"><%=GF_Traducir(".")%></td>
				<%end if%>				
			<%end if%>							
		</tr>

			<!--*********************** RESULTADOS ******************************-->
		<% 		
				while ((not rsPar.eof) and (CInt(reg) < CInt(mostrar)))				
					reg = reg + 1
					%>
					<tr>
						<td class="reg_header_navdos" align="center"><%=rsPar("CDPARAMETRO")%></td>
						
						<td class="reg_header_navdos" align="center"><%=rsPar("DSPARAMETRO")%></td>		
						
						<td class="reg_header_navdos" align="center"><%=rsPar("VLPARAMETRO")%></td>
						<%						
						if(tieneParametrosExtra(rsPar("CDPARAMETRO"), pto))then
							set v_Idpuesto = traerParametrosEditables(rsPar("CDPARAMETRO"), pto)%>							
							<td class="reg_header_navdos" align="center"><%=obtenerNombrePuesto(v_Idpuesto("PUESTO"),pto)%></td>
						<%else
							v_Idpuesto = SIN_PUESTO_OCUPADO%>	
							<td class="reg_header_navdos" align="center"></td>
						<%end if%>										
						
						<%if(esmodificable = TASK_PARAM_ADMIN)then
						'BORRAR: si es SOLO administrador puede dar de baja un parametro
						%>
							<td class="reg_header_navdos" align="center" onclick="javascript:eliminarParametro('<%=rsPar("CDPARAMETRO")%>')">
								<img title="Eliminar" src="images/cancel-16x16.png" style="cursor:pointer">
							</td>
							
						<%end if%>		
						<%if(validarEditables(esmodificable,pto,rsPar("CDPARAMETRO"))) then
						'EDITAR:valido para que cuando sea usuario, muestre aquellos parametros que 
						'sean modificables
						%>	
						<td class="reg_header_navdos" align="center" onclick="javascript:modificarParametro('<%=rsPar("CDPARAMETRO")%>')">						
							<img title="Editar" src="images/edit-16x16.png" style="cursor:pointer">
						</td>						
						<%else%>	
							<td class="reg_header_navdos" align="center"></td>							
						<%end if%>							
					</tr>
					<% 
					rsPar.movenext
				wend 
			else%>
			<tr class="TDNOHAY"><td colSpan="4"><% =GF_TRADUCIR("No hay informacion disponible en estos momentos") %></td></tr>		
			<%end if%>			
	</table>
	
	<INPUT TYPE="HIDDEN" ID="pto" NAME="pto" VALUE=<%=pto%>>

	<INPUT TYPE="HIDDEN" ID="accion" NAME="accion" VALUE=<%=accion%>>
</form>	
</BODY>
</HTML>
