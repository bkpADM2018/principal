<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->

<%
Call initAccessInfo(RES_INV_SM)
Dim txtCdEquipo, txtDsEquipo, myWhere, strSQL, rs, conn, ref, flagMaster
Dim idEquipoActivo, idEquipo, cdEquipo, dsEquipo, idDivision, dsDivision, idSector, dsSector, idUbicacion, dsUbicacion, cdActivoFijo

idEquipo = GF_PARAMETROS7("idEquipo", 0 ,6)
idEquipoActivo = GF_PARAMETROS7("idEquipoActivo", 0 ,6)
idGrupo = GF_PARAMETROS7("idGrupo", 0 ,6)
idGrupoActivo = GF_PARAMETROS7("idGrupoActivo", 0 ,6)
accion = GF_PARAMETROS7("accion", "" ,6) 
ref = GF_PARAMETROS7("ref", "" ,6)  
'------------------------------------------------------------------------------------------------------------------------
if idEquipoActivo <> 0 then
	flagMaster = false
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_ID", idEquipoActivo)
	if not rs.eof then
		   idEquipo = rs("IDEQUIPMENT")
		   cdEquipo = rs("CDEQUIPMENT")
		   dsEquipo = rs("DSEQUIPMENT") 
		   idDivision = rs("IDDIVISION") 
		   dsDivision = rs("DSDIVISION") 
		   idSector = rs("IDSECTOR") 
		   dsSector = rs("DSSECTOR") 
		   if tipoOperacion = "M" or tipoOperacion = "A" then
				cdActivacion = trim(right(rs("CDACTIVATION"),len(rs("CDACTIVATION"))-2))
		   else	
		   		cdActivacion = trim(rs("CDACTIVATION"))
		   end if
   		   dsActivacion = trim(rs("DSACTIVATION")) 
		   cdActivoFijo = trim(rs("CDACTIVECODE"))
		end if
else	
	flagMaster = true
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_GET_BY_PARAMETERS", idEquipo & "|| || || " & ESTADO_ACTIVO & " || ")
	if not rs.eof then
		   idEquipo = rs("IDEQUIPMENT")
		   cdEquipo = rs("CDEQUIPMENT")
		   dsEquipo = rs("DSEQUIPMENT") 
	end if	   
end if	
%>

<html>
<head>
<title>Sistema de Mantenimiento - Despiece</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">
<link rel="stylesheet" href="css/MagicSearch.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/MagicSearchObj.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>

<script type="text/javascript">
	var ch = new channel();	
	subIndice = 1;
        function addComponent(pIdEquipo, pIdEquipoActivo) {
            document.getElementById("componentsToSave").value = document.getElementById("componentsToSave").value + '||' + pIdEquipo + '|' + pIdEquipoActivo + '|component_' + subIndice; 
            var className = "thicon";
            $("#COMPONENT_TABLE")
				.find('tfoot:last')
                    .append($('<tr>')
                        .addClass(className)
                        .append($('<td>'))
                        .append($('<td>'))
                        .append($('<td align=left>')
                            .append($('<input type=\"text\" id=\"component_' + subIndice + '\" name=\"text\">')
                                .attr('size', 150)
                            )
                        )
                        .append($('<td>'))
                    );
           $('table#COMPONENT_TABLE tr:last').after($('#ACTION_ROW'));     
           subIndice = subIndice + 1;   
        }	
       
        function addSubComponent(pIdEquipo, pIdEquipoActivo, pIdGrupo) {
			document.getElementById("idGrupoActivo").value = pIdGrupo;
			document.getElementById("subComponentsToSave").value = document.getElementById("subComponentsToSave").value + '||' + pIdEquipo + '|' + pIdEquipoActivo + '|' + pIdGrupo + '|subComponent_' + subIndice; 
            $("#TABLE_ID_" + pIdGrupo)
                .find('tfoot')
                    .append($('<tr>')
                        .append($('<td>'))
                        .append($('<td align=left>')
                            .append($('<input type=\"text\" id=\"subComponent_' + subIndice + '\" name=\"text\">')
                                .attr('size', 150)
                            )
                        )
                        .append($('<td>'))
                        
                    );
           $('table#TABLE_ID_' + pIdGrupo + ' tr:last').after($('#ACTION_ROW_' + pIdGrupo));        
           subIndice = subIndice + 1;
        }	
	function saveComponent(pIdEquipo,pIdEquipoActivo, pIdGrupo, pObjDs){
		document.getElementById("idGrupoActivo").value = pIdGrupo;
		ch.bind("mantenimientoComponentesABMAJAX.asp?idEquipo=" + pIdEquipo + "&idEquipoActivo=" + pIdEquipoActivo + "&idGrupo=" + pIdGrupo + "&dsComponente=" + pObjDs.value + "&tipoOperacion=A", "saveComponent_Callback('" + pIdGrupo + "')");
		ch.send();			
	}
	function saveComponent_Callback(pIdGrupoActivo){
		if (ch.response()!='') document.getElementById("divError").innerHTML = ch.response();
	}	
	
	function updateComponent(pIdEquipo, pIdEquipoActivo, pIdComponente, pObjDs, pIdGrupo){
		document.getElementById("idGrupoActivo").value = pIdGrupo;
		ch.bind("mantenimientoComponentesABMAJAX.asp?idEquipo=" + pIdEquipo + "&idEquipoActivo=" + pIdEquipoActivo + "&idComponente=" + pIdComponente + "&idGrupo=" + pIdGrupo + "&dsComponente=" + pObjDs.value + "&tipoOperacion=M", "updateComponent_Callback('" + pIdGrupo + "')");
		ch.send();			
	}
	function updateComponent_Callback(pIdGrupoActivo){
		if (ch.response()!='') document.getElementById("divError").innerHTML = ch.response();
	}	
	function bodyOnLoad(){
		var toolBarEquipos = new Toolbar('toolBarEquipos', 5, "images/");	
		toolBarEquipos.addButtonRETURN("Volver", "irA('<%=ref%>')");
		<% if isAdminInAny then %>
			toolBarEquipos.addButtonSAVE("Guardar", "saveAll();");	
		<% end if %>
		toolBarEquipos.addButtonREFRESH("Refrescar", "submitInfo();");		
		toolBarEquipos.draw();		
		pngfix();		
 		armarListaComponentes(<%=idGrupoActivo%>);
	}
	function saveAll(){
		//Buscar componentes por guardar
		var components = document.getElementById("componentsToSave").value;
		var listOfComponents = components.split("||");
		var detailOfComponent;
		for (i=1;i<listOfComponents.length;i++){
			detailOfComponent = listOfComponents[i].split("|");
			if (document.getElementById(detailOfComponent[2]).value != '') saveComponent(detailOfComponent[0],detailOfComponent[1], 0, document.getElementById(detailOfComponent[2]));
		}
		//Buscar subcomponentes por guardar
		var subComponents = document.getElementById("subComponentsToSave").value;
		var listOfSubComponents = subComponents.split("||");
		var detailOfSubComponent;
		for (i=1;i<listOfSubComponents.length;i++){
			detailOfSubComponent = listOfSubComponents[i].split("|");
			if (document.getElementById(detailOfSubComponent[3]).value != '') saveComponent(detailOfSubComponent[0], detailOfSubComponent[1], detailOfSubComponent[2], document.getElementById(detailOfSubComponent[3]));
		}
		//Buscar componentes por actualizar
		subComponents = document.getElementById("componentsToEdit").value;
		listOfSubComponents = subComponents.split("||");
		var rtrn = 0;
		for (i=1;i<listOfSubComponents.length;i++){
			detailOfSubComponent = listOfSubComponents[i].split("|");
			if (document.getElementById(detailOfSubComponent[3]).value != '') 
				rtrn = updateComponent(detailOfSubComponent[0], detailOfSubComponent[1],detailOfSubComponent[2], document.getElementById(detailOfSubComponent[3]),detailOfSubComponent[4]);
		}		
		armarListaComponentes(document.getElementById("idGrupoActivo").value);
	}
	function submitInfo() {		
		document.getElementById("frmSel").submit();
	}
	function irA(pLink) {
		location.href = pLink;
	}	
	function armarListaComponentes(pIdGrupoActivo){
 		habilitarLoading("visible","relative")
		ch.bind("mantenimientoListaComponentesAJAX.asp?idEquipo=" + document.getElementById("idEquipo").value + "&idEquipoActivo=" + document.getElementById("idEquipoActivo").value + "&idGrupoActivo=" + pIdGrupoActivo, "armarListaComponentes_Callback()");
		ch.send();	
	}
	function armarListaComponentes_Callback(){
 		habilitarLoading("hidden","absolute")
		document.getElementById("results").innerHTML = ch.response();
	}
	
	function habilitarLoading(pVisibility, pPosition){
		document.getElementById("imgLoading").style.position = pPosition;
		document.getElementById("imgLoading").style.visibility  = pVisibility;
		document.getElementById("lblLoading").style.position = pPosition;
		document.getElementById("lblLoading").style.visibility  = pVisibility;
	}
	function showTR(pImg,pTrName){
		if (document.getElementById(pTrName).className == "troculto"){
			document.getElementById(pTrName).className = "trvisible";
			pImg.src = "images/menos.gif";
		}
		else{
			document.getElementById(pTrName).className = "troculto";
			pImg.src = "images/mas.gif";
		}
	}
	function habilitarItem(pIdComponente, pIdGrupoActivo){
			document.getElementById("componentImHab" + pIdComponente).title = "Habilitando";
			document.getElementById("componentImHab" + pIdComponente).src = "images/loader.gif"	
			ch.bind("mantenimientoComponentesABMAJAX.asp?idComponente=" + pIdComponente + "&tipoOperacion=H", "habilitarItem_Callback('" + pIdGrupoActivo + "')");
			ch.send();
	}
	function habilitarItem_Callback(pIdGrupoActivo){
		if (ch.response()==''){
			if (pIdGrupoActivo == 0){
				document.getElementById("idGrupoActivo").value = pIdGrupoActivo;			
				armarListaComponentes(pIdGrupoActivo);
				//submitInfo();
			}
			else{
				armarListaComponentes(pIdGrupoActivo);
				document.getElementById("divError").innerHTML = "";
			}		
		}else{
			document.getElementById("divError").innerHTML = ch.response();		
		}		
	}	

	function deshabilitarItem(pIdComponente, pIdGrupoActivo){
			document.getElementById("componentImHab" + pIdComponente).title = "Quitando";
			document.getElementById("componentImHab" + pIdComponente).src = "images/loader.gif";
			ch.bind("mantenimientoComponentesABMAJAX.asp?idComponente=" + pIdComponente + "&tipoOperacion=B", "deshabilitarItem_Callback('" + pIdGrupoActivo + "')");
			ch.send();	
	}
	function deshabilitarItem_Callback(pIdGrupo){
		if (ch.response()==''){
			if (pIdGrupo == 0){
				document.getElementById("idGrupoActivo").value = pIdGrupo;			
				armarListaComponentes(pIdGrupo);
				//submitInfo();
			}
			else{
				armarListaComponentes(pIdGrupo);
				document.getElementById("divError").innerHTML = "";
			}
		}else{
			document.getElementById("divError").innerHTML = ch.response();
		}		
	}	
	function editComponent(pIdEquipo, pIdEquipoActivo, pIdComponent, pObjId, pIdGrupo){
		document.getElementById("componentsToEdit").value = document.getElementById("componentsToEdit").value + '||' + pIdEquipo + '|' + pIdEquipoActivo + '|' + pIdComponent + '|' + pObjId + '|' + pIdGrupo + '|M'; 		
	}	
	function enableItem(pIdComponente){
			document.getElementById("componentDs" + pIdComponente).type = "text";
			document.getElementById("componentFn" + pIdComponente).style.visibility = "hidden";
			document.getElementById("componentFn" + pIdComponente).style.position = "absolute";
	}
	function disableItem(pIdComponente){
			document.getElementById("componentDs" + pIdComponente).type = "hidden";
			document.getElementById("componentFn" + pIdComponente).innerHTML = "<i>" + document.getElementById("componentDs" + pIdComponente).value + "</i>"; 
			document.getElementById("componentFn" + pIdComponente).style.visibility = "visible";
			document.getElementById("componentFn" + pIdComponente).style.position = "relative";		
	}
	function castBool(str) {
	    if (str.toLowerCase() === 'true') {
	        return true;
	    } else if (str.toLowerCase() === 'false') {
	        return false;
	    }
	    return ERROR;
	}	
	function AbrirUploader(pId, pCd, pIdComponent, pIdSubcomponent){
		var puw = new winPopUp('popupEquipo','mantenimientoEquipoFiles.asp?idEquipo=' + pId + '&idComponent=' + pIdComponent + '&idSubComponent=' + pIdSubcomponent,'780','350','Archivos del Master: ' + pCd, "");
	}	
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<div id="toolBarEquipos"></div>
<form id="frmSel" name="frmSel">

<% if flagMaster then %>
	<div class="tableaside size100"> 
		<h3><%=GF_Traducir("Datos del Master")%></h3>
	  
	    <div class="tableasidecontent">
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("ID") %> </div>
	        <div class="col26"> <% =idEquipo%> </div>
	       
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Código") %> </div>
	        <div class="col26"> <% =cdEquipo %> </div>
	        
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripción") %> </div>
	        <div class="col36"> <% =dsEquipo%> </div>
		</div>
	</div>
	<div class="col66"></div>	
<% else %>
	<div class="tableaside size100"> 
		<h3><%=GF_Traducir("Datos de Activación")%></h3>
	  
	    <div class="tableasidecontent">
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Código") %> </div>
	        <div class="col26">
	        		<%
					if tipoOperacion = "M" or tipoOperacion = "A" then
						Response.Write "<b>" & left(cdEquipo,2) & " </b>"
						%>
						<input type="text" size="7" maxlength="8" id="cdActivacion" name="cdActivacion" value="<%=cdActivacion%>">
					<%else
						Response.Write cdActivacion
						%>
						<input type="hidden" name="cdActivacion" id="cdActivacion" value="<%=cdActivacion%>">
					<%end if%>
					<input type="hidden" name="cdActivacionPrefijo" id="cdActivacionPrefijo" value="<%=left(cdEquipo,2)%>">
			</div>
	       
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripción") %> </div>
	        <div class="col26"> 
	   				<%if tipoOperacion = "M" or tipoOperacion = "A" then%>
						<input type="text" size="30" maxlength="100" id="dsActivacion" name="dsActivacion" value="<%=dsActivacion%>">
					<%else
						Response.Write dsActivacion
						%>
						<input type="hidden" name="dsActivacion" id="dsActivacion" value="<%=dsActivacion%>">
					<%end if%>
			</div>
	        
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("División") %> </div>
	        <div class="col26"> 
					<%if tipoOperacion = "M" or tipoOperacion = "A" then
						%>
						<select name="idDivision" id="idDivision">
							<%
							
							call executeProcedureDb(DBSITE_SQL_INTRA, rsSel, "TBLDIVISIONES_GET_BY_LIST", getListaCargosAdmin())
							while not rsSel.eof
								if not isAuditor(rsSel("IDDIVISION")) then
								%>	
									<option value="<%=rsSel("IDDIVISION")%>" <%if cint(idDivision)=cint(rsSel("IDDIVISION")) then Response.Write "Selected"%>><%=rsSel("DSDIVISION")%></option>
								<%	
								end if
								rsSel.movenext
							wend	
							%>
						</select>
					<%else
						Response.Write dsDivision
						%>
						<input type="hidden" name="idDivision" id="idDivision" value="<%=idDivision%>">
					<%end if%>
	        </div>
	        
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Sector") %> </div>
	        <div class="col26"> 
					<%if tipoOperacion = "M" or tipoOperacion = "A" then%>
						<select name="idSector" id="idSector">
							<%
							call executeProcedureDb(DBSITE_SQL_INTRA, rsSel, "TBLSMSECTOR_GET", "")
							while not rsSel.eof
								%>	
									<option value="<%=rsSel("IDSECTOR")%>" <%if cint(idSector)=cint(rsSel("IDSECTOR")) then Response.Write "Selected"%>><%=rsSel("DSSECTOR")%></option>
								<%	
								rsSel.movenext
							wend	
							%>
						</select>				
					<%else
						Response.Write dsSector
						%>
						<input type="hidden" name="idSector" id="idSector" value="<%=idSector%>">
					<%end if%>
	        </div>
	        
	        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Activo Fijo") %> </div>
	        <div class="col26"> 
					<%if tipoOperacion = "M" or tipoOperacion = "A" then%>
						<input type="text" size="9" maxlength="9" id="cdActivoFijo" name="cdActivoFijo" value="<%=cdActivoFijo%>">
					<%else
						Response.Write cdActivoFijo
						%>
						<input type="hidden" name="cdActivoFijo" id="cdActivoFijo" value="<%=cdActivoFijo%>">
					<%end if%>
			</div>
	        
		</div>
	</div>
<% end if %>
<div class="col66"></div>
<table align="center" width="90%" border="0">
	<tr>
		<td align="center">
			<img style="position:absolute;visibility:hidden;" id="imgLoading" src="images/Loading4.gif">
			<div style="position:absolute;visibility:hidden;" id="lblLoading"><b><br>Aguarde por favor...</b></div>
					
		</td>
	</tr>
</table>      	

<div id=results></div>
<div id="divError"></div>	    
<input type="HIDDEN" name="idGrupoActivo" id="idGrupoActivo" value="">	
<input type="HIDDEN" name="idEquipo" id="idEquipo" value="<%=idEquipo%>">
<input type="HIDDEN" name="idEquipoActivo" id="idEquipoActivo" value="<%=idEquipoActivo%>">
	
	
	
<input type="HIDDEN" name="ref" id="ref" value="<%=ref%>">	
<input type="HIDDEN" name="accion" id="accion">
</form>	
</body>
</html>