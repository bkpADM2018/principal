<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<%
Call initAccessInfo(RES_INV_SM)
Dim strSQL, rs, conn, idPreviousComponentGroup, cont, modoEdicion, myPreviousKey, mySplittedKey, mySplittedItem, myId, myDs
dim compDicc
idEquipo = GF_PARAMETROS7("idEquipo", 0 ,6)
idEquipoActivo = GF_PARAMETROS7("idEquipoActivo", 0 ,6)
idGrupoActivo = GF_PARAMETROS7("idGrupoActivo", 0 ,6)
modoEdicion = isAdminInAny()
set compDicc = setComponentsDicc(idEquipo, idEquipoActivo)

'-------------------------------------------------------------------------------------
function setComponentsDicc(pIdEquipo, pIdEquipoActivo)
dim dicc, rsComp, myValue
call executeProcedureDb(DBSITE_SQL_INTRA, rsComp, "TBLSMCOMPONENT_GET_LIST_BY_IDEQUIPMENT", pIdEquipo & "||" & pIdEquipoActivo)
Set dicc = server.CreateObject("Scripting.Dictionary")
while not rsComp.eof
	if rsComp("IDCOMPONENTGROUP") = 0 then 
		call dicc.Add (rsComp("IDCOMPONENT") & "|" & rsComp("DSCOMPONENT") & "|" & rsComp("CDSTATE"),"")
	else
		if dicc.Exists(rsComp("IDCOMPONENTGROUP") & "|" & rsComp("DSCOMPONENTGROUP") & "|" & rsComp("CDSTATECOMPONENTGROUP")) = true then 'Si ya existe el grupo
			dicc(rsComp("IDCOMPONENTGROUP") & "|" & rsComp("DSCOMPONENTGROUP") & "|" & rsComp("CDSTATECOMPONENTGROUP")) = dicc(rsComp("IDCOMPONENTGROUP") & "|" & rsComp("DSCOMPONENTGROUP") & "|" & rsComp("CDSTATECOMPONENTGROUP")) & "||" & rsComp("IDCOMPONENT") & "|" & rsComp("DSCOMPONENT") & "|" & rsComp("CDSTATE")
		else 'No existe grupo
			myValue = ""
			myValue = "||" & rsComp("IDCOMPONENT") & "|" & rsComp("DSCOMPONENT") & "|" & rsComp("CDSTATE")
			call dicc.Add (rsComp("IDCOMPONENTGROUP") & "|" & rsComp("DSCOMPONENTGROUP") & "|" & rsComp("CDSTATECOMPONENTGROUP"),myValue)
		end if	
	end if	
	rsComp.movenext
wend
set setComponentsDicc = dicc
end function
%>
<table class="datagrid datagridlv1" width="100%" align="center" id="COMPONENT_TABLE" align="center">
	<thead>
		<tr>
			<th class="thicon"><%=GF_Traducir("")%></th>
			<th class="thicon" align="center"><%=GF_Traducir("ID")%></th>
			<th><%=GF_Traducir("Descripción")%></th>
			<th class="thiconac" align="center"><%=GF_Traducir("-")%></th>
		</tr>	
	</thead>
	<%
	if (compDicc.Count > 0) then 'Hay componentes Cargados
		myKey=compDicc.Keys 'Grupos
		myItem=compDicc.Items 'Lista de Componentes
		for index = 0 To compDicc.Count -1
			if myPreviousKey <> myKey(index) then 'Cambio el grupo
				myPreviousKey = myKey(index)
				mySplittedKey  = split(myKey(index),"|") 'Obtengo datos del grupo
				myIdGrupo = mySplittedKey(0)
				myDsGrupo = mySplittedKey(1)
				myStateGrupo = mySplittedKey(2)
				
				myOnClick = "showTR(this,'TR_ID_" & myIdGrupo & "')"
				if cint(idGrupoActivo) = cint(myIdGrupo) then 
					myImage = "images/menos.gif"
					myClassTR = "trvisible"
				else
					myImage = "images/mas.gif"
					myClassTR = "troculto"
				end if
				%>
				<tbody>  
   				<tr>	
					<td class="thicon">
						<img src="<%=myImage%>" onclick="<%=myOnClick%>" style="cursor:pointer;">
					</td>
					<%
					myStyleBaja = ""
					if cint(myStateGrupo) = cint(ESTADO_BAJA) then myStyleBaja = "color:gray;font-style:italic;"
					%>
					<td align="center">
						<font  style="<%=myStyleBaja%>"><%=myIdGrupo%></font>
					</td>
					<td>	
						<div id="componentFn<%=myIdGrupo%>"><font  style="<%=myStyleBaja%>"><%=myDsGrupo%></font></div>
						<input size="100" maxlength="100" type="hidden" name="componentDs<%=myIdGrupo%>" id="componentDs<%=myIdGrupo%>" value="<%=myDsGrupo%>">
					</td>
					<!--<td width="2%">&nbsp;</td>-->
					<%
					if modoEdicion then
						if cint(myStateGrupo) = cint(ESTADO_BAJA) then 
							%>
							<td class="thiconac" width="80">
								<img id="componentImHab<%=myIdGrupo%>" onclick="habilitarItem('<%=myIdGrupo%>','0')" title="Habilitar" style="cursor:pointer;" src="images/checkmark-16.png"></td>
							</td>	
							<%		
						else
							%>
							<td class="thiconac" width="80">	
								<img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivos" onclick="AbrirUploader('<% =idEquipo %>','','<%=myIdGrupo%>','0')">
								<img id="componentIm<%=myIdGrupo%>" onclick="enableItem('<%=myIdGrupo%>'); editComponent('<%=idEquipo%>', '<%=idEquipoActivo%>', '<%=myIdGrupo%>','componentDs<%=myIdGrupo%>', '0')" title="Editar" style="cursor:pointer;" src="images/edit-16.png">
								<img id="componentImHab<%=myIdGrupo%>" onclick="deshabilitarItem('<%=myIdGrupo%>','0')" title="Deshabilitar" style="cursor:pointer;" src="images/cross-16.png">
							</td>
							<%  
						end if
					end if
					%>
				</tr>
				</tbody>  	
				<%
			end if
			mySplittedItem  = split(myItem(index),"||") 'Obtener lista de sub componetes asociados
			%>
			
			<tr class="<%=myClassTR%>" id="TR_ID_<%=myIdGrupo%>">
				<td colspan="4">
					<table class="datagridlv2" width="90%" align="center" border="0" id="TABLE_ID_<%=myIdGrupo%>" style="<%=myStyle%>">
						<thead>
							<tr>
								<th class="thicon"><%=GF_Traducir("ID")%></th>
								<th><%=GF_Traducir("Descripción")%></th>
								<th class="thiconac"><%=GF_Traducir("-")%></th>
							</tr>	
						</thead>	
					<%	
					for j=1 to ubound(mySplittedItem)
						mySplittedSubItem = split(mySplittedItem(j),"|") 'Datos de sub-compoennte
						myId = mySplittedSubItem(0)
						myDs = mySplittedSubItem(1)
						myState = mySplittedSubItem(2)
						
						myStyleBaja = ""
						if cint(myState) = cint(ESTADO_BAJA) then myStyleBaja = "color:gray;font-style:italic;"
						%>	
					    <tbody>    
   							<tr>
								<td align="center">
									<font  style="<%=myStyleBaja%>"><%=myId%></font>
								</td>							
								<td>	
									<div id="componentFn<%=myId%>"><font style="<%=myStyleBaja%>"><%=myDs%></font></div>
									<input size="100" maxlength="100" type="hidden" name="componentDs<%=myId%>" id="componentDs<%=myId%>" value="<%=myDs%>">
								</td>
   						
								<%
								if modoEdicion and myStateGrupo <> 2 then
									if cint(myState) = cint(ESTADO_BAJA) then 
										%>
										<td class="thiconac" width="80">
											<img id="componentImHab<%=myId%>" onclick="habilitarItem('<%=myId%>','<%=myIdGrupo%>')" title="Habilitar" style="cursor:pointer;" src="images/checkmark-16.png">
										</td>
										<%
									else
										%>
										<td class="thiconac" width="80">	
											<img src="images/adjunto-16.png" style="cursor: pointer" title="Adjuntar Archivos" onclick="AbrirUploader('<% =idEquipo %>','','0','<%=myId%>')"> 
											<img id="componentIm<%=myId%>" onclick="enableItem('<%=myId%>'); editComponent('<%=idEquipo%>','<%=idEquipoActivo%>', '<%=myId%>','componentDs<%=myId%>', '<%=myIdGrupo%>')" title="Editar" style="cursor:pointer;" src="images/edit-16.png">
											<img id="componentImHab<%=myId%>" onclick="deshabilitarItem('<%=myId%>','<%=myIdGrupo%>')" title="Deshabilitar" style="cursor:pointer;" src="images/cross-16.png">
										</td>
										<%  
									end if
								else
								%>
									<td class="thiconac" width="80">&nbsp;</td>	
								<%
								end if
								%>
							</tr>
						</tbody>	
						<%
					next
					if cint(myStateGrupo) <> cint(ESTADO_BAJA) and modoEdicion then 
					%>
					<tfoot>
						<tr id="ACTION_ROW_<%=myIdGrupo%>">
							<td colspan="3" align="right">
						        <a class="btnmore" href="javascript:addSubComponent('<%=idEquipo%>','<%=idEquipoActivo%>','<%=myIdGrupo%>')"><img src="images/plus-16.png"><%=GF_Traducir("Agregar SubComponente")%></a>
						    </td>
						</tr>					
					</tfoot>	
					<%end if%>	
					</table>
				</td>
			</tr>
			<%				
		next
		if modoEdicion then
		%>
		<tfoot id="last">
		<tr id="ACTION_ROW">
			<td colspan="4" align="right" >
				<a class="btnmore" href="javascript:addComponent('<%=idEquipo%>','<%=idEquipoActivo%>')"><img src="images/plus-16.png"><%=GF_Traducir("Agregar Componente")%></a>
		    </td>
		</tr>	
		</tfoot>	
		<%
		end if
	else
		if modoEdicion then
		%>
		<tfoot id="last">
		<tr id="ACTION_ROW">
			<td colspan="4" align="right" >
		        <a class="btnmore" href="javascript:addComponent('<%=idEquipo%>','<%=idEquipoActivo%>')"><img src="images/plus-16.png"><%=GF_Traducir("Agregar Componente")%></a>
		    </td>
		</tr>	
		</tfoot>
		<%
		end if
	end if
%>	
<%= showMessages%>
<input type="hidden" id="componentsToSave" value="" size="150">
<input type="hidden" id="subComponentsToSave" value="" size="150">
<input type="hidden" id="componentsToEdit" value="" size="150">
</table>
