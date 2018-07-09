<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosUser.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<%
'ProcedimientoControl "GVPERMISOS"
dim cdProducto, cdCliente, cdCamionesDe, cdCosecha, tipo, cdAviso, cdAvisoAnt, cdBuque, dsBuque, kilos,cdProductoAnt, kilosPrevios, mySelected, kilosSobrantes, kilosInformados, msjEmbarcados
dim myTableHTML, dicErr, accion, mySaveText, puerto,  kilosFinal, controlsState, cargaExclusivaCosecha, secuencia, kilosReales, listOfProducts, listOfHarvest
dim rskilosXcosecha 

Set dicErr = Server.CreateObject("Scripting.Dictionary")
puerto = Request("Pto")
cdProducto = GF_Parametros7("cdProducto", "", 6)
if cdProducto <> "" then	
	strSQL="Select CDCOSECHA from COSECHAS where CDCOSECHA <> 0 and CDPRODUCTO = " & cdProducto & " order by CDCOSECHA desc"
	Call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
	if (not rs.eof) then
		listOfHarvest = rs.GetString(2,,, ";")
		listOfHarvest = left(listOfHarvest, Len(listOfHarvest)-1) 'Saco el último ';'
	end if
end if
cdCliente = GF_Parametros7("cdCliente", 0, 6)
cdCosecha = GF_Parametros7("cdCosecha", 0, 6)
if cdproducto <> 0 then	
	accion = ACCION_SUBMITIR	
	set rskilosXcosecha = getKilosDisponibles(cdCliente,cdproducto)	
end if	
if cdproducto = 0 then setError(PRODUCTO_REQUERIDO)
'----------------------------------------------------------------------------------------
Function getKilosDisponibles(pcdCliente, pcdProducto)
dim strSQL,rs
If pcdCliente <> 0 Then 
	auxWhere1 = " AND CDCLIENTE=" & pcdCliente
	auxWhere2 = " AND HCD.CDCLIENTE=" & pcdCliente
end if	
fechaInicio = "2010-03-01"
strSQL = "SELECT  distinct cdcosecha as COSECHA ,SUM(KILOSNETOS) AS KILOS_DISPONIBLES FROM ( "
strSQL = strSQL &	"SELECT TG.DTCONTABLE, TG.IDCAMION, TG.CDCHAPACAMION, TG.CDCHAPAACOPLADO, TG.CDCOSECHA, " 
strSQL = strSQL &	"	TG.CARTAPORTE, TG.CTG, TG.DSCLIENTE, TG.DSPRODUCTO, CASE WHEN TG.KILOSCARGADOS IS NULL THEN TG.KILOSNETOS ELSE TG.KILOSNETOS-TG.KILOSCARGADOS END AS KILOSNETOS " 
strSQL = strSQL &	"	  FROM " 
strSQL = strSQL &	"	( "
strSQL = strSQL &	"	    SELECT HCD.DTCONTABLE, HCD.IDCAMION, HC.CDCHAPACAMION, HC.CDCHAPAACOPLADO, HCD.CDCOSECHA, " 
strSQL = strSQL &	"	        RTRIM(HCD.NUCARTAPORTE) + RTRIM(HCD.NUCTAPTEDIG) AS CARTAPORTE, HCD.CTG, C.DSCLIENTE, P.DSPRODUCTO, " 
'Para SQL SERVER - strSQL = strSQL &	"	        RTRIM(RTRIM(HCD.NUCARTAPORTE)+''+RTRIM(HCD.NUCTAPTEDIG)) AS CARTAPORTE, HCD.CTG, C.DSCLIENTE, P.DSPRODUCTO, " 
strSQL = strSQL &	"	        ( "
strSQL = strSQL &	"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) " 
strSQL = strSQL &	"	            -  "
strSQL = strSQL &	"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2))  " 
strSQL = strSQL &	"	            -  "
strSQL = strSQL &	"	            ( SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION)) "
strSQL = strSQL &	"	        ) KILOSNETOS , EMBARCADOS.KILOSCARGADOS "
strSQL = strSQL &	"	    FROM HCAMIONESDESCARGA HCD "
strSQL = strSQL &	"	        LEFT JOIN  "
strSQL = strSQL &	"	            (SELECT IDCAMION, DTCONTABLE, SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS GROUP BY IDCAMION, DTCONTABLE) "
strSQL = strSQL &	"	                EMBARCADOS ON HCD.IDCAMION = EMBARCADOS.IDCAMION AND HCD.DTCONTABLE=EMBARCADOS.DTCONTABLE "
strSQL = strSQL &	"	        LEFT JOIN HCAMIONES HC ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE=HCD.DTCONTABLE  "
strSQL = strSQL &	"	        LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO  "
strSQL = strSQL &	"	        LEFT JOIN CLIENTES C ON C.CDCLIENTE = HCD.CDCLIENTE "
strSQL = strSQL &	" WHERE HCD.DTCONTABLE >='" & fechaInicio & "'"
strSQL = strSQL &    auxWhere2
strSQL = strSQL &   " AND HC.CDPRODUCTO = " & pcdProducto
strSQL = strSQL &	" AND HC.CDESTADO IN (6,8) "
strSQL = strSQL &	" ) TG)TA WHERE KILOSNETOS> 0 "
strSQL = strSQL &	"  group by cdcosecha "
strSQL = strSQL &	"  order by cdcosecha desc "

call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
set getKilosDisponibles = rs
end function
'----------------------------------------------------------------------------------------
Function armarFormatoCosecha(cdCosecha)
	armarFormatoCosecha = left(cdCosecha,4) & "-" & right(cdCosecha,4)
End function
'----------------------------------------------------------------------------------------
%>
<HTML>
<HEAD>
   <TITLE>CTGs Disponibles</TITLE>
</HEAD>
<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">
<link rel="stylesheet" href="../../css/ActisaIntra-1.css" type="text/css">
<script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

<script type="text/javascript">
	var ch = new channel();	
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function bodyOnLoad(){
		refPopUpDisponibilidad = getObjPopUp('popupDisponibles');
	}
	function updateHarvests(pObj){
		ch.bind("getListOfHarvest_AJAX.asp?cdProducto=" + pObj.options[pObj.selectedIndex].value + "&Pto=<%=puerto%>", "updateHarvests_Callback()");
		ch.send();			
	}
	function updateHarvests_Callback(){
		var newOption;
		var refSelect;
		var index;
		var listOfHarvests;
		
		listOfHarvests = ch.response();
		refSelect = document.getElementById("cdCosecha"); 
		refSelect.options.length=0;     		
		listOfHarvests = listOfHarvests.split(";");
		for(index=0;index<listOfHarvests.length;index++){
			newOption=document.createElement("OPTION");
			newOption.value=listOfHarvests[index];
			newOption.text=listOfHarvests[index];
			refSelect.options.add (newOption);
		}
	}
		
	function submitInfo(accion){	
		if (pAccion=='SEARCH'){
			if (document.getElementById("results")){
				document.getElementById("results").innerHTML = "";
			}
			if (document.getElementById("loading")){
				document.getElementById("loading").style.visibility = "visible";
				document.getElementById("loading").style.position = "relative";
			}
		}
		document.getElementById("accion").value = pAccion;
		document.form1.submit();	
	}
	function abrirDetalleCosecha(cosecha,kilos,pcdProducto,pcdCliente)
	{				 			
		window.open("CamionesXCosechaCTGEmbarcados.asp?cosecha="+cosecha+"&kilos="+kilos+"&producto="+pcdProducto+"&cliente="+pcdCliente+"&accion=1&pto=<%=puerto%>");		
	}
	function abrirDetalleCosecha_callBack()
	{
		alert("abrirDetalleCosecha_callBack")
	}

</script>	
<BODY onload="bodyOnLoad()">
<br>
<FORM action="CtgDisponibles.asp" method=POST id=form1 name=form1>

	<table id="tblBusqueda" width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="../images/marcos/marco_r1_c1.gif"></td>
           <td width="25%"><img src="../images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="../images/marcos/marco_r1_c3.gif"></td>
           <td width="75%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="../images/marcos/marco_r2_c1.gif"></td>
           <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
           <td width="8"><img src="../images/marcos/marco_r2_c3.gif"></td>
           <td></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="../images/marcos/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td><img src="../images/marcos/marco_c_s_d.gif" height="8" width="8"></td>
           <td><img src="../images/marcos/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="../images/marcos/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="../images/marcos/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="100%" align="center" border="0">
							<tr>
								<td align="right"><% = GF_TRADUCIR("Producto") %>:</td>
								<td>
									<select onchange="updateHarvests(this);" style="z-index:-1;" id="cdProducto" name="cdProducto" <%=controlsState%>>
										<option value="0" SELECTED>- Seleccione -</option>
										<%
										strSQL = "SELECT CDPRODUCTO, DSPRODUCTO FROM dbo.PRODUCTOS ORDER BY DSPRODUCTO"
										Call executeQueryDb(puerto, rs, "OPEN", strSQL)
										while not rs.eof 
											if cint(cdProducto) = cint(rs("CDPRODUCTO")) then
												mySelected = "SELECTED"
											else
												mySelected = ""
											end if	
												%>
												<option value="<%=rs("CDPRODUCTO")%>" <%=mySelected%>><%=rs("DSPRODUCTO")%></option>
												<%			
											rs.movenext
										wend	
										%>							
									</select>
								</td>
							</tr>
							
							<tr>	
								<td align="right"><% = GF_TRADUCIR("Camiones de") %>:</td>
								<td>
									<select style="z-index:-1;" name="cdCliente" <%=controlsState%>>
										<option value="0" ><%=GF_Traducir("Cualquiera...")%></option>
										<%
										strSQL = "SELECT CDCLIENTE, DSCLIENTE FROM dbo.CLIENTES ORDER BY DSCLIENTE"
										Call executeQueryDb(puerto, rs, "OPEN", strSQL)
										while not rs.eof 
											if cdCliente = rs("CDCLIENTE") then
												mySelected = "SELECTED"
											else
												mySelected = ""
											end if												
												%>
												<option value="<%=rs("CDCLIENTE")%>" <%=mySelected%>><%=rs("DSCLIENTE")%></option>
												<%			
											rs.movenext
										wend	%>							
									</select>
								</td>					
							</tr>
                            <tr>
								<td colspan="2" align="center">
									<input type="SUBMIT" value="Buscar..." id=cmdSearch name=cmdSearch onclick="submitInfo('SEARCH');">
								</td>		
                            </tr>								                            
                     </table>
	           </td>
	           <td height="100%"><img src="../images/marcos/marco_r2_c3.gif" width="8" height="100%"></td>
	       </tr>
	       <tr>
	           <td width="8"><img src="../images/marcos/marco_r3_c1.gif"></td>
	           <td width="100%" align=center colspan="3"><img src="../images/marcos/marco_r3_c2.gif" width="100%" height="8"></td>
	           <td width="8"><img src="../images/marcos/marco_r3_c3.gif"></td>
	       </tr>
	</table>
	<br>
	<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
		<tr>
			<td colspan=3>
			<%
			if hayError() then 
				call showErrors()
			end if
			%>
			</td>
		</tr>	
	</table>			

    <INPUT type="hidden" id="Pto" name="Pto" value=<%= Request("Pto")%>>
    <INPUT type="hidden" id="accion" name="accion">
    <INPUT type="hidden" id="tipo" name="tipo" value=<%=tipo%>>
	<%
	Dim TotalKilosCosecha
	if(accion = ACCION_SUBMITIR)then 
		if(rskilosXcosecha.eof)then%> 
			<table id="results" border=0 align="center" width="85%" class="reg_Header" cellpadding=1 cellspacing=0>
				<tr>
					<td align="center"><% =GF_TRADUCIR("No se encontraron resultados")%></td>				
				</tr>
			</table>	
		<%else%>
			<table id="results" align="center" width="75%" class="reg_Header" cellspacing="1" cellpadding="2">			
				<tr class="reg_Header_nav">
					<td align="center" width="45%"><% =GF_TRADUCIR("Cosecha")%></td>
					<td align="center" width="45%"><% =GF_TRADUCIR("Kilos")%></td>					
					<td align="center" ></td>					
				</tr>	
				
				<%while not rskilosXcosecha.eof%>
				<tr class="reg_Header_navdos">
					<td align="center">
						<%=armarFormatoCosecha(rskilosXcosecha("COSECHA"))%>
					</td>
					<td align="right">
						<%=GF_EDIT_DECIMALS(cdbl(rskilosXcosecha("KILOS_DISPONIBLES"))*100,2)%>
					</td>
					<td align='right'>
						<img style='cursor:pointer;' onclick="abrirDetalleCosecha(<%=rskilosXcosecha("COSECHA")%>,<%=rskilosXcosecha("KILOS_DISPONIBLES")%>,<%=cdProducto%>,<%=cdCliente%>)" title='Verdetalle' src='../images/see_all-16x16.png'>
					</td>
				</tr>
				<%TotalKilosCosecha = TotalKilosCosecha + cdbl(rskilosXcosecha("KILOS_DISPONIBLES"))%>
				<%rskilosXcosecha.movenext
				wend%>		
				<tr class="reg_Header_nav">
					<td width="42%" align="center"><% =GF_TRADUCIR("Total")%></td>
					<td width="45%" align="right"><%=GF_EDIT_DECIMALS(cdbl(TotalKilosCosecha)*100,2)%></td>					
					<td></td>
				</tr>		
			</table>					
		<%end if %>
	<%end if %>
</form>
</body>
</html>

