<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
Dim pIdObra, pAccion, esModificable, rsBudget, indiceSecciones, indiceLineas, obraFechaBudget
Dim obraCD, obraDS, obraDivID, obraDivDS, obraImporte, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS
Dim dicDesc, dicValores, dicOperaciones,DsDetalle,dicDs, dicCuenta, dicCC,dicImpTrim, tipoFormulario
Dim totalTrim(5)

Const EMPTY_LINE_CODE = "*"
Const FIRST_CUSTOM_CODE = 1000

Const KEY_CODIGO = "cd_"
Const KEY_DESCRIPCION = "ds_"
Const KEY_CUENTA = "cuenta_"
Const KEY_CCOSTO = "cc_"
Const KEY_IMPORTE = "bg_"
Const KEY_TRIMESTRES = "trim_"
Const KEY_DETALLE = "ta_"
Const KEY_DIV = "div_"

CONST TRIMESTRE_1 = 0
CONST TRIMESTRE_2 = 1
CONST TRIMESTRE_3 = 2
CONST TRIMESTRE_4 = 3

Const TRIM_SPLIT_CHAR = "|"
Const TRIM_SIN_IMPORTE = -1
Const TRIM_PROCESADO = -1

Call comprasControlAccesoCM(RES_OBR)
'----------------------------------------------------------------------------------------
Function LeerParametros(ByRef pDicParamDesc, ByRef pDicParamValor, ByRef pDicParamDs, ByRef pDicParamCuenta, Byref pDicParamCC, Byref pDicImpTrim)
	Dim auxTrim1,auxTrim2,auxTrim3,auxTrim4
	
	Set pDicParamDesc = server.CreateObject("Scripting.Dictionary")
	Set pDicParamValor = server.CreateObject("Scripting.Dictionary")
	Set pDicParamDs = server.CreateObject("Scripting.Dictionary")
	Set pDicParamCuenta = server.CreateObject("Scripting.Dictionary")	
	Set pDicParamCC = server.CreateObject("Scripting.Dictionary")		
	Set pDicImpTrim = server.CreateObject("Scripting.Dictionary")		
	'El criterio para terminar cada ciclo sera encontrar un codigo = EMPTY_LINE_CODE
	'Si el importe esta vacio, se ignora la linea pero se sigue procesando.
	area = 0
	stopAreas=false	
	while (not stopAreas)
		areaCodigo = GF_PARAMETROS7(KEY_CODIGO & area & "_0", "", 6)		
		importe = 0
		if ((areaCodigo <> EMPTY_LINE_CODE) and (areaCodigo <> ""))then
			descripcion = GF_PARAMETROS7(KEY_DESCRIPCION & area & "_0", "", 6)
			dsDetalle = GF_PARAMETROS7(KEY_DETALLE & claveFila, "", 6)			
			Call pDicParamDesc.Add(areaCodigo & "_0", descripcion)			
			Call pDicParamValor.Add(areaCodigo & "_0", 0)
			Call pDicParamCuenta.add(areaCodigo & "_0", "")
			Call pDicParamCC.add(areaCodigo & "_0", 0)
			fila = 1
			'Se lee la siguiente linea del area
			stopDetalle=false	
			while (not stopDetalle)				
				claveFila = area & "_" & fila
				codigo = GF_PARAMETROS7(KEY_CODIGO & claveFila, "", 6)	
				if ((codigo = EMPTY_LINE_CODE) or (codigo = "")) then 
					stopDetalle=true					
				else
					'Hay codigo => Hay un registro que analizar!
					descripcion = GF_PARAMETROS7(KEY_DESCRIPCION & claveFila, "", 6)					
					dsDetalle = GF_PARAMETROS7(KEY_DETALLE & claveFila, "", 6)	
					cuenta = GF_PARAMETROS7(KEY_CUENTA & claveFila, "", 6)
					ccosto = GF_PARAMETROS7(KEY_CCOSTO & claveFila, "", 6)
					if (descripcion <> "") then
						'Hay descripcion, se lee el importe				
						importe = GF_PARAMETROS7(KEY_IMPORTE & claveFila, "", 6)						
						if (importe = "") then importe = 0
						
						'importe trimestrales
						auxTrim1 = GF_PARAMETROS7(KEY_TRIMESTRES & "_" & TRIMESTRE_1 & "_" & claveFila, "", 6)
						auxTrim2 = GF_PARAMETROS7(KEY_TRIMESTRES & "_" & TRIMESTRE_2 & "_" & claveFila, "", 6)
						auxTrim3 = GF_PARAMETROS7(KEY_TRIMESTRES & "_" & TRIMESTRE_3 & "_" & claveFila, "", 6)
						auxTrim4 = GF_PARAMETROS7(KEY_TRIMESTRES & "_" & TRIMESTRE_4 & "_" & claveFila, "", 6)
						
						if (not pDicImpTrim.Exists(areaCodigo & "_" & codigo)) then Call pDicImpTrim.add(areaCodigo & "_" & codigo,auxTrim1&TRIM_SPLIT_CHAR&auxTrim2&TRIM_SPLIT_CHAR&auxTrim3&TRIM_SPLIT_CHAR&auxTrim4)
						'Hay importe, leo la cuenta y la valido.							
						if (not pDicParamDesc.Exists(areaCodigo & "_" & codigo)) then Call pDicParamDesc.Add(areaCodigo & "_" & codigo, Trim(descripcion))
						if (not pDicParamValor.Exists(areaCodigo & "_" & codigo)) then Call pDicParamValor.Add(areaCodigo & "_" & codigo, importe)
						if (not pDicParamDS.Exists(areaCodigo & "_" & codigo)) then Call pDicParamDS.Add(areaCodigo & "_" & codigo, replace(Trim(dsDetalle) , chr(13)&chr(10), ENTER_SYMBOL))							
						if (not pDicParamCuenta.Exists(areaCodigo & "_" & codigo)) then Call pDicParamCuenta.Add(areaCodigo & "_" & codigo, Trim(cuenta))
						if (not pDicParamCC.Exists(areaCodigo & "_" & codigo)) then Call pDicParamCC.Add(areaCodigo & "_" & codigo, ccosto)
					
					end if
					
					
				end if			
				fila = fila+1
			wend
		else
			if (areaCodigo = "") then stopAreas=true
		end if
		area = area+1
	wend
	
End Function
'----------------------------------------------------------------------------------------
Function extraerCambios(pIdObra, pDicParamDesc,pDicParamDs, pDicParamValor, pDicParamCuenta, pDicParamCC, ByRef pDicOperaciones, tc,ByRef pDicImpTrim)
	Dim rs, strSQL, conn, clave, importePesos, importeDolares, k, aKey,auxTrim,auxClave
	
	Set pDicOperaciones = server.CreateObject("Scripting.Dictionary")
	'Se recorren todos los registros grabado buscando cuales hay que eliminar y cuales hay que modificar.
	Set rs = leerBudget(pIdObra)
	
	while (not rs.eof)
		'Se recorren todas las clave de la base buscando cambios y ausencias en la info de la pagina.
		clave= rs("IDAREA") & "_" & rs("IDDETALLE")
		if (pDicParamDesc.Exists(clave)) then
			'Se compara descripción e importe.						
			dim aux
			aux = rs("DSDETALLE")
			if (isnull(aux)) then aux = ""
			
			importeDolares = CDbl(pDicParamValor(clave))
			importePesos = importeDolares * tc
			
			if ((pDicParamDesc(clave) <> Trim(rs("DSBUDGET")))or _
				(CDbl(rs("DLBUDGET")) <> importeDolares*100) or _
				(CDbl(rs("PSBUDGET")) <> importePesos*100) or _
				(rs("CDCUENTA") <> pDicParamCuenta(clave)) or _
				(rs("CCOSTOS") <> pDicParamCC(clave)) or _
				(pDicParamDs(clave) <> aux)) then					'Cambio algo, actualizo.
				
				strSQL = "Update TBLBUDGETOBRAS Set DSBUDGET='" & pDicParamDesc(clave) & "', PSBUDGET=" & importePesos*100 & ", DLBUDGET=" & importeDolares*100 & ", TIPOCAMBIO=" & tc & ", CDCUENTA='" & pDicParamCuenta(clave) & "', CCOSTOS='" & pDicParamCC(clave) & "', CDUSUARIO='" & session("Usuario") & "', MOMENTO=" & session("MmtoDato") & ", DSDETALLE='"&pDicParamDs(clave)&"'"
				strSQL = strSQL & " where IDOBRA=" & pIdObra & " and IDAREA=" & rs("IDAREA") & " and IDDETALLE=" & rs("IDDETALLE") 
				
				Call pDicOperaciones.Add(clave, strSQL)		
				
			end if			
			Call pDicParamDesc.Remove(clave)
		else
			'No existe la clave, se borra el registro.
			strSQL = "Delete from TBLBUDGETOBRAS where IDOBRA=" & pIdObra & " and IDAREA=" & rs("IDAREA") & " and IDDETALLE=" & rs("IDDETALLE")
			Call pDicOperaciones.Add(clave, strSQL)
		end if
		
		rs.MoveNext()
	wend
	
	'Se recorren las claves que quedaron en el diccionario, son lineas nuevas del budget.
	for each k in pDicParamDesc.Keys		
		aKey = split(k,"_")
		importeDolares = CDbl(pDicParamValor(k))
		importePesos = importeDolares * tc
		strSQL="Insert into TBLBUDGETOBRAS(IDOBRA, IDAREA, IDDETALLE, DSBUDGET,DSDETALLE, PSBUDGET, DLBUDGET, TIPOCAMBIO, CDCUENTA, CCOSTOS, CDUSUARIO, MOMENTO)"
		strSQL=strSQL & " values(" & pIdObra & ", " & aKey(0) & ", " & aKey(1) & ", '" & pDicParamDesc(k) & "','" & pDicParamDs(k) & "'," & importePesos*100 & ", " & importeDolares*100 & ", " & tc & ", '" & pDicParamCuenta(k) & "', '" & pDicParamCC(k) & "', '" & session("Usuario") & "'," & session("MmtoDato") & ")"
		'Response.Write strSQL & "<BR>"
		Call pDicOperaciones.Add(k, strSQL)
	Next
	
	Call extraerCambiosDetalle(pIdObra, pDicOperaciones, pDicImpTrim)
End Function
'----------------------------------------------------------------------------------------
Function extraerCambiosDetalle(pIdObra,byref pDicOperaciones,byref pDicImpTrim)
	Dim myImportes,i,impTrim,auxClave,clave,auxImp
	
	if (tipoFormulario = OBRA_FORM_TRIM) then 
		Set impTrim = obtenerImportesTrim(pIdObra)
		
		
		while not impTrim.EoF
			clave= impTrim("IDAREA") & "_" & impTrim("IDDETALLE")	
			if (pDicImpTrim.Exists(clave)) then
				
				myImportes = split(pDicImpTrim(clave),TRIM_SPLIT_CHAR)
				'verifico si hubo cambios
				if (cdbl(impTrim("importe")) <> cdbl(myImportes(impTrim("periodo"))*100)) then
					strSQL = "update tblbudgetobrasdetalle set dlbudget = " & myImportes(impTrim("periodo"))*100 & ", psbudget = " & myImportes(impTrim("periodo")) * 100 * pTipoCambio & ", tipocambio=" & pTipoCambio & ", cdusuario='"& session("Usuario") & "', momento=" & session("mmtosistema") & " where periodo = " & impTrim("periodo") & " and idarea = " &impTrim("idarea")& " and iddetalle = " &impTrim("iddetalle")& " and idobra = " &pIdObra
					auxClave = clave & "_trim_" & impTrim("periodo")
					
					Call pDicOperaciones.Add(auxClave, strSQL)	
				end if
				
				'indico que el trimestre ya fue procesado
				myImportes(impTrim("periodo")) = TRIM_PROCESADO
				auxImp = ""
				for i = TRIMESTRE_1 to TRIMESTRE_4
					auxImp = auxImp & myImportes(i) & TRIM_SPLIT_CHAR
				next
				pDicImpTrim(clave) = auxImp
				
			else
				'la clave no existe, la borro de la base de datos
				strSQL = "Delete from tblbudgetobrasdetalle where idobra=" & pIdObra & " and IDAREA=" & impTrim("IDAREA") & " and IDDETALLE=" & impTrim("IDDETALLE")
				Call pDicOperaciones.Add(clave&"_trimDel"&impTrim("periodo"), strSQL)
			end if
			impTrim.MoveNext
		wend
		'Se recorren las claves que quedaron en el diccionario, son lineas nuevas del budget.
		for each k in pDicImpTrim.Keys
			aKey = split(k,"_")
			myImportes = split(pDicImpTrim(k),TRIM_SPLIT_CHAR)
			for i = TRIMESTRE_1 to TRIMESTRE_4
				if ( cdbl(myImportes(i)) > 0 and cdbl(myImportes(i)) <> TRIM_PROCESADO) then
					strSQL = "Insert into tblbudgetobrasdetalle (idobra,idarea,iddetalle,periodo,psbudget,dlbudget,tipocambio,cdusuario,momento) "
					strSQL = strSQL & " values("&pIdObra&","&aKey(0)&","&aKey(1)&","&i&","&myImportes(i)*100*pTipoCambio&","&myImportes(i)*100&","&pTipoCambio&",'"&session("Usuario")&"',"&session("mmtoSistema")&")"
					Call pDicOperaciones.Add(k+"_trimInsert_"&i, strSQL)
				end if
			next
		next
		
	end if
End Function
'----------------------------------------------------------------------------------------
Function obtenerImportesTrim(pIdObra)
	Dim strSQL,rs,conn
	
	strSQL = "select idarea,iddetalle,periodo,dlbudget importe from tblbudgetobrasdetalle where idobra = "&pIdObra&" order by idarea,iddetalle,periodo"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	Set obtenerImportesTrim = rs
	
End Function
'----------------------------------------------------------------------------------------
Function actualizaBudget(pDicOperaciones)
	DIm k,rs, conn
	for each k in pDicOperaciones.Keys
		'Response.Write pDicOperaciones(k) & "<BR><BR><BR>"			
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", pDicOperaciones(k))
	Next
	'response.end
End Function
'----------------------------------------------------------------------------------------
'	COMIENZO DE LA PAGINA
'----------------------------------------------------------------------------------------

call GP_ConfigurarMomentos()

pIdObra = GF_PARAMETROS7("idobra", 0, 6)
pAccion = GF_PARAMETROS7("accion", "", 6)
pTipoCambio = getTipoCambioBudget(pIdObra)

tipoFormulario = getFormTypeByIdObra(pIdObra)

if (not checkControlObra(pIdObra)) then
	response.redirect "comprasAccesoDenegado.asp"
end if

if ((pAccion = ACCION_GRABAR) or (pAccion = ACCION_CONFIRMAR)) then	
	Call LeerParametros(dicDesc, dicValores,dicDs, dicCuenta, dicCC,dicImpTrim)
	Call extraerCambios(pIdObra, dicDesc,dicDs, dicValores, dicCuenta, dicCC, dicOperaciones, pTipoCambio,dicImpTrim)
	Call actualizaBudget(dicOperaciones)
	if (pAccion = ACCION_CONFIRMAR) then Call confirmarBudget(pIdObra)
end if

Call loadDatosObra(pIdObra, obraCD, obraDS, obraDivID, obraDivDS, obraImporte, obraFechaBudget, obraMonedaID, obraFechaInicio, obraFechaFin, obraFechaAjustada, obraRespCD, obraRespDS)
if (not isBudgetProvisorio(obraFechaBudget)) then 
	if (puedeReasignarBudget(obraRespCD, obraDivID)) then
		Response.Redirect "comprasBudgetReasignaciones.asp?idObra=" & pIdObra
	else
		Response.Redirect "comprasBudgetObraPrint.asp?idObra=" & pIdObra				
	end if
end if

strSQL="Select * from TBLBUDGETOBRAS where IDOBRA=" & pIdObra & " and iddetalle = 0 Order by IDAREA, IDDETALLE"
Call executeQueryDb(DBSITE_SQL_INTRA, rsBudget, "OPEN", strSQL)
'Set rsBudget = leerBudget(pIdObra)
'Se toma el valor de codigo mas alto de los detalle, dado que vienen por ajax, las areas luego se comparan con este numero y se tomara el mayor absoluto de toda la partida
strSQL="Select MAX(IDDETALLE) CODIGO from TBLBUDGETOBRAS where IDOBRA=" & pIdObra
Call executeQueryDb(DBSITE_SQL_INTRA, rsMax, "OPEN", strSQL)
if (isNull(rsMax("CODIGO"))) then
	maxCode = FIRST_CUSTOM_CODE
else
	maxCode = CInt(rsMax("CODIGO"))
end if

%>

<html>
<head>
<title>Compras - Administración de Presupuesto</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="CSS/MagicSearch.css" type="text/css">
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
.cursorStyle {
	cursor: pointer;
}
.deshabilitado{
	background-color:#DDD;
}
</style>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/magicSearchObj.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>

<script type="text/javascript">		
	var ch = new channel();	
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");
	var MIN_CUST_CODE = 1000;
	var nextRow = 3;					//Proxima fila fisica en la tabla.		
	var sectionIndex = 0;				//Indice de codigos en cada seccion.
	var sectionCounter = new Array();	//Indice de filas en cada seccion.
	var msArray = new Array();			//Vector con los MS utilizados en el presupuesto.
	var custCode = <% =FIRST_CUSTOM_CODE %>;
	var shadowItems = new Array();
	var idsTrimestres = new Array();
	var idsImportesTotales = new Array();
	var totalTrim0 = 0;
	var totalTrim1 = 0;
	var totalTrim2 = 0;
	var totalTrim3 = 0;
	var totalTrim5 = 0;

	
	function sacarImagenDel(area,fila) {
		var colI = document.getElementById("colI_" + area + "_" + fila);
		var img = document.getElementById("imgDel_" + area + "_" + fila);
		if (img) colI.removeChild(img);
	}
	function modificarTotalGeneral(me,areaDetalle)
	{
		
		var importe = 0;
		var aux1 = 0;
		for (var idsTotales in idsImportesTotales)
		{
			aux1 = document.getElementById('<% =KEY_IMPORTE %>'+ idsImportesTotales[idsTotales]).value;
			if (isNaN(aux1) || aux1 == '') aux1 = 0
			importe += parseFloat(aux1);
		}
		
		total = document.getElementById('totalAnual');
		total.value = editarImporte(String(importe));
		
		total.readonly = true;
		total.style.textAlign = 'right';
	}
	
	function modificarTotal(me,areaDetalle){
		if (controlCampo(me, 'I'))
		{
			var trim1, trim2, trim3, trim4
			var tempTrim1, tempTrim2, tempTrim3, tempTrim4
			var totalesTrim = new Array();
			totalesTrim[1] = 0;
			totalesTrim[2] = 0;
			totalesTrim[3] = 0;
			totalesTrim[4] = 0;
			//Suma de todo el BUDGET - Todas las lineas del BUDGET
			for (var idtrim in idsTrimestres)
			{
				tempTrim1 = Math.round(parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_1%>'+idsTrimestres[idtrim]).value), 2);
				tempTrim2 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_2%>'+idsTrimestres[idtrim]).value);
				tempTrim3 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_3%>'+idsTrimestres[idtrim]).value);
				tempTrim4 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_4%>'+idsTrimestres[idtrim]).value);
				if (isNaN(tempTrim1) || tempTrim1 == '') tempTrim1 = 0;
				if (isNaN(tempTrim2) || tempTrim2 == '') tempTrim2 = 0;
				if (isNaN(tempTrim3) || tempTrim3 == '') tempTrim3 = 0;
				if (isNaN(tempTrim4) || tempTrim4 == '') tempTrim4 = 0;
				totalesTrim[1] += tempTrim1;
				totalesTrim[2] += tempTrim2;
				totalesTrim[3] += tempTrim3;
				totalesTrim[4] += tempTrim4;					
			}
			//Suma de los trimestres del area en cuestion - Total de detalle
			trim1 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_1%>'+"_"+areaDetalle).value)
			trim2 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_2%>'+"_"+areaDetalle).value)
			trim3 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_3%>'+"_"+areaDetalle).value)
			trim4 = parseFloat(document.getElementById('<%=KEY_TRIMESTRES & "_" & TRIMESTRE_4%>'+"_"+areaDetalle).value)
			if (isNaN(trim1) || trim1 == '') trim1 = 0
			if (isNaN(trim2) || trim2 == '') trim2 = 0
			if (isNaN(trim3) || trim3 == '') trim3 = 0
			if (isNaN(trim4) || trim4 == '') trim4 = 0
			document.getElementById('<% =KEY_IMPORTE %>'+areaDetalle).value  = editarImporte(String(trim1+trim2+trim3+trim4));			
			//Suma de los totales del BUDGET - Total General
			for (var i=1;i<=4;i++)
			{
				document.getElementById('totalTrim'+i).value = editarImporte(String(totalesTrim[i]));
				document.getElementById('totalTrim'+i).readonly = true;
				document.getElementById('totalTrim'+i).style.textAlign = 'right';
			}
			document.getElementById('totalAnual').value = editarImporte(String(totalesTrim[1]+totalesTrim[2]+totalesTrim[3]+totalesTrim[4]));
			document.getElementById('totalAnual').readonly = true;
			document.getElementById('totalAnual').style.textAlign = 'right';
		}
	}
	
	function agregarImagenDel(area, fila) {
		var colI = document.getElementById("colI_" + area + "_" + fila);
		var img = document.getElementById("imgDel_" + area + "_" + fila);
		if (!img) {
			img = document.createElement("img");
			img.id= "imgDel_" + area + "_" + fila;
			img.src="images/compras/remove-16x16.png";
			img.className="cursorStyle";
			if (isFirefox) {
				img.setAttribute('onclick', "eliminarLinea1(" + fila + "," + area + ")");
			} else {
				img['onclick'] = new Function("eliminarLinea1(" + fila + "," + area + ")");
			}
			colI.appendChild(img);
		}
	}
	
	function seleccionarItem(area, fila, ms) {
		var desc = ms.getSelectedItem();
		var kCode = "<% =KEY_CODIGO %>" + area + "_" + fila;
		var kDiv = "<% =KEY_DIV %>" + area + "_" + fila;
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById(kCode).value = arr[0];
			document.getElementById(kDiv).innerHTML = arr[0];
			ms.setValue(arr[1]);
			shadowItems[area + "_" + fila] = arr[1];
		} else {
			if (desc == "") {
				document.getElementById(kCode).value = "<% =EMPTY_LINE_CODE %>";
				document.getElementById(kDiv).innerHTML = "<% =EMPTY_LINE_CODE %>";
				shadowItems[area + "_" + fila] = "";
			} else {				
				ms.setValue(desc);
				if (shadowItems[area + "_" + fila] != desc) {
					shadowItems[area + "_" + fila] = desc;
					if ((document.getElementById(kCode).value == '<% =EMPTY_LINE_CODE %>') ||
						(document.getElementById(kCode).value < MIN_CUST_CODE)){						
						document.getElementById(kCode).value = custCode;
						document.getElementById(kDiv).innerHTML = custCode;
						custCode++;
					}
				}
			}			
		}			
		if (fila > 0) {	//Solo para los items
			if (desc != "") {		
				//Nueva descripcion, se da la chance de borrarla.					
				agregarImagenDel(area, fila);
			} else {
				//Borro el texto, elimina la imagen.
				sacarImagenDel(area,fila);
			}
		}
	}
	
	function agregarLinea(rowNbr, level, area) {
	
		var idFila = sectionCounter[area];
		sectionCounter[area]++;
		var tb = document.getElementById("budgetTable");
		var row = tb.insertRow(rowNbr);		
		//Se determina la primer columna
		var columnIdx = level;
		//Es una linea del nivel 1?
		if (columnIdx > 0) row.insertCell(0);			
		//Codigo del budget.
		var colCodigo = row.insertCell(columnIdx);				
				
		colCodigo.className = "reg_Header_nav";						
		var divCodigo = document.createElement("div");
		divCodigo.id = "<% =KEY_DIV %>" + area + "_" + idFila;
		divCodigo.name = "<% =KEY_DIV %>" + area + "_" + idFila;
		divCodigo.innerHTML='*';
		divCodigo.style.textAlign='center';
		colCodigo.appendChild(divCodigo);
		
		var inputCodigo = document.createElement("input");		
		inputCodigo.type="hidden";
		inputCodigo.id = "<% =KEY_CODIGO %>" + area + "_" + idFila;
		inputCodigo.name = "<% =KEY_CODIGO %>" + area + "_" + idFila;				
		inputCodigo.value = '*';
		colCodigo.appendChild(inputCodigo);				
		
		//Descripcion del budget.		
		columnIdx++;
		var colDescripcion = row.insertCell(columnIdx);
		var inputDesc = document.createElement('div');
		inputDesc.id = "<% =KEY_DESCRIPCION %>" + area + "_" + idFila;		
		colDescripcion.appendChild(inputDesc);				
		var link = "comprasStreamElementos.asp?tipo=";
		link += (level==0)? "aBudget": "dBudget";
		var claveDS = "<% =KEY_DESCRIPCION %>" + area + "_" + idFila;
		msArray[claveDS] = new MagicSearch('', "<% =KEY_DESCRIPCION %>" + area + "_" + idFila , 75, 2, link);
		msArray[claveDS].setToken(";");
		msArray[claveDS].onBlur = "seleccionarItem(" + area + ", " + idFila + ")";
				
		columnIdx++;
		var colBudget = row.insertCell(columnIdx);
		if (level > 0) {			
			//Valor asignado al item.
			var inputBudget = document.createElement("input");
			inputBudget.name = "<% =KEY_IMPORTE %>" + area + "_" + idFila;
			inputBudget.id = "<% =KEY_IMPORTE %>" + area + "_" + idFila;
			inputBudget.size=11;
			inputBudget.style.textAlign = 'right';
			
			<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
				inputBudget.className='deshabilitado';
			<% end if %>
			
			if (isFirefox) {
				
				<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
					inputBudget.setAttribute('readonly', "readonly");	
				<% else %>
					inputBudget.setAttribute('onkeypress', "return controlDatos(this, event, 'I')");
					inputBudget.setAttribute('onchange', "return controlCampo(this, 'I')");	
					inputBudget.setAttribute('onblur', "return modificarTotalGeneral(this,'"+area + "_" + idFila+"')");	
				<% end if %>
			} else {
				<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
					inputBudget['readonly'] = new Function("readonly");
				<% else %>
					inputBudget['onkeypress'] = new Function("return controlDatos(this, event, 'I')");
					inputBudget['onchange'] = new Function("return controlCampo(this, 'I')");
					inputBudget['onblur'] = new Function("return modificarTotalGeneral(this,'"+area + "_" + idFila+"')");
				<% end if %>
			}	
				
			colBudget.appendChild(inputBudget);	
			
			idsImportesTotales[nextRow] = area + "_" + idFila
			//Trimestres
			<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
				var colTrim = new Array(3);
				var inputTrim = new Array(3);
				for (var i=1;i<=4;i++)
				{
					idsTrimestres[nextRow] = "_" + area + "_" + idFila
					
					columnIdx++;
					colTrim[i] = row.insertCell(columnIdx)
					inputTrim[i] = document.createElement("input");
					inputTrim[i].name = "<% =KEY_TRIMESTRES %>_"+(i-1)+"_" + area + "_" + idFila;
					inputTrim[i].id = "<% =KEY_TRIMESTRES %>_"+(i-1)+"_" + area + "_" + idFila;
					inputTrim[i].size=11;
					inputTrim[i].value=editarImporte("0");
					inputTrim[i].style.textAlign = 'right';
					if (isFirefox)
					{
						inputTrim[i].setAttribute('onkeypress', "return controlDatos(this, event, 'I')");
						inputTrim[i].setAttribute('onblur', "return modificarTotal(this,'"+area + "_" + idFila+"')");
					}
					else
					{
						inputTrim[i]['onkeypress'] = new Function("return controlDatos(this, event, 'I')");
						inputTrim[i]['onblur'] = new Function("return modificarTotal(this,'"+area + "_" + idFila+"')");
					}
					
					colTrim[i].appendChild(inputTrim[i]);	
				}
			//alert(idsTrimestres[nextRow]);
			<% end if %>
			
			//Cuenta
			columnIdx++;
			var colCuenta = row.insertCell(columnIdx);
			var inputCuenta = document.createElement("input");
			inputCuenta.name = "<% =KEY_CUENTA %>" + area + "_" + idFila;
			inputCuenta.id = "<% =KEY_CUENTA %>" + area + "_" + idFila;
			inputCuenta.size=11;
			colCuenta.appendChild(inputCuenta);
		
			<% if (tipoFormulario <> OBRA_FORM_ANUAL) then %>
				//Centro de Costo
				columnIdx++;
				var colCC = row.insertCell(columnIdx);
				var inputCC = document.createElement("input");
				inputCC.name = "<% =KEY_CCOSTO %>" + area + "_" + idFila;
				inputCC.id = "<% =KEY_CCOSTO %>" + area + "_" + idFila;
				inputCC.size=4;
				if (isFirefox) {
					inputCC.setAttribute('onkeypress', "return controlDatos(this, event, 'N')");
					inputCC.setAttribute('onblur', "return controlCampo(this, 'N')");
				} else {
					inputCC['onkeypress'] = new Function("return controlDatos(this, event, 'N')");
					inputCC['onblur'] = new Function("return controlCampo(this, 'N')");
				}					
				colCC.appendChild(inputCC);
			<% end if %>
		}
		
		//Accion para ver detalle del item.
		columnIdx++;
		var colDs = row.insertCell(columnIdx);
		if (level != 0) {
			var imgDs = document.createElement("img");
			imgDs.id = "imgDs_" + area + "_" + idFila;
			imgDs.src = "images/compras/budget_item_detail_16x16.png";
			colDs.setAttribute('onclick', "showHide('ta_" + area + "_" + idFila + "');");
			imgDs.className="cursorStyle"		
			colDs.appendChild(imgDs);
		}
				
		
		//Se agrega la imagen para la accion disponible en la linea.				
		columnIdx++;
		var colImagen = row.insertCell(columnIdx);
		colImagen.id ="colI_" + area + "_" + idFila; 			
		
		columnIdx++;	
		var colImagen2 = row.insertCell(columnIdx);
		colImagen2.id ="colI2_" + area + "_" + idFila; 
		var img = document.createElement("img");
		img.id = "img_" + area + "_" + idFila;
		img.src = "images/compras/add-16x16.png";
		img.className="cursorStyle";
        <% if (tipoFormulario = OBRA_FORM_TRIM) then %>
			colImagen2.colSpan=5;
		<% end if %>
		colImagen2.align="right";
		colImagen2.appendChild(img);
		if (isFirefox) {
			if (level == 0)
				img.setAttribute('onclick', "nuevaArea()");
			else
				img.setAttribute('onclick', "nuevoDetalle(" + area + ")");
		} else {
			if (level == 0)
				img['onclick'] = new Function("nuevaArea()");
			else
				img['onclick'] = new Function("nuevoDetalle(" + area + ")");
		}
				
		//Se da formato a las columnas		
		colCodigo.width      ="5%";
		colDescripcion.width ="80%";		
		if (level==0) colDescripcion.colSpan=2;
		colBudget.width  ="10%";
		colImagen.width  ="5%";
		colImagen2.width ="5%";	
		nextRow++;
		
		//Se crea la linea con el area para ingresar detalles.
		if (level != 0) {
			var detailRow = tb.insertRow(rowNbr+1);
			detailRow.insertCell(0);
			detailRow.insertCell(1);			
			var detailCell = detailRow.insertCell(2);
			var detailText = document.createElement("textarea");		
			detailText.name = "<% =KEY_DETALLE %>" + area + "_" + idFila;
			detailText.id = "<% =KEY_DETALLE %>" + area + "_" + idFila;
			detailText.cols = 82;
			detailText.style.visibility = 'hidden';
			detailText.style.position = 'absolute';
			if (isFirefox) {
				detailText.setAttribute('onkeypress', "controlMaxChars(this,254)");
			} else {
				detailText['onkeypress'] = new Function("controlMaxChars(this,254)");
			}		
			detailCell.appendChild(detailText);				
			nextRow++;	//Se incrementa el indice a la ultima fila de la tabla.
		}
	}

	function showHide(pName){
		
		var text  = document.getElementById(pName)
		if (text.style.visibility == 'visible'){
			text.style.visibility = 'hidden';
			text.style.position   = 'absolute';
		}else{
			text.style.visibility = 'visible';
			text.style.position   = 'relative';
		}
		
		
	}
	function controlMaxChars(campo,limite){
		if(campo.value.length>=limite){
			campo.value=campo.value.substring(0,limite);
		}
	}

	//Elimina una linea del nivel 1
	function eliminarLinea1(idx, area) {
		var idxNext = idx+1;
		var codigoNext = document.getElementById("<% =KEY_CODIGO %>" + area + "_" + idxNext);		
		while (codigoNext) {
			msArray["<% =KEY_DESCRIPCION %>" + area + "_" + idx].setValue(msArray["<% =KEY_DESCRIPCION %>" + area + "_" + idxNext].getSelectedItem());													
			document.getElementById("<% =KEY_CODIGO %>" + area + "_" + idx).value = document.getElementById("<% =KEY_CODIGO %>" + area + "_" + idxNext).value;
			document.getElementById("<% =KEY_DIV %>" + area + "_" + idx).innerHTML = document.getElementById("<% =KEY_DIV %>" + area + "_" + idxNext).innerHTML;						
			document.getElementById("<% =KEY_IMPORTE %>" + area + "_" + idx).value = document.getElementById("<% =KEY_IMPORTE %>" + area + "_" + idxNext).value;
			
			document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_1 %>_" + area + "_" + idx).value = document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_1%>_" + area + "_" + idxNext).value;
			document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_2%>_" + area + "_" + idx).value = document.getElementById("<% =KEY_TRIMESTRES & "_"  & TRIMESTRE_2%>_" + area + "_" + idxNext).value;
			document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_3%>_" + area + "_" + idx).value = document.getElementById("<% =KEY_TRIMESTRES & "_"  & TRIMESTRE_3%>_" + area + "_" + idxNext).value;
			document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_4%>_" + area + "_" + idx).value = document.getElementById("<% =KEY_TRIMESTRES& "_"   & TRIMESTRE_4 %>_" + area + "_" + idxNext).value;

			document.getElementById("<% =KEY_CUENTA %>" + area + "_" + idx).value = document.getElementById("<% =KEY_CUENTA %>" + area + "_" + idxNext).value;
            <% if (tipoFormulario <> OBRA_FORM_ANUAL) then %>
				document.getElementById("<% =KEY_CCOSTO %>" + area + "_" + idx).value = document.getElementById("<% =KEY_CCOSTO %>" + area + "_" + idxNext).value;
			<% end if %>
			document.getElementById("<% =KEY_DETALLE %>" + area + "_" + idx).value = document.getElementById("<% =KEY_DETALLE %>" + area + "_" + idxNext).value;
			shadowItems[area + "_" + idx] = shadowItems[area + "_" + idxNext];
			idx++;
			idxNext++;
			codigoNext = document.getElementById("<% =KEY_CODIGO %>" + area + "_" + idxNext);
		}			
		//Se borra la ultima linea de la sección
		msArray["<% =KEY_DESCRIPCION %>" + area + "_" + idx].setValue("");
		shadowItems[area + "_" + idx] = "";
		
		document.getElementById("<% =KEY_CODIGO %>" + area + "_" + idx).value = "<%=EMPTY_LINE_CODE%>";
		document.getElementById("<% =KEY_DIV %>" + area + "_" + idx).innerHTML = "<%=EMPTY_LINE_CODE%>";
		
		document.getElementById("<% =KEY_IMPORTE %>" + area + "_" + idx).value = "";
		document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_1%>_" + area + "_" + idx).value = "";
		document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_2%>_" + area + "_" + idx).value = "";
		document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_3%>_" + area + "_" + idx).value = "";
		document.getElementById("<% =KEY_TRIMESTRES & "_" & TRIMESTRE_4%>_" + area + "_" + idx).value = "";
		document.getElementById("<% =KEY_CUENTA %>" + area + "_" + idx).value = "";
		
		<% if (tipoFormulario <> OBRA_FORM_ANUAL) then %>
			document.getElementById("<% =KEY_CCOSTO %>" + area + "_" + idx).value = "";
		<% end if %>
		document.getElementById("<% =KEY_DETALLE %>" + area + "_" + idx).value = "";
		sacarImagenDel(area,idx);
	
	}
	
	function agregarArea() {
		var s = sectionIndex-1;
		var imgId = "img_" + s + "_0"
		var img = document.getElementById(imgId);
		if (img) img.parentNode.removeChild(img);		
		var area = sectionIndex;		
		sectionIndex++;
		sectionCounter.push(0);		
		agregarLinea(nextRow, 0, area);							
		return area;
	}
	
	function nuevaArea() {
		var area = agregarArea();	
		agregarLinea(nextRow, 1, area);
	}
	
	function nuevoDetalle(area) {			
		var fila = sectionCounter[area]-1;		
		var colI2 = document.getElementById("colI2_" + area + "_" + fila);
		var row = colI2.parentNode.rowIndex + 1;
		if (fila > 0) {			
			var imgAdd = document.getElementById("img_" + area + "_" + fila);							
			colI2.removeChild(imgAdd);		
			row += 1;//Por que la linea anterior tiene textarea para detalle.
		}		
		agregarLinea(row, 1, area);
	}
		
	function grabar() {
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel2").style.visibility = 'visible';		
		document.getElementById("frmSel").submit();
	}
	
	function cargar() {
		document.getElementById("actionLabel").style.visibility = 'visible';
		document.getElementById("actionLabel2").style.visibility = 'visible';		
	}
	
	function ocultar_cargar() {
		document.getElementById("actionLabel").style.visibility = 'hidden';
		document.getElementById("actionLabel2").style.visibility = 'hidden';		
	}
	
	
	
	function confirmarBudget() {
		var resp = confirm("Esta seguro que desea confirmar este pedido, esto hara que el mismo sea DEFINITIVO y no podrá volver a modificarse?");
		if (resp) {
			document.getElementById("accion").value="<% =ACCION_CONFIRMAR %>";
			document.getElementById("mensaje").value = "Salvando";
			grabar();
		}
	}
	function getAreasCallBack(cantidad)
	{
		eval(ch.response());
		
		if (cantidad == <%=rsBudget.recordCount%>)
		{
			ocultar_cargar();
			
			document.getElementById('totalAnual').value = totalTrim5 ;
			<%if (tipoFormulario = OBRA_FORM_TRIM) then
				for i = 1 to 4 %>
					document.getElementById("totalTrim<%=i%>").value = editarImporte(String(totalTrim<%=i-1%>));
				<%next
			end if%>
		}
	}
	
	<% 	if (not rsBudget.eof) then %>

	function loadBudget() {
		var theArea;
		var cantidad = 1;
		cargar();
		<%	flagEmptyItem = false			
			while (not rsBudget.eof)
				flagEmptyItem = false
				if (CDbl(rsBudget("IDDETALLE")) = 0) then 
					indiceLineas = 0 
					flagEmptyItem = true	%>
					theArea = agregarArea();
					document.getElementById("<% =KEY_CODIGO %>" + theArea + "_<% =indiceLineas %>").value = "<% =rsBudget("IDAREA") %>";
					document.getElementById("<% =KEY_DIV %>" + theArea + "_<% =indiceLineas %>").innerHTML = "<% =rsBudget("IDAREA") %>";
					msArray["<% =KEY_DESCRIPCION %>" + theArea + "_<% =indiceLineas %>"].setValue("<% =rsBudget("DSBUDGET") %>");
					shadowItems[theArea + "_<% =indiceLineas %>"] = "<% =rsBudget("DSBUDGET") %>";
					
					<%indiceLineas = indiceLineas +1%>
					ch.bind("comprasbudgetObraAjax.asp?theArea="+theArea+"&idobra=<%=pIdObra%>&idArea=<% =rsBudget("IDAREA") %>&indicelineas=<%=indiceLineas%>&tipoFormulario=<%=tipoFormulario%>", "getAreasCallBack("+cantidad+")");
					ch.send();
					
					cantidad += 1;
					
			<%	end if	%>										
			
			<%	if (maxCode < rsBudget("IDAREA")) then maxCode=rsBudget("IDAREA")				
				indiceLineas = indiceLineas + 1
				rsBudget.MoveNext()
			wend	
			if (flagEmptyItem) then		%>
				nuevoDetalle(theArea);								
		<%  end if					
			if (maxCode >= FIRST_CUSTOM_CODE) then	%>
			custCode = <% =maxCode+1 %>;
		<%	end if		%>			
		
	}
	<%	end if %>
	
	function imprimir() {
		window.open("comprasbudgetobrafilter.asp?idobra=<% =pIdObra %>");		
	}
	
	function bodyOnLoad() {
		<% if (pAccion= ACCION_CONFIRMAR) then%>
		<% end if %>
		var tb = new Toolbar('toolbar', 10, 'images/compras/');
		var tb2 = new Toolbar('toolbar2', 10, 'images/compras/');
		tb.addButtonSAVE("Grabar", "grabar()");
		tb2.addButtonSAVE("Grabar", "grabar()");
		tb.addButtonCANCEL("Cancelar", "window.close()");		
		tb2.addButtonCANCEL("Cancelar", "window.close()");
		tb.addButton("budget_confirm-16x16.png", "Confirmar", "confirmarBudget()");
		tb2.addButton("budget_confirm-16x16.png", "Confirmar", "confirmarBudget()");
		tb.addButton("printer-16x16.png", "Imprimir", "imprimir()");
		tb2.addButton("printer-16x16.png", "Imprimir", "imprimir()");		
		tb.draw();
		tb2.draw();
		<% 	if (rsBudget.RecordCount > 0) then %>
		loadBudget();
		<%	else	%>
		nuevaArea();
		<%	end if %>
		
	}		
</script>
</head>
<body onload="bodyOnLoad()">
	<input type="hidden" id="mensaje" name="menasje" value="Cargando">
	<div id="toolbar"></div>
	<div align="center"><div id="actionLabel" class="round_border_bottom round_border_top TDSUCCESS" style="width:70%;visibility:hidden;"><script> document.write(document.getElementById("mensaje").value);</script>...</div></div><br>
	<form name="frmSel" id="frmSel" method="POST" action="comprasBudgetObra.asp">	
	<table class="reg_Header" id="budgetTable" align="center" width="70%" border="0" cellpadding="0" cellspacing="1">
		<tr>
			<% if (tipoFormulario = OBRA_FORM_ANUAL) then %>
				<td colspan="8" class="reg_header_navdos">
			<% else 
			    if (tipoFormulario = OBRA_FORM_SERVICIOGRAL) then %>
				<td colspan="9" class="reg_header_navdos">
			<%  else  %>
				<td colspan="13" class="reg_header_navdos">
			<%  end if 
			   end if%>
			<h2><% =GF_TRADUCIR("Responsable") %>:&nbsp;<% =obraRespDS %></h2>
			</td>			
		</tr>
		<tr>
			<% if (tipoFormulario <> OBRA_FORM_ANUAL) then %>
				<td colspan="4" style="background-color:#ffff99">
			<% else %>
				<td colspan="3" style="background-color:#ffff99">
			<% end if%>
			
				<img src="images/compras/warning-16x16.png" align="absMiddle">&nbsp;<b><% =GF_TRADUCIR("ATENCIÓN: Todos los importes informados se deben expresar en Dólares Estadounidenses.") %></b>
			</td>
			
			<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
				<td colspan="6" class="reg_header_nav" style="text-align: right;"><% =GF_TRADUCIR("T. Cambio") %>:</td>
			<% else %>
				<td colspan="3" class="reg_header_nav" style="text-align: right;"><% =GF_TRADUCIR("T. Cambio") %>:</td>
			<% end if%>
			<td colspan="3"><input type="text" id="tipoCambio" name="tipoCambio" size="5" value="<% =pTipoCambio %>" onKeyPress="return controlDatos(this, event, 'I')"></td>
		</tr>		
		<tr class="reg_Header_nav">
			<td align="center"></td>
			<td colspan="2"><%=GF_TRADUCIR("Detalle") %></td>				
			<td><% =GF_TRADUCIR("Total") %></td>
			<% if (tipoFormulario = OBRA_FORM_TRIM) then %>
				<td><% =GF_TRADUCIR("1er Trim.") %></td>
				<td><% =GF_TRADUCIR("2do Trim.") %></td>
				<td><% =GF_TRADUCIR("3er Trim.") %></td>
				<td><% =GF_TRADUCIR("4to Trim.") %></td>
			<% end if %>
			
			<% if (tipoFormulario = OBRA_FORM_ANUAL) then %>
				<td><% =GF_TRADUCIR("Subproyecto") %></td>
			<% else %>
				<td><% =GF_TRADUCIR("Cuenta") %></td>
				<td><% =GF_TRADUCIR("C.Costo") %></td>
			<% end if %>
			<td align="center">.</td>
			<td align="center">.</td>
			<td align="center">.</td>
			
		</tr>
		
		
		<tr class="reg_Header_nav">
			<td colspan="3"><%=GF_TRADUCIR("Total") %></td>				
				<td><input type="text" id="totalAnual" name="totalAnual" value="" size="11" class="deshabilitado" readonly style="text-align:right;"></td>
			<% if (tipoFormulario = OBRA_FORM_TRIM) then %>			
				<td><input type="text" id="totalTrim1" name="totalTrim1" value="" size="11" class="deshabilitado" readonly style="text-align:right;"></td>
				<td><input type="text" id="totalTrim2" name="totalTrim2" value="" size="11" class="deshabilitado" readonly style="text-align:right;"></td>
				<td><input type="text" id="totalTrim3" name="totalTrim3" value="" size="11" class="deshabilitado" readonly style="text-align:right;"></td>
				<td><input type="text" id="totalTrim4" name="totalTrim4" value="" size="11" class="deshabilitado" readonly style="text-align:right;"></td>
			<% end if %>
		</tr>
		
		
	</table>
	<input type="hidden" name="idObra" value="<% =pIdObra %>">
	<input type="hidden" id="accion" name="accion" value="<% =ACCION_GRABAR %>">
	</form>
	<br><div align="center"><div id="actionLabel2" class="round_border_bottom TDSUCCESS" style="width:70%;visibility:hidden;"><script> document.write(document.getElementById("mensaje").value);</script>...</div></div>
	<div id="toolbar2"></div>
</body>
</html>