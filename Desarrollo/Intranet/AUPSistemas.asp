<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/GF_MGSRADD.asp"-->

<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Modificacion:Henzel Pavlo, 01/08/08               -->
'<!--               Fecha      : 20/12/2007                         -->
'<!--               Pagina     : AUPSistemas.ASP                    -->
'<!--               Descripcion: Seleccion de sistemas a asignar    -->
'<!------------------------------------------------------------------->
ProcedimientoControl "AUPSIS"
Dim strSQL, rs, rsPersonas, oConn, strUbicacion
dim myApellido, myNombre, myNroLegajo, mySRO1KR, myTTKC, myTTDS, myChecked
dim myDescomposicion, myDescomposicionDES, myDescomposicionDES2, myProfesional, myProfesionalDS 
dim myO1KRPS, index, myBody, myRecipients, myMAILKR
dim FrmDic, titu, myAccion, myAccionLoad, myEgresoValido
dim myAlta, myBaja, MyUsuario

'Se crea el diccionario de parametros.
myChecked = ""
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.Form 
   FrmDic.Add  i,Request.Form(i).item
Next

	'KR del usuario logueado	
	call GF_MGC ("SG", session("Usuario"), myProfesional, myProfesionalDS)
	'call GF_MGC ("SG", "JAS", myProfesional, myProfesionalDS)
	
	'Guardar datos del credencial, si se ingreso uno...
	if FrmDic("nroCredencial") <> "" then
		if esCredencialNuevo(FrmDic("nroCredencial"), FrmDic("IdPersona")) then
			Call guardarCredencialNuevo(FrmDic("nroCredencial"), FrmDic("IdPersona"))
		end if
	end if
	
	'Obtener datos del empleado seleccionado
	'strSQL= "select Apellido, Nombre, NroLegajo from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional where P.idprofesional=" & FrmDic("IdPersona") & " and EgresoValido='F'"
	'call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
	'if not rsPersonas.eof then
	strSQL= "select Apellido, Nombre, NroLegajo , mg.mg_kc as Usuario, EgresoValido from Profesionales P inner join Personas Pe on Pe.idpersona=P.idProfesional inner join mg on p.idprofesional=mg.mg_kr where P.idprofesional=" & FrmDic("IdPersona") 
	'Response.Write strsql
	call GF_BD_CONTROL (rsPersonas,oConn,"OPEN",strSQL)
	if not rsPersonas.eof then
		myApellido = rsPersonas("Apellido")
		myNombre = rsPersonas("Nombre")
		myUsuario = rsPersonas("Usuario")
		myNroLegajo = rsPersonas("NroLegajo")
		myEgresoValido = rsPersonas("EgresoValido")		
	end if
	call GF_MGC("SR","TSTT",mySRO1KR,"")
	call GF_MGC("SR","PRTT",myO1KRPS,"")
	call GF_MGC("SD","SGEMAIL",myMAILKR,"")

	'Cargar lista de receptores del mail
	myRecipients = GF_DT1KR("READ", myMAILKR, "", "", myO1KRPS)
	if session("Usuario") = "EAB" then myRecipients = "BacariniE@acti.de"

	'Accion a realizar
	'Agregar
	myDescomposicion = split("/" & FrmDic("Relaciones") & "/","//")
	for index=1 to ubound(myDescomposicion)-1
		myDescomposicion2 = split(myDescomposicion(index),"=")
		myValor = myDescomposicion2(1)
		if myValor = "X" then myValor = "*"
		call GF_MGSRADD(myO1KRPS, FrmDic("IdPersona"), myDescomposicion2(0), myValor,  "")
	next
	'Quitar
	myDescomposicion = split("/" & FrmDic("RelacionesDel") & "/","//")
	for index=1 to ubound(myDescomposicion)-1
		myDescomposicion2 = split(myDescomposicion(index),"=")
		call GF_MGSRADD(myO1KRPS, FrmDic("IdPersona"), myDescomposicion2(0), "*", "")
	next

	titu = myProfesionalDS & "(" & session("Usuario") & ")" & " ha solicitado la modificación del perfil del siguiente usuario." & chr(10) & chr(13) & chr(10) & chr(13) & "USUARIO:" & chr(10) & chr(13) & "Apellido...: " & myapellido & chr(10) & chr(13) & "Nombre.....: " & MyNombre & chr(10) & chr(13) & "Legajo nro.: " & myNroLegajo
	
	myDescomposicionDES = split("/" & FrmDic("RelacionesDES") & "/","//")
	if ubound(myDescomposicionDES)-1 >0 then myBody = myBody & chr(10) & chr(13) & chr(10) & chr(13) & "TAREAS CON SOLICITUD DE ALTA:"
	for index=1 to ubound(myDescomposicionDES) - 1
		myDescomposicionDES2 = split(myDescomposicionDES(index),"##")
		if myDescomposicionDES2(0) = myAntSis then
			myBody = myBody & chr(10) & chr(13) & prepareWord(myDescomposicionDES2(1))
		else
			myAntSis = myDescomposicionDES2(0)
			myBody = myBody & chr(10) & chr(13) & prepareWord(myDescomposicionDES2(0))
			myBody = myBody & chr(10) & chr(13) & prepareWord(myDescomposicionDES2(1))
		end if
	next
	if myAntSis <> "" then myAlta = true
	myAntSis = ""
	myDescomposicionDES = split("/" & FrmDic("RelacionesDELDES") & "/","//")
	if ubound(myDescomposicionDES)-1 >0 then myBody = myBody & chr(10) & chr(13) & chr(10) & chr(13) & "TAREAS CON SOLICITUD DE BAJA:"
	for index=1 to ubound(myDescomposicionDES) - 1
		myDescomposicionDES2 = split(myDescomposicionDES(index),"##")
		if myDescomposicionDES2(0) = myAntSis then
			myBody = myBody & chr(10) & chr(13) & prepareWord(myDescomposicionDES2(1))
		else
			myAntSis = myDescomposicionDES2(0)
			myBody = myBody & chr(10) & chr(13) & replace(prepareWord(myDescomposicionDES2(0)),"a los que Tendr","a los que ya no Tendr")
			myBody = myBody & chr(10) & chr(13) & prepareWord(myDescomposicionDES2(1))
		end if
	next


	if myAntSis <> "" then myBaja = true

    if FrmDic("Accion") <> "" then
		if myAlta then myAccionLoad = "A"
		if myBaja then myAccionLoad = "D"
		if myBaja and myAlta then myAccionLoad = "M"
	end if

    if len(myBody) > 0 then
		myBody = myBody & chr(10) & chr(13) & chr(10) & chr(13) & chr(10) & chr(13) & replace("Si desea imprimir el reporte haga click en el siguiente link."," ",chr(32))
		myBody = myBody & chr(10) & chr(13) & chr(10) & "http://bai-vm-intra-1/ActisaIntra/AUPReporte2.asp?pIdPersona=" & FrmDic("IdPersona") & "&pOP=" & myAccionLoad
		call GP_ENVIAR_MAIL("SOLICITUD DE ABM DE PERFILES DE USUARIOS.",titu & myBody,"ABMPerfiles@toepfer.com",myRecipients)
    end if

function prepareWord(pWord)
dim rtrn
	rtrn = replace(pWord," ",chr(32))
	rtrn = replace(rtrn,"%20",chr(32))
	prepareWord = rtrn
end function
%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<Link REL=stylesheet href="CSS/calendarioPagos.css" type="text/css">
<Link REL=stylesheet href="CSS/Iwin.css" type="text/css">
<title>Intranet ActiSA - Asignación de Sistemas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="Scripts/channel.js"></script>  
<script language="JavaScript" src="Scripts/iwin.js"></script>  
<script language="JavaScript">
	var pp;
	var pedidos = new Array();
	var ch = new channel();
	function fcnResaltar(P_objFila)
	{
		P_objFila.style.color = "#0000FF";
		P_objFila.style.cursor = "hand";
	}
	function fcnNormal(P_objFila)
	{
		P_objFila.style.color= "#000000";
	}
	function fcnExpand(p_obj,p_img)
	{
		var myObj, myImg;
		myObj = document.getElementById(p_obj); 
		myImg = document.getElementById(p_img); 
		if (myObj.style.visibility == "hidden"){
				//Expansion.		 
				myObj.style.visibility = "visible";
				myObj.style.position = "relative";
				myImg.src = "images/TMinus.gif";
		}
		else{
				//Colapso.
				myObj.style.visibility = "hidden";
 				myObj.style.position = "absolute";
 				myImg.src = "images/Tplusik.gif";
		}
	}   
	function cargarLista(pSRO3, pSisDS, pTTKC, pTTDS, pValue, pAccion)
	{
		var myObj, myStr, myStrCC, myRegExp;
		var myObjDES, myStrDES, myStrCCDES, myRegExpDES;
		if (pAccion == 'A'){
			myObj = document.getElementById("Relaciones"); 
			myObjDES = document.getElementById("RelacionesDES"); 
		}
		else{
			myObj = document.getElementById("RelacionesDel"); 
			myObjDES = document.getElementById("RelacionesDELDES"); 
		}	
		myStr = myObj.value;
		var myStrAux = new String(); 
		var myStrAux2 = new String(); 
		var myStrCCAux = new String();
		myStrCCAux = "/" + pSRO3 + "=" + pValue + "/";
		myStrCC = new RegExp(pSRO3 + "=");
		//alert("Busco(" + myStrCC + ") en (" + myStr + ") dev(" + myStr.search(myStrCC, myStr) + ")")
		if (myStr.search(myStrCC, myStr) == -1){
			myStr = myStr.concat(myStrCCAux);
		}
		else{
			if (pValue == '1') {
				myStr = myStr.replace("/" + pSRO3 + "=" + pValue + "/","");
			}
			else{

					myStrAux = myStr.substring(0,myStr.search("/" + pSRO3 + "="));
					myStrAux2 = myStr.substring(myStr.indexOf("/" + pSRO3 + "=") + myStrCCAux.length,myStr.length);
					myStr = myStrAux + myStrCCAux + myStrAux2
			}
		}
		myObj.value = myStr;
		//Para las descripciones
		myStrDES = myObjDES.value;
		myStrCCDES = new RegExp(pSisDS + " ## " + pTTDS);	
		if (myStrDES.search(myStrCCDES, myStrDES) == -1){
			myStrDES = myStrDES.concat(myStrCCDES);
		}
		else{
			if (pValue != 'Y' && pValue != 'A' && pValue != 'U'){
				myStrDES = myStrDES.replace("/" + pSisDS + " ## " + pTTDS + "/","");
			}
		}
		myObjDES.value = myStrDES;
	}
	function cargarListaPre(pObjCmb, pSRO3, pSisDS, pTTKC, pTTDS, pAccion){
	    var indice = pObjCmb.selectedIndex; 
		var valor = pObjCmb.options[indice].value;
		cargarLista(pSRO3, pSisDS, pTTKC, pTTDS, valor, pAccion)
	}

	function cargarListaPrev(pSRO3, pSisDS, pTTKC, pTTDS, pAccion)

	{
		var myObj, myStr, myStrCC, myRegExp;
		var myObjDES, myStrDES, myStrCCDES, myRegExpDES;
		if (pAccion == 'A'){
			myObj = document.getElementById("Relaciones"); 
			myObjDES = document.getElementById("RelacionesDES"); 
		}
		else{
			myObj = document.getElementById("RelacionesDel"); 
			myObjDES = document.getElementById("RelacionesDELDES"); 
		}	
		
		myStr = myObj.value;
		myStrCC = new RegExp(pSRO3);
		if (myStr.search(myStrCC, myStr) == -1){
			myStr = myStr.concat(myStrCC);
		}
		else{
			myStr = myStr.replace("/" + pSRO3 + "/","");
		}
		myObj.value = myStr;
		//Para las descripciones
		myStrDES = myObjDES.value;
		myStrCCDES = new RegExp(pSisDS + " ## " + pTTDS);	
		if (myStrDES.search(myStrCCDES, myStrDES) == -1){
			myStrDES = myStrDES.concat(myStrCCDES);
		}
		else{
			myStrDES = myStrDES.replace("/" + pSisDS + " ## " + pTTDS + "/","");
		}
		myObjDES.value = myStrDES;
	}

	function loadAndSubmit(pAccion)
	{
		frmMain.Accion.value = pAccion;
		frmMain.submit();
	}   
	function printReport(pTipo, pAccion)
	{
		var myObjA, myObjD, msg;
		myObjA = document.getElementById("Relaciones"); 
		myObjD = document.getElementById("RelacionesDel"); 
		if (myObjA.value||myObjD.value){
			msg = "Existen modificaciones que aun no han sido guardadas! \nDesea visualizar la ventana de impresión de todas modos?";
			if (!confirm(msg)) {return;}
		}
		window.open (pTipo + ".asp?pIdPersona=" + frmMain.IdPersona.value + "&pOP=" + pAccion,"Reporte","toolbar=yes,menubar=yes,type=fullWndow,resizable=yes,scrollbars=1");
	}   
	function traerSector_callback(pSector)
	{		
		document.getElementById(pSector).innerHTML = ch.response(pedidos[pSector]);
	}
	function traerSector(pIdPersona, pSisKR, pSisDS, pImg, pStatus) 
	{			
		var flag = false;
		var element, myImg;
		myImg = document.getElementById(pImg);
		if (document.getElementById(pSisKR).innerHTML == "") {			
			//Buscar si ya esta cargado en el array
			for (element in pedidos){
				if (element==pSisKR){
					document.getElementById(pSisKR).innerHTML = pedidos[element];
					flag = true;
				}
			}
			//no esta cargado en el array pedirselo al canal
			if (!flag){	
				var link = "AUPGetSistemas.asp?IdPersona=" + pIdPersona + "&pSisKR=" + pSisKR + "&pSisDS=" + pSisDS + "&pEgreso=" + pStatus;;
				var param = "traerSector_callback(" + pSisKR + ")";
				ch.bind(link, param);
				pedidos[pSisKR] = ch.send();
			}	
			myImg.src = "images/TMinus.gif";
		}
		else{
			pedidos[pSisKR] = document.getElementById(pSisKR).innerHTML;
			//if (document.all){
			document.getElementById(pSisKR).innerHTML = "";
			//}
			myImg.src = "images/Tplusik.gif";
		}
	}   
	function openWin(pPath){
		window.open(pPath,"","top=1, left=1, scrollbars=yes,status=no,resizable=yes,toolbar=no,location=no,menu=no,width=700,height=700");
	}


	function mostrarEvento(p_obj) {
		 var x,y;
		 var vecCoords;
			 
		 //document.getElementById('ifrmDetalleEvento').src = 'AUPProcesos.asp';
         var divEvento = document.getElementById('divEvento');

		 vecCoords = findPos(p_obj);
		 divEvento.style.left = vecCoords[0] + 15 + 'px'; //15 es el width del span, pero no lo toma por HTML DOM
		 if ((vecCoords[1] - document.body.scrollTop) < 450) {
         	divEvento.style.top = vecCoords[1] + 'px';
		 } else {
            divEvento.style.top = 450 + document.body.scrollTop + 'px';
		 }
		 divEvento.className = 'evento visible';
   	}

    function ocultarEvento() {
		document.getElementById('divEvento').className = 'evento oculto';
	}
    function findPos(obj) {
		var curleft = curtop = 0;
		if (obj.offsetParent) {
			curleft = obj.offsetLeft
			curtop = obj.offsetTop
			while (obj = obj.offsetParent) {
				curleft += obj.offsetLeft
				curtop += obj.offsetTop
			}
		}
		return [curleft,curtop];
	}		
	function createPopUpWindow(){
		pp = new PopUpWindow('Lista de Procesos', 'AUPProcesos.asp', '340', '350');
	}
</script>
</head>
<body>
<form name="frmMain" action="AUPSistemas.asp" method="post">
<input type="hidden" id="Relaciones" name="Relaciones" size="150">
<input type="hidden" id="RelacionesDel" name="RelacionesDel" size="150">
<input type="hidden" id="RelacionesDES" name="RelacionesDES" size="150">
<input type="hidden" id="RelacionesDELDES" name="RelacionesDELDES" size="150">
<input type="hidden" id="Accion" name="Accion">
<input type="hidden" name="IdPersona" value=<%=FrmDic("IdPersona")%>>
<%=GF_TITULO("Usuarios.gif", "<b>" & myApellido & ", " & myNombre & " (" & myUsuario & ")" & " - Legajo: " & myNroLegajo & "</b>")%>
<%  
'Obtener lista de sistemas
strSQL= "select * from MG where mg_km='TS' order by mg_ds"
call GF_BD_CONTROL (rsSistemas,oConn,"OPEN",strSQL)
%>
<table align="center" width="80%" border="0" cellspacing=0> 
<TR>
 <TD> 
 		<a href="AUPSectores.asp"><img align=absMiddle src="images/Anterior.gif" alt="<% =GF_TRADUCIR("Volver") %>" border="0">&nbsp;<% =GF_TRADUCIR("Volver") %></a> |		
			<a href="javascript:loadAndSubmit('G')"><img align=absMiddle src ="images/Guardar.gif" name="btnGuardar" alt="<% =gf_traducir("Guardar")%>">&nbsp;<% =GF_TRADUCIR("Guardar") %></a> |
		<% if ucase(session("AUPUSER"))="ADMIN" then %>
			<!--<a href="javascript:printReport('AUPReporte','P')"><img align=absMiddle src="images/printer.gif" alt="<% =GF_TRADUCIR("Imprimir Reporte 1") %>" border="0">&nbsp;<% =GF_TRADUCIR("Reporte 1") %></a> |-->
			<a href="javascript:printReport('AUPReporte2','P')"><img align=absMiddle src="images/printer.gif" alt="<% =GF_TRADUCIR("Imprimir Reporte") %>" border="0">&nbsp;<% =GF_TRADUCIR("Reporte") %></a> |
		<% end if %>
		<a target="_new" href="AUPAuditoria.asp?pUser=<%=FrmDic("IdPersona")%>&pHistorica=S" title="<%=GF_Traducir("Reporte para Auditoria")%>"><img align=absMiddle src="Images/history.gif">&nbsp;<% =GF_TRADUCIR("Reporte Auditoria") %></a>		
 </td>
</TR> 
</table>

<table border="0" cellspacing="0" cellpadding="0" width="80%" align="center" bordercolor="#fffaf0">
	<% 
	while not rsSistemas.eof 
		%>		
		<tr style="cursor:pointer;" onClick="traerSector(<%=FrmDic("IdPersona")%>,<%=rsSistemas("MG_KR")%>,'<%=rsSistemas("MG_DS")%>','IMG_<%=rsSistemas("MG_KR")%>','<%=myEgresoValido%>')">
   			<td class="titu_Header" align="left">
				<img id="IMG_<%=rsSistemas("MG_KR")%>" style="position:absolute;" src="Images/TPlusik.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=GF_TRADUCIR(rsSistemas("MG_DS"))%>
			</td>
		</tr>
		<tr>
			<td align="center">
				<div id="<%=rsSistemas("MG_KR")%>"></div>
			</td>
		</tr>						
		<%	
		rsSistemas.movenext
	wend 
	call GF_BD_CONTROL (rsSistemas,oConn,"CLOSE",strSQL)	
	%>
</table>
</form>
<div id="divEvento" class="evento oculto">
   	<!--<iframe id="ifrmDetalleEvento" src="aupProcesos.asp?" frameborder=0 scrolling="auto" style="border-width:1px;border-style:solid;width:400px;height:150px;"></iframe>-->
   	<iframe id="ifrmDetalleEvento" src="aupProcesos.asp?" frameborder=0 scrolling="auto" style="border-width:1px;border-style:solid;width:340px;height:340px;"></iframe>
</div>
</body>
</html>
<%
function esCredencialNuevo(p_nroCredencial, p_idPersona)
	dim sql, oConn, rsCredencial
	sql="Select NroCredencial from Profesionales where idProfesional=" & p_idPersona
	'response.write sql
	'response.end
	call GF_BD_CONTROL (rsCredencial,oConn,"OPEN",sql)
	if rsCredencial.eof then
		response.write "no esta registrado este profesional en el sistema"
	else
		if isNull(rsCredencial("NroCredencial")) or not (CLng(rsCredencial("NroCredencial"))=CLng(p_nroCredencial)) then
			esCredencialNuevo = true
		else
			esCredencialNuevo = false
		end if
	end if
end function
function guardarCredencialNuevo(p_nroCredencial, p_idPersona)
	dim sql, oConn, rsCredencial
	sql="update Profesionales set NroCredencial =" & p_nroCredencial & " where IdProfesional=" & p_idPersona
	call GF_BD_CONTROL (rsCredencial,oConn,"EXEC",sql)
end function
%>