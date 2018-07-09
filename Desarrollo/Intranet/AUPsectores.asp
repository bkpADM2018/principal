<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 20/12/2007                         -->
'<!--               Pagina     : AUPSecores.ASP                     -->
'<!--               Descripcion: Listado de personal por sector     -->
'<!------------------------------------------------------------------->
ProcedimientoControl "AUPSEC"
Dim strSQL, rs, rsPersonas, oConn, strUbicacion, myProfesional, myProfesionalDS
dim myPersona, mySector, mySectorDS, MySectorARR, myEsJefeDeKR
dim srEXEC, srTREE, srUSER, index, myClassPend, myTextConf, krUruguay, dsUruguay
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary") 
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
'---------------------------------------------------------------------------------------
Function GF_MGSR2(P_o1kr, p_o2kr, P_o3kr, byref P_3okr)
DIM CON, RS, strSQL
  strSQL = "SELECT * FROM RelacionesConsulta where sro1kr = " & p_o1kr & "  and sro2kr = " & p_o2kr & " and sro3kr = " & p_o3kr & " order by SRMMDT desc"
  GF_BD_Control rs,con, "OPEN", strSQL
  P_3OKR = ""
  GF_MGSR2 = false
  if not RS.EOF then 
	if rs("srValor") <> "*" then
		p_3okr = rs("sr3okr")
		GF_MGSR2 = true
	end if	
  end if	   
  GF_BD_Control rs,con,"CLOSE",strSQL 
end function
'---------------------------------------------------------------------------------------
'-- COMIENZO DE PROGRAMA
'---------------------------------------------------------------------------------------
session("AUPUSER") = ""
	'call GF_MGC ("UP", session("Usuario"), srUSER, "")
	'call GF_MGC ("SR", "EXEC", srEXEC, "")
	'call GF_MGC ("UC", "TREE001", srTREE, "")
	if (session("Usuario") = "JAS") then 
		myProfesional = "ALL"
		myProfesionalDS = "ADMINISTRADOR"
		session("AUPUSER") = "ADMIN"
	else
		'KR del usuario logueado	
		call GF_MGC ("SG", session("Usuario"), myProfesional, myProfesionalDS)
		call GF_MGC ("SR", "EsJefeDe", myEsJefeDeKR, "")
		'Response.Write "(" & myEsJefeDeKR & ")"
		'call GF_MGC ("SG", "JAS", myProfesional, myProfesionalDS)
	end if	
'---------------------------------------------------------------------------------------
function getConfPendientes(pSector)
dim strSql, rs, cn, rtrn
rtrn = 0
strSql = "Select ConfPendientes from ConfirmacionesPermisos where ConfPendientes > 0 and krSector =" & pSector & " order by mmtoenvio desc"
call GF_BD_CONTROL (rs,oConn,"OPEN",strSQL)
	if not rs.eof then rtrn = rs("ConfPendientes")
call GF_BD_CONTROL (rs,oConn,"CLOSE",strSQL)
getConfPendientes = rtrn
end function
'---------------------------------------------------------------------------------------
%>

<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<title>Intranet ActiSA - Empleados de los sectores a los que pertenece <%=myProfesionalDS%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="Scripts/channel.js"></script>  
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script language="JavaScript">
	var pedidos = new Array();
	var ch = new channel();
	
	function fcnResaltar(P_objFila)
	{
		P_objFila.style.background= "#d3d3a3"
		P_objFila.style.cursor="hand";
	}
	function fcnNormal(P_objFila)
	{
		P_objFila.style.background= "#dcdcdc"
	}
	
	function fcnResaltar2(P_objFila)
	{
		P_objFila.style.background = "#ffcc33";
	}
	function fcnNormal2(P_objFila)
	{
		P_objFila.style.background= "#FFFFFF"	
	}	
	function loadAndSubmit(pPersona)
	{
		frmMain.IdPersona.value = pPersona;
		frmMain.submit();
	}
	function traerSector_callback(pSector, pKey) {		
		document.getElementById(pKey).innerHTML = ch.response(pedidos[pKey]);
		pngfix();
	}
	function traerSector(pSector,pImg, pModo) {	
		var flag = false;
		var element, myImg, myKey;
		myKey = pSector + pModo;
		myImg = document.getElementById(pImg);		
		if (document.getElementById(myKey).innerHTML == "") {			
			//Buscar si ya esta cargado en el array
			var aux = new String();
			var des = new Array();
			for (element in pedidos){
				if (element==myKey){
					document.getElementById(myKey).innerHTML = ch.response(pedidos[myKey]);
					flag = true
				}
			}
			//no esta cargado en el array pedirselo al canal
			if (!flag){	
				var link = "AUPGetSector2.asp?SECTOR=" + pSector + "&MODO=" + pModo;
				var param = "traerSector_callback(" + pSector + ",'" + myKey + "')";
				ch.bind(link, param);
				pedidos[myKey] = ch.send();
			}	
			if (pModo=='F') { myImg.src = "images/TMinus.gif"; }
		}
		else{
			document.getElementById(myKey).innerHTML = ""
			if (pModo=='F') { myImg.src = "images/Tplusik.gif"; }
		}
	}
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, "");
		tb.addButton("printer1.gif", "Confirmaciones", "irRepConf()");		
		tb.draw();	
	}
	function irRepConf(){
		window.open("AUPAuditoriaAll.asp");
		//location.href = "AUPAuditoriaAll.asp"
	}
</script>

</head>
<body onLoad="bodyOnLoad()">
<form name="frmMain" action="AUPSistemas.asp" method="post">
<input type=hidden name="IdPersona">
<% =GF_TITULO("Usuarios.gif","Empleados de los sectores a los que pertence: <b>" & myProfesionalDS & "</b>") %>
<div id="toolbar"></div>
<br>
<table border="0" cellspacing="0" cellpadding="0" width="80%" align="center">
<%  
	
	''Obtengo el KR de Uruguay - No se debe mostrar Uruguay!!!!
	Call GF_MGKS("SS", "20", krUruguay, dsUruguay)
	'Obtener sector del usuario logueado 
	if myProfesional = "ALL" then
		'strSQL = "select Sector, C.confPendientes as ConfPendientes, mg_ds as SectorDS from Profesionales P inner join MG on P.sector=MG.mg_kr left join ConfirmacionesPermisos C on P.sector = C.krSector group by sector, mg_ds, confPendientes order by mg_ds"
		strSQL= "select Sector, mg_ds as SectorDS from Profesionales P inner join MG on P.sector=MG.mg_kr group by sector, mg_ds order by mg_ds"
	else		
		'JAS- Antes estaba estas: strSQL= "select Sector, mg_ds as SectorDS from Profesionales P inner join MG on P.sector=MG.mg_kr where idProfesional=" & myProfesional
		strSQL= "select sro3kr as Sector, sro3ds as SectorDS from RelacionesConsulta where sro1kr=" & myEsJefeDeKR & " and sro2kr=" & myProfesional
	end if	
	'Response.Write strsql
	call GF_BD_CONTROL (rs,oConn,"OPEN",strSQL)
	mySector = 0
	if rs.eof then 
		Response.Write "<center><font color='red'><b>No tiene asigando ningún sector el usuario logueado</b></font></center>"
		Response.End 
	end if	
	while not rs.eof	
		mySector=rs("Sector")
		mySectorDS=rs("SectorDS")	
		if (CLng(krUruguay) <> CLng(mySector)) then
			myConfPendiente = getConfPendientes(CLng(mySector))
			if myConfPendiente = "0" then
				myClassPend = "titu_headerW"
				myTextConf = "&nbsp;"
			else
				myClassPend = "titu_headerX"
				myTextConf = "[Confirmar Permisos]"
			end if
			%>
			
				<tr class="<%=myClassPend%>" id="TR(<%=MySector%>)">
					<td onClick="traerSector(<%=MySector%>,'IMG_<%=MySector%>','F');" width="73%" style="cursor:pointer;" align="left">
						<img style="position:absolute;" id="IMG_<%=MySector%>" src="Images/TPlusik.gif">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=GF_TRADUCIR(mySectorDS)%>
					</td>
					<td>
							<a href="AUPAuditoria.asp?pSector=<%=MySector%>&p_Responsable=<%=myProfesional%>&p_accion=CONFIRMA"><font color="white" class="small"><%=myTextConf%></font></a>
					</td>
					<td style="cursor:pointer;" align=right>
						<!--<a onClick="traerSector(<%=MySector%>,'IMGE_<%=MySector%>','E');"><img title="Ver externos" style="position:absolute;" id="IMGE_<%=MySector%>" src="Images/kopeteaway.png"></a>-->
						&nbsp;&nbsp;&nbsp;&nbsp;
						<a onClick="traerSector(<%=MySector%>,'IMGV_<%=MySector%>','V');"><img title="Ver empleados dados de baja en el sector: <%=GF_TRADUCIR(mySectorDS)%>" style="position:absolute;" id="IMGV_<%=MySector%>" src="Images/button_cancel.png"></a>
						&nbsp;&nbsp;&nbsp;&nbsp;


					<% if myProfesional = "ALL" then %>
						<a target="_new" href="AUPAuditoria.asp?pSector=<%=MySector%>" title="<%=GF_Traducir("Reporte para Auditoria por Sector")%>"><img src="Images/printer1.gif"></a>
					<% end if %>
					</td>
				</tr>
				<tr>
					<td colspan="3">
						<div id="<%=MySector%>F"></div>
					</td>
				</tr>
				<tr>
					<td colspan="3">
						<div id="<%=MySector%>V"></div>
					</td>
				</tr>
				<tr>
					<td colspan="3" height="1px"></td>
				</tr>	
			
		<%
		end if
		rs.movenext
	wend
	call GF_BD_CONTROL (rs,oConn,"CLOSE",strSQL)	
%>	
</table>	
</form>
</body>
<script language="javascript" src="scripts/pngfix.js"></script>

</html>
