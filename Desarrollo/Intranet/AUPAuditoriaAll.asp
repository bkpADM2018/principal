<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosseguridad.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 22/01/2008                         -->
'<!--               Pagina     : AUPAuditoria.ASP                   -->
'<!--               Descripcion: Listado para Auditoria             -->
'<!--               Modificacion: Henzel Pavlo              -->
'<!--               Fecha      : 22/01/2008                         -->
'<!------------------------------------------------------------------->
dim rsAuditoria, oConn, strSQL, myColorIndex, myColor
Dim Sector,Desde,Hasta, myWhere,primero,FechaDesde,FechaHasta,dtConf

Sector   = GF_Parametros7("sector","",6)
Desde    = GF_Parametros7("Desde","",6)
Hasta    = GF_Parametros7("Hasta","",6)
Detalles = GF_Parametros7("Detalles","",6)

FechaDesde  = GF_DTE2FN(Desde)
FechaHasta  = GF_DTE2FN(Hasta)

if Sector <> "" or Desde <> "" or Hasta <> "" or detalles <> "false" then primero = 1
%>
<html>
<head>
<Link REL=stylesheet href="CSS/ActisaIntra-1.css" type="text/css">
<link rel="stylesheet" type="text/css" media="all" href="CSS/calendar-win2k-2.css" title="win2k-2" />
<title>Intranet ActiSA - Reporte de Situación Actual de Confirmacion de Permisos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Script del calendario -->
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-<% =GF_GET_IDIOMA() %>.js"></script>
<script language=javascript>
	/**
	* Funciones adicionales para el calendario
	*/
	function SeleccionarCalDesde(cal,date)
	{
		
	    var str= new String(date);
	    str = str.replace("-","");
	    str = str.replace("-","");
		cal.hide();
		document.getElementById("Desde").value = date
		
		
    }
	function SeleccionarCalHasta(cal,date)
	{
		
	    var str= new String(date);
	    str = str.replace("-","");
	    str = str.replace("-","");
		cal.hide();
		document.getElementById("Hasta").value = date
		
		
    }
	
	function CerrarCal(cal)
	{
		cal.hide();
	}
	
	function MostrarCalendario(p_objID, funcSel) {
		

		var dte= new Date();		    	    
		var elem= document.getElementById(p_objID);
		

		
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, funcSel, CerrarCal);
	    
		cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	function Sumitir(){
		//var sector  = document.getElementById("sector").value;
		var sectores = getMultiple(document.getElementById("Sectores"));
		var Desde   = document.getElementById("Desde").value;
		var Hasta   = document.getElementById("Hasta").value;
		var Detalles = document.getElementById('Detalles').checked;
		
		
		f1=new Date(Desde);
		f2=new Date(Hasta);
		if (f1>f2) {
			alert("La fecha Desde debe ser menor que la fecha Hasta")  
		}else{
			if (sectores==""){
				alert("Debe seleccionar almenos un sector")  
			}else{
			document.location.href="AUPAuditoriaPrint.asp?sector=" +sectores+ "&Desde="+Desde+"&Hasta="+Hasta+"&Detalles="+Detalles
			}			
		}
		
		
		
		
		
		
		//abrir("AUPAuditoriaPrint.asp?sector=" +sector+ "&Desde="+Desde+"&Hasta="+Hasta+"&Detalles="+Detalles,0,1,1,1,1,1,1,100,100,100,100,1);
		//window.open("AUPAuditoriaPrint.asp?sector=" +sector+ "&Desde="+Desde+"&Hasta="+Hasta+"&Detalles="+Detalles)
	}
	
	function abrir(direccion, pantallacompleta, herramientas, direcciones, estado, barramenu, barrascroll, cambiatamano, ancho, alto, izquierda, arriba, sustituir){
    var opciones = "fullscreen=" + pantallacompleta +
                 ",toolbar=" + herramientas +
                 ",location=" + direcciones +
                 ",status=" + estado +
                 ",menubar=" + barramenu +
                 ",scrollbars=" + barrascroll +
                 ",resizable=" + cambiatamano +
                 ",width=" + ancho +
                 ",height=" + alto +
                 ",left=" + izquierda +
                 ",top=" + arriba;
	var ventana = window.open(direccion,"venta",opciones,sustituir);
	}           

	function getMultiple(ob) { 
		var aux = "";

		
		while (ob.selectedIndex != -1) { 
			if (ob.selectedIndex != -1) {
				if (aux==""){
					aux = ob.options[ob.selectedIndex].value;
				}else{
					aux = aux + "," + ob.options[ob.selectedIndex].value;
				}
				
			}
			ob.options[ob.selectedIndex].selected = false; 

		}
		 return aux;
	}	
	
</script>
</head>
<body class="print" <%if primero <> 1 then%> onload="alert('Recuerde configurar la página de la siguiente manera: \n -Orientación: Horizontal\n -Encabezado y Pie de página: En Blanco\n -Bordes: 1 cm.');"<%end if%>>
<form name="frmMain" method="post">
<%
Response.Write GF_TITULO("Usuarios.gif","Reporte de Situación Actual de confirmaciones de permisos")
'strSQL = "SELECT M.MG_KC AS USERKC, M.MG_DS AS USERDS, C.MMTOCONF as DATECONF, M2.MG_KR AS SECTORKR, M2.MG_DS AS SECTORDS FROM CONFIRMACIONESPERMISOS C INNER JOIN MG M ON C.KRULTIMOUSUARIO=M.MG_KR INNER JOIN MG M2 ON C.KRSECTOR=M2.MG_KR ORDER BY M2.MG_DS"
if Sector <> "" then
		if isNumeric(Sector) then
			call AgregarWhere(myWhere,"m2.mg_kr = " & Sector )
		else
			call AgregarWhere(myWhere,"m2.mg_ds like '%" & Sector & "%'")
		end if
end if
If Desde <> "" then
		call AgregarWhere(myWhere,"c.mmtoconf >= '" & FechaDesde & "'")
end if
If Hasta <> "" then
		call AgregarWhere(myWhere,"c.mmtoconf <= '" & FechaHasta & "'")
end if

strSQL = "SELECT   m.mg_kc    AS userkc, " & _
		"	         m.mg_kr    AS userkr, " & _
		"	         m.mg_ds    AS userds, " & _
		"	         c.mmtoconf AS dateconf, " & _
		"	         m2.mg_kr   AS sectorkr, " & _
		"	         m2.mg_ds   AS sectords " & _
		"	FROM     confirmacionespermisos c " & _
		"	         INNER JOIN mg m " & _
		"	           ON c.krultimousuario = m.mg_kr " & _
		"	         INNER JOIN mg m2 " & _
		"	           ON c.krsector = m2.mg_kr " & _
		"	         INNER JOIN (SELECT   krsector, " & _
		"	                              Max(mmtoconf) AS maxmmto " & _
		"	                     FROM     confirmacionespermisos c " & _
		"	                     GROUP BY krsector) s1 " & _
		"	           ON s1.krsector = c.krsector " & _
		"	              AND (s1.maxmmto = c.mmtoconf " & _
		"	                    OR s1.maxmmto IS NULL) " & _
		myWhere & _
		"	ORDER BY m2.mg_ds "


'Response.Write strSQL
call GF_BD_CONTROL (rsAuditoria,oConn,"OPEN",strSQL)
'---------------------------------------------------
function AgregarWhere(byref p_where,p_agregar)

    if p_where = "" then
        p_where = " WHERE " & p_agregar & vbCrLf
    else
        p_where = p_where & " AND " & p_agregar  & vbCrLf
    end if

end function

dim rsSectores
strsql = "select * from mg where mg_km='ss' order by mg_ds"
call GF_BD_CONTROL (rsSectores,oConn,"OPEN",strSQL)

%>
<FORM method='post' action='AUPAuditoriaAll.asp'>
	<INPUT type='hidden' name='primero' VALUE="<%=Request("primero")%>">
	<table cellpadding=0 cellspacing=0 align=center width='250px'>
		<tr>
			<td width="8"><img border=0 src="images/marco_r1_c1.gif" width="8" height="8"></td>
			<td background="images/marco_r1_c2.gif"><img src="images/marco_r1_c2.gif"></td>
			<td width="8"><img border=0 src="images/marco_r1_c3.gif" width="8" height="8"></td>
		</tr>
			<td background="images/marco_r2_c1.gif"></td>
			<td>
			<TABLE align='left' cellspacing='2' cellpadding='2' border = 0 width='100%' >
				<TR bgcolor='green'>
					<TD align='center' colspan ='2'>
						<font color='white'><B>Busqueda</B></font>
					</TD>
				</TR>
				<TR>
					<TD colspan='2' align='Center'>
						Sector <BR>
						
						<select multiple="multiple" size="5" id='Sectores' >
							<option value="-1">Todos</option>
							<%while not rsSectores.eof %>
							<option value="<%=rsSectores("MG_KR")%>"><%=rsSectores("MG_DS")%></option>
							<%
							rsSectores.movenext
							wend%>
						</select>
						
					</TD>
				</TR>
				<TR align='center'>
					<TD>
						<table border='0'>
							<TR>
								<TD colspan='2' align ='center'>
									Fecha Desde
								</TD>
							</TR>
							<TR>
								<TD>
									<div id="imgCal" style="cursor:hand;" ><IMG src='images/date.gif' border='0' alt='Cambiar Fecha' title='Cambiar Fecha' onClick="javascript:MostrarCalendario('Desde',SeleccionarCalDesde);">
									</div>
								</TD>
								<TD>	
									
									<INPUT type='text' id='Desde' name='Desde' VALUE="<%=Request("Desde")%>" disabled="disabled">
									<INPUT type='hidden' id='FechaDesde' name='FechaDesde' VALUE="<%=Request("FechaDesde")%>">
								</TD>
							</TR>
						</table>
					</TD>
				</TR>
				<TR align = 'center'>
					<TD>
						<table border='0'>
							<TR>
								<TD colspan='2' align = 'center'>
									Fecha Hasta
								</TD>
							</TR>
							<TR>
								<TD>
									<div id="imgCal2" style="cursor:hand;" ><IMG src='images/date.gif' border='0' alt='Cambiar Fecha' title='Cambiar Fecha' onClick="javascript:MostrarCalendario('Hasta',SeleccionarCalHasta);">
									</div>
								</TD>
								<TD>	
									<INPUT type='text' id='Hasta' name='Hasta' VALUE="<%=Request("Hasta")%>" disabled="disabled">
									<INPUT type='hidden' id='FechaHasta' name='FechaHasta' VALUE="<%=Request("FechaHasta")%>">
								</TD>
							</TR>
						</table>
						
						
					</TD>
				</TR>
				<TR align='center'>
					<TD>
						<INPUT Value="1" type='checkbox' name='Detalles' id='Detalles' <%if Detalles then response.write "Checked" %>> Mostrar detalles
					</TD>
				</TR>
				<TR align='center' >	
					
					<TD colspan=2>
						<INPUT type='button' name='btnSend' value='<%=GF_TRADUCiR("Buscar")%>' onclick="javascript:Sumitir();">
					</TD>
					
				</TR>
				
	        </TABLE>
			</td>
			<td background="images/marco_r2_c3.gif"></td>
		</tr>
		<tr>
			<td width="8"><img border=0 src="images/marco_r3_c1.gif" width="8" height="8"></td>
			<td background="images/marco_r3_c2.gif"><img src="images/marco_r3_c2.gif"></td>
			<td width="8"><img border=0 src="images/marco_r3_c3.gif" width="8" height="8"></td>
		</tr>
	</table>
</form>	
<BR>

</form>
</body>
</html>