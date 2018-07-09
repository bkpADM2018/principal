<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosGemini.asp"-->
<!--#include file="Includes/procedimientosRoles.asp"-->
<%
'--------------------------------------------------------------------------------------------------------------------------------------
Function cargarFirmas(pIdTarea)
	Dim rsFirmas
	Set rsFirmas = getIssueSign(pIdTarea)
	'Vienen ordenadas por secuencia
    if (not rsFirmas.eof) then
        '1) Tester / Tecnico
		cdTester = rsFirmas("CDUSUARIO")
		txTester = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MMTO"))
        rsFirmas.MoveNext()
    end if
    if (not rsFirmas.eof) then
        '2) Solicitante
	    cdSolicitante = rsFirmas("CDUSUARIO")
		txSolicitante = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MMTO"))
        rsFirmas.MoveNext()
    end if
    if (not rsFirmas.eof) then
        '3) Publicador
		cdPublicador = rsFirmas("CDUSUARIO")
		txPublicador = armarTextoPlanoFirma(rsFirmas("HKEY"), rsFirmas("MMTO"))
        rsFirmas.MoveNext()
	end if
End Function

'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
Dim idTarea, accion, cdTester,txTester,cdSolicitante,txSolicitante,cdPublicador,txPublicador,solicitante, errFirma

idTarea = GF_Parametros7("idTarea",0,6)
accion = GF_PARAMETROS7("accion","",6)
errFirma = GF_PARAMETROS7("errFirma","",6)

if (errFirma <> "") then Call setError(errFirma)

Call GP_CONFIGURARMOMENTOS

Set rs = getIssue(idTarea)
if (rs.eof) then 
    response.redirect "comprasAccesoDenegado.asp"
else
    tituloTarea = Trim(rs("SUMMARY"))
    estadoTarea = rs("ISSUESTATUSID")
    geminiUser = UCase(Trim(rs("USERNAME")))
    mailSolicitante = Trim(rs("MAILUSUARIO"))
    issueBody = rs("LONGDESC")
    Call cargarFirmas(idTarea)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Firma de Tarea") %></title>
<link rel="stylesheet" href="css/main.css" type="text/css" />
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
    
	var link = "geminiFirmarIssue.asp?idTarea=<% =idTarea %>";
	var hkey0 = new Hkey('hk0', link , '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', link, '<% =HKEY() %>', 'check_callback()');
	<%  if (CInt(rs("templateid")) = TEMPLATE_BUG_TRACKING) then %>
	var hkey2 = new Hkey('hk2', link , '<% =HKEY() %>', 'check_callback()');
    <%  end if %>
	function check_callback(resp) {
        if (resp != "<% =RESPUESTA_OK %>") document.getElementById("errFirma").value = resp;
		document.getElementById("frmSel").submit();
	}
	
	function bodyOnLoad(){
        hkey0.start();
		hkey1.start();
		hkey2.start();
	}					
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<div class="tableaside size100" style="display:inline-table;"> 
	<div class="tableasidecontent"><% call showErrors() %></div>
	<h3><%=GF_Traducir("Datos de la Tarea")%></h3>    
    <table class="datagrid" width="90%" align="center">
        <thead>
            <tr>
                <th align="center" style="width:10%;"><% =GF_TRADUCIR("Tarea") %></th>
                <th align="center" style="width:45%;"><% =GF_TRADUCIR("Titulo") %></th>
                <th align="center" style="width:15%;"><% =GF_TRADUCIR("Estado") %></th>
                <th align="center" style="width:30%;"><% =GF_TRADUCIR("Desarrollador/Técnico") %></th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td align="center"><%= getGeminiTaskCode(idTarea) %></td>
                <td align="center"><%= tituloTarea %></td>
                <td align="center"><%= getIssueStatusDs(estadoTarea) %></td>
                <td align="center"><%= getUserAndIssue(idTarea) %></td>
            </tr>
            <tr>
                <td colspan="4"><% =issueBody %></td>
            </tr>
        </tbody>
    </table>
</div>
<form name="frmSel" id="frmSel" method="post" action="geminiIssueFirma.asp">
    <div class="tableaside size100">     
    <h3><%=GF_Traducir("Firmas")%></h3>    
        <table border=0 width='90%' class="datagrid" align="center">
            <thead>
		        <tr>
		        <%  if (CInt(rs("templateid")) = TEMPLATE_BUG_TRACKING) then %>
			        <th align='center' style="width:33%"><font size='2'><% =GF_Traducir("Tester") %></font></th>
				    <th align='center' style="width:33%"><font size='2'><% =GF_Traducir("Solicitante") %></font></th>
				    <th align='center' style="width:33%"><font size='2'><% =GF_Traducir("Publicador") %></font></th>
                <%  else    %>
                    <th align='center' style="width:33%"><font size='2'><% =GF_Traducir("T&eacute;cnico") %></font></th>
				    <th align='center' style="width:33%"><font size='2'><% =GF_Traducir("Solicitante") %></font></th>				        				    
                <%  end if  %>
                </tr>
            </thead>
            <tbody>
                <tr>
                 <% Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsEst, "TBLESTADOSTRANSICION_GET_BY_PARAMETERS", SEC_SYS_GEMINI &"||"& RES_GEM_TAREAS &"||"& estadoTarea &"||"& EVENTO_FIRMA &"||||"& getRolFirma(Session("Usuario"),SEC_SYS_GEMINI) &"||") %>

                 <td align="center">
                    <% if (Trim(txTester) <> "") then %>
		                <img src="images/firmas/<%=obtenerFirma(cdTester)%>"><br><%=getUserDescription(cdTester)%><BR><%=txTester%>
	                <% else %>
                        <% if (not rsEst.Eof) then %>
                               <div id='hk0'></div>
                           <%  rsEst.MoveNext()
                           else %>
                               <br><br/><br><br/>
                        <% end if %>                        
                    <% end if %>
                  </td>
                  <td align="center">
                    <% if (Trim(txSolicitante) <> "") then %>
		                <img src="images/firmas/<%=obtenerFirma(cdSolicitante)%>"><br><%=getUserDescription(cdSolicitante)%><BR><%=txSolicitante%>
	                <% else %>                        
                             <% if (((CInt(estadoTarea) = ISSUE_STATUS_SIGN) or (CInt(estadoTarea) = ISSUE_STATUS_HD_APR_TECH)) and ((geminiUser = Session("Usuario"))) or (mailSolicitante = Ucase(Trim(getUserMail(Session("Usuario")))))) then %>
                                  <div id='hk1'></div>
                             <% end if %>                  
                     <% end if %>
                   </td>
                   <%  if (CInt(rs("templateid")) = TEMPLATE_BUG_TRACKING) then %>
                       <td align="center">
                       <% if (Trim(txPublicador) <> "") then %>
		                    <img src="images/firmas/<%=obtenerFirma(cdPublicador)%>"><br><%=getUserDescription(cdPublicador)%><BR><%=txPublicador%>
	                   <% else %>
                          <% if (not rsEst.Eof) then %>
                                <div id='hk2'></div>
                            <%  rsEst.MoveNext()
                             else %>
                                 <br><br/><br><br/>
                          <% end if %>
                       <% end if %>
                       </td>
                   <% end if %>
                </tr>
            </tbody>
	    </table>
    </div>
    <input type="hidden" name="errFirma" id="errFirma" />
    <input type="hidden" name="accion" id="accion" value="<% =ACCION_CONFIRMAR %>" />
    <input type="hidden" id="idTarea" name="idTarea" value="<%=idTarea%>" />
</form>
</body>
</html>