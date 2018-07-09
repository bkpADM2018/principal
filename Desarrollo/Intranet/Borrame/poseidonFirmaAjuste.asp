<!--#include file="Includes/procedimientosHKEY.asp"-->

<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Poseidon/ajusteAutorizacionPrint.asp"-->
<%
'----------------------------------------------------------------------------------------------
' Función:	  sendEmailDraftAuthorized
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  04/11/2013
' Objetivo:	  Enviar mail notificando que el Director ha firmado el Draft Survey, en esta instancia se cierra el proceso 
'			  de autorización. La lista que recibirán el mail se encuentra en una Lista de Correo, la constante
'			  para identificar a la lista es:
'			  		LISTA_DRAFT_AUTORIZADO
'			  Además en el mail va a ir ajunto el archivo PDF autorizado
' Parametros:
'			  [int]		pIdAjuste - Id Ajuste	
'			  [string]	pCdAjuste - Codigo Ajuste
'			  [int]		pPto	  - Puerto
' Devuelve:
'			  -
'----------------------------------------------------------------------------------------------
Function sendEmailDraftAuthorized(pIdAjuste, pCdAjuste, pPto)
	Dim pathPDF, asunto, mensaje, origen, auxIdDivision, listaMails
	pathPDF = armadoPDF(pPto,pIdAjuste,PDF_FILE_MODE)
	select case Ucase(pPto)
	    case TERMINAL_ARROYO:
	       auxIdDivision  = getDivisionID(CODIGO_ARROYO)
        case TERMINAL_TRANSITO:
	       auxIdDivision  = getDivisionID(CODIGO_TRANSITO)
        case TERMINAL_PIEDRABUENA:
	       auxIdDivision  = getDivisionID(CODIGO_PIEDRABUENA)
    end select
	Set rs = getListMail(auxIdDivision, LISTA_DRAFT_AUTORIZADO)
	If not rs.Eof then
		asunto  = GF_TRADUCIR("DRAFT SURVEY AUTORIZADO") & " - Buque: " & gDSBuque
		mensaje = GF_TRADUCIR("El ajuste por "& getDsCodigoAjustePuerto(pCdAjuste) &" - N°: "& pIdAjuste &" fué autorizado por todos los responsables, se adjunta el informe.")
		origen  = obtenerMail(CD_TOEPFER)
		while not rs.Eof
			if(Len(Trim(rs("EMAIL"))) > 0)then listaMails = listaMails & Trim(rs("EMAIL")) & ";"
			rs.MoveNext()
		wend		
		if(Len(listaMails) > 0)then 
			'Saco el ultimo punto y coma.
			listaMails = left(listaMails, len(listaMails)-1)
			'Envio los mails.
			Call GP_ENVIAR_MAIL_ATTACHMENT(asunto, mensaje, origen, listaMails, pathPDF)
		end if
	end if
End Function
'----------------------------------------------------------------------------------------------
' Función:	  esUltimaFirmaAjustePto
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  23/08/2013
' Objetivo:	  Comprueba si queda pendiente una firma para un ajuste
' Parametros: 
'			  pIdAjuste  [int]
'			  pPto  [string]
' Devuelve:
'			  True  : si se completaron todas las firmas 
'			  False : si quedan firmas pendientes
'----------------------------------------------------------------------------------------------
Function esUltimaFirmaAjustePto(pIdAjuste, pPto)
	Dim strSQL,rs,rtrn
	rtrn = false
	strSQL = "SELECT COUNT(*) CANT FROM TBLAJUSTESFIRMAS WHERE IDAJUSTE = "& pIdAjuste &" AND HKEY IS NULL"
	Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	if cdbl(rs("CANT")) = 0 then rtrn = true	
	esUltimaFirmaAjustePto = rtrn
End Function
'----------------------------------------------------------------------------------------------
' Función:	  firmarAjustePto
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  23/08/2013
' Objetivo:	  Graba la firma del ajuste. Ademas si es la ultima firma actuliza el estado del Ajuste
' Parametros: 
'			  pIdAjuste  [int]
'			  pLlave  [string]
'			  pSecuencia  [int]
'			  pCdAjuste  [string]
'			  pOrigen  [int]
'			  pPto  [string]
' Devuelve:		-
'----------------------------------------------------------------------------------------------
Function firmarAjustePto(pIdAjuste, pLlave, pSecuencia, pCdAjuste,pOrigen, pPto)	
	Dim rtrn, rs, strSQL
	rtrn = false
	Call GP_CONFIGURARMOMENTOS()
	strSQL = "UPDATE TBLAJUSTESFIRMAS SET HKEY = '"& pLlave &"', MMTO = " & session("mmtoSistema") & ", cdusuario = '"&session("Usuario")&"' " &_
		     " WHERE IDAJUSTE = " & pIdAjuste & " AND SECUENCIA = "& pSecuencia
	Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	rtrn = true
	if (esUltimaFirmaAjustePto(pIdAjuste, pPto))  then
		'Si es la ultima firma, actualizo el estado del Ajuste
		strSQL = "UPDATE TBLAJUSTES SET ESTADO = " & ESTADO_AUTORIZADO & " WHERE IDAJUSTE = " & pIdAjuste
		Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
		if (pCdAjuste = AJUSTE_DRAFT_SURVEY) then ' actualizo el draft
			strSQL = "UPDATE TBLEMBARQUESDRAFTSURVEY SET CDESTADO = " & ESTADO_AUTORIZADO & " WHERE IDDRAFT = " & pOrigen
			Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
			Call sendEmailDraftAuthorized(pIdAjuste,pCdAjuste,pPto)
		end if 		
	else
	    Call sendNotifyEmail(pIdAjuste, pCdAjuste, pSecuencia+1, pPto)
	end if
	firmarAjustePto = rtrn
End function
'----------------------------------------------------------------------------------------------
' Función:	  sendNotifyEmail
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  17/09/2013
' Objetivo:	  Envia el mail notificando al siguiente firmante en la lista que ya puede autorizar el ajuste.
' Parametros: 
'			  pIdAjuste  [int]  ID del ajuste que se firmo
'             pSecuencia [int]  Secuencia del firmante a notificar
'			  pPto  [string]    Puerto para el cual corresponde el ajuste.
' Devuelve:		-
'----------------------------------------------------------------------------------------------
Function sendNotifyEmail(pIdAjuste, pCdAjuste, pSecuencia, pPto)
    Dim rtrn, rs, strSQL, strTo, msg, emailToepfer
    
    'Se determina quien es proximo firmante y se le envia un mail
    'Los mails solo se envian a los controller y a los directores ya que el primero en firmar nunca recibe un mail salvo por el enviado en los avisos automaticos.	    
    if (pSecuencia = AJS_FIRMA_DIRECTOR) then
        'Se toman los directores	        
        strSQL = "select * from TOEPFERDB.TBLREGISTROFIRMAS where rolfirma = " & FIRMA_ROL_DIRECTOR	        
   else
        'Se toman los usuarios que son controllers
        strSQL = "select * from TOEPFERDB.TBLREGISTROFIRMAS where rolfirma = " & FIRMA_ROL_CONTROLLER	        
    end if
    'Se pide que el responsable tenga acceso a los ajustes del puerto en cuestion.
    select case UCase(pPto)
        case TERMINAL_ARROYO: 
	        strSQL = strSQL & " and AJPTOARROYO = 1"
        case TERMINAL_TRANSITO: 	
            strSQL = strSQL & " and AJPTOTRANSITO = 1"
        case TERMINAL_PIEDRABUENA: 	
	        strSQL = strSQL & " and AJPTOPIEDRABUENA = 1"
    end select		    	            
    Call executeQuery(rs, "OPEN", strSQL)
    if (not rs.eof) then
        'Tengo destinatarios, mando el mail.
        emailToepfer = obtenerMail(CD_TOEPFER)
        strTo = ""
        while (not rs.eof)
            strTo = strTo & getUserMail(rs("CDUSUARIO")) & ";"
            rs.MoveNext()
        wend
        msg = "Está disponible para su firma un ajuste de stock realizado en el puerto de " & pPto & vbCrLf
        msg = msg & "Motivo del Ajuste: " & getDsCodigoAjustePuerto(pCdAjuste) & vbCrLf
        msg = msg & "Código del Ajuste: " & pIdAjuste        
        Call GP_ENVIAR_MAIL("POSEIDON - Sistema de Puertos - Ajuste de Stock", msg, emailToepfer, strTo)
    end if
End Function
'----------------------------------------------------------------------------------------------
' Función:	  cargarFirmas
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  23/08/2013
' Objetivo:	  Carga todas las firmas que tiene el Ajuste registradas hasta el momento
' Parametros: 
'			  pIdAjuste  [int]
'			  pPto  [string]
' Devuelve:		-
'----------------------------------------------------------------------------------------------
Function cargarFirmas(pIdAjuste, pPto)
	Dim rsFirmas, connFirmas, strSQL
	strSQL = "SELECT * FROM TBLAJUSTESFIRMAS WHERE IDAJUSTE=" & pIdAjuste & " ORDER BY SECUENCIA"	
	Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	while not rs.eof
		select case cint(rs("SECUENCIA"))
			case AJS_FIRMA_GERENTE_PUERTOS
				member1Cd = rs("CDUSUARIO")
				member1 = getUserDescription(member1Cd)				
				if (rs("HKEY") <> "") then member1Firma = armarTextoFirma(rs("HKEY"), rs("MMTO"))
			case AJS_FIRMA_CONTROLLER
				member2Cd = rs("CDUSUARIO")
				member2 = getUserDescription(member2Cd)
				if (rs("HKEY") <> "") then member2Firma = armarTextoFirma(rs("HKEY"), rs("MMTO"))
			case AJS_FIRMA_DIRECTOR
				member3Cd = rs("CDUSUARIO")
				member3 = getUserDescription(member3Cd)
				if (rs("HKEY") <> "") then member3Firma = armarTextoFirma(rs("HKEY"), rs("MMTO"))
		end select
		rs.movenext
	wend	
End Function
'----------------------------------------------------------------------------------------------
' Función:	  writeObservacionesDrafSurvey
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  23/08/2013
' Objetivo:	  Muestra las observaciones de un Ajuste de Draft Survey
' Parametros: 
'			  pIdOrigen  [int] (idDraft)
'			  pPto  [string]
' Devuelve:	  Observaciones 
'----------------------------------------------------------------------------------------------
Function writeObservacionesDrafSurvey(pIdOrigen, pto)
	Dim msg,rs, diff, kg3ros, tipoDiffDraft, tipoDiff3ros
	strSQL = "SELECT A.*, B.CDBUQUE, C.DSBUQUE " &_
			 "FROM (SELECT * " &_
			 "		FROM TBLEMBARQUESDRAFTSURVEY WHERE IDDRAFT = " & pIdOrigen &") A " &_
			 "	 INNER JOIN EMBARQUES B ON A.CDAVISO = B.CDAVISO " &_
			 "	 INNER JOIN BUQUES C ON C.CDBUQUE = B.CDBUQUE " 			 
	Call GF_BD_Puertos (pto, rs, "OPEN",strSQL)	
	if not rs.EoF then
	    kg3ros = 0
	    diff = CDbl(rs("TOTALDRAFT")) - CDbl(rs("TOTALBALANZA"))	    
	    tipoDiffDraft = "Faltante"
	    tipoDiff3ros = "Sobrante"
	    if (diff > 0) then
	        tipoDiffDraft = "Sobrante"
	        tipoDiff3ros = "Faltante"
	    end if
	    'Se toma el nombre del buque para la notificación de autorizacion una vez que está totalmente firmado.
	    gDSBuque = rs("DSBUQUE")
		msg = "<table>" &_
			  "		<tr><td><b>C&oacutedigo y Nombre del Buque</b></td><td width='2px'>:</td><td>" & rs("CDBUQUE") &" - "& rs("DSBUQUE") & "</td></tr>" &_			  
			  "		<tr><td><b>C&oacutedigo de Aviso de Embarque</b></td><td>:</td><td>" & rs("CDAVISO") & "</td></tr>" &_
			  "		<tr><td><b>DATOS DE BALANZA</b></td></tr>"
        if (not isNull(rs("KGBZATOEPFER"))) then
		      msg = msg & "<tr><td align='right'><b>Cargas de Toepfer</b></td><td>:</td><td align='right'>" & GF_EDIT_DECIMALS(CDbl(rs("KGBZATOEPFER")),0) &" Kg." & "</td></tr>"
		      kg3ros = CDbl(rs("TOTALBALANZA")) - CDbl(rs("KGBZATOEPFER"))
		      msg = msg & "<tr><td align='right'><b>Cargas de 3ros</b></td><td>:</td><td align='right'>" & GF_EDIT_DECIMALS(kg3ros,0) &" Kg." & "</td></tr>"		      
		end if
		msg = msg & "<tr><td align='right'><b>Carga Total Balanza</b></td><td>:</td><td align='right' style='border-top:thin solid #000000;'><b>" & GF_EDIT_DECIMALS(CDbl(rs("TOTALBALANZA")),0) &" Kg." & "</b></td></tr>" &_			  
		      "		<tr><td colspan='3'>&nbsp;</td></tr>" &_		      		      
		      "		<tr><td align='right'><b>Resultado de Draft Survey</b></td><td>:</td><td align='right'><b>" & GF_EDIT_DECIMALS(CDbl(rs("TOTALDRAFT")),0) &" Kg." & "</b></td></tr>" &_		      
		      "		<tr><td>&nbsp;</td></tr>" &_
              "		<tr><td><b>RESULTADO</b></td></tr>" &_
              "		<tr><td align='right'><b>" & tipoDiffDraft & " x D.Survey</b></td><td>:</td><td align='right'><b>" & GF_EDIT_DECIMALS(diff,0) &" Kg. </b></td><td><b>(" &  GF_EDIT_DECIMALS(diff*100*100/CDbl(rs("TOTALBALANZA")),2) & "%)</b></td></tr>"
        if (kg3ros <> 0) then
              diff3ros = -diff * (CDbl(rs("TOTALBALANZA")) - CDbl(rs("KGBZATOEPFER")))/CDbl(rs("TOTALBALANZA"))
              msg = msg & "<tr><td align='right'><b>" & tipoDiff3ros & " x Cargas de 3ros</b></td><td>:</td><td align='right'><b>" & GF_EDIT_DECIMALS(diff3ros,0) &" Kg.</b></td></tr>"
        end if
		msg = msg & "</table>"			  
	end if
	writeObservacionesDrafSurvey = msg
End Function
'----------------------------------------------------------------------------------------------
' Función:	  leerRegistroFirmas
' Autor: 	  CNA - Ajaya Nahuel
' Fecha: 	  23/08/2013
' Objetivo:	  Comprueba si el firmante y su llave son correctos
' Parametros: -
' Devuelve:	  -
'----------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
	ret = false
	if (HK_isKeyReady()) then
		strSQL = "Select * from TOEPFERDB.TBLREGISTROFIRMAS where HKEY='" & HK_readKey() & "'"		
		Call executeQuery(rs, "OPEN", strSQL)
		if (not rs.eof) then
			gCdUsuario = rs("CDUSUARIO")
			if (session("Usuario") = gCdUsuario) then ret = true			
		else
			gCdUsuario = ""			
		end if
	end if		
	leerRegistroFirmas = ret
End Function
'*************************************************************************************************
'										INICIO DE LA PAGINA
'*************************************************************************************************
Dim strSQL, pto,gCdUsuario,cdAjuste,secuencia,accion, member1Firma, member2Firma, member3Firma, member1Cd, member2Cd, member3Cd
Dim member3, member2, member1, flagUltimoAjuste, origen, filename, extFile, rolFirma
Dim gDSBuque

pto   = GF_PARAMETROS7("pto","",6)
idAjuste = GF_PARAMETROS7("idAjuste",0,6)
errFirma = GF_PARAMETROS7("errFirma","",6)
accion = GF_PARAMETROS7("accion","",6)
secuencia = GF_PARAMETROS7("secuencia",0,6)
cdAjuste = GF_PARAMETROS7("cdAjuste","",6)
origen = GF_PARAMETROS7("origen",0,6)
g_strPuerto = pto

if (errFirma <> "") then Call setError(errFirma)
rolFirma = getRolFirma(session("Usuario"), SEC_SYS_POSEIDON)

if (accion = ACCION_GRABAR) then	
	ret = LLAVE_NO_CORRESPONDE
	if (leerRegistroFirmas()) then
		if (secuencia = AJS_FIRMA_GERENTE_PUERTOS) then
			if(firmarAjustePto(idAjuste, HK_readKey(), secuencia, cdAjuste, origen, pto)) then ret = RESPUESTA_OK
		end if
		if (secuencia = AJS_FIRMA_CONTROLLER) then
			if(firmarAjustePto(idAjuste, HK_readKey(), secuencia, cdAjuste, origen, pto)) then ret = RESPUESTA_OK
		end if
		if (secuencia = AJS_FIRMA_DIRECTOR) then
			if(firmarAjustePto(idAjuste, HK_readKey(), secuencia, cdAjuste, origen, pto)) then ret = RESPUESTA_OK
		end if
	end if
	Call HK_sendResponse(ret)
end if
		
Call cargarFirmas(idAjuste, pto)
strSQL = " select * from tblajustes where idajuste = "& idAjuste
call GF_BD_Puertos (pto, rs, "OPEN",strSQL)

	
%>
<html>
<head>
<title>Autorizaci&oacuten de  Ajustes</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/hkey.js"></script>
<script type="text/javascript">
	
	// Se determina el explorador.	
	isFirefox=true; //FF
	if (navigator.userAgent.indexOf("MSIE")>=0) isFirefox=false; //IE
	var hkey0 = new Hkey('hk0', "poseidonFirmaAjuste.asp?pto=<% =pto %>&accion=<%=ACCION_GRABAR%>&idAjuste=<%=idAjuste%>&secuencia=<%=AJS_FIRMA_GERENTE_PUERTOS%>&cdAjuste=<%=rs("CDAJUSTE")%>&origen=<%=rs("IDORIGEN")%>", '<% =HKEY() %>', 'check_callback()');
	var hkey1 = new Hkey('hk1', "poseidonFirmaAjuste.asp?pto=<% =pto %>&accion=<%=ACCION_GRABAR%>&idAjuste=<%=idAjuste%>&secuencia=<%=AJS_FIRMA_CONTROLLER%>&cdAjuste=<%=rs("CDAJUSTE")%>&origen=<%=rs("IDORIGEN")%>", '<% =HKEY() %>', 'check_callback()');
	var hkey2 = new Hkey('hk2', "poseidonFirmaAjuste.asp?pto=<% =pto %>&accion=<%=ACCION_GRABAR%>&idAjuste=<%=idAjuste%>&secuencia=<%=AJS_FIRMA_DIRECTOR%>&cdAjuste=<%=rs("CDAJUSTE")%>&origen=<%=rs("IDORIGEN")%>", '<% =HKEY() %>', 'check_callback()');
	
	
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
<form name="frmSel" id="frmSel" method="POST" action="poseidonFirmaAjuste.asp?idAjuste=<% =idAjuste %>&pto=<% =pto %>">
	<table width="70%" align="center" class="reg_header">		
		<tr>	
			<td>
				<table width="100%" border="0" cellpadding="1" cellspacing="1" class="reg_header">
					<tr>						
						<td align="left" class="reg_header_nav" width="20%"><%=GF_TRADUCIR("Concepto:")%></td>
						<td align="left" class="reg_header_navdos" width="60%"><%=getDsCodigoAjustePuerto(rs("CDAJUSTE")) &" ("&rs("CDAJUSTE")&")"%></td>
						<td align="left" class="reg_header_nav" width="10%"><%=GF_TRADUCIR("Nro. Ajuste")%></td>						
						<td align="left" class="reg_header_navdos" width="10%"><%=rs("IDAJUSTE") %></td>
		            </tr> 
		            <tr>						
						<td align="left" class="reg_header_nav"><%=GF_TRADUCIR("Puerto:")%></td>
						<td align="left" colspan="3" class="reg_header_navdos"><%= UCase(g_strPuerto)%></td>
		            </tr>
		            <tr>
						<td align="left" class="reg_header_nav"><%=GF_TRADUCIR("Fecha/Período:")%></td>
						<td align="left" colspan="3" class="reg_header_navdos">
						<% if (Cdbl(rs("FECHADESDE")) = Cdbl(rs("FECHAHASTA"))) then
								Response.Write GF_FN2DTE(rs("FECHADESDE"))
						   else
								Response.Write "Desde " & GF_FN2DTE(rs("FECHADESDE")) & "  hasta " & GF_FN2DTE(rs("FECHAHASTA"))
						   end if %>
						</td>
					</tr>	
					<tr>						
						<td align="left" class="reg_header_nav"><%=GF_TRADUCIR("Producto:")%></td>
						<td align="left" colspan="3" class="reg_header_navdos"><%=rs("CDPRODUCTO")&" - "& getDsProducto(rs("CDPRODUCTO"))%></td>						
					</tr>					
		            <tr>    
						<td align="left" class="reg_header_nav"><%=GF_TRADUCIR("Kilos:")%></td>		                
						<td align="left" colspan="3" class="reg_header_navdos <%if CDbl(rs("KILOSAJUSTE")) < 0 then Response.write "reg_header_rejected"  %>"><%=GF_EDIT_DECIMALS(CDbl(rs("KILOSAJUSTE"))*100,2) & " Kg."%></td>
					</tr>
					<tr>
						<td align="left" class="reg_header_nav"><%=GF_TRADUCIR("Observaciones:")%></td>				
						<td colspan="3" align="left" class="reg_header_navdos">
      					<%	filename = ""
      					    select case rs("CDAJUSTE")
      							case AJUSTE_DRAFT_SURVEY
	      							Response.Write writeObservacionesDrafSurvey(rs("IDORIGEN"), pto)
	      							'Tomo el adjunto
	      							strSQL = "Select * from TBLEMBARQUESDRAFTSURVEY where idDraft = "& rs("IDORIGEN")	      							
                                    Call GF_BD_Puertos(g_strPuerto, rsSurvey, "OPEN", strSQL)
                                    if (not rsSurvey.eof) then
                                        if (not isnull(rsSurvey("NAMEFILE"))) then 
                                            filename = rsSurvey("NAMEFILE")
                                            extFile = rsSurvey("EXTFILE")
                                        end if
                                    end if                                    
      							case AJUSTE_MANIPULEO
	      							Response.Write "Ajuste estimativo de Merma por Manipuleo"
      						end select	%>
      		         </td>					 	
					</tr>
					<tr>
					    <td class="reg_header_nav">Documentaci&oacuten del Proceso</td>
					    <td>					        
					        <% if (filename <> "") then %>
					        <a href='Documentos/Draft Survey/<%=g_strPuerto%>/<% =filename %>.<% =extFile %>' target='_blank'>
								<img title='Descargar adjunto' src='images/compras/download.png'> <% =filename %>
							</a>
							<% else   
							    Response.Write GF_TRADUCIR("No se ha adjuntado ninguna documentación respaldatoria.")
						       end if %>	
					    </td>
					</tr>				  
				</table>
				<input type="hidden" name="cdAjuste" id="cdAjuste" value="<%=rs("CDAJUSTE")%>">
				<input type="hidden" name="idAjuste" id="idAjuste" value="<%=idAjuste%>">
				<input type="hidden" name="secuencia" id="secuencia" value="<%=pSecuencia%>">				
				<input type="hidden" name="origen" id="origen" value="<%=rs("IDORIGEN")%>">								
			</td>
       </tr>       
      <tr>
          <td>
			  <table width="100%" border="0" cellpadding="1" cellspacing="1" class="reg_header">
				 <tr>
					<td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Gerente Puerto")%></td>
					<td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Controller")%></td>
					<td width="33%" align="center" class="reg_header_nav"><%=GF_TRADUCIR("Director")%></td>
				</tr>
				<tr>
					<td height="100" align="center" class="recuadro round_border_bottom_left">
						<% if (member1Firma  <> "" ) then     %>
						      <img src="images/firmas/<% =obtenerFirma(member1Cd) %>"><br>
						      <% =member1Firma %>
						 <%	else						 
								if (rolFirma = FIRMA_ROL_RESP_PUERTO) then	
									if (member1Cd <> session("Usuario")) then	%>
										<br><div id="hk0"></div><br>
								<%	else	
										response.write GF_TRADUCIR("Usted ya ha firmado como Gerente de Puerto.")
									end if	
								else %>
									<br><br><br>	
							<%	end if		
						    end if	%>
						  ________________________________________<br />
						<%if (member1Firma <> "") then%>
							  	<%=member1%>
						<%else%>
								<br />
						<%end if%>
				    </td>
					<td align="center" class="recuadro" >
						<%	if (member2Firma  <> "") then %>
						    <img src="images/firmas/<% =obtenerFirma(member2Cd) %>"><br>
						    <% =member2Firma %>
						<%	else
						        if (rolFirma = FIRMA_ROL_CONTROLLER) then	
						           if (member2Cd <> session("Usuario")) then%>
  						               <br><div id="hk1"></div><br>
						         <%else
										response.write GF_TRADUCIR("Usted ya ha firmado como Controller.")
                				   end if%>
						    <%	else	%>
						           <br><br><br>
						     <%	end if	
						    end if	%>
						________________________________________<br />
                		<%if (member2Firma <> "") then%>
							<%=member2%>
						<%end if%>
		            </td>
				    <td align="center" class="recuadro" >
						<%	if (member3Firma  <> "") then %>
						    <img src="images/firmas/<% =obtenerFirma(member3Cd) %>"><br>
						    <% =member3Firma %>
						<%	else								
						        if (rolFirma = FIRMA_ROL_DIRECTOR) then	                            
						           if (member3Cd <> session("Usuario")) then%>                                   
  						                <br><div id="hk2"></div><br>
						         <%else
										response.write GF_TRADUCIR("Usted ya ha firmado como Director.")
                					end if%>
						     <%	else	%>
						            <br><br><br>
						     <%	end if	
						    end if	%>
						________________________________________<br />
                		<%if (member3Firma <> "") then%>
							<%=member3%>
						<%end if%>
		            </td>
				</tr>          
			 </table>			
		 </td>
    </tr>
</table>
<input type="hidden" name="errFirma" id="errFirma">
</form>
</body>
</html>
