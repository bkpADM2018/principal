<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientossql.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientosSeguridad.asp"-->
<%
Call initTaskAccessInfo(TASK_POS_INFO_ANALISIS, session("DIVISION_PUERTO"))

'----------------------------------------------------------------------------------------------------------
Dim Conn, g_rsRD
Dim g_strPuerto, g_idCamion, g_strSector, countResultados
Dim g_rsInfoCamion, g_rsPesadas, g_rsCaladas, g_rsHumedimetro, g_ctaPorte, g_dtContable,g_dsObservaciones,g_NroAnalisis
dim g_fltPromedioHumedad, g_fltPromedioPesoHect, g_fltPromedioTemp, g_fltMaxHumedad, g_fltMinPesoHect, g_fltMaxTemp, g_rsResultados



g_strPuerto = GF_Parametros7("Pto","",6)
g_CdDestino = GF_Parametros7("destino",0,6)
if (g_strPuerto = "") then 
    g_strPuerto = getDsPuertoByNro(g_CdDestino)
else
    g_CdDestino  = getNumeroPuerto(g_strPuerto)
end if
g_dtContable = GF_Parametros7("dtContable","",6)
'if (g_dtContable <> "") then g_dtContable = Left(g_dtContable,4) & "-" &mid(g_dtContable,5,2) &"-"& Right(g_dtContable,2)
g_ctaPorte = GF_Parametros7("ctaPorte","",6)
g_NroAnalisis = GF_Parametros7("nroAnalisis",0,6)
g_idCamion = GF_Parametros7("camion","",6)
Call loadAnalisisCamionPto(g_dtContable,g_ctaPorte,g_strPuerto,g_idCamion,g_NroAnalisis)
%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Informacion de Cálidad de Camion </TITLE>
	<link href="../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="../css/main.css" />	
	
	<script type="text/javascript" src="../scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript" src="../scripts/channel.js"></script>
	<SCRIPT LANGUAGE="JavaScript">
		var ch= new channel();
		
		function verCaladaCamion(pIdCamion,pDtContable,pSqCalada){
			var pElement = document.getElementById("divCaladaCamion_" + pSqCalada); 
			if (document.getElementById("trCaladaCamion_" + pSqCalada).className == "troculto") {
				document.getElementById("trCaladaCamion_" + pSqCalada).className = "trvisible";
				var iImgAddArt = document.createElement('img');
				iImgAddArt.id = "loading_"  + pSqCalada;
				iImgAddArt.name = "loading_"  + pSqCalada;
				iImgAddArt.src = "../images/Loading4.gif";
				iImgAddArt.title = "Agregar Articulo";
				iImgAddArt.setAttribute('style', "cursor:pointer;");
				pElement.align = "center";
				pElement.appendChild(iImgAddArt);				
				ch.bind("infoAnalisisCamionAjax.asp?Pto=<%=g_strPuerto%>&idCamion=" +  pIdCamion + "&dtContable=" + pDtContable +"&sqCalada="+pSqCalada+"&cartaporte=<%=g_ctaPorte%>" ,"CallBack_verCalada("+pSqCalada+")");
				ch.send();
				document.getElementById("imgVerCalada_"+pSqCalada).src = "../images/Menos.gif"
			}
			else{				
				document.getElementById("trCaladaCamion_" + pSqCalada).className = "troculto";
				removeAllChilds(pElement);
				document.getElementById("imgVerCalada_"+pSqCalada).src = "../images/Mas.gif"
			}
		}
		function CallBack_verCalada(pSqCalada){
			var padre = document.getElementById("loading_" + pSqCalada).parentNode;
			padre.removeChild(document.getElementById("loading_" + pSqCalada));
			var respuesta = ch.response();
			document.getElementById("divCaladaCamion_" + pSqCalada).style.display = "";
			document.getElementById("divCaladaCamion_" + pSqCalada).innerHTML = respuesta;
		}
		function removeAllChilds(a){			
			while(a.hasChildNodes()){
				a.removeChild(a.firstChild);
			}	
		}
		function lightOn(tr) {
			tr.className = "reg_Header_navdosHL";
		}
		function lightOff(tr) {
			tr.className = "reg_Header_navdos";
		}
		
	</script>
</HEAD>	
<BODY >
	<div class="col66"></div>
	<INPUT type="hidden" id="Pto" name="Pto" value = <%= g_strPuerto %>>
	<INPUT type="hidden" id="Camion" name="Camion" value = <%= g_idCamion%>>
    <div class="tableasidecontent">
        <%  if (g_idCamion <> "") then%>
        <div class="col26 reg_header_navdos"> Fecha Descarga </div>
        <div class="col26"> <% =GF_FN2DTE(g_dtContable) %>  </div>
        
        <div class="col26 reg_header_navdos"> ID Camion </div>
        <div class="col26">  <% =g_idCamion %> </div>
        <% end if %>
        <div class="col26 reg_header_navdos"> Carta Porte </div>
        <div class="col26"> <% =GF_EDIT_CTAPTE(g_ctaPorte) %>  </div>
        
    </div>
	<div class="col66"></div>
     <% if (g_strPuerto <> "") then %>     
		<div class="tableaside size100">
			<h3>Resultados Camara</h3>
			<table class="datagrid datagridlv1" width="100%">
                <%	Set g_rsRC = getResultadosCamaraByCamion(g_dtContable, g_ctaPorte, g_strPuerto) %>	 
                <%  if not g_rsRC.eof then  %>
				<thead>
			    	<tr> 
			    		<th width="15%"><%=GF_Traducir("Fecha Certificado")%></th>			    		
			    		<th width="15%"><%=GF_Traducir("Certificado")%></th>
			        	<th width="10%"><%=GF_Traducir("Nro. Solicitud")%></th>
			        	<th width="15%"><%=GF_Traducir("Nro. Muestra")%></th>			        	
			        </tr>
			    </thead>
				<tbody>
				   		<tr>
				   			<td align="center"><% =g_rsRC("FECHACERTIFICADO") %></td>
				   			<td align="center"><% =g_rsRC("NUCERTIFICADO") %></td>
				   			<td align="center"><% =g_rsRC("NUINFOANALISIS") %></td>
				   			<td align="center"><% =g_rsRC("NUBARRAS") %></td>
				   		</tr>
				   		<tr>
				   		    <td colspan="4">
				   		    <%	Set g_rsRD = getResultadosDetalle(g_rsRC("IDANALISIS"), g_strPuerto) %>	                            
				   		        <table class="datagrid datagridlv1" width="80%" align="center">
				   		            <thead>
			    	                    <tr> 
			    		                    <th colspan="2"><%=GF_Traducir("Ensayo")%></th>
			        	                    <th width="25%"><%=GF_Traducir("Resultado")%></th>
			        	                    <th width="25%"><%=GF_Traducir("Fuera Std?")%></th>			        	
			                            </tr>			                            
			                        </thead> 
                             <%  while (not g_rsRD.eof)  %>			                            
			                        <tbody>
				   		                <tr>
				   			                <td align="center" width="10%"><% =g_rsRD("CDENSAYO") %></td>
				   			                <td width="40%"><% =g_rsRD("DSENSAYO") %></td>
				   			                <td align="center"><% =g_rsRD("RESULTADO") %></td>				   			                
				   			                <td align="center"><% =g_rsRD("FUERAESTANDAR") %></td>
				   		                </tr>   
				   		            </tbody>				   		        
				   		    <%      g_rsRD.MoveNext()
				   		        wend %>
				   		        </table>
				   		    </td>
				   		</tr>				   		
			   </tbody>
               <% else %>
               <tbody><tr><td colspan="5" align="center"><%=GF_TRADUCIR("No se encontraron resultados") %></td></tr></tbody>
               <% end if %>
			</table>
		</div>	 
	    <div class="tableaside size100">
		<h3>Datos Calada</h3>
			<table class="datagrid datagridlv1" width="100%">
                <%Set g_rsCaladas = getCaladasCamion (g_dtContable, g_idCamion,g_strPuerto) %>
		        <%if not g_rsCaladas.eof then%>		
			        <thead>
			            <tr>
			                <th width="5%" align="center"> . </th>
			                <th width="10%" align="center"> <%=GF_Traducir("Secuencia Calada")%> </th>
			                <th width="50%" align="center"> <%=GF_Traducir("Recibidor")%> </th>
			                <th width="35%" align="center"> <%=GF_Traducir("Momento")%> </th>
			            </tr>
			        </thead>
			        <tbody>
				    <%while not g_rsCaladas.eof %>
					    <tr>
						    <td align="center"><img style="cursor:pointer;" title="Ver información Calada" id="imgVerCalada_<%=g_rsCaladas("sqcalada")%>" onClick="verCaladaCamion('<%=g_idCamion%>','<%=g_dtContable%>',<%=g_rsCaladas("sqcalada")%>)" src="../images/Mas.gif"></td>
						    <td align="center"><%=g_rsCaladas("sqcalada")%></td>
						    <td align="left"><%=g_rsCaladas("dsname")%>&nbsp;<%=g_rsCaladas("dslastname")%></td>
						    <td align="left"><%=GF_FN2DTE(g_rsCaladas("dtcalada"))%></td>
					    </tr>
					    <tr>
			        	    <td id="trCaladaCamion_<%=g_rsCaladas("sqcalada")%>" name="trCaladaCamion_<%=g_rsCaladas("sqcalada")%>" colspan="4" class="troculto">
			            	    <div id="divCaladaCamion_<%=g_rsCaladas("sqcalada")%>"></div>
			                </td>
					    <tr>
				    <%	g_rsCaladas.MoveNext()
				     wend %>	
				    </tbody>
                <% else %>	
                    <tbody><tr><td colspan="4" align="center"></td></tr></tbody>
                <%end if%>
			</table>	  
		<br>
	</div>	
     <% end if %>
</BODY>
</HTML>
<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getCaladasCamion(p_dtContable, p_idCamion,pPto)
	dim strSql, diaHoy, rs, auxDTcontable
	diaHoy = Year(Now()) &"-"& GF_nDigits(Month(Now()), 2) &"-"& GF_nDigits(Day(Now()), 2)
    auxDTcontable = Left(p_dtContable,4) &"-"& Mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)

	strSql = "Select sqcalada,dsname, dslastname,AC.dsaceptacion, " &_
             "       CAST(FECHACALADA as BIGINT)*1000000 + right('000000' + cast(HORACALADA AS varchar(6)), 6) AS dtcalada "&_
			 " from ((Select '" & diaHoy & "' DTCONTABLE, IDCAMION, SQCALADA,CDUSERNAME,CDACEPTACION,  "&_
             "              ((Year(DTCALADA) * 10000) + (Month(DTCALADA) * 100) + Day(DTCALADA)) AS FECHACALADA,  "&_
             "              ((DATEPART(HOUR, DTCALADA) * 10000) + (DATEPART(MINUTE, DTCALADA) * 100) + DATEPART(SECOND, DTCALADA)) AS HORACALADA "&_
             "        from caladadecamiones where idCamion = '" & p_idCamion & "')" &_
			 "		union" &_
			 "      (Select DTCONTABLE,IDCAMION, SQCALADA,CDUSERNAME,CDACEPTACION, "&_
             "              ((Year(DTCALADA) * 10000) + (Month(DTCALADA) * 100) + Day(DTCALADA)) AS FECHACALADA,  "&_
             "              ((DATEPART(HOUR, DTCALADA) * 10000) + (DATEPART(MINUTE, DTCALADA) * 100) + DATEPART(SECOND, DTCALADA)) AS HORACALADA "&_
             "       from hcaladadecamiones where DTCONTABLE ='" & auxDTcontable & "' and idCamion = '" & p_idCamion & "')) CC" &_
			 " inner join accounts ACC on UPPER(CC.CDUSERNAME) = UPPER(ACC.CDUSERNAME)"&_
			 " inner join aceptacioncalidad AC on CC.CDACEPTACION=AC.CDACEPTACION"&_
			 " where DTCONTABLE='" & auxDTcontable & "' ORDER BY SQCALADA DESC"
	
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSql)
	Set getCaladasCamion = rs
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'p_dtContable : formato aaaammdd
Function getResultadosCamaraByCamion(p_dtContable, p_NuCartaPorte, p_Pto)
	dim strSql, diaHoy, rs, auxDTcontable	
    'Se deja el formato aaaa-mm-dd por que luego de la subconsulta hace un inner join con otra resultadosCamara que tiene 
    'ese formato. Luego para mostrar la fecha si lo hace aaaammdd
	diaHoy = Year(Now()) &"-"& GF_nDigits(Month(Now()), 2) &"-"& GF_nDigits(Day(Now()), 2)
    auxDTcontable = Left(p_dtContable,4) &"-"& Mid(p_dtContable,5,2) &"-"& Right(p_dtContable,2)

    Call executeSP_Puertos(rs, p_Pto, "RESULTADOSCAMARACABECERA_GET_BY_NUCARTAPORTE", auxDTcontable & "||" & p_NuCartaPorte)    
	Set getResultadosCamaraByCamion = rs
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getResultadosDetalle(p_idAnalisis, p_Pto)	    
    Call executeSP_Puertos(rs, p_Pto, "RESULTADOSCAMARADETALLE_GET_BY_IDANALISIS", p_idAnalisis)    
	Set getResultadosDetalle = rs
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Carga los datos del puerto(Analisis y Camion) en caso que no vengan por parametro
Function loadAnalisisCamionPto(ByRef p_dtContable, p_CtaPte, p_Pto, ByRef p_IdCamion, ByRef p_NroAnalisis)
	dim strSql, diaHoy, rs
    if (p_Pto <> "") then 
	    diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	    strSql = "Select COALESCE(IDCAMION,'') IDCAMION, COALESCE(NUINFOANALISIS,'') NUINFOANALISIS, DTCONTABLE "&_
			     " from (Select " & diaHoy & " DTCONTABLE, IDCAMION,NUINFOANALISIS from CAMIONESDESCARGA where NUCARTAPORTE = '"& p_CtaPte &"'"&_
			     "      union " &_
			     "      Select (YEAR(DTCONTABLE )*10000 + Month(DTCONTABLE )*100 + DAY(DTCONTABLE )) DTCONTABLE, IDCAMION,NUINFOANALISIS from HCAMIONESDESCARGA WHERE NUCARTAPORTE = '" & p_CtaPte & "' ) CC "	    
        
        Call GF_BD_Puertos(p_Pto, rs, "OPEN", strSql)
        if (not rs.Eof) then 
            If (p_IdCamion = "") then p_IdCamion = rs("IDCAMION")
            if (p_NroAnalisis = "") then p_NroAnalisis = rs("NUINFOANALISIS")
            if (p_dtContable = "") then p_dtContable = rs("DTCONTABLE")
        end if
    end if
End Function
%>