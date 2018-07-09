<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientossql.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosfechas.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->
<%

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Función que devuelve el resultado de cada uno de los rubros cargados en la calada para cada una de sus secuencias.
Function getResultadosCalada(p_dtContable, p_idVagon, p_sqCalada)
    Dim strSql, rs, myDtContable
    
    myDtContable = GF_FN2DTCONTABLE(p_dtContable)

    strSQL= "Select * from " & _
            "((Select "& Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2) &"  DTCONTABLE, CDVAGON, SQCALADA, A.CDRUBRO, DSRUBRO, VLBONREBAJA, CDSUPERVISOR,ICINGMANUAL " & _
            "from RUBROSVISTEOVAGONES A " & _
	        "INNER JOIN RUBROS B on A.CDRUBRO=B.CDRUBRO " & _	        
            "where SQCALADA=" & p_sqCalada & " and CDVAGON='" & p_idVagon & "'" & _
            ") UNION (" & _
            "Select ((YEAR(A.DTCONTABLE)*10000) + (MONTH(A.DTCONTABLE)*100) + DAY(A.DTCONTABLE)) DTCONTABLE, CDVAGON, SQCALADA, A.CDRUBRO, DSRUBRO, VLBONREBAJA, CDSUPERVISOR,ICINGMANUAL " & _
            "from HRUBROSVISTEOVAGONES A " & _
	        "INNER JOIN RUBROS B on A.CDRUBRO=B.CDRUBRO " & _	        
            "where DTCONTABLE='" & myDtContable & "' and SQCALADA=" & p_sqCalada & " and CDVAGON='" & p_idVagon & "')) TABLA " & _            
            "Order by SQCALADA DESC, CDRUBRO"
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSql)
    'Response.Write strSql
    Set getResultadosCalada = rs    
End Function

'--------------------------------------------------------------------------------------------------------
Dim Conn,g_cartaPorte,g_cdOperativo
Dim g_strPuerto, g_strSector, countResultados
Dim g_rsPesadas, g_rsCaladas, g_rsHumedimetro, g_dtContable,g_dsObservaciones
dim g_fltPromedioHumedad, g_fltPromedioPesoHect, g_fltPromedioTemp, g_fltMaxHumedad, g_fltMinPesoHect, g_fltMaxTemp, g_rsResultados

g_strPuerto = GF_Parametros7("Pto","",6)
g_dtContable = GF_Parametros7("dtContable","",6)
g_idVagon = GF_Parametros7("nroVagon","",6)
g_cdOperativo = GF_Parametros7("cdOperativo","",6)
g_cartaPorte = GF_Parametros7("cartaPorte","",6)

Set g_rsCaladas = getCaladasVagones(g_dtContable, g_idVagon,g_cdOperativo,g_cartaPorte)  
if not g_rsCaladas.eof Then 
while not g_rsCaladas.eof %>

	<link rel="stylesheet" type="text/css" href="../../css/main.css" />

                	
	<table class="datagridlv2" align="center" width="80%" id="tblCalada_<%=g_idVagon%>_<%=g_rsCaladas("sqcalada")%>">
    	<thead>
        	<tr>
        	    <th>&nbsp;</th>
            	<th colspan="2"><%=GF_Traducir("Secuencia Calada: ") & g_rsCaladas("sqcalada")%></th>
            	<th>&nbsp;</th>
            </tr>
        </thead>
        
        <tbody>
            <tr>
                <td><b><%=GF_Traducir("Camara")%>:</b></td>
                <td><%=g_rsCaladas("iccamara")%></td>
                <td><b><%=GF_Traducir("Usuario")%>:</b></td>
                <td><%=g_rsCaladas("cdusername")%>&nbsp;<%=g_rsCaladas("dslastname")%></td>
            </tr>
            <tr>
                <td><b><%=GF_Traducir("Humedad")%>.:</b></td>
                <td><%=g_rsCaladas("vlhumedad")%></td>
                <td><b><%=GF_Traducir("Proteina")%>.:</b></td>
                <td><%=g_rsCaladas("vlproteina")%></td>
            </tr>
            <tr>			
                <td><b><%=GF_Traducir("Aceptacion")%>.:</b></td>
                <td><%=g_rsCaladas("dsaceptacion")%></td>
                <td><b><%=GF_Traducir("Merma")%>.:</b></td>
                <td><%=g_rsCaladas("pcmerma")%></td>
            </tr>
            <tr>
                <td><b><%=GF_Traducir("Momento calada")%>.:</b></td>
                <td><%=Right(GF_FN2DTE(g_rsCaladas("dtcalada")),8) %></td>
                <td><b><%=GF_Traducir("Tipo Calada")%>.:</b></td>
                <td><%=g_rsCaladas("ictipocalada")%></td>
            </tr>            
            <% if (g_rsCaladas("CDMOTIVORECHAZO") > 0) then %>
            <tr>			
                <td><b><%=GF_TRADUCIR("Motivo Rechazo")%>.:</b></td>
                <td colspan="3"><%= getDSMotivoRechazo(g_rsCaladas("CDMOTIVORECHAZO"))%></td>			
            </tr>
            <% end if %>            
            <tr>
                <td><b><%=GF_TRADUCIR("Grado")%>.:</b></td>                
                <% myGradoParticular =  VerGrado (g_strPuerto,g_rsCaladas("CDACEPTACION"), g_rsCaladas("NUBARRAS"), g_rsCaladas("DTCONTABLE")) %>
                <td><%=myGradoParticular%></td>
                <td><b><%=GF_TRADUCIR("Sticker")%>.:</b></td>
                <td><%=  g_rsCaladas("NUBARRAS")  %></td>                
            </tr>
            <tr>			
                <td><b><%=GF_TRADUCIR("Observaciones")%>.:</b></td>
                <% g_dsObservaciones = "-"
                   if (Len(Trim(g_rsCaladas("DSOBSERVACIONES"))) > 0) then g_dsObservaciones = Trim(g_rsCaladas("DSOBSERVACIONES")) %>
                <td colspan="3"><%= g_dsObservaciones%></td>
            </tr>
        </tbody>
    </table>
	
    <div class="col66"></div><br/>
    	   
       <table align="center" width="80%" class="datagridlv2">
           <thead>
               <tr>
                  <th align="center" width="60%">Rubro</th>
                  <th align="center" width="20%">Valor</th>
                  <th align="center" width="20%">Autom./Manual</th>
               </tr>
           </thead>
       <%  Set g_rsResultados = getResultadosCalada(g_dtContable, g_idVagon, g_rsCaladas("sqcalada"))
           if (not g_rsResultados.eof) then
                while (not g_rsResultados.eof)	%>
           <tbody>
                <tr>
                    <td><% =GF_TRADUCIR(g_rsResultados("DSRUBRO")) %></td>
                    <td align="center"><% =g_rsResultados("VLBONREBAJA") %></td>
                    <td align="center">
                    <%  if (g_rsResultados("ICINGMANUAL") = "N") then
                            response.Write "Automatico"
                        else
                            Response.Write "Manual"
                        end if %>
                    </td>
                </tr>
           </tbody>
                <%	g_rsResultados.MoveNext()
                wend			   	
           else	%>
           <tfoot>
                <tr>
                    <td colspan="4" align="center"><%=GF_Traducir("Rubros para este Vagon no disponible")%></td>
                </tr>
           </tfoot>
    <%	   end if	%>
       </table>

	
    <div class="col66"></div>
    <br /><br /><br />
	
    
<%	g_rsCaladas.MoveNext()
wend
else
	Response.Write "Calada para este Vagon no disponible"
end if
%>		
