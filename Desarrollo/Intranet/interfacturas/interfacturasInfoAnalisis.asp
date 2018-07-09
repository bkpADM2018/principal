<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<%
	'NOTA: Solo brindará informacion de Analisis a aquellas facturas que :
	'		- Se migraron de los puertos (se guarda las descargas historicas).
	'		- Tipo gastos acodicionamiento (GP)
    Dim reg, rs
    reg = GF_PARAMETROS7("id","",6)
    'Para identificar si es camion o vagon busco el tipo de transporte en la tabla MER311F6
	strSQL =  "SELECT a.CPORR6,a.IDCAR6,a.FECDR6,b.CTRAR6,a.DVCDR6,a.CPROR6  "&_
			  "FROM MERFL.MER711L1 A "&_
			  "INNER JOIN (SELECT CPORR6,CTRAR6 "&_
			  "			   FROM MERFL.MER311F6 "&_
			  "			   WHERE CTRAR6 IN ("&TIPO_TRANSPORTE_CAMION&","&TIPO_TRANSPORTE_VAGON&") GROUP BY CPORR6,CTRAR6)B "&_
			  "	   ON B.CPORR6 = A.CPORR6 "&_
			  "WHERE FCRGR6 = "& reg &_
			  " group by a.CPORR6,a.IDCAR6,a.FECDR6,b.CTRAR6,a.DVCDR6,a.CPROR6 "&_
			  "	order by a.FECDR6 "			  
    Call GF_BD_COMPRAS(rs,conn,"OPEN",strSQL)

%>
<html>
<body>
<div>
	<hr /><h3>INFORMACION ANALISIS</h3><hr />
</div>

	<table class="datagrid" align="center" width="80%">
	
	    <thead>
	        <tr>
	           <th align="center" width="25%">FECHA</th>
	           <th align="center" width="35%">CARTA PORTE</th>
	           <th align="center" width="35%">ID</th>
	           <th align="center" width="5%">.</th>
	        </tr>
	    </thead>
		<tbody>
		<%	if (not rs.eof) then %>
		<%		while (not rs.eof)	%>
		            <tr>
		                <td align="center"><%=GF_FN2DTE(rs("FECDR6"))%></td>
		                <td align="center"><%=GF_EDIT_CTAPTE(GF_nDigits(rs("CPORR6"),12))%></td>
		                <td align="center"><%=rs("IDCAR6")%></td>
		                <td align="center">
		                <% 'Solo si es definido como camion o vagon se podrá ver los analisis, debido a que se tiene que saber que pagina abrir (camiones o vagones)							
							if (Cdbl(rs("CTRAR6")) = TIPO_TRANSPORTE_CAMION)or(Cdbl(rs("CTRAR6")) = TIPO_TRANSPORTE_VAGON) then 
								'Se validan los formatos de los paramtros quese vana pasar a los PopUp de camiones y vagones
								auxCartaPorte = GF_EDIT_CTAPTE(GF_nDigits(rs("CPORR6"),12))
								auxFecha	  = rs("FECDR6")
								auxId		  = rs("IDCAR6")
								if (Cdbl(rs("CTRAR6")) = TIPO_TRANSPORTE_VAGON) then									
									auxCartaPorte = GF_nChars(GF_nDigits(rs("CPORR6"),12),16,"0",CHR_AFT)
									auxFecha	  = Right(rs("FECDR6"),2)&"/"&Mid(rs("FECDR6"),5,2)&"/"&Left(rs("FECDR6"),4)
									auxId = GF_nChars(Right(rs("CPORR6"),Len(rs("CPORR6"))-1),12,"0",CHR_AFT)
								end if %>
								<img src="images/analisis-2-16.png" title="Ver analisis" style="cursor:pointer;" onclick="abrirAnalisis('<%=auxFecha%>','<%=getDsPuertoByLetra(rs("DVCDR6"))%>','<%=auxCartaPorte%>','<%=auxId%>',<%=rs("CTRAR6")%>,'<%=rs("CPROR6")%>')">
						<%  end if %>
						</td>
		            </tr>
		<%			rs.MoveNext()
			    wend %>
		<%  else  %>
			    <tr><td colspan="4" align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td></tr>
		<%  end if  %>
		</tbody>
	</table>
</body>
</html>
