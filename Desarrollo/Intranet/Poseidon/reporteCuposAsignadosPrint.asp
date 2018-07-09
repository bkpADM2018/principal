<!--#include file="../Includes/procedimientossql.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="../Includes/procedimientosExcel.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->

<%

Const CUPO_CANTIDAD_DIAS_SEMANA = 7
Const CUPO_PRIMERA_SEMANA = 1
Const CUPO_SEGUNDA_SEMANA = 2
Const CUPO_PESO_CAMION = 30  'Toneladas
Const SECCION_CUPO = 1
Const SECCION_SIN_CUPO = 2

'Esta funcion devuelve los cupos que fueron asignados y recibidos correctamente con su respectivo numero de cupo
Function armarSQLConCupos(pFechaDesde, pFechaHasta, pCdProducto)
    Dim strSQL, fechaDesde, fechaHasta, auxFechaDesde, auxFechaHasta
	
    fechaDesde = GF_nDigits(Year(pFechaDesde),4) & GF_nDigits(Month(pFechaDesde),2) & GF_nDigits(Day(pFechaDesde),2)
	fechaHasta = GF_nDigits(Year(pFechaHasta),4) & GF_nDigits(Month(pFechaHasta),2) & GF_nDigits(Day(pFechaHasta),2)	
    
    auxFechaDesde = GF_nDigits(Year(pFechaDesde),4) &"-"& GF_nDigits(Month(pFechaDesde),2) &"-"& GF_nDigits(Day(pFechaDesde),2)
    auxFechaHasta = GF_nDigits(Year(pFechaHasta),4) &"-"& GF_nDigits(Month(pFechaHasta),2) &"-"& GF_nDigits(Day(pFechaHasta),2)
    
    strSQL = "  SELECT CASE WHEN CRR.CDCORREDOR IS NULL THEN 0 ELSE CRR.CDCORREDOR END AS CDCORREDOR, "&_
             "         CASE WHEN CRR.DSCORREDOR IS NULL THEN '' ELSE CRR.DSCORREDOR END AS DSCORREDOR, "&_
             "         CASE WHEN V.CDVENDEDOR IS NULL THEN 0 ELSE V.CDVENDEDOR END AS CDVENDEDOR, "&_
             "         CASE WHEN V.DSVENDEDOR IS NULL THEN '' ELSE V.DSVENDEDOR END AS DSVENDEDOR, "&_
             "         CASE WHEN CLI.CDCLIENTE IS NULL THEN 0 ELSE CLI.CDCLIENTE END AS CDCLIENTE, "&_
             "         CASE WHEN CLI.DSCLIENTE IS NULL THEN '' ELSE CLI.DSCLIENTE END AS DSCLIENTE, "&_
             "         CUP.FECHA, "&_
             "         PRO.DSPRODUCTO, "&_
	         "         COUNT(CUP.CODIGOCUPO) AS ASIGNADOS, "&_
	         "         COUNT(CAM.CUPO) AS INGRESADOS "&_
             " FROM   (SELECT case when cdcorredor = 0 or cdcorredor = 15 or cdcorredor=cdvendedor then 0 else cdcorredor end as CORREDOR, "&_
			 "                  case when cdvendedor = 0 or cdvendedor = 15  then 0 else cdvendedor end as VENDEDOR, "&_
			 "                  (select min(cdcliente) from clientes where rtrim(nucuit) = rtrim(cuitcliente)) CLIENTE, "&_
			 "                  fechacupo  FECHA, "&_
			 "                  cdproducto PRODUCTO, "&_
			 "                  CODIGOCUPO  "&_
		     "           FROM codigoscupo  "&_
		     "           WHERE  FECHACUPO >= "& fechaDesde &" AND FECHACUPO < "& fechaHasta &" AND ESTADO > " & CUPO_CANCELADO & " and cdproducto = "& pCdProducto &" ) CUP "&_
             "   LEFT JOIN 	 "&_
	         "       (SELECT CORREDOR,VENDEDOR,CLIENTE,FECHA,PRODUCTO,CUPO,IDCAMION "&_
	         "        FROM (( SELECT CD.cdcorredor CORREDOR, "&_
			 "		                 CD.cdvendedor VENDEDOR, "&_
			 "		                 CD.CDCLIENTE CLIENTE, "&_
			 "		                 ( Year(CD.dtcontable) * 10000 + Month(CD.dtcontable) * 100 + Day(CD.dtcontable) ) FECHA, "&_
			 "		                 C.cdproducto  PRODUCTO, "&_
			 "		                 C.NUCUPO AS CUPO, "&_
			 "		                 C.IDCAMION "&_
			 "                 FROM   hcamiones C "&_
			 "	                INNER JOIN hcamionesdescarga CD ON C.idcamion = CD.idcamion AND C.dtcontable = CD.dtcontable "&_
			 "                  WHERE  CD.dtcontable >= '"& auxFechaDesde &"' AND CD.dtcontable < '"& auxFechaHasta &"' AND C.cdestado <> "& CAMIONES_ESTADO_BAJA &" AND C.cdproducto = "& pCdProducto &")"&_
			 "                  UNION "&_
			 "                  (SELECT CD.cdcorredor CORREDOR, "&_
		 	 "		                  CD.cdvendedor VENDEDOR, "&_
			 "		                  CD.CDCLIENTE  CLIENTE, "&_
			 "		                   "& Left(session("MmtoDato"),8) &" AS FECHA, "&_
			 "		                   C.cdproducto  PRODUCTO, "&_
			 "		                   C.NUCUPO AS CUPO, "&_
			 "		                   C.IDCAMION "&_
			 "	                FROM   camiones C "&_
			 "		                INNER JOIN camionesdescarga CD ON C.idcamion = CD.idcamion "&_
			 "	                WHERE  C.cdproducto = "& pCdProducto &" AND C.cdestado <> "& CAMIONES_ESTADO_BAJA &" ) )TT "&_
             "          WHERE  TT.fecha >= "& fechaDesde &" AND TT.fecha < "& fechaHasta &_
             "         ) CAM "&_
             "         ON  CAM.CUPO =  CUP.CODIGOCUPO AND CAM.FECHA = CUP.FECHA  "&_
             "     LEFT JOIN CLIENTES CLI ON CLI.CDCLIENTE = CUP.CLIENTE "&_
             "     INNER JOIN PRODUCTOS PRO ON PRO.CDPRODUCTO = CUP.PRODUCTO "&_
             "     LEFT JOIN CORREDORES CRR ON CRR.CDCORREDOR = CUP.CORREDOR  "&_
             "     LEFT JOIN VENDEDORES V ON V.CDVENDEDOR = CUP.VENDEDOR "&_
             "    GROUP BY CRR.CDCORREDOR,CRR.DSCORREDOR,V.CDVENDEDOR,V.DSVENDEDOR,CLI.CDCLIENTE,CLI.DSCLIENTE, CUP.FECHA, PRO.DSPRODUCTO "&_
             "    ORDER BY CRR.CDCORREDOR,CRR.DSCORREDOR,V.CDVENDEDOR,V.DSVENDEDOR,CLI.CDCLIENTE,CLI.DSCLIENTE, CUP.FECHA, PRO.DSPRODUCTO  "
            
    'Response.Write strSQL &"<BR><BR>"
    Call GF_BD_Puertos(gv_pto, rs, "OPEN", strSQL)
	Set armarSQLConCupos = rs
End function
'--------------------------------------------------------------------------------------------
Function armarSQLSinCupos(pFechaDesde, pFechaHasta, pCdProducto)
    Dim strSQL, fechaDesde, fechaHasta, auxFechaDesde, auxFechaHasta
	
    fechaDesde = GF_nDigits(Year(pFechaDesde),4) & GF_nDigits(Month(pFechaDesde),2) & GF_nDigits(Day(pFechaDesde),2)
	fechaHasta = GF_nDigits(Year(pFechaHasta),4) & GF_nDigits(Month(pFechaHasta),2) & GF_nDigits(Day(pFechaHasta),2)	
    
    auxFechaDesde = GF_nDigits(Year(pFechaDesde),4) &"-"& GF_nDigits(Month(pFechaDesde),2) &"-"& GF_nDigits(Day(pFechaDesde),2)
    auxFechaHasta = GF_nDigits(Year(pFechaHasta),4) &"-"& GF_nDigits(Month(pFechaHasta),2) &"-"& GF_nDigits(Day(pFechaHasta),2)


    strSQL = "  SELECT CASE WHEN CRR.CDCORREDOR IS NULL or CRR.CDCORREDOR=V.CDVENDEDOR THEN 0 ELSE CRR.CDCORREDOR END AS CDCORREDOR, "&_    
             "         CASE WHEN CRR.DSCORREDOR IS NULL THEN '' ELSE CRR.DSCORREDOR END AS DSCORREDOR, "&_
             "         CASE WHEN V.CDVENDEDOR IS NULL THEN 0 ELSE V.CDVENDEDOR END AS CDVENDEDOR, "&_
             "         CASE WHEN V.DSVENDEDOR IS NULL THEN '' ELSE V.DSVENDEDOR END AS DSVENDEDOR, "&_
             "         CASE WHEN CLI.CDCLIENTE IS NULL THEN 0 ELSE CLI.CDCLIENTE END AS CDCLIENTE, "&_
             "         CASE WHEN CLI.DSCLIENTE IS NULL THEN '' ELSE CLI.DSCLIENTE END AS DSCLIENTE, "&_
             "         CAM.FECHA, "&_
             "         PRO.DSPRODUCTO, "&_
	         "         COUNT(CAM.CUPO) AS INGRESADOS "&_	         
             "   FROM    "&_
	         "       (SELECT case when TT.CORREDOR = 0 or TT.CORREDOR = 15 then 0 else TT.CORREDOR end as CORREDOR, "&_
			 "               case when TT.vendedor = 0 or TT.vendedor = 15 then 0 else TT.vendedor end as VENDEDOR, "&_
			 "               CLI.CDCLIENTE CLIENTE, "&_
			 "               FECHA,PRODUCTO,CUPO,IDCAMION "&_
	         "        FROM (( SELECT CD.cdcorredor CORREDOR, "&_
			 "		                CD.cdvendedor VENDEDOR, "&_
			"		                CD.CDCLIENTE CLIENTE, "&_
			"		                ( Year(CD.dtcontable) * 10000 + Month(CD.dtcontable) * 100 + Day(CD.dtcontable) ) FECHA, "&_
			"		                C.cdproducto  PRODUCTO, "&_
			"		                C.NUCUPO AS CUPO, "&_
			"		                C.IDCAMION "&_
			 "                 FROM   hcamiones C "&_
			"	                INNER JOIN hcamionesdescarga CD ON C.idcamion = CD.idcamion AND C.dtcontable = CD.dtcontable "&_
			"                  WHERE  CD.dtcontable >= '"& auxFechaDesde &"' AND CD.dtcontable < '"& auxFechaHasta &"' AND C.cdestado <> "& CAMIONES_ESTADO_BAJA &" AND C.cdproducto = "& pCdProducto &")"&_
			"                  UNION "&_
			"                  (SELECT CD.cdcorredor CORREDOR, "&_
			"		                   CD.cdvendedor VENDEDOR, "&_
			"		                   CD.CDCLIENTE  CLIENTE, "&_
			"		                   "& Left(session("MmtoDato"),8) &" AS FECHA, "&_
			"		                   C.cdproducto  PRODUCTO, "&_
			"		                   C.NUCUPO AS CUPO, "&_
			"		                   C.IDCAMION "&_
			"	                FROM   camiones C "&_
			"		                INNER JOIN camionesdescarga CD ON C.idcamion = CD.idcamion "&_
			"	                WHERE  C.cdproducto = "& pCdProducto &" AND C.cdestado <> "& CAMIONES_ESTADO_BAJA &" ) )TT "&_
		    "           LEFT JOIN CLIENTES CLI ON TT.CLIENTE = CLI.CDCLIENTE "&_
            "          WHERE  TT.fecha >= "& fechaDesde &" AND TT.fecha < "& fechaHasta &_
            "         ) CAM "&_
            "     LEFT JOIN CODIGOSCUPO CUP ON  CAM.CUPO =  CUP.CODIGOCUPO AND CAM.FECHA = CUP.FECHACUPO "&_
            "     LEFT JOIN CLIENTES CLI ON CLI.CDCLIENTE = CAM.CLIENTE "&_
            "     INNER JOIN PRODUCTOS PRO ON PRO.CDPRODUCTO = CAM.PRODUCTO "&_
            "     LEFT JOIN CORREDORES CRR ON CRR.CDCORREDOR = CAM.CORREDOR  "&_
            "     LEFT JOIN VENDEDORES V ON V.CDVENDEDOR = CAM.VENDEDOR "&_
            "    WHERE (CUP.CODIGOCUPO IS NULL or CUP.ESTADO = " & CUPO_CANCELADO & ")"&_
            "    GROUP BY CRR.CDCORREDOR,CRR.DSCORREDOR,V.CDVENDEDOR,V.DSVENDEDOR,CLI.CDCLIENTE,CLI.DSCLIENTE,CAM.FECHA, PRO.DSPRODUCTO "&_
            "    ORDER BY CRR.CDCORREDOR,CRR.DSCORREDOR,V.CDVENDEDOR,V.DSVENDEDOR,CLI.CDCLIENTE,CLI.DSCLIENTE,CAM.FECHA, PRO.DSPRODUCTO "
    'Response.Write strSQL &"<BR><BR>"
   ' Response.End
    Call GF_BD_Puertos(gv_pto, rs, "OPEN", strSQL)
	Set armarSQLSinCupos = rs
End Function
'--------------------------------------------------------------------------------------------
Function armarSQLCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto)
    Dim strSQL, fechaDesde, fechaHasta     	
	fechaDesde = GF_nDigits(Year(pFechaDesde),4) & GF_nDigits(Month(pFechaDesde),2) & GF_nDigits(Day(pFechaDesde),2)
	fechaHasta = GF_nDigits(Year(pFechaHasta),4) & GF_nDigits(Month(pFechaHasta),2) & GF_nDigits(Day(pFechaHasta),2)		
	
	strSQL = "	Select T.*, " &_
			 "    		CASE WHEN CLI.DSCLIENTE IS NULL THEN '' ELSE CLI.DSCLIENTE END AS DSCLIENTE "&_
			 "	From ( " &_
			 "		SELECT 	CUP.CDCORREDOR, "&_
			 "      	   	CASE WHEN CRR.DSCORREDOR IS NULL THEN '' ELSE CRR.DSCORREDOR END AS DSCORREDOR, " &_
	         "		       	CUP.CDVENDEDOR, "&_
			 "      		CASE WHEN V.DSVENDEDOR IS NULL THEN '' ELSE V.DSVENDEDOR END AS DSVENDEDOR, "&_
             "             	(SELECT MIN(CDCLIENTE) FROM CLIENTES WHERE NUCUIT = CUP.CUITCLIENTE) CLIENTE, "&_			 
	         "		       	CUP.FECHACUPO , "&_
	         "	           	CUP.CDPRODUCTO, "&_
             "             	CUP.CONDICION , "&_
	         "		       	CUP.QTINGRESADOS, "&_
	         "             	CUP.QTASIGNADOS "&_
	         "      FROM  CODIGOSCUPOESPECIALES AS CUP "&_	        
			 " 		INNER JOIN PRODUCTOS PRO ON CUP.CDPRODUCTO = PRO.CDPRODUCTO "&_
             " 		LEFT JOIN CORREDORES CRR ON CUP.CDCORREDOR = CRR.CDCORREDOR "&_
             " 		LEFT JOIN VENDEDORES V ON CUP.CDVENDEDOR = V.CDVENDEDOR "&_
			 "      WHERE CUP.FECHACUPO >= " & fechaDesde & " AND CUP.FECHACUPO < " & fechaHasta & " AND CUP.CDPRODUCTO = "& pCdProducto &_	                                   
             "		) T " &_
			 " 		LEFT JOIN CLIENTES CLI ON CLI.CDCLIENTE = T.CLIENTE "&_	         
             " ORDER BY T.DSCORREDOR, T.DSVENDEDOR, DSCLIENTE, T.CONDICION, T.FECHACUPO"
			 
    Call GF_BD_Puertos(gv_pto, rs, "OPEN", strSQL)
	Set armarSQLCuposEspeciales = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	armarSQLTotalCuposEspeciales
' Autor: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	18/01/13
' Objetivo:	
'			Arma la Sql para totalizar los cupos Especiales
' Parametros:
'			pFechaDesde 	[date] 	fecha Inicio
'			pFechaHasta 	[date] 	fecha Final
'Devuelve : 
'			RecordSet con los cupos Especiales				
'--------------------------------------------------------------------------------------------
Function armarSQLTotalCuposEspeciales(pFechaDesde, pFechaHasta,pProducto)
    Dim strSQL, fechaDesde, fechaHasta     	
	fechaDesde = GF_nDigits(Year(pFechaDesde),4) & GF_nDigits(Month(pFechaDesde),2) & GF_nDigits(Day(pFechaDesde),2)
	fechaHasta = GF_nDigits(Year(pFechaHasta),4) & GF_nDigits(Month(pFechaHasta),2) & GF_nDigits(Day(pFechaHasta),2)		
	
	strSQL = "			SELECT 	CUP.FECHACUPO FECHA,"
	strSQL = strSQL & "		Sum(CUP.qtasignados)  QTASIGNADOS, "		
	strSQL = strSQL & "     Sum(CUP.qtingresados) QTINGRESADOS "	
	strSQL = strSQL & " FROM   CODIGOSCUPOESPECIALES CUP "
	strSQL = strSQL & " WHERE   CUP.FECHACUPO >= " & fechaDesde & " AND CUP.FECHACUPO < " & fechaHasta
	strSQL = strSQL & "		AND CUP.cdproducto = " & pProducto 
	strSQL = strSQL & " GROUP  BY CUP.FECHACUPO "
	
	Call GF_BD_Puertos(gv_pto, rs, "OPEN", strSQL)
	Set armarSQLTotalCuposEspeciales = rs
End Function
'--------------------------------------------------------------------------------------------
' Función:	armarCabeceraCupos
' Autor: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	14/01/13
' Objetivo:	
'			Arma la cabecera con los titulos de la Tabla
' Parametros:
'			pFechaDesde 	[date] 	fecha Inicio
'			pFechaHasta 	[date] 	fecha Final
'--------------------------------------------------------------------------------------------
Function armarCabeceraCupos(pFechaDesde, pFechaHasta, pIsNotEspecial, pMedia)
	Dim myFecha, myCuposEspeciales
	myCuposEspeciales = "No"
	if(gv_cuposEspeciales)then myCuposEspeciales = "Si"
	%>	
	<TABLE width="100%">
	<TBODY class='xls_border_left' >
		<TR style="font-size:14;">
			<TD align=left colspan="17"><B><%=GF_TRADUCIR("ACTI AR")%></B></TD>
			<TD align=left colspan="4"><B><%= GF_TRADUCIR("Fecha: " & Left(GF_FN2DTE(session("MmtoSistema")),10))  %></B></TD>
		</TR>
		<TR style="font-size:14;">
			<TD align=left colspan="17"><B><%=GF_TRADUCIR("Pto: " & Ucase(gv_pto))%></B></TD>
			<TD align=left colspan="4"><B><%= GF_TRADUCIR("Hora: " & Right(GF_FN2DTE(session("MmtoSistema")),8))  %></B></TD>
		</TR>		
		<BR></BR>
		<TR style="font-size:12;">
			<TD colspan="2" align="center" ><B>
				<% if(pIsNotEspecial)then 
                        if (pMedia = TIPO_AFIRMACION) then  %>
                            <h1><%=GF_TRADUCIR("CONSULTA DE CUPOS : " & getDsProducto(gv_cdProducto)) %></h1>
                        <% else %>
					        <%=GF_TRADUCIR("REPORTE DE CUPOS ASIGNADOS") %>
				<%      end if
                    else 
                        if (pMedia = TIPO_AFIRMACION) then  %>
                            <h1><%=GF_TRADUCIR("CUPOS ESPECIALES") %></h1>
                        <% else %>
					        <%=GF_TRADUCIR("CUPOS ESPECIALES") %>
                        <% end if 
				   end if %>	
			</B></TD>					
		</TR>        
		<% if ((pIsNotEspecial) and (pMedia <> TIPO_AFIRMACION)) then %>
		<TR style="font-size:12;line-height:50%">
			<TD align=left><B><%=GF_TRADUCIR("FECHA DESDE: ") %></B></TD>
			<TD align=left colspan="2"><B><%= GF_nDigits(Day(pFechaDesde), 2) & "/" & GF_nDigits(Month(pFechaDesde), 2) & "/" & GF_nDigits(Year(pFechaDesde), 4) %></B></TD>					
		</TR>		
		<TR style="font-size:12;line-height:50%">		
			<TD align=left><B><%=GF_TRADUCIR("FECHA HASTA: ") %></B></TD>
			<TD align=left colspan="2"><B><%= GF_nDigits(Day(pFechaHasta) - 1, 2) & "/" & GF_nDigits(Month(pFechaHasta), 2) & "/" & GF_nDigits(Year(pFechaHasta), 4) %></B></TD>			
		</TR>		
		<TR style="font-size:12;line-height:50%">
			<TD align=left><B><%=GF_TRADUCIR("PRODUCTO: ")  %></B></TD>			
			<TD align=left colspan="2"><B><%=gv_cdProducto & " - " & getDsProducto(gv_cdProducto) %></B></TD>
		</TR>		
		<TR style="font-size:12;line-height:50%">			
			<TD align=left><B><%=GF_TRADUCIR("CUPOS ESPECIALES: ")  %></B></TD>			
			<TD align=left colspan="2"><B><%=myCuposEspeciales %></B></TD>
		</TR>				
		<% end if %>
	</TBODY>		
	</TABLE>
<%	
End Function
'-----------------------------------------------------------------------------------------------------------------
Function dibujarTotalesPorSemana(pTotalCumplidos, pTotalAsignados)
	writeXLS("<TD align=right style='width:75px;font-size:12; text-align:center;' bgcolor='#F2F2F2'>" & cstr(pTotalCumplidos & " | " & pTotalAsignados) & "</TD>")
	writeXLS("<TD align=right style='width:75px;font-size:12; text-align:center;' bgcolor='#F2F2F2'>" & cstr(pTotalCumplidos * CUPO_PESO_CAMION & " | " & pTotalAsignados * CUPO_PESO_CAMION) & "</TD>")	
End Function
'--------------------------------------------------------------------------------------------
Function dibujarLineaCupo(p_row, ByRef p_arrayCupoAsignados,ByRef p_arrayCupoRecibidos,p_dsCorredor,p_dsVendedor,p_dsCliente,pIsEspecial,p_Contrato)
    Dim htmlTable,parcialRecibidosSemana,parcialAsignadosSemana,myCupoDia, myClase
    parcialAsignadosSemana = 0
    parcialRecibidosSemana = 0
    gv_linea = gv_linea + 1
    myClase = "lineacupo"
    if ((gv_linea mod 2) = 0) then myClase = "lineacupo2"
    %> 
     <tr class="<% =myClase %>">
      <% if (pIsEspecial) then %>
         <td align=left style="width:120px; font-size:12;"><%=p_Contrato%></td>
      <% end if %>
        <td align=left style="width:120px; font-size:12;"><%=p_dsCorredor%></td>
        <td align=left style="width:120px; font-size:12;"><%=p_dsVendedor%></td>
        <td align=left style="width:120px; font-size:12;"><%=p_dsCliente%></td> 
    <%  for i = 0 to Ubound(p_arrayCupoAsignados, 2) -1
            myCupoDia = ""
            if ((not IsEmpty(p_arrayCupoAsignados(p_row, i)))or(not IsEmpty(p_arrayCupoRecibidos(p_row, i)))) then 
                myCupoDia = p_arrayCupoRecibidos(p_row, i) & " | " & p_arrayCupoAsignados(p_row, i)
                parcialAsignadosSemana = parcialAsignadosSemana + CDbl(p_arrayCupoAsignados(p_row, i))
				parcialRecibidosSemana = parcialRecibidosSemana + Cdbl(p_arrayCupoRecibidos(p_row, i))
                g_arrayTotalCuposAsignadosDia(i) = Cdbl(g_arrayTotalCuposAsignadosDia(i)) + Cdbl(p_arrayCupoAsignados(p_row, i))    
                g_arrayTotalCuposRecibidosDia(i) = Cdbl(g_arrayTotalCuposRecibidosDia(i)) + Cdbl(p_arrayCupoRecibidos(p_row, i))
                
            end if %>
            <TD align=right  style="font-size:12; text-align:center;" style="width:70px" <% if (Cdbl(p_arrayCupoRecibidos(p_row, i)) > Cdbl(p_arrayCupoAsignados(p_row, i))) then%> class="cls_cupos" <% end if %>><%=myCupoDia %></TD>
        <%  p_arrayCupoAsignados(p_row, i) = empty
            p_arrayCupoRecibidos(p_row, i) = empty 
            'Si el array llega a cumpli la semana entonces se dibuja los totales de la primera semana
            if(i = (CUPO_CANTIDAD_DIAS_SEMANA - 1))then	
			    Call dibujarTotalesPorSemana(parcialRecibidosSemana, parcialAsignadosSemana)
				parcialAsignadosSemana = 0
				parcialRecibidosSemana = 0
            end if
        next
        'Se dibujan los totales de la segunda semana
        Call dibujarTotalesPorSemana(parcialRecibidosSemana, parcialAsignadosSemana) %>
    </tr>
<%
End function
'--------------------------------------------------------------------------------------------
Function dibujarTotalesPorFecha(p_Titulo, p_ArrayTotalRecibidos, p_ArrayTotalAsignados,p_IsEspecial)
    Dim sumTotalRecibidos,sumTotalAsignados,sumTotalRecibidosTon,sumTotalAsignadosTon
    'Primero se procesa la primera semana
    sumTotalRecibidos = 0
    sumTotalAsignados= 0  %>
    <TR style='background-color:#F2F2F2'>
	    <TD align=left style="font-size:12;" rowspan="2" colspan='<% if (p_IsEspecial) then response.write "3" else response.write "2" end if %>' ><%="<B>"& p_Titulo &"</B>"%></TD>    
        <TD align=left style="font-size:12;"><B>Recibidos</B></TD>
    <%  for i = 0 to (CUPO_CANTIDAD_DIAS_SEMANA - 1)
	        if(p_ArrayTotalRecibidos(i) = "")then p_ArrayTotalRecibidos(i)= 0 %>
			<TD align=right style="font-size:14; text-align: center; font-weight: bold;"><%=Cstr(p_ArrayTotalRecibidos(i))%></TD>
    <%		sumTotalRecibidos = sumTotalRecibidos + p_ArrayTotalRecibidos(i)            
		next
        sumTotalRecibidosTon = sumTotalRecibidos * CUPO_PESO_CAMION  %>
        <TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%= cstr(sumTotalRecibidos) %></TD>
		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%= cstr(sumTotalRecibidosTon) %></TD>

    <%  
        'Luego se procesa la segunda semana
        sumTotalRecibidos = 0
        for i = CUPO_CANTIDAD_DIAS_SEMANA  to ((CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)- 1)
            if(p_ArrayTotalRecibidos(i) = "")then p_ArrayTotalRecibidos(i)= 0 %>
            <TD align=right style="width:20px;font-size:14; text-align: center; font-weight: bold;"><%=Cstr(p_ArrayTotalRecibidos(i)) %></TD>
	<%		sumTotalRecibidos = sumTotalRecibidos + p_ArrayTotalRecibidos(i)            
	    next
		sumTotalRecibidosTon = sumTotalRecibidos * CUPO_PESO_CAMION %>

		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%= cstr(sumTotalRecibidos) %></TD>
		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%= cstr(sumTotalRecibidosTon) %></TD>
    </TR>    
    <TR style='background-color:#F2F2F2'>
        <TD align=left style="font-size:12;"><B>Asignados</B></TD>	    
    <%  for i = 0 to (CUPO_CANTIDAD_DIAS_SEMANA - 1)	        
            if(p_ArrayTotalAsignados(i) = "")then p_ArrayTotalAsignados(i)= 0 %>
			<TD align=right style="font-size:14; text-align: center; font-weight: bold;"><%=Cstr(p_ArrayTotalAsignados(i))%></TD>
    <%		sumTotalAsignados = sumTotalAsignados + p_ArrayTotalAsignados(i)
		next    
        sumTotalAsignadosTon = sumTotalAsignados * CUPO_PESO_CAMION  %>
        <TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%=cstr(sumTotalAsignados) %></TD>
		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%=cstr(sumTotalAsignadosTon) %></TD>

    <%  
        'Luego se procesa la segunda semana        
        sumTotalAsignados= 0
        for i = CUPO_CANTIDAD_DIAS_SEMANA  to ((CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)- 1)
            if(p_ArrayTotalAsignados(i) = "")then p_ArrayTotalAsignados(i)= 0  %>
            <TD align=right style="width:20px;font-size:14; text-align: center; font-weight: bold;"><%=Cstr(p_ArrayTotalAsignados(i))%></TD>
	<%		sumTotalAsignados = sumTotalAsignados + p_ArrayTotalAsignados(i)
	    next		
        sumTotalAsignadosTon = sumTotalAsignados * CUPO_PESO_CAMION %>

		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%=cstr(sumTotalAsignados) %></TD>
		<TD align=right style='width:30px;font-size:14; text-align: center; font-weight: bold;' ><%=cstr(sumTotalAsignadosTon) %></TD>
    </TR>
    <%
End function
'--------------------------------------------------------------------------------------------
'Obtiene los contratos especiales para una determinada fecha y producto
Function procesarTotalesCuposEspeciales(pFechaDesde, pFechaHasta, pProducto)
    Dim rsCupEspecial,arrayCupoRecibidosEspeciales(),arrayCupoAsignadosEspeciales(),indexArr
    
    Redim arrayCupoRecibidosEspeciales(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    Redim arrayCupoAsignadosEspeciales(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    
    Set rsCupEspecial = armarSQLTotalCuposEspeciales(pFechaDesde, pFechaHasta, pProducto)
    while (not rsCupEspecial.EoF)
        'Obtengo la clave de la fecha para el indice del array
        indexArr = g_dicFechas.Item(Right(rsCupEspecial("FECHA"),4))
        arrayCupoAsignadosEspeciales(indexArr) = CDbl(rsCupEspecial("QTASIGNADOS"))
        arrayCupoRecibidosEspeciales(indexArr) = CDbl(rsCupEspecial("QTINGRESADOS"))
        rsCupEspecial.MoveNext()
    wend
    Call dibujarTotalesPorFecha("TOTAL CUPOS ESPECIALES", arrayCupoRecibidosEspeciales,arrayCupoAsignadosEspeciales,false)
End function 
'--------------------------------------------------------------------------------------------
Function procesarCupoArray(p_Fecha, p_cdCorredor, p_cdVendedor, p_cdCliente, p_dsCorredor, p_dsVendedor, p_dsCliente, p_CuposRecibidos, p_CuposAsignados, ByRef p_arrayCupoRecibidos, ByRef p_arrayCupoAsignados, ByRef p_arrayParticipantes)
    Dim indexCol, indexRow	
	'On Error resume next	
    indexCol = g_dicFechas.Item(Right(p_Fecha,4))	
	if (not g_dicParticipantes.Exists(p_cdCorredor & "_" & p_cdVendedor & "_" & p_cdCliente)) then		
		Call g_dicParticipantes.Add(p_cdCorredor & "_" & p_cdVendedor & "_" & p_cdCliente, g_idxParticipantes)
		p_arrayParticipantes(g_idxParticipantes, 1) = p_dsCorredor
		p_arrayParticipantes(g_idxParticipantes, 2) = p_dsVendedor
		p_arrayParticipantes(g_idxParticipantes, 3) = p_dsCliente
		g_idxParticipantes = g_idxParticipantes + 1		
	end if	
	indexRow = g_dicParticipantes.Item(p_cdCorredor & "_" & p_cdVendedor & "_" & p_cdCliente)
    p_arrayCupoRecibidos(indexRow, indexCol) = CLng(p_arrayCupoRecibidos(indexRow, indexCol)) + CLng(p_CuposRecibidos)
    p_arrayCupoAsignados(indexRow, indexCol) = CLng(p_arrayCupoAsignados(indexRow, indexCol)) + CLng(p_CuposAsignados)	
End Function
'--------------------------------------------------------------------------------------------
' Función:	armarExelCupos
' Autor: 	CNA - Ajaya Cesar Nahuel
' Autor modificacion: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	14/01/13
' Fecha modificacion: 	14/04/14
' Objetivo:	
'			Arma la Tabla de Cupos.
'			Se modifico para que ademas de imprimir todos los cupos que se asignaron y cumplieron, haga tambien 
'			lo mismo para aquellos ingresos que tuvo la planta sin tener cupo necesariamente. Igualmente para
'			la suma de los asignados por semana se consideran aquellos que tengan cupo
' Parametros:
'			pFechaDesde 	[date] 	fecha Inicio
'			pFechaHasta 	[date] 	fecha Final
'--------------------------------------------------------------------------------------------
Function armarExelCupos(pFechaDesde, pFechaHasta, pCdProducto, pMedia)
	Dim rsCupos, arrayCupoRecibidos(),arrayCupoAsignados(), arrayParticipantes(), rsCuposCC, rsCuposSC, r
    Dim cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,dsCorredor,dsVendedor,dsCliente

	g_idxParticipantes = 1
    Call loadPeriodoFecha(pFechaDesde)

	Set rsCuposCC = armarSQLConCupos(pFechaDesde, pFechaHasta, pCdProducto) 
    Set rsCuposSC = armarSQLSinCupos(pFechaDesde, pFechaHasta, pCdProducto)
    
	Call armarCabeceraCupos(pFechaDesde, pFechaHasta, true, pMedia)
	
    'creao el array que tendra los valores de cada fila
    Redim arrayCupoRecibidos(500, CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    Redim arrayCupoAsignados(500, CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)	
    Redim arrayParticipantes(500, 3)
    'creo los array para totalizar por planta o buenos aires el dia
    Redim g_arrayTotalCuposRecibidosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
	Redim g_arrayTotalCuposAsignadosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)   
	
    if not rsCuposCC.eof or not rsCuposSC.eof then %>
        <TABLE width="100%">
		<TBODY class='xls_border_left'; font-size:6;" >                  
		    <TR style='background-color:#BDBDBD;'>
                <TD align=left style="width:200px;font-size:12; font-weight: bold; "><%=GF_TRADUCIR("Corredor")%></TD>
			    <TD align=left style="width:200px;font-size:12; font-weight: bold;"><%=GF_TRADUCIR("Vendedor")%></TD>			
                <TD align=left style="width:200px;font-size:12; font-weight: bold;"><%=GF_TRADUCIR("Destinatario")%></TD>
	    <%	For z = 0 to CUPO_CANTIDAD_DIAS_SEMANA - 1
			    myFecha = DateAdd("d", z, pFechaDesde) %>
			    <TD align=center style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & " / " & GF_nDigits(Month(myFecha),2))%></TD>						
	    <%	Next	%>				
		    <TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR("Total Sem.1")%> </TD>
		    <TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR("Total Ton.1")%>	</TD>	
		
	    <%	For z = CUPO_CANTIDAD_DIAS_SEMANA  to (CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA) - 1
			    myFecha = DateAdd("d", z, pFechaDesde) %>
			    <TD align=center style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & " / " & GF_nDigits(Month(myFecha),2))%></TD>						
	    <%	Next	%>				
		    <TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;"><%=GF_TRADUCIR("Total Sem.2")%> </TD>
		    <TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;"><%=GF_TRADUCIR("Total Ton.2")%>	</TD>	
		    </TR>	    
	<%  while (not rsCuposCC.eof)			
			Call procesarCupoArray(rsCuposCC("FECHA"), rsCuposCC("CDCORREDOR"), rsCuposCC("CDVENDEDOR"), rsCuposCC("CDCLIENTE"), rsCuposCC("DSCORREDOR"), rsCuposCC("DSVENDEDOR"), rsCuposCC("DSCLIENTE"), rsCuposCC("INGRESADOS"),rsCuposCC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados, arrayParticipantes)
			rsCuposCC.MoveNext()
		wend		
		while (not rsCuposSC.eof)
			Call procesarCupoArray(rsCuposSC("FECHA"), rsCuposSC("CDCORREDOR"), rsCuposSC("CDVENDEDOR"), rsCuposSC("CDCLIENTE"), rsCuposSC("DSCORREDOR"), rsCuposSC("DSVENDEDOR"), rsCuposSC("DSCLIENTE"), rsCuposSC("INGRESADOS"), 0, arrayCupoRecibidos, arrayCupoAsignados, arrayParticipantes)
			rsCuposSC.MoveNext()
		wend
        'Dibujo los cupos 			
		for r = 1 to g_idxParticipantes
			Call dibujarLineaCupo(r, arrayCupoAsignados,arrayCupoRecibidos,arrayParticipantes(r, 1),arrayParticipantes(r, 2),arrayParticipantes(r, 3),false,"")
		Next
        Call dibujarTotalesPorFecha("TOTAL CUPOS", g_arrayTotalCuposRecibidosDia,g_arrayTotalCuposAsignadosDia,false)
        Call procesarTotalesCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto) 
        %>
    </TBODY></TABLE>
<%  else  %>
        <TABLE class='xls_border_left'  style="width:100%; font-size:12;">
    		<TR >
			    <TD align=center colspan="20" ><%= GF_TRADUCIR("<B>NO SE ENCONTRARON RESULTADOS</B>")%></TD>		
		    </TR>
	    </TABLE> 
<%  end if %>
<%      
    
End Function
'-----------------------------------------------------------------------------------------------
Function armarExelCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto, pMedia)
	Dim rsCupos, arrayCupoRecibidos(),arrayCupoAsignados(), rsCuposCC, rsCuposSC, condicion
    Dim cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,dsCorredor,dsVendedor,dsCliente
	
    Set rsCupos = armarSQLCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto)	
	
	Call armarCabeceraCupos(pFechaDesde, pFechaHasta, false, pMedia)
    'creo los array para totalizar por planta o buenos aires el dia
    Redim g_arrayTotalCuposRecibidosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
	Redim g_arrayTotalCuposAsignadosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)   

	if(not rsCupos.EoF)then	%>
		<TABLE>
			<TBODY class='xls_border_left'  style="width:80%; font-size:6;" >				
				<TR style='background-color:#BDBDBD;'>
				    <TD align=left style="width:100px;font-size:12;"><%=GF_TRADUCIR("Contrato")%></TD>			    
					<TD align=left style="width:200px;font-size:12; font-weight: bold; "><%=GF_TRADUCIR("Corredor")%></TD>
					<TD align=left style="width:200px;font-size:12; font-weight: bold;"><%=GF_TRADUCIR("Vendedor")%></TD>			
					<TD align=left style="width:200px;font-size:12; font-weight: bold;"><%=GF_TRADUCIR("Destinatario")%></TD>
					<%	For z = 0 to CUPO_CANTIDAD_DIAS_SEMANA - 1
							myFecha = DateAdd("d", z, pFechaDesde) %>
							<TD align=center style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & " / " & GF_nDigits(Month(myFecha),2))%></TD>						
					<%	Next	%>				
						<TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR("Total Sem.1")%> </TD>
						<TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR("Total Ton.1")%>	</TD>	
					
					<%	For z = CUPO_CANTIDAD_DIAS_SEMANA  to (CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA) - 1
							myFecha = DateAdd("d", z, pFechaDesde) %>
							<TD align=center style="width:50px;font-size:12; font-weight: bold; text-align:center;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & " / " & GF_nDigits(Month(myFecha),2))%></TD>						
					<%	Next	%>				
					<TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;"><%=GF_TRADUCIR("Total Sem.2")%> </TD>
					<TD align=left style="width:50px;font-size:12; font-weight: bold; text-align:center;"><%=GF_TRADUCIR("Total Ton.2")%>	</TD>	
		    </TR>	   
            <% 	while(not rsCupos.EoF) 
					Redim arrayCupoRecibidos(1, CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
					Redim arrayCupoAsignados(1, CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
                    vendedor_old = Cdbl(rsCupos("CDVENDEDOR"))
					corredor_old = Cdbl(rsCupos("CDCORREDOR"))						
                    cliente_old = Cdbl(rsCupos("CLIENTE"))
                    condicion = cstr(rsCupos("CONDICION"))
                    dsCorredor = Trim(rsCupos("DSCORREDOR"))
                    dsVendedor = Trim(rsCupos("DSVENDEDOR"))
                    dsCliente = Trim(rsCupos("DSCLIENTE"))
					while(controlCorredorVendedor(vendedor_old, corredor_old, cliente_old, condicion, rsCupos))
					    indexArr = g_dicFechas.Item(Right(rsCupos("FECHACUPO"),4))
                        arrayCupoRecibidos(1, indexArr) = CDbl(rsCupos("QTIngresados"))
                        arrayCupoAsignados(1, indexArr) = CDbl(rsCupos("QTAsignados"))
                        rsCupos.MoveNext()
                    wend					                    
					Call dibujarLineaCupo(1, arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,true,condicion)						 
                wend
                Call dibujarTotalesPorFecha("TOTAL CUPOS ESPECIALES",g_arrayTotalCuposRecibidosDia ,g_arrayTotalCuposAsignadosDia,true) %>
			</TBODY>
		</TABLE>	
<%	else %>
	<TABLE class='xls_border_left'  style="width:80%; font-size:12;">
		<TR >
			<TD align=center colspan=22><%= GF_TRADUCIR("<B>NO SE ENCONTRARON RESULTADOS</B>")%></TD>		
		</TR>
	</TABLE>
<%	end if

End Function


'--------------------------------------------------------------------------------------------
' Función:	controlCorredorVendedor
' Autor: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	15/01/13
' Objetivo:	
'			Realiza el corte de Control en caso de que no coincidan los parametros
' Parametros:
'			vendedor_old 	[Int] 	IdVendedor viejo
'			corredor_old 	[Int] 	IdCorredor viejo
'			vendedor_new 	[Int] 	IdVendedor actual
'			corredor_new 	[Int] 	IdCorredor actual
'Devuelve: 
'			True = si son los mismos registros  ,	false = si son distintos
'--------------------------------------------------------------------------------------------
Function controlCorredorVendedor(vendedor_old, corredor_old, cliente_old, condicion, pRs)
	Dim rtrn
	rtrn  = false
	if not pRs.Eof then
		if((Cdbl(pRs("CDVENDEDOR")) = vendedor_old)and(CDbl(pRs("CDCORREDOR")) = corredor_old)and(CDbl(pRs("CLIENTE")) = Cdbl(cliente_old)) and (Trim(condicion) = Trim(pRs("CONDICION"))))then rtrn = true
	end if
	controlCorredorVendedor = rtrn
End Function
'--------------------------------------------------------------------------------------------
' Función:	loadPeriodoFecha
' Autor: 	CNA - Ajaya Cesar Nahuel
' Fecha: 	16/01/13
' Objetivo:	
'			Carga al Vector el periodo de Fechas que se va a trabajar
' Parametros:
'			pFechaDesde 	[string] 	Fecha de Inicio
' Devuelve:
'			--
'--------------------------------------------------------------------------------------------
Function loadPeriodoFecha(pFechaDesde)
	Dim myFecha 
    Set g_dicFechas = Server.CreateObject("Scripting.Dictionary")
	for i = 0 to CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA
		myFecha = DateAdd("d", i, pFechaDesde)
	    myKey = GF_nDigits(Month(myFecha),2) & GF_nDigits(Day(myFecha),2)	
        Call g_dicFechas.Add(myKey, i)
	Next	
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA ***********************************
'*************************************************************************************
Dim gv_fecha, gv_cuposEspeciales, gv_cdProducto, fileName,gv_pto, gv_Dia,gv_Mes,gv_Anio,gv_fecha1
Dim fechaCierre, fechaDesde, fechaHasta ,cupEsp, myfecha,g_dicFechas, g_dicParticipantes, g_arrayTotalCuposAsignadosDia(), g_arrayTotalCuposRecibidosDia()
Dim media, gv_linea, g_idxParticipantes

media = GF_Parametros7("media", "", 6)
cupEsp = GF_Parametros7("chkCuposEspeciales", 0, 6)
gv_cuposEspeciales = false
if(cupEsp = 1) then gv_cuposEspeciales = true
gv_pto = GF_Parametros7("pto", "", 6)
g_strPuerto = gv_pto
gv_cdProducto = GF_Parametros7("cdProducto", 0, 6)
gv_fecha = GF_Parametros7("fecha", 0, 6)
myfecha = DateSerial(Left(gv_fecha,4),Mid(gv_fecha,5,2),Right(gv_fecha,2) ) 
'Para que el DateAdd tome correctamente la fecha se le debe dar el formato mm/dd/aaaa
fechaDesde = DateAdd("d", - (Weekday(myfecha,2) - 1),myfecha)
fechaHasta = DateAdd("d", 14, fechaDesde)
fileName = "Cupos_" & gv_pto & "_" & gv_fecha
if (media <> TIPO_AFIRMACION) then Call GF_createXLS(fileName)    
gv_linea = 0
%>
<html>
<head>
	<style type="text/css">
		.xls_border_left { 
			border-color:#000000; 
			border-style:solid; 
			border-width:thin;
		}
		.xls_align_center { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
			text-align: center;
		}
		.xls_align_right { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
			text-align: right;
		}
		.xls_precioUC_tabla
		{
			BACKGROUND-COLOR: #ffff80;
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;

		}
		.cls_cupos
		{
			COLOR: #FE2E2E;
		}        
        .lineacupo {
            background-color:#FFFFFF;
            cursor:pointer;
        }
        .lineacupo2 {
            background-color:#EEEEEE;
            cursor:pointer;
        }        
		</style>
        <script type="text/javascript">
            function bodyOnLoad() {
                <% if (media = TIPO_AFIRMACION) then %>
                setTimeout("submitMe()", 60000);
                <% end if %>
            }

            function submitMe() {
                document.getElementById('frmSel').submit();
            }
        </script>
	</head>
	<body onload="bodyOnLoad()">	
	<form name="frmSel" method="post" action="reporteCuposAsignadosPrint.asp">
        <input type="hidden" name="media" value ="<% =media %>">
        <input type="hidden" name="chkCuposEspeciales" value ="<% =cupEsp %>">
        <input type="hidden" name="pto" value ="<% =gv_pto %>">
        <input type="hidden" name="fecha" value ="<% =gv_fecha %>">
        <input type="hidden" name="cdProducto" value ="<% =gv_cdProducto %>">
	</form>
<%
Set g_dicParticipantes = Server.CreateObject("Scripting.Dictionary")
Call armarExelCupos(fechaDesde, fechaHasta, gv_cdProducto, media)
if(gv_cuposEspeciales)then Call armarExelCuposEspeciales(fechaDesde, fechaHasta, gv_cdProducto, media)
if (media <> TIPO_AFIRMACION) then Call closeXLS()

%> 
</body>
</html>
