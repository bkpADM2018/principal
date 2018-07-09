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
             " FROM   (SELECT case when cuitcliente <> "& CUIT_TOEPFER &" then 0 else cdcorredor end as CORREDOR, "&_
			 "                  case when cuitcliente <> "& CUIT_TOEPFER &" then 0 else cdvendedor end as VENDEDOR, "&_
			 "                  case when cuitcliente <> "& CUIT_TOEPFER &" then (select min(cdcliente) from clientes where rtrim(nucuit) = rtrim(cuitcliente)) else 0 end as CLIENTE, "&_
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


    strSQL = "  SELECT CASE WHEN CRR.CDCORREDOR IS NULL THEN 0 ELSE CRR.CDCORREDOR END AS CDCORREDOR, "&_    
             "         CASE WHEN CRR.DSCORREDOR IS NULL THEN '' ELSE CRR.DSCORREDOR END AS DSCORREDOR, "&_
             "         CASE WHEN V.CDVENDEDOR IS NULL THEN 0 ELSE V.CDVENDEDOR END AS CDVENDEDOR, "&_
             "         CASE WHEN V.DSVENDEDOR IS NULL THEN '' ELSE V.DSVENDEDOR END AS DSVENDEDOR, "&_
             "         CASE WHEN CLI.CDCLIENTE IS NULL THEN 0 ELSE CLI.CDCLIENTE END AS CDCLIENTE, "&_
             "         CASE WHEN CLI.DSCLIENTE IS NULL THEN '' ELSE CLI.DSCLIENTE END AS DSCLIENTE, "&_
             "         CAM.FECHA, "&_
             "         PRO.DSPRODUCTO, "&_
	         "         COUNT(CAM.CUPO) AS INGRESADOS, "&_
	         "         COUNT(CUP.CODIGOCUPO) AS ASIGNADOS "&_
             "   FROM    "&_
	         "       (SELECT case when RTRIM(CLI.NUCUIT) <> '"& CUIT_TOEPFER &"' then 0 else TT.CORREDOR end as CORREDOR, "&_
			 "               case when RTRIM(CLI.NUCUIT) <> '"& CUIT_TOEPFER &"' then 0 else TT.vendedor end as VENDEDOR, "&_
			 "               case when RTRIM(CLI.NUCUIT) <> '"& CUIT_TOEPFER &"' then CLI.CDCLIENTE ELSE 0 end as CLIENTE, "&_
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
            "    WHERE CUP.CODIGOCUPO IS NULL "&_
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
	
	strSQL =   "SELECT T.CORREDOR, "&_
               "      CASE WHEN CRR.DSCORREDOR IS NULL THEN '' ELSE CRR.DSCORREDOR END AS DSCORREDOR, "&_
               "       T.VENDEDOR, "&_
               "      CASE WHEN V.DSVENDEDOR IS NULL THEN '' ELSE V.DSVENDEDOR END AS DSVENDEDOR, "&_
               "       T.CLIENTE, "&_
               "       CASE WHEN CLI.DSCLIENTE IS NULL THEN '' ELSE CLI.DSCLIENTE END AS DSCLIENTE, "&_
               "       FECHA, "&_
               "       T.PRODUCTO, "&_
               "       PRO.DSPRODUCTO, "&_
               "       T.CONTRATO, "&_
               "       T.ORIGEN, "&_
               "       T.DESTINO "&_
             " FROM ( "&_
             "      SELECT CUP.CDCORREDOR CORREDOR, "&_
	         "		       CUP.CDVENDEDOR VENDEDOR, "&_
             "             (SELECT MIN(CDCLIENTE) AS CDCLIENTE FROM CLIENTES WHERE NUCUIT = CUP.CUITCLIENTE) CLIENTE, "&_
	         "		       CUP.FECHACUPO       FECHA, "&_
	         "	           CUP.CDPRODUCTO   PRODUCTO, "&_
             "             CUP.CONDICION    CONTRATO, "&_
	         "		       CUP.QTINGRESADOS  DESTINO, "&_
	         "             CUP.QTASIGNADOS ORIGEN "&_
	         "      FROM  CODIGOSCUPOESPECIALES AS CUP "&_	        
	         "      WHERE CUP.FECHACUPO >= " & fechaDesde & " AND CUP.FECHACUPO < " & fechaHasta & " AND CUP.CDPRODUCTO = "& pCdProducto &_	         
             "      ) T "&_	        
             " INNER JOIN PRODUCTOS PRO ON T.PRODUCTO = PRO.CDPRODUCTO "&_
             " LEFT JOIN CORREDORES CRR ON T.CORREDOR = CRR.CDCORREDOR "&_
             " LEFT JOIN VENDEDORES V ON T.VENDEDOR = V.CDVENDEDOR "&_
             " LEFT JOIN CLIENTES CLI ON CLI.CDCLIENTE = T.CLIENTE "&_
             " ORDER BY CRR.DSCORREDOR, V.DSVENDEDOR, FECHA, T.PRODUCTO,CONTRATO "
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
	
	strSQL = "			SELECT 	CUP.DTCUPO FECHA,"
	strSQL = strSQL & "		Sum(CUP.qtasignados)  QTASIGNADOS, "		
	strSQL = strSQL & "     Sum(CUP.qtingresados) QTINGRESADOS "	
	strSQL = strSQL & " FROM   asignacioncupos as CUP "
	strSQL = strSQL & "		RIGHT JOIN contratosespeciales CE ON CE.IDCUPO = CUP.IDCUPO "
	strSQL = strSQL & " WHERE   CUP.dtcupo >= " & fechaDesde & " AND CUP.dtcupo < " & fechaHasta
	strSQL = strSQL & "		AND CUP.cdproducto = " & pProducto 
	strSQL = strSQL & " GROUP  BY CUP.DTCUPO "
	
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
Function armarCabeceraCupos(pFechaDesde, pFechaHasta, pIsNotEspecial)
	Dim myFecha, myCuposEspeciales
	myCuposEspeciales = "No"
	if(gv_cuposEspeciales)then myCuposEspeciales = "Si"
	%>	
	<TABLE>
	<TBODY class='xls_border_left'  style="width:80%; " >
		<TR style="font-size:12;line-height:50%">
			<TD align=left colspan="17"><B><%=GF_TRADUCIR("ACTI AR")%></B></TD>
			<TD align=left colspan="4"><B><%= GF_TRADUCIR("Fecha: " & Left(GF_FN2DTE(session("MmtoSistema")),10))  %></B></TD>
		</TR>
		<TR style="font-size:12;line-height:50%">
			<TD align=left colspan="17"><B><%=GF_TRADUCIR("Pto: " & Ucase(gv_pto))%></B></TD>
			<TD align=left colspan="4"><B><%= GF_TRADUCIR("Hora: " & Right(GF_FN2DTE(session("MmtoSistema")),8))  %></B></TD>
		</TR>		
		<BR></BR>
		<TR style="font-size:14;line-height:120%">			
			<TD colspan=21 align=center ><B>
				<% if(pIsNotEspecial)then %>
					<%=GF_TRADUCIR("REPORTE DE CUPOS ASIGNADOS") %>
				<% else %>
					<%=GF_TRADUCIR("CUPOS ESPECIALES") %>
				<% end if %>	
			</B></TD>					
		</TR>
		<% if(pIsNotEspecial)then %>				
		<TR style="font-size:10;line-height:50%">
			<TD align=left><B><%=GF_TRADUCIR("FECHA DESDE: ") %></B></TD>
			<TD align=left colspan="2"><B><%= GF_nDigits(Day(pFechaDesde), 2) & "/" & GF_nDigits(Month(pFechaDesde), 2) & "/" & GF_nDigits(Year(pFechaDesde), 4) %></B></TD>					
		</TR>		
		<TR style="font-size:10;line-height:50%">		
			<TD align=left><B><%=GF_TRADUCIR("FECHA HASTA: ") %></B></TD>
			<TD align=left colspan="2"><B><%= GF_nDigits(Day(pFechaHasta) - 1, 2) & "/" & GF_nDigits(Month(pFechaHasta), 2) & "/" & GF_nDigits(Year(pFechaHasta), 4) %></B></TD>			
		</TR>		
		<TR style="font-size:10;line-height:50%">
			<TD align=left><B><%=GF_TRADUCIR("PRODUCTO: ")  %></B></TD>			
			<TD align=left colspan="2"><B><%=gv_cdProducto & " - " & getDsProducto(gv_cdProducto) %></B></TD>
		</TR>		
		<TR style="font-size:10;line-height:50%">			
			<TD align=left><B><%=GF_TRADUCIR("CUPOS ESPECIALES: ")  %></B></TD>			
			<TD align=left colspan="2"><B><%=myCuposEspeciales %></B></TD>
		</TR>				
		<% end if %>
	</TBODY>
	<br><br>
	<TBODY class='xls_border_left' style="width:80%; font-size:6;">
		<TR style='background-color:#BDBDBD;'>
			<%if(not pIsNotEspecial)then %>
				<TD align=left style="width:100px;font-size:10;"><%=GF_TRADUCIR("Contrato")%></TD>
			<%end if%>
            <TD align=left style="width:100px;font-size:10;"><%=GF_TRADUCIR("Corredor")%></TD>
			<TD align=left style="width:100px;font-size:10;"><%=GF_TRADUCIR("Vendedor")%></TD>			
            <TD align=left style="width:100px;font-size:10;"><%=GF_TRADUCIR("Destinatario")%></TD>
	<%	For z = 0 to CUPO_CANTIDAD_DIAS_SEMANA - 1
			myFecha = DateAdd("d", z, pFechaDesde) %>
			<TD align=center style="width:20px;font-size:10;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & "/" & GF_nDigits(Month(myFecha),2))%></TD>						
	<%	Next	%>				
		<TD align=left style="width:30px;font-size:10;" ><%=GF_TRADUCIR("Total Sem.1")%> </TD>
		<TD align=left style="width:30px;font-size:10;" ><%=GF_TRADUCIR("Total Ton.1")%>	</TD>	
		
	<%	For z = CUPO_CANTIDAD_DIAS_SEMANA  to (CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA) - 1
			myFecha = DateAdd("d", z, pFechaDesde) %>
			<TD align=center style="width:20px;font-size:10;" ><%=GF_TRADUCIR(GF_nDigits(Day(myFecha),2) & "/" & GF_nDigits(Month(myFecha),2))%></TD>						
	<%	Next	%>				
		<TD align=left style="width:30px;font-size:10;"><%=GF_TRADUCIR("Total Sem.2")%> </TD>
		<TD align=left style="width:30px;font-size:10;"><%=GF_TRADUCIR("Total Ton.2")%>	</TD>	
		</TR>
	</TBODY>
	</TABLE>
<%	
End Function
'--------------------------------------------------------------------------------------------
'Busco si el ultimo vendedor-corredor-cliente se sigue procesando para el nuevo registro del recordset,
' esto indica si le quedaron pendientes cupos de otros dias para procesar al momento del corte o no
Function seguirProcesandoCupoParaEmpresa(p_cdVendedorCCOld,p_cdVendedorSCOld,p_cdCorredorCCOld,p_cdCorredorSCOld,p_cdClienteCCOld,p_cdClienteSCOld,p_rs)
    Dim ret
    ret = false
    if not p_rs.Eof then
        'Como no se sabe donde se proceso la ultima vez que corto el ciclo se consulta por las variables viejas de los CON CUPO(CC) y los SIN CUPO(SC)
        'Response.Write "<BR>(("& p_rs("CDVENDEDOR") &" = "& p_cdVendedorCCOld &" or "&p_rs("CDVENDEDOR")&" = "&p_cdVendedorSCOld&")and("&p_rs("CDCORREDOR")&" = "&p_cdCorredorCCOld&" or "&p_rs("CDCORREDOR")&" = "&p_cdCorredorSCOld&")and("& p_rs("CDCLIENTE") &" = "&p_cdClienteCCOld&" or "& p_rs("CDCLIENTE") &" = "& p_cdClienteSCOld &"))<BR>"
        if ((Cdbl(p_rs("CDVENDEDOR")) = Cdbl(p_cdVendedorCCOld) or Cdbl(p_rs("CDVENDEDOR")) = Cdbl(p_cdVendedorSCOld))and(Cdbl(p_rs("CDCORREDOR")) = Cdbl(p_cdCorredorCCOld) or Cdbl(p_rs("CDCORREDOR")) = Cdbl(p_cdCorredorSCOld))and(Cdbl(p_rs("CDCLIENTE")) = Cdbl(p_cdClienteCCOld) or Cdbl(p_rs("CDCLIENTE")) = Cdbl(p_cdClienteSCOld))) then
            ret = true
        end if
    end if
    seguirProcesandoCupoParaEmpresa = ret
End Function
'--------------------------------------------------------------------------------------------
'Si es -1 entonces proceso el de SC (camiones sin cupo)
'Si es 0 entonces proceso los dos SC y CC
'Si es 1 entonces proceso el de CC (camiones con cupo)
Function compararClaveCupo(p_CdCorredorCC,p_CdCorredorSC,p_CdVendedorCC,p_CdVendedorSC,p_CdClienteCC,p_CdClienteSC) ',p_FechaCC,p_FechaSC,p_CdCorredorCCOld,p_CdCorredorSCOld,p_CdVendedorCCOld,p_CdVendedorSCOld,p_CdClienteCCOld,p_CdClienteSCOld,p_FechaCCOld,p_FechaSCOld
    dim rtrn
    rtrn = 1
    if cdbl(p_CdCorredorCC) = cdbl(p_CdCorredorSC) then
	    if cdbl(p_CdVendedorCC) = cdbl(p_CdVendedorSC) then
		    if cdbl(p_CdClienteCC) = cdbl(p_CdClienteSC) then
                rtrn = 0
		    else	
			    if cdbl(p_CdClienteCC) > cdbl(p_CdClienteSC) then rtrn = -1
		    end if						
	    else	
		    if cdbl(p_CdVendedorCC) > cdbl(p_CdVendedorSC) then rtrn = -1
	    end if						
    else
	    if cdbl(p_CdCorredorCC) > cdbl(p_CdCorredorSC) then rtrn = -1
    end if	
    compararClaveCupo = rtrn
end function
'--------------------------------------------------------------------------------------------
'Verifico que los registros sean iguales a los que se estan procesadno
Function corteControlEmpresasCupos(p_CdCorredorCC,p_CdCorredorSC,p_CdVendedorCC,p_CdVendedorSC,p_CdClienteCC,p_CdClienteSC,p_rsCuposCC,p_rsCuposSC,p_CdCorredorCCOld,p_CdCorredorSCOld,p_CdVendedorCCOld,p_CdVendedorSCOld,p_CdClienteCCOld,p_CdClienteSCOld)
    'Response.Write "-----------INGRESA IF: "&ret&"<BR>"                       
    Dim ret 
    ret = false      
    if not p_rsCuposCC.Eof and not p_rsCuposSC.Eof then
        'Response.Write "(("& p_CdCorredorCC &" = "& p_CdCorredorCCOld &" and "& p_CdVendedorCC &" = "& p_CdVendedorCCOld &")and("&p_CdClienteCC&" = "& p_CdClienteCCOld &"))<BR>"
        if ((cdbl(p_CdCorredorCC) = cdbl(p_CdCorredorCCOld))and(cdbl(p_CdVendedorCC) = cdbl(p_CdVendedorCCOld))and(cdbl(p_CdClienteCC) = cdbl(p_CdClienteCCOld))) then
            'Response.Write "(("& p_CdCorredorSC &" = "& p_CdCorredorSCOld &" and "& p_CdVendedorSC &" = "& p_CdVendedorSCOld &")and("&p_CdClienteSC&" = "& p_CdClienteSCOld &"))<BR>"
            if ((cdbl(p_CdCorredorSC) = cdbl(p_CdCorredorSCOld))and(cdbl(p_CdVendedorSC) = cdbl(p_CdVendedorSCOld))and(cdbl(p_CdClienteSC) = cdbl(p_CdClienteSCOld))) then ret = true
        end if
    end if
    'Response.Write "-----------DEVULEVE: "&ret&"<BR>"
    corteControlEmpresasCupos = ret
End function
'--------------------------------------------------------------------------------------------
'Permite llevar un corte de control cuado alguna empresa (Corredor,Vendedor,Cliente) cambia con respecto a la anterior, de esta manera permite dibujar la fila
Function corteControlCupos(p_rsCupos,p_cdCorredorOld,p_cdVendedorOld,p_cdClienteOld)
    Dim ret
    ret = false
    if (not p_rsCupos.Eof) then
        if ((cdbl(p_rsCupos("CDCORREDOR")) = cdbl(p_cdCorredorOld))and(cdbl(p_rsCupos("CDVENDEDOR")) = cdbl(p_cdVendedorOld))and(cdbl(p_rsCupos("CDCLIENTE")) = cdbl(p_cdClienteOld))) then ret = true
    end if
    corteControlCupos = ret
End function
'-----------------------------------------------------------------------------------------------------------------
Function dibujarTotalesPorSemana(pTotalCumplidos, pTotalAsignados)
	writeXLS("<TD align=right style='width:30px;font-size:10;' bgcolor='#F2F2F2'>" & cstr(pTotalCumplidos & "|" & pTotalAsignados) & "</TD>")
	writeXLS("<TD align=right style='width:30px;font-size:10;' bgcolor='#F2F2F2'>" & cstr(pTotalCumplidos * CUPO_PESO_CAMION & "|" & pTotalAsignados * CUPO_PESO_CAMION) & "</TD>")	
End Function
'--------------------------------------------------------------------------------------------
Function dibujarLineaCupo(ByRef p_arrayCupoAsignados,ByRef p_arrayCupoRecibidos,p_dsCorredor,p_dsVendedor,p_dsCliente,pIsEspecial,p_Contrato)
    Dim htmlTable,parcialRecibidosSemana,parcialAsignadosSemana,myCupoDia
    parcialAsignadosSemana = 0
    parcialRecibidosSemana = 0
    %> 
     <tr>
      <% if (pIsEspecial) then %>
         <td align=left bgcolor='#F2F2F2' style="width:90px; font-size:10;"><%=p_Contrato%></td>
      <% end if %>
        <td align=left bgcolor='#F2F2F2' style="width:90px; font-size:10;"><%=p_dsCorredor%></td>
        <td align=left bgcolor='#F2F2F2' style="width:90px; font-size:10;"><%=p_dsVendedor%></td>
        <td align=left bgcolor='#F2F2F2' style="width:90px; font-size:10;"><%=p_dsCliente%></td> 
    <%  for i = 0 to Ubound(p_arrayCupoAsignados) -1 
            myCupoDia = ""
            if ((not IsEmpty(p_arrayCupoAsignados(i)))or(not IsEmpty(p_arrayCupoRecibidos(i)))) then 
                myCupoDia = p_arrayCupoRecibidos(i) &"|"& p_arrayCupoAsignados(i)
                parcialAsignadosSemana = parcialAsignadosSemana + CDbl(p_arrayCupoAsignados(i))
				parcialRecibidosSemana = parcialRecibidosSemana + Cdbl(p_arrayCupoRecibidos(i))
                g_arrayTotalCuposAsignadosDia(i) = Cdbl(g_arrayTotalCuposAsignadosDia(i)) + Cdbl(p_arrayCupoAsignados(i))    
                g_arrayTotalCuposRecibidosDia(i) = Cdbl(g_arrayTotalCuposRecibidosDia(i)) + Cdbl(p_arrayCupoRecibidos(i))
                
            end if %>
            <TD align=right  style="width:20px; font-size:10;" <% if (Cdbl(p_arrayCupoRecibidos(i)) > Cdbl(p_arrayCupoAsignados(i))) then%> class="cls_cupos" <% end if %>><%=myCupoDia %></TD>
        <%  p_arrayCupoAsignados(i) = empty
            p_arrayCupoRecibidos(i) = empty 
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
	    <TD align=left style="font-size:10;" colspan='<% if (p_IsEspecial) then response.write "4" else response.write "3" end if %>' ><%="<B>"& p_Titulo &"</B>"%></TD>
    <%  for i = 0 to (CUPO_CANTIDAD_DIAS_SEMANA - 1)
	        if(p_ArrayTotalAsignados(i) = "")then p_ArrayTotalAsignados(i)= 0
            if(p_ArrayTotalRecibidos(i) = "")then p_ArrayTotalRecibidos(i)= 0 %>
			<TD align=right style="width:20px;font-size:10;"><%=Cstr(p_ArrayTotalRecibidos(i)) &"|"& Cstr(p_ArrayTotalAsignados(i))%></TD>
    <%		sumTotalRecibidos = sumTotalRecibidos + p_ArrayTotalRecibidos(i)
            sumTotalAsignados = sumTotalAsignados + p_ArrayTotalAsignados(i)
		next
        sumTotalRecibidosTon = sumTotalRecibidos * CUPO_PESO_CAMION  
        sumTotalAsignadosTon = sumTotalAsignados * CUPO_PESO_CAMION  %>
        <TD align=right style='width:30px;font-size:10;' ><%= cstr(sumTotalRecibidos) &"|"& cstr(sumTotalAsignados) %></TD>
		<TD align=right style='width:30px;font-size:10;' ><%= cstr(sumTotalRecibidosTon) &"|"& cstr(sumTotalAsignadosTon) %></TD>

    <%  
        'Luego se procesa la segunda semana
        sumTotalRecibidos = 0
        sumTotalAsignados= 0
        for i = CUPO_CANTIDAD_DIAS_SEMANA  to ((CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)- 1)
            if(p_ArrayTotalAsignados(i) = "")then p_ArrayTotalAsignados(i)= 0
            if(p_ArrayTotalRecibidos(i) = "")then p_ArrayTotalRecibidos(i)= 0 %>
            <TD align=right style="width:20px;font-size:10;"><%=Cstr(p_ArrayTotalRecibidos(i)) &"|"& Cstr(p_ArrayTotalAsignados(i))%></TD>
	<%		sumTotalRecibidos = sumTotalRecibidos + p_ArrayTotalRecibidos(i)
            sumTotalAsignados = sumTotalAsignados + p_ArrayTotalAsignados(i)
	    next
		sumTotalRecibidosTon = sumTotalRecibidos * CUPO_PESO_CAMION  
        sumTotalAsignadosTon = sumTotalAsignados * CUPO_PESO_CAMION %>

		<TD align=right style='width:30px;font-size:10;' ><%= cstr(sumTotalRecibidos) &"|"& cstr(sumTotalAsignados) %></TD>
		<TD align=right style='width:30px;font-size:10;' ><%= cstr(sumTotalRecibidosTon) &"|"& cstr(sumTotalAsignadosTon) %></TD>
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
Function procesarCupoArray(p_Fecha, p_CuposRecibidos, p_CuposAsignados, ByRef p_arrayCupoRecibidos, ByRef p_arrayCupoAsignados)
    Dim indexArr
    indexArr = g_dicFechas.Item(Right(p_Fecha,4))
    p_arrayCupoRecibidos(indexArr) = Cdbl(p_CuposRecibidos)
    p_arrayCupoAsignados(indexArr) = Cdbl(p_CuposAsignados)
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
Function armarExelCupos(pFechaDesde, pFechaHasta, pCdProducto)
	Dim rsCupos, arrayCupoRecibidos(),arrayCupoAsignados(), rsCuposCC, rsCuposSC
    Dim cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,dsCorredor,dsVendedor,dsCliente

    Call loadPeriodoFecha(pFechaDesde)

	Set rsCuposCC = armarSQLConCupos(pFechaDesde, pFechaHasta, pCdProducto) 
    Set rsCuposSC = armarSQLSinCupos(pFechaDesde, pFechaHasta, pCdProducto)
    
	Call armarCabeceraCupos(pFechaDesde, pFechaHasta, true)
	
    'creao el array que tendra los valores de cada fila
    Redim arrayCupoRecibidos(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    Redim arrayCupoAsignados(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    'creo los array para totalizar por planta o buenos aires el dia
    Redim g_arrayTotalCuposRecibidosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
	Redim g_arrayTotalCuposAsignadosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)   

    if not rsCuposCC.eof or not rsCuposSC.eof then %>
        <TABLE>
		<TBODY class='xls_border_left'  style="width:80%; font-size:6;" >      
    <%  'Si hay registros en los dos recordset los procesa juntos
        if not rsCuposCC.eof and not rsCuposSC.eof then
            cdVendedorCCOld = rsCuposCC("CDVENDEDOR")
            cdVendedorSCOld = rsCuposSC("CDVENDEDOR")
            cdCorredorCCOld = rsCuposCC("CDCORREDOR")
            cdCorredorSCOld = rsCuposSC("CDCORREDOR")
            cdClienteCCOld  = rsCuposCC("CDCLIENTE")
            cdClienteSCOld  = rsCuposSC("CDCLIENTE")
            while (not rsCuposCC.eof and not rsCuposSC.eof )

                    SELECT CASE compararClaveCupo(rsCuposCC("CDCORREDOR"),rsCuposSC("CDCORREDOR"),rsCuposCC("CDVENDEDOR"),rsCuposSC("CDVENDEDOR"),rsCuposCC("CDCLIENTE"),rsCuposSC("CDCLIENTE"))  'cdCorredorCCOld,cdCorredorSCOld,cdVendedorCCOld,cdVendedorSCOld,cdClienteCCOld,cdClienteSCOld,fechaCCOld,fechaSCOld
                        CASE -1
                            if (((Cdbl(cdVendedorSCOld) <> Cdbl(rsCuposSC("CDVENDEDOR")))or(Cdbl(cdCorredorSCOld) <> Cdbl(rsCuposSC("CDCORREDOR"))))or(Cdbl(cdClienteSCOld) <> Cdbl(rsCuposSC("CDCLIENTE")))) then Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
                            'Proceso los Sin Cupo 
                            cdVendedorCCOld = 0
                            cdVendedorSCOld = rsCuposSC("CDVENDEDOR")
                            cdCorredorCCOld = 0
                            cdCorredorSCOld = rsCuposSC("CDCORREDOR")
                            cdClienteCCOld  = 0
                            cdClienteSCOld  = rsCuposSC("CDCLIENTE")
                            dsCorredor = rsCuposSC("DSCORREDOR")
                            dsVendedor = rsCuposSC("DSVENDEDOR")
                            dsCliente = rsCuposSC("DSCLIENTE")
                            'Ahora proceso todas las fechas que tiene el recordset sin cupo para el Corredor-Vendedor-Empresa
                            while corteControlEmpresasCupos(0,rsCuposSC("CDCORREDOR"),0,rsCuposSC("CDVENDEDOR"),0,rsCuposSC("CDCLIENTE"),rsCuposCC,rsCuposSC,0,cdCorredorSCOld,0,cdVendedorSCOld,0,cdClienteSCOld)
                                Call procesarCupoArray(rsCuposSC("FECHA"),rsCuposSC("INGRESADOS"),rsCuposSC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                                rsCuposSC.MoveNext()
                            wend
                        CASE 0
                            if (((Cdbl(cdVendedorSCOld) <> Cdbl(rsCuposSC("CDVENDEDOR")))or(Cdbl(cdCorredorSCOld) <> Cdbl(rsCuposSC("CDCORREDOR"))))or(Cdbl(cdClienteSCOld) <> Cdbl(rsCuposSC("CDCLIENTE")))) then 
                                Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
                            else
                                if (((Cdbl(cdVendedorCCOld) <> Cdbl(rsCuposCC("CDVENDEDOR")))or(Cdbl(cdCorredorCCOld) <> Cdbl(rsCuposCC("CDCORREDOR"))))or(Cdbl(cdClienteCCOld) <> Cdbl(rsCuposCC("CDCLIENTE")))) then Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
                            end if
                            cdVendedorCCOld = rsCuposCC("CDVENDEDOR")
                            cdVendedorSCOld = rsCuposSC("CDVENDEDOR")
                            cdCorredorCCOld = rsCuposCC("CDCORREDOR")
                            cdCorredorSCOld = rsCuposSC("CDCORREDOR")
                            cdClienteCCOld  = rsCuposCC("CDCLIENTE")
                            cdClienteSCOld  = rsCuposSC("CDCLIENTE")
                            dsCorredor = rsCuposSC("DSCORREDOR")
                            dsVendedor = rsCuposSC("DSVENDEDOR")
                            dsCliente = rsCuposSC("DSCLIENTE")
                            'Ahora proceso todas las fechas que tiene el recordset sin cupo para el Corredor-Vendedor-Empresa
                            while corteControlEmpresasCupos(rsCuposCC("CDCORREDOR"),rsCuposSC("CDCORREDOR"),rsCuposCC("CDVENDEDOR"),rsCuposSC("CDVENDEDOR"),rsCuposCC("CDCLIENTE"),rsCuposSC("CDCLIENTE"),rsCuposCC,rsCuposSC,cdCorredorCCOld,cdCorredorSCOld,cdVendedorCCOld,cdVendedorSCOld,cdClienteCCOld,cdClienteSCOld)
                                'busco la fecha menor de ambas listas para empezar a calcular
                                if (Cdbl(rsCuposSC("FECHA")) < Cdbl(rsCuposCC("FECHA"))) then 
                                    Call procesarCupoArray(rsCuposSC("FECHA"),rsCuposSC("INGRESADOS"),rsCuposSC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                                    rsCuposSC.MoveNext()
                                else if(Cdbl(rsCuposSC("FECHA")) = Cdbl(rsCuposCC("FECHA"))) then
                                        Call procesarCupoArray(rsCuposSC("FECHA"),Cdbl(rsCuposSC("INGRESADOS")) + Cdbl(rsCuposCC("INGRESADOS")),Cdbl(rsCuposSC("ASIGNADOS")) + Cdbl(rsCuposCC("ASIGNADOS")), arrayCupoRecibidos, arrayCupoAsignados)
                                        rsCuposSC.MoveNext()
                                        rsCuposCC.MoveNext()
                                     else
                                        Call procesarCupoArray(rsCuposCC("FECHA"),rsCuposCC("INGRESADOS"),rsCuposCC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                                        rsCuposCC.MoveNext()
                                     end if
                                end if
                            wend
                        CASE 1      
                            'Proceso los que ya vienen con cupos
                            if (((Cdbl(cdVendedorCCOld) <> Cdbl(rsCuposCC("CDVENDEDOR")))or(Cdbl(cdCorredorCCOld) <> Cdbl(rsCuposCC("CDCORREDOR"))))or(Cdbl(cdClienteCCOld) <> Cdbl(rsCuposCC("CDCLIENTE")))) then 
                                Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
                            end if
                            'Proceso los Sin Cupo 
                            cdVendedorCCOld = rsCuposCC("CDVENDEDOR")
                            cdVendedorSCOld = 0
                            cdCorredorCCOld = rsCuposCC("CDCORREDOR")
                            cdCorredorSCOld = 0
                            cdClienteCCOld  = rsCuposCC("CDCLIENTE")
                            cdClienteSCOld  = 0
                            dsCorredor = rsCuposCC("DSCORREDOR")
                            dsVendedor = rsCuposCC("DSVENDEDOR")
                            dsCliente = rsCuposCC("DSCLIENTE")
                            'Ahora proceso todas las fechas que tiene el recordset sin cupo para el Corredor-Vendedor-Empresa
                            while corteControlEmpresasCupos(rsCuposCC("CDCORREDOR"),0,rsCuposCC("CDVENDEDOR"),0,rsCuposCC("CDCLIENTE"),0,rsCuposCC,rsCuposSC,cdCorredorCCOld,0,cdVendedorCCOld,0,cdClienteCCOld,0)
                                Call procesarCupoArray(rsCuposCC("FECHA"),rsCuposCC("INGRESADOS"),rsCuposCC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                                rsCuposCC.MoveNext()
                            wend
                    END SELECT
                wend 
        end if      
        'Dibujo la ultima linea totalizada
        Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
        'Verifico si quedaron resgistros Con Cupo sin procesar (esto puede pasar debido a que termine el recordset Sin Cupo antes y genera el corde del ciclo sin terminar de procesar el otro recordset)
        if (not rsCuposCC.eof) then
            if (not seguirProcesandoCupoParaEmpresa(cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,rsCuposCC)) then Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
            'Proceso los registros que faltan                
            while (not rsCuposCC.eof)
                cdVendedorCCOld = rsCuposCC("CDVENDEDOR")
                cdCorredorCCOld = rsCuposCC("CDCORREDOR")
                cdClienteCCOld  = rsCuposCC("CDCLIENTE")
                dsCorredor = rsCuposCC("DSCORREDOR")
                dsVendedor = rsCuposCC("DSVENDEDOR")
                dsCliente = rsCuposCC("DSCLIENTE")
                'Ahora proceso todas las fechas que tiene el recordset sin cupo para el Corredor-Vendedor-Empresa
                while corteControlCupos(rsCuposCC,cdCorredorCCOld,cdVendedorCCOld,cdClienteCCOld)
                    Call procesarCupoArray(rsCuposCC("FECHA"),rsCuposCC("INGRESADOS"),rsCuposCC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                    rsCuposCC.MoveNext()
                wend
                Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
            wend
        end if
        if (not rsCuposSC.eof) then
            if (not seguirProcesandoCupoParaEmpresa(cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,rsCuposSC)) then Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
            'Proceso los registros que faltan                
             while (not rsCuposSC.eof)
                cdVendedorCCOld = rsCuposSC("CDVENDEDOR")
                cdCorredorCCOld = rsCuposSC("CDCORREDOR")
                cdClienteCCOld  = rsCuposSC("CDCLIENTE")
                dsCorredor = rsCuposSC("DSCORREDOR")
                dsVendedor = rsCuposSC("DSVENDEDOR")
                dsCliente = rsCuposSC("DSCLIENTE")
                'Ahora proceso todas las fechas que tiene el recordset sin cupo para el Corredor-Vendedor-Empresa
                while corteControlCupos(rsCuposSC,cdCorredorCCOld,cdVendedorCCOld,cdClienteCCOld)
                    Call procesarCupoArray(rsCuposSC("FECHA"),rsCuposSC("INGRESADOS"),rsCuposSC("ASIGNADOS"), arrayCupoRecibidos, arrayCupoAsignados)
                    rsCuposSC.MoveNext()
                wend
                Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,false,"")
            wend
        end if
        Call dibujarTotalesPorFecha("TOTAL CUPOS", g_arrayTotalCuposRecibidosDia,g_arrayTotalCuposAsignadosDia,false)
        Call procesarTotalesCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto) 
        %>
    </TBODY></TABLE>
<%  else  %>
        <TABLE class='xls_border_left'  style="width:80%; font-size:10;">
    		<TR >
			    <TD align=center colspan="20" ><%= GF_TRADUCIR("<B>NO SE ENCONTRARON RESULTADOS</B>")%></TD>		
		    </TR>
	    </TABLE> 
<%  end if %>
<%      
    
End Function
'-----------------------------------------------------------------------------------------------
Function armarExelCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto)
	Dim rsCupos, arrayCupoRecibidos(),arrayCupoAsignados(), rsCuposCC, rsCuposSC
    Dim cdVendedorCCOld,cdVendedorSCOld,cdCorredorCCOld,cdCorredorSCOld,cdClienteCCOld,cdClienteSCOld,dsCorredor,dsVendedor,dsCliente
	
    Set rsCupos = armarSQLCuposEspeciales(pFechaDesde, pFechaHasta, pCdProducto)	
	
	Call armarCabeceraCupos(pFechaDesde, pFechaHasta, false)
    'creao el array que tendra los valores de cada fila
    Redim arrayCupoRecibidos(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    Redim arrayCupoAsignados(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
    'creo los array para totalizar por planta o buenos aires el dia
    Redim g_arrayTotalCuposRecibidosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)
	Redim g_arrayTotalCuposAsignadosDia(CUPO_CANTIDAD_DIAS_SEMANA * CUPO_SEGUNDA_SEMANA)   

	if(not rsCupos.EoF)then	%>
		<TABLE>
			<TBODY class='xls_border_left'  style="width:80%; font-size:6;" >				
            <% while(not rsCupos.EoF) 
                    vendedor_old = Cdbl(rsCupos("VENDEDOR"))
					corredor_old = Cdbl(rsCupos("CORREDOR"))						
                    cliente_old = Cdbl(rsCupos("CLIENTE"))
                    contrato = cstr(rsCupos("CONTRATO"))
                    dsCorredor = Trim(rsCupos("DSCORREDOR"))
                    dsVendedor = Trim(rsCupos("DSVENDEDOR"))
                    dsCliente = Trim(rsCupos("DSCLIENTE"))
					while(controlCorredorVendedor(vendedor_old, corredor_old, cliente_old, rsCupos))
					    indexArr = g_dicFechas.Item(Right(rsCupos("FECHA"),4))
                        arrayCupoRecibidos(indexArr) = CDbl(rsCupos("DESTINO"))
                        arrayCupoAsignados(indexArr) = CDbl(rsCupos("ORIGEN"))
                        rsCupos.MoveNext()
                    wend
                    Call dibujarLineaCupo(arrayCupoAsignados,arrayCupoRecibidos,dsCorredor,dsVendedor,dsCliente,true,contrato)
                wend
                Call dibujarTotalesPorFecha("TOTAL CUPOS ESPECIALES",g_arrayTotalCuposRecibidosDia ,g_arrayTotalCuposAsignadosDia,true) %>
			</TBODY>
		</TABLE>	
<%	else %>
	<TABLE class='xls_border_left'  style="width:80%; font-size:10;">
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
Function controlCorredorVendedor(vendedor_old, corredor_old, cliente_old, pRs)
	Dim rtrn
	rtrn  = false
	if not pRs.Eof then
		if((Cdbl(pRs("VENDEDOR")) = vendedor_old)and(CDbl(pRs("CORREDOR")) = corredor_old)and(CDbl(pRs("CLIENTE")) = Cdbl(cliente_old)))then rtrn = true
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
Dim fechaCierre, fechaDesde, fechaHasta ,cupEsp, myfecha,g_dicFechas, g_arrayTotalCuposAsignadosDia(), g_arrayTotalCuposRecibidosDia()

cupEsp = GF_Parametros7("chkCuposEspeciales", 0, 6)
gv_cuposEspeciales = false
if(cupEsp = 1)then gv_cuposEspeciales = true
gv_pto = GF_Parametros7("pto", "", 6)
g_strPuerto = gv_pto
gv_cdProducto = GF_Parametros7("cdProducto", 0, 6)
gv_fecha = GF_Parametros7("fecha", 0, 6)
myfecha = DateSerial(Left(gv_fecha,4),Mid(gv_fecha,5,2),Right(gv_fecha,2) ) 
'Para que el DateAdd tome correctamente la fecha se le debe dar el formato mm/dd/aaaa
fechaDesde = DateAdd("d", - (Weekday(myfecha,2) - 1),myfecha)
fechaHasta = DateAdd("d", 14, fechaDesde)
fileName = "Cupos_" & gv_pto & "_" & gv_fecha
%>
<html>
<head>
	<style type="text/css">
		.xls_border_left { 
			border-color:#666666; 
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
		</style>
	</head>
	<body>	
	
<%
Call armarExelCupos(fechaDesde, fechaHasta, gv_cdProducto)
Call armarExelCuposEspeciales(fechaDesde, fechaHasta, gv_cdProducto)
%> 
</body>
</html>
