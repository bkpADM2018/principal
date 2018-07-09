<%
Dim g_Puerto, g_fechaDesde, g_fechaHasta, g_Producto, g_Vendedor, g_Destinatario, g_Entregador, g_Coordinado, g_cPorte, g_idCamion
Dim myWhere, strSQL, separaFecha1,separaFecha2, fileCode
Dim cdProducto,cdVendedor,dsVendedor,cdDestinatario,dsDestinatario,cdCoordinado,dsCoordinado,cdEntregador,dsEntregador
Dim RPT_Division, RPT_Month, RPT_Year, RPT_Filtro, RPT_accion, strPathAdm, rsMuestras, cdMuestra
Dim rs, conn, filename, ultimaLinea, stringCamion, arch, fs, strPath, arrAlignCamion, arrAlignVisteo, rsRecarga
	
'Se reciben los parametros.
g_Puerto = GF_PARAMETROS7("pto", "", 6)

g_Producto = GF_PARAMETROS7("cdProducto", 0, 6)	
g_Vendedor = GF_PARAMETROS7("cdVendedor", 0, 6)
g_Destinatario = GF_PARAMETROS7("cdDestinatario", 0, 6)
g_Coordinado = GF_PARAMETROS7("cdCoordinado", 0, 6)
	
fecContableD = GF_PARAMETROS7("fecContableD", "", 6)
fecContableM = GF_PARAMETROS7("fecContableM", "", 6)
fecContableA = GF_PARAMETROS7("fecContableA", "", 6)
Call GF_STANDARIZAR_FECHA(fecContableD, fecContableM, fecContableA)

fecContableDH = GF_PARAMETROS7("fecContableDH", "", 6)
fecContableMH = GF_PARAMETROS7("fecContableMH", "", 6)
fecContableAH = GF_PARAMETROS7("fecContableAH", "", 6)
Call GF_STANDARIZAR_FECHA(fecContableDH, fecContableMH, fecContableAH)

g_fechaDesde = fecContableA & "-" & fecContableM & "-" & fecContableD
g_fechaHasta = fecContableAH & "-" & fecContableMH & "-" & fecContableDH
	
g_strPuerto = g_Puerto
dsProducto = getDsProducto(g_Producto)
 


strSQL = "		SELECT  C.CdProducto,				 "
strSQL = strSQL & "		P.DSProducto,				 "
strSQL = strSQL & "		C.NUAUTSALIDA Remito,		 "
strSQL = strSQL & "		(YEAR(C.DTCONTABLE)*10000 + Month(C.DTCONTABLE)*100 + DAY(C.DTCONTABLE))  Fecha, "
strSQL = strSQL & "		C.SQTURNO Turno,			 "
strSQL = strSQL & "		C.IDCAMION IdCamion,		 "
strSQL = strSQL & "		CD.NUCARTAPORTE CP,			 "
'strSQL = strSQL & "		E.DSEMPRESA Coordinador,	 "
strSQL = strSQL & "		CL.DSCLIENTE Coordinado,	 "
strSQL = strSQL & "		CMP.DsComprador Destinatario,"
strSQL = strSQL & "		VEN.DSVENDEDOR Vendedor,     "
strSQL = strSQL & "		C.CdChapaCamion Chapa,		 "
strSQL = strSQL & "     (							 "
strSQL = strSQL & "       SELECT pc.vlPesada		 "
strSQL = strSQL & "		  FROM dbo.HPesadasCamion pc    "
strSQL = strSQL & "		  WHERE pc.dtContable = c.dtContable "
strSQL = strSQL & "				AND pc.Idcamion = c.Idcamion "
strSQL = strSQL & "			    AND pc.cdPesada = 1  "
strSQL = strSQL & "			    AND pc.sqpesada = (SELECT MAX(sqPesada)			    " 
strSQL = strSQL & "								   FROM dbo.HPesadasCamion     "
strSQL = strSQL & "								   WHERE dtcontable = pc.DtContable "
strSQL = strSQL & "										AND pc.Idcamion = Idcamion  "
strSQL = strSQL & "										AND cdPesada = 1 )			"
strSQL = strSQL & "		 ) AS Bruto,		 "
strSQL = strSQL & "		(					 "
strSQL = strSQL & "       SELECT pc.vlPesada " 
strSQL = strSQL & "		  FROM dbo.HPesadasCamion pc	  "
strSQL = strSQL & "		  WHERE  pc.dtContable = c.dtContable "
strSQL = strSQL & "				AND pc.Idcamion = c.Idcamion  "
strSQL = strSQL & "				AND pc.cdPesada = 2			  "
strSQL = strSQL & "				AND pc.sqpesada = (SELECT MAX(sqPesada)				"
strSQL = strSQL & "								   FROM dbo.HPesadasCamion     "
strSQL = strSQL & "								   WHERE dtcontable = pc.DtContable "
strSQL = strSQL & "										AND Idcamion = pc.Idcamion  "
strSQL = strSQL & "										AND cdPesada = 2)           "
strSQL = strSQL & "		 ) AS Tara,		"
strSQL = strSQL & "		0 AS Merma, "
strSQL = strSQL & "     1 AS Imprimir	"
strSQL = strSQL & " FROM HCAMIONES C	"
strSQL = strSQL & "		JOIN HCAMIONESCARGA CD ON C.DTCONTABLE= CD.DTCONTABLE and C.IDCAMION = CD.IDCAMION "
strSQL = strSQL & "		LEFT JOIN Productos p on c.cdproducto = p.cdproducto	"
'strSQL = strSQL & "		LEFT JOIN Empresas e on CD.cdempresa = e.cdempresa		"
strSQL = strSQL & "		LEFT JOIN Clientes cl on cd.cdcliente = cl.cdcliente	"
strSQL = strSQL & "		LEFT JOIN Compradores cmp on cd.cddestinatario = cmp.cdcomprador "
strSQL = strSQL & "		LEFT JOIN Vendedores VEN on cd.CDVENDEDOR = VEN.CDVENDEDOR	"
strSQL = strSQL & " WHERE C.CDESTADO NOT IN ( 12, 7, 14 ) "


If (g_Coordinado > 0) Then	  strSQL = strSQL & " AND CL.CDCLIENTE = " & g_Coordinado
If (g_Destinatario > 0) Then  strSQL = strSQL & " AND CMP.CDCOMPRADOR = " & g_Destinatario
If (g_Vendedor > 0) Then	  strSQL = strSQL & " AND  VEN.CDVENDEDOR = " & g_Vendedor
If (g_Producto > 0) Then	  strSQL = strSQL & " AND C.CDPRODUCTO = " & g_Producto
If (Len(g_fechaDesde) > 8) Then  strSQL = strSQL & " AND  C.DtContable >= '" & g_fechaDesde & "'"
If (Len(g_fechaHasta) > 8) Then  strSQL = strSQL & " AND  C.DtContable <= '" & g_fechaHasta & "'"

strSQL = strSQL & " ORDER BY  P.DSPRODUCTO, C.NUAUTSALIDA, C.DTCONTABLE"
'Response.Write "<BR></BR>" & strSQL & "<BR></BR>"

Call GF_BD_Puertos(g_Puerto, rsRecarga, "OPEN", strSQL)		  



%>
