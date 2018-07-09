<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosuser.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteBoletinesAnalisisCommon.asp"-->
<%
'--------------PARAMETROS PUERTOS-------------
'contiene el codigo de Parametro de los Ensayos cámara
CONST PARAM_CD_ENSAYO_CAMARA = "VLCORES"
'contiene el codigo de Parametro del Ensayo de Análisis de grado
CONST PARAM_CD_ENSAYO_GRADO = "CDENSAYOANGRA"
'-----------------------------------------------------------------------------------------------------------------
'Encargada de extraer los valores(items) concatenados de una Clave del Diccionario de Cabecera 
Function loadCamposCabecera(pIndex, ByRef pCalador, ByRef pTipo, ByRef pRebaja)
	Dim str,posLast,posFirst
	pCalador = ""
	pTipo    = ""
	pRebaja  = ""
	if (oDiccAnalisis.Exists(pIndex)) then	
		str = oDiccAnalisis.Item(cdbl(pIndex))	
		posFirst = InStr(1,str,";")
		posLast  = InStrRev(str,";")		
		pCalador = Left(str,posFirst-1)
		pTipo	 = Mid(str,posFirst+1,(posLast-posFirst)-1)
		pRebaja  = Right(str,Len(str)-posLast)		
   end if
End Function
'-----------------------------------------------------------------------------------------------------------------
'Encargada de extraer los valores(items) concatenados de una Clave del Diccionario del Detalle
Function loadCamposDetalle(key, ByRef pCdResultado, ByRef pDsEnsayo,ByRef pCdEnsayo)	
	if (oDiccDetalle.Exists(key)) then	
		str = oDiccDetalle.Item(cdbl(key))	
		posFirst = InStr(1,str,";")
		posLast  = InStrRev(str,";")		
		pCdEnsayo = Left(str,posFirst-1)
		pCdResultado = Mid(str,posFirst+1,(posLast-posFirst)-1)
		pDsEnsayo  = Right(str,Len(str)-posLast)		
   end if
End Function
'-----------------------------------------------------------------------------------------------------------------
'Verifica si el Boletin a mostrar tiene el grado seleccionado en el filtro
Function isDetalleByGrado(pauxGrado)
	Dim auxCdResultado,pCdEnsayo,count
	flagOK = false	
	count = 0	
	for each key in oDiccDetalle.Keys		
		Call loadCamposDetalle(key,auxCdResultado,pDsEnsayo,pCdEnsayo)		
		if (Trim(pCdEnsayo) = Trim(valGrado))then
			if (pauxGrado = GRADO_CAMARA_FE)then
				if((Cdbl(auxCdResultado) = 0)or(Cdbl(auxCdResultado) >= GRADO_CAMARA_FE))then flagOK = true
			else
				if (Cdbl(auxCdResultado) = Cdbl(pauxGrado))then flagOK = true			
			end if
		else
			count = count + 1	
		end if
	next
	if ((pauxGrado = GRADO_CAMARA_FE)and(oDiccDetalle.Count = count))then flagOK = true			
	isDetalleByGrado = flagOK
End function
'------------------------------------------------------------------------------------------------------------
Function generateWhereBoletines()			 
   Dim strWhere
   'La SQL de cabcera de Boletines del RMD solo trabaja con la ultima SQCalada del HcaladaCamiones (F.Contable-Id Camion-Max SQCalada),
   'de esta manera habrá solo un Usuario de Carga, esto me permite poder filtrar por Calador desde esta SQL sabiendo que solo puede 
   'variar el Tipo y Bonif/Rebaja pero no el Calador. 
   strWhere = " WHERE CLD.SQCALADA = (SELECT MAX(SQCALADA) "&_
			  "				 FROM HCALADADECAMIONES "&_
			  "				 WHERE IDCAMION = CLD.IDCAMION  AND DTCONTABLE = CLD.DTCONTABLE) "
  if (g_Producto <> 0)	then Call mkWhere(strWhere, "C.CDPRODUCTO", g_Producto, "=", 1)
  if (g_Coordinador <> "") then Call mkWhere(strWhere, "CD.CDEMPRESA", g_Coordinador, "=", 1)
  if (g_Coordinado <> "") then Call mkWhere(strWhere, "CD.CDCLIENTE", g_Coordinado, "=", 1)
  if (g_FechaDesde <> "") then Call mkWhere(strWhere, "C.DTCONTABLE", g_FechaDesde, ">=", 0)
  if (g_FechaHasta <> "")	then Call mkWhere(strWhere, "C.DTCONTABLE", g_FechaHasta, "<=", 0)
  if (g_sticker <> "") then Call mkWhere(strWhere, "CLD.NUBARRAS", g_sticker, "=", 3)
  if (g_Certificado <> "") then Call mkWhere(strWhere, "RC.nuCertificado", g_Certificado, "=", 3)
  if (g_Calador <> "") then Call mkWhere(strWhere, "Cld.cdUserName", g_Calador, "=", 3)
  generateWhereBoletines = strWhere
End Function
'------------------------------------------------------------------------------------------------------------
Function armarSQLCabecera()
	Dim strSQL
	strSQL = "SELECT TG.BRUTO, "&_
			 "		 TG.TARA, "&_
			 "		 CASE WHEN TG.MERMA IS NULL THEN 0 ELSE TG.MERMA END AS MERMA, "&_
			 "		 TG.DSPRODUCTO, "&_
			 "		 TG.CDPRODUCTO, "&_
			 "		 CASE WHEN Aceptacion IS NULL THEN '' ELSE Aceptacion END AS Aceptacion, "&_
			 "		 (YEAR(Fecha)*10000 + Month(Fecha)*100 + DAY(Fecha)) as Fecha , "&_
			 "		 CASE WHEN Barras IS NULL THEN '' ELSE Barras END AS Barras, "&_
			 "		 CASE WHEN Certificado IS NULL THEN '' ELSE Certificado END AS Certificado, "&_
			 "		 CASE WHEN Recibo IS NULL THEN 0 ELSE Recibo END AS Recibo, "&_
             "		 CASE WHEN TG.CDEMPRESA IS NULL THEN 0 ELSE TG.CDEMPRESA END AS CDEMPRESA, "&_
             "		 CASE WHEN TG.DSEMPRESA IS NULL THEN '' ELSE TG.DSEMPRESA END AS DSEMPRESA, "&_			 
			 "		 CASE WHEN TG.CDCLIENTE IS NULL THEN 0 ELSE TG.CDCLIENTE END AS CDCLIENTE, "&_
			 "		 CASE WHEN TG.DSCLIENTE IS NULL THEN '' ELSE TG.DSCLIENTE END AS DSCLIENTE, "&_
			 "		 CASE WHEN TG.CDVENDEDOR IS NULL THEN 0 ELSE TG.CDVENDEDOR END AS CDVENDEDOR, "&_
			 "		 CASE WHEN TG.DSVENDEDOR IS NULL THEN '' ELSE TG.DSVENDEDOR END AS DSVENDEDOR, "&_
			 "		 CASE WHEN TG.CDCORREDOR IS NULL THEN 0 ELSE TG.CDCORREDOR END AS CDCORREDOR, "&_
			 "		 CASE WHEN TG.DSCORREDOR IS NULL THEN '' ELSE TG.DSCORREDOR END AS DSCORREDOR, "&_
			 "		 CARTAPORTE, "&_
			 "		 CARTAPORTEDIG, "&_
			 "		 TG.Camion, "&_
			 "		 CASE WHEN RECEP IS NULL THEN 0 ELSE RECEP END AS RECEP, "&_
			 "		 CASE WHEN Turno IS NULL THEN 0 ELSE Turno END AS Turno, "&_
			 "		 CASE WHEN Calador IS NULL THEN '' ELSE Calador END AS Calador, "&_
			 "		 Tipo, "&_
			 "		 CASE WHEN Rebaja IS NULL THEN 0 ELSE Rebaja END AS Rebaja  "&_
			 " FROM( "&_
			 "		 SELECT (  "&_
			 "			SELECT pc.vlPesada "&_
			 "			FROM dbo.HPesadasCamion pc "&_
			 "			WHERE pc.dtContable = c.dtContable "&_
			 "				AND pc.Idcamion = c.Idcamion AND pc.cdPesada = 1 "&_
			 "				AND pc.sqpesada = (SELECT Max(sqPesada) "&_
			 "								   FROM dbo.HPesadasCamion "&_
			 "								   WHERE dtcontable = pc.DtContable "&_
			 "									   AND pc.Idcamion = Idcamion AND cdPesada = 1)) Bruto, "&_
			 "       (SELECT pc.vlPesada "&_
			 "		  FROM dbo.HPesadasCamion pc "&_
			 "		  WHERE pc.dtContable = c.dtContable "&_
			 "			  AND pc.Idcamion = c.Idcamion AND pc.cdPesada = 2 "&_
			 "			  AND pc.sqpesada = (SELECT Max(sqPesada) "&_
			 "								 FROM dbo.HPesadasCamion "&_
			 "								 WHERE dtcontable = pc.DtContable "&_
			 "									AND Idcamion = pc.Idcamion AND cdPesada = 2)) Tara, "&_
			 "		 (SELECT mc.vlmermakilos "&_
			 "		  FROM dbo.HMermascamiones mc"&_
			 "		  WHERE mc.idcamion = c.idcamion AND mc.dtcontable = c.dtcontable "&_
			 "			  AND mc.sqpesada = (SELECT max(sqpesada) "&_
			 "								 FROM dbo.HPesadasCamion "&_
			 "								 WHERE DtContable = mc.dtContable  "&_
			 "									AND idcamion = mc.idcamion AND cdPesada = 2 )) Merma, "&_
			 "		  p.cdProducto AS CDPRODUCTO,"&_	
			 "		  Dsproducto AS DSPRODUCTO, "&_
			 "		  AC.DSACEPTACION  AS Aceptacion,"&_
			 "		  C.DTCONTABLE AS Fecha,"&_
			 "		  CLD.NUBARRAS AS Barras,"&_       
			 "		  RC.nuCertificado AS Certificado,"&_
			 "		  Cd.NuRecibo AS Recibo,"&_
			 "		  e.cdEmpresa AS CDEMPRESA, "&_
			 "		  e.DsEmpresa AS DSEMPRESA, "&_
			 "		  cl.cdCliente AS CDCLIENTE, "&_
			 "		  cl.DsCliente AS DSCLIENTE, "&_
			 "		  V.cdVendedor AS CDVENDEDOR, "&_
			 "		  V.DsVendedor AS DSVENDEDOR,"&_
			 "		  crr.cdCorredor AS CDCORREDOR, "&_
			 "		  crr.DsCorredor AS DSCORREDOR ,"&_
			 "        cd.nuctaptedig AS CARTAPORTEDIG, "&_
			 "		  cd.nucartaporte AS CARTAPORTE, "&_
			 "		  C.IDCAMION AS CAMION, "&_
			 "		  C.NUAUTSALIDA AS RECEP,"&_
			 "		  C.SQTURNO AS TURNO,"&_
			 "		 Cld.cdUserName AS CALADOR, "&_
			 "       CASE WHEN RC.ICTIPO = 1 THEN 'R' ELSE 'B' END AS Tipo, "&_
			 "		 RC.Vlbonifrebaja  AS Rebaja"&_
			 " FROM dbo.HCamionesDescarga CD "&_
			 "    Join dbo.HCamiones C ON C.IdCamion = CD.IdCamion AND C.DTCONTABLE = CD.DTCONTABLE "&_
			 "    Join dbo.HCaladaDeCamiones Cld On C.IdCamion = Cld.IdCamion AND C.DTCONTABLE = CLD.DTCONTABLE and nubarras <> ''"&_
			 "    Left Join dbo.Productos p on c.cdproducto = p.cdproducto "&_
			 "    Left Join dbo.Empresas e on CD.cdempresa = e.cdempresa"&_
			 "    Left Join dbo.Clientes cl on cd.cdcliente = cl.cdcliente"&_
			 "    Left Join dbo.Corredores crr on cd.cdcorredor = crr.cdcorredor"&_
			 "    Left Join dbo.Vendedores v on cd.cdvendedor = v.cdvendedor"&_
			 "    Left Join dbo.AceptacionCalidad AC on CLD.CdAceptacion = AC.CdAceptacion"&_
			 "    Join (select cast(rc2.Nubarras as bigint) as NUBARRAS, rc2.dtcontable, rc2.nuCertificado,RC2.Vlbonifrebaja, rc2.ictipo "&_
             "          from dbo.ResultadosCamara rc2 where rc2.nubarras <> '' ) as RC  "&_
             "      on cast(substring(Cld.Nubarras,1,9) as bigint) = rc.Nubarras And cld.dtcontable = rc.dtcontable  "& generateWhereBoletines() &_
			 " )TG GROUP BY TG.BRUTO, "&_
			 "		 TG.TARA,TG.MERMA,TG.DSPRODUCTO,TG.CDPRODUCTO,Aceptacion,Fecha,Barras,Certificado,Recibo,TG.CDEMPRESA, "&_
			 "		 TG.DSEMPRESA,TG.CDCLIENTE,TG.DSCLIENTE,TG.CDVENDEDOR,TG.DSVENDEDOR,TG.CDCORREDOR,TG.DSCORREDOR,CARTAPORTE, "&_
			 "		 CARTAPORTEDIG, TG.Camion, RECEP, Turno, Calador, Tipo, Rebaja "&_
			 " Order by Barras, Fecha "			 
	Call GF_BD_Puertos(g_Pto, rs, "OPEN", strSQL)	
	Set armarSQLCabecera = rs
End Function
'-------------------------------------------------------------------------------------------------------------
Function armarSQLDetalleDB2(pBarra, pFecha)
	Dim strSQL, auxFecha
    auxFecha = Left(pFecha,4) &"-"& Mid(pFecha,5,2) & "-" & Right(pFecha,2)
	strSQL = " SELECT CASE WHEN RC.cdResultado IS NULL THEN 0 ELSE RC.cdResultado END as cdResultado , "&_
			 "		  CASE WHEN RC.dsResultado IS NULL THEN '' ELSE RC.dsResultado END as dsResultado , "&_
			 "		  CASE WHEN RC.cdEnsayo IS NULL THEN '' ELSE RC.cdEnsayo END as cdEnsayo , "&_
			 "		  CASE WHEN RC.icFueraStandard = 1 THEN 'Fuera de Standard' else '' END AS FStandar, "&_ 			 
			 "		  case when E.DSENSAYO is null then 'Código Desconocido' else E.DSENSAYO end AS Descripcion "&_
			 " FROM ResultadosCamara Rc "&_
			 "	   Left Join Ensayos E ON RC.CDENSAYO = E.CDENSAYO "&_
			 " WHERE cast(substring(Nubarras,1,9) as bigint) = cast(substring('" & pBarra & "',1,9) as bigint)" & _
			 "	 AND rc.DtContable = '"& auxFecha &"'"
    Call GF_BD_Puertos(g_Pto, rs, "OPEN", strSQL)
	Set armarSQLDetalleDB2 = rs
End Function
'----------------------------------------------------------------------------------------
Function imprimirDatosCabcera(pFecha,pCertificado,pSticker,pProducto,pCoordinador,pCoordinado,pCorredor,pVendedor,pCtaPte)
	imprimirDatosCabcera = arrTitulosCabecera(0) & VALUE_TOKEN & GF_FN2DTE(pFecha)
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(1) & VALUE_TOKEN & Trim(pCertificado)
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(2) & VALUE_TOKEN & Trim(pSticker)
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(3) & VALUE_TOKEN & Trim(pProducto)
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(4) & VALUE_TOKEN & Trim(pCoordinador) &" / "& Trim(pCoordinado)	
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(5) & VALUE_TOKEN & Trim(pCorredor) & " / "& Trim(pVendedor)	
	imprimirDatosCabcera = imprimirDatosCabcera & FIELD_TOKEN & arrTitulosCabecera(6) & VALUE_TOKEN & GF_EDIT_CTAPTE(pCtaPte)	
End Function
'-----------------------------------------------------------------------------------------
Function imprimirDatosTotal(pGrado,pNeto)
	imprimirDatosTotal = arrTitulosTotal(0) & VALUE_TOKEN & GF_EDIT_DECIMALS(pNeto,0) & " Kg."
	imprimirDatosTotal = imprimirDatosTotal & FIELD_TOKEN & arrTitulosTotal(1) & VALUE_TOKEN & getDsGrado(Cdbl(pGrado))
End function
'-----------------------------------------------------------------------------------------
Function imprimirDatosDetalle(pCdEnsayo,pDsEnsayo,pCdResultado,pCalador,pTipo,pRebaja,pRecep,pTurno,pNeto)
	imprimirDatosDetalle = arrTitulosDetalle(0) & VALUE_TOKEN & Trim(pRecep)
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(1) & VALUE_TOKEN & Trim(pTurno)
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(2) & VALUE_TOKEN & GF_EDIT_DECIMALS(pNeto,0) & " Kg."
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(3) & VALUE_TOKEN & Trim(pCdEnsayo) & " - " & Trim(pDsEnsayo)	
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(4) & VALUE_TOKEN & Trim(pCdResultado)
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(5) & VALUE_TOKEN & pCalador
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(6) & VALUE_TOKEN & pTipo
	imprimirDatosDetalle = imprimirDatosDetalle & FIELD_TOKEN & arrTitulosDetalle(7) & VALUE_TOKEN & pRebaja	
End Function
'*****************************************************************************************
'*	COMIENZO DE PAGINA
'*	ETAPA 1 - GENERACION DEL ARCHIVO DE TEXTO TEMPORAL
'*
'*	Se procesaran los datos entre las fecha de inicio y fin, trabajando de a un día por 
'*	vez. Cada día será considerado un segmento de la información que será almacenado en un 
'*	archivo individual hasta que se completen todos los segmentos y los mismos sean
'*	unificados.
'*****************************************************************************************
fechaDesdeD = GF_PARAMETROS7("fecContableDS", "", 6)
fechaDesdeM = GF_PARAMETROS7("fecContableMS", "", 6)
fechaDesdeA = GF_PARAMETROS7("fecContableAS", "", 6)
Call GF_STANDARIZAR_FECHA(fechaDesdeD, fechaDesdeM, fechaDesdeA)
fechaHastaD = fechaDesdeD
fechaHastaM = fechaDesdeM
fechaHastaA = fechaDesdeA

Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arch = fs.OpenTextFile(strPath, 8, true)
Set fs = nothing

g_FechaDesde = fechaDesdeA & "-" & fechaDesdeM & "-" & fechaDesdeD
g_FechaHasta = fechaHastaA & "-" & fechaHastaM & "-" & fechaHastaD
			 
flagCargarCabecera = true
index = 0
'Obtengo los valores de los Parametros necesarios para calcular el detalle 
valCore  = getValueParametro(PARAM_CD_ENSAYO_CAMARA,g_Pto)
valGrado = getValueParametro(PARAM_CD_ENSAYO_GRADO,g_Pto)
Set oDiccAnalisis  = createObject("Scripting.Dictionary")			
Set oDiccDetalle  = createObject("Scripting.Dictionary")
Set rsCab = armarSQLCabecera()
while not rsCab.eof
	if flagCargarCabecera then
		'GENERO LA CADENA CON LOS CAMPOS=VALOR DE LA CABECERA		
		myTara = 0
		myBruto = 0
		if not isNull(rsCab("BRUTO")) then myBruto = Cdbl(rsCab("BRUTO"))
		if not isNull(rsCab("TARA")) then myTara = Cdbl(rsCab("TARA"))
		netoAnterior = myBruto - myTara
		barraAnterior = rsCab("BARRAS")
		fechaAnterior = rsCab("FECHA")
		recepcionAnterior = rsCab("RECEP")
		turnoAnterior = rsCab("TURNO")
		productoAnterior = rsCab("CDPRODUCTO") & "-" & rsCab("DSPRODUCTO")
		certificadoAnterior = rsCab("CERTIFICADO")
		empresaAnterior = rsCab("CDEMPRESA") &"-"& rsCab("DSEMPRESA")
		clienteAnterior = rsCab("CDCLIENTE") &"-"& rsCab("DSCLIENTE")
		corredorAnterior = rsCab("CDCORREDOR") &"-"& rsCab("DSCORREDOR")
		vendedorAnterior = rsCab("CDVENDEDOR") &"-"& rsCab("DSVENDEDOR")
		ctaPteAnterior = rsCab("CARTAPORTE")
		if not isNull(rsCab("CARTAPORTEDIG")) then ctaPteAnterior = rsCab("CARTAPORTE") & rsCab("CARTAPORTEDIG")
	end if
	oDiccAnalisis.Add index, rsCab("CALADOR") &";"& rsCab("TIPO") &";"& rsCab("REBAJA")
	index = index + 1
	rsCab.MoveNext	
	flagCargarDetalle = false
	flagCargarCabecera =  false
	'Verifico si el proximo registro tiene el mismo Sticker, en ese caso no se dibuja la Cabecera y Detalle
	if not rsCab.Eof then
		if (barraAnterior <> rsCab("BARRAS")) then flagCargarDetalle = true
	else
		flagCargarDetalle = true
	End If
	if flagCargarDetalle then
		Set rsDet = armarSQLDetalleDB2(barraAnterior,fechaAnterior)
		if not rsDet.Eof then 
			i = 0
			while not rsDet.Eof
				'Voy guardando los registros del detalle en el diccionario 				
				auxResultado = CDbl(rsDet("CDRESULTADO")) / valCore
				oDiccDetalle.Add i , rsDet("CDENSAYO")&";"& auxResultado &";"& rsDet("DESCRIPCION")
				i = i + 1
				rsDet.MoveNext()
			wend
		end if
		flagVer = true
		if (auxGrado > 0 )then flagVer = isDetalleByGrado(auxGrado)
		if flagVer then
			stringBoletines = imprimirDatosCabcera(fechaAnterior,certificadoAnterior,barraAnterior,productoAnterior,empresaAnterior,clienteAnterior,corredorAnterior,vendedorAnterior,ctaPteAnterior) & SECTOR_TOKEN			
			valorGrado = 0
			for each key in oDiccDetalle.Keys
				Call loadCamposCabecera(key, auxCalador, auxTipo, auxRebaja)
				Call loadCamposDetalle(key,auxCdResultado,auxDsEnsayo,auxCdEnsayo)
				if (Trim(auxCdEnsayo) = Trim(valGrado)) then valorGrado = auxCdResultado
				stringBoletines = stringBoletines & imprimirDatosDetalle(auxCdEnsayo,auxDsEnsayo,auxCdResultado,auxCalador,auxTipo,auxRebaja,recepcionAnterior,turnoAnterior,netoAnterior) & DETAIL_TOKEN
			next
			stringBoletines = left(stringBoletines,len(stringBoletines)-3)
			stringBoletines = stringBoletines & SECTOR_TOKEN & imprimirDatosTotal(valorGrado,netoAnterior)
			arch.WriteLine(stringBoletines)			
		end if
		oDiccDetalle.RemoveAll
		oDiccAnalisis.RemoveAll
		flagCargarCabecera = true
		index = 0		
	end if	
wend
arch.close()
Set arch = Nothing
	 
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript">
	parent.generateSegment_callback();
</script>
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
