<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosMG.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosuser.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteVisteosCaladaCommon.asp"-->
<%
'-----------------------------------------------------------------------------------------
Function leerDatosCamiones(pidcamion, pto, pnuCartaPorte, pemision, phasta,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
	dim strSQL, rs, myWhere, diaHoy	, filtroProd, filtroEstado
	'Analizo los filtros.
	if (pnuCartaPorte <> "") 	then Call mkWhere(myWhere, "A.NUCARTAPORTE", pnuCartaPorte, "LIKE", 3)								  	
	if (pidcamion > 0)			then Call mkWhere(myWhere, "A.IDCAMION", pidcamion, "=", 0)
	if (pemision <> "") 		then Call mkWhere(myWhere, "A.DTCONTABLE", pemision, ">=", 0)
	if (phasta <> "") 			then Call mkWhere(myWhere, "A.DTCONTABLE", phasta, "<=", 0)	
	filtroProd = ""
	if (pcdProducto <> 0) 		then filtroProd = " and CDPRODUCTO =" & pcdProducto
	if (pcdCliente <> 0)	    then Call mkWhere(myWhere, "A.CDCLIENTE", pcdCliente, "=", 1)
	if (pcdCorredor <> 0) 		then Call mkWhere(myWhere, "A.CDCORREDOR", pcdCorredor, "=", 1)
	if (pcdVendedor <> 0) 		then Call mkWhere(myWhere, "A.CDVENDEDOR", pcdVendedor, "=", 1)
	if (pcdEntregador <> 0)		then Call mkWhere(myWhere, "A.CDENTREGADOR", pcdEntregador, "=", 1)
	filtroEstado = CAMIONES_ESTADO_EGRESADOOK & ", " & CAMIONES_ESTADO_PESADOTARA
	if (CInt(pcdEstado) <> 0)	then filtroEstado = pcdEstado
		
	'Armo la fecha del dia la SQL diaria.
    diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2) 
	'Consulta a la tabla diaria.
	' o - o - o - o - o - o
	'CONSIDERACIONES
	' Para esta SQL se tuvieron en cuenta las siguiente hipotesis
	'	1.- Si un cmion tiene multiples pesadas de bruto y/o tara, siempre se toma la que posee SQPESADA mas grande, tanto para bruto como para tara.
	'	2.- Si un camion tiene merma, esta estara registrada en el registro con SQPESADA mas grande.
	' o - o - o - o - o - o
			
	strSQL = "Select"
	strSQL = strSQL & "	A.NUCARTAPORTE,"
	strSQL = strSQL & " A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & " E.DSCLIENTE,"	
	strSQL = strSQL & "	A.CDCHAPACAMION,"
	strSQL = strSQL & " A.CDCHAPAACOPLADO,"
	strSQL = strSQL & " A.CDPRODUCTO,"
	strSQL = strSQL & " C.DSPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"	
	strSQL = strSQL & "	A.DTCONTABLE,"	
	strSQL = strSQL & " D.DSCORREDOR,A.CDCORREDOR," 
	strSQL = strSQL & "	F.DSVENDEDOR,A.CDVENDEDOR,"
	strSQL = strSQL & " G.DSENTREGADOR,A.CDENTREGADOR,"
	strSQL = strSQL & "	CASE WHEN A.BRUTO is Null THEN 0 else A.BRUTO END BRUTO,"	
	strSQL = strSQL & "	CASE WHEN A.TARA is Null THEN 0 else A.TARA END TARA,"		
	strSQL = strSQL & "	A.MERMA"
	strSQL = strSQL & " from"
	
	strSQL = strSQL & "((SELECT '" & diaHoy & "' AS DTCONTABLE," 
	strSQL = strSQL & "	A.NUCARTAPORTE,A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & "	B.CDCHAPACAMION,B.CDCHAPAACOPLADO,"
	strSQL = strSQL & " B.CDPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"	
	strSQL = strSQL & "	A.CDCORREDOR,"
	strSQL = strSQL & "	A.CDVENDEDOR,"
	strSQL = strSQL & " A.CDENTREGADOR,"
	strSQL = strSQL & " (Select top 1 VLPESADA PESO from PESADASCAMION PC where CDPESADA=1 and PC.IDCAMION=A.IDCAMION  order by PC.IDCAMION, SQPESADA DESC ) BRUTO," 
    strSQL = strSQL & " (Select top 1 VLPESADA PESO from PESADASCAMION PC where CDPESADA=2 and PC.IDCAMION=A.IDCAMION  order by PC.IDCAMION, SQPESADA DESC ) TARA,"    
    strSQL = strSQL & " (Select top 1 VLMERMAKILOS MERMA from MERMASCAMIONES MC where MC.IDCAMION=A.IDCAMION  order by MC.IDCAMION, MC.SQPESADA DESC ) MERMA"
	strSQL = strSQL & " FROM dbo.CAMIONESDESCARGA A"
	strSQL = strSQL & "	INNER JOIN dbo.CAMIONES B ON A.IDCAMION = B.IDCAMION " & filtroProd & " and B.CDESTADO in (" & filtroEstado 	
	strSQL = strSQL & "))"
	strSQL = strSQL & " UNION " 	
	strSQL = strSQL & " (SELECT (YEAR(A.DTCONTABLE)*10000 + Month(A.DTCONTABLE)*100 + DAY(A.DTCONTABLE)) DTCONTABLE,"
	strSQL = strSQL & "	A.NUCARTAPORTE,A.NUCTAPTEDIG,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & "	B.CDCHAPACAMION,B.CDCHAPAACOPLADO,"
	strSQL = strSQL & "	B.CDPRODUCTO,"
	strSQL = strSQL & "	A.IDCAMION,"
	strSQL = strSQL & "	A.CDCORREDOR,"
	strSQL = strSQL & "	A.CDVENDEDOR,"
	strSQL = strSQL & " A.CDENTREGADOR,"	
	strSQL = strSQL & " (Select top 1 VLPESADA PESO from HPESADASCAMION HPC where CDPESADA=1 and HPC.IDCAMION=A.IDCAMION and HPC.DTCONTABLE=A.DTCONTABLE order by HPC.DTCONTABLE, HPC.IDCAMION, SQPESADA DESC ) BRUTO,"
	strSQL = strSQL & " (Select top 1 VLPESADA PESO from HPESADASCAMION HPC where CDPESADA=2 and HPC.IDCAMION=A.IDCAMION and HPC.DTCONTABLE=A.DTCONTABLE order by HPC.DTCONTABLE, HPC.IDCAMION, SQPESADA DESC ) TARA,"
	strSQL = strSQL & " (Select top 1 VLMERMAKILOS MERMA from HMERMASCAMIONES HMC where HMC.DTCONTABLE=A.DTCONTABLE and HMC.IDCAMION=A.IDCAMION  order by HMC.DTCONTABLE, HMC.IDCAMION, HMC.SQPESADA DESC ) MERMA"
	strSQL = strSQL & " FROM dbo.HCAMIONESDESCARGA A"
	strSQL = strSQL & "	INNER JOIN dbo.HCAMIONES B ON A.IDCAMION = B.IDCAMION AND A.DTCONTABLE =B.DTCONTABLE " & filtroProd & " and B.CDESTADO in (" & filtroEstado 
	strSQL = strSQL & "))) A"		
	strSQL = strSQL & "	INNER JOIN dbo.PRODUCTOS C ON A.CDPRODUCTO = C.CDPRODUCTO"			
	strSQL = strSQL & " INNER JOIN dbo.CORREDORES D ON A.CDCORREDOR = D.CDCORREDOR"
	strSQL = strSQL & "	INNER JOIN dbo.CLIENTES E ON A.CDCLIENTE = E.CDCLIENTE"
	strSQL = strSQL & "	INNER JOIN dbo.VENDEDORES F ON A.CDVENDEDOR = F.CDVENDEDOR"
	strSQL = strSQL & "	LEFT JOIN dbo.ENTREGADORES G ON A.CDENTREGADOR= G.CDENTREGADOR"
	
	strSQL = strSQL & myWhere & " ORDER BY DTCONTABLE, IDCAMION"
	
	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	Set leerDatosCamiones = rs
end function
'-----------------------------------------------------------------------------------------
Function leerVisteosCamionDetalle(psqcalada, pidcamion,pdtcontable,ppto) 		
	dim rs,myWhere,strSQL,diaHoy
	diaHoy = GF_nDigits(Year(Now()), 4) & GF_nDigits(Month(Now()), 2)  & GF_nDigits(Day(Now()), 2) 
	Call mkWhere(myWhere, "IDCAMION", pidcamion, "=", 0)		
	Call mkWhere(myWhere, "SQCALADA", psqcalada, "=", 1)
	Call mkWhere(myWhere, "DTCONTABLE", pdtcontable, "=", 0)		
	strSQL = "Select * from "
	strSQL = strSQL & " ((select '" & diaHoy & "' AS DTCONTABLE ,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro,idcamion,sqcalada "
	strSQL = strSQL & " FROM dbo.audrubrosvisteocamiones A) "
	strSQL = strSQL & " UNION "
	strSQL = strSQL & " (select (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro,idcamion,sqcalada "
	strSQL = strSQL & " FROM dbo.haudrubrosvisteocamiones A)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY cdrubro ASC"
    
	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
Set leerVisteosCamionDetalle = rs
end function
'-----------------------------------------------------------------------------------------
Function leerVisteosCamionCabecera(v_idcamion,pto,fecha)
	dim rs,myWhere,strSQL	
	if(v_idcamion > 0)then Call mkWhere(myWhere, "idcamion", v_idcamion, "=", 0)	
	if (fecha <> "") then Call mkWhere(myWhere, "DTCONTABLE", fecha, "=", 0)		 
	'Armo la fecha del dia la SQL diaria.
    diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	  
    strSQL = "Select Tabla.DTCONTABLE,Tabla.sqauditoria,Tabla.sqcalada,Tabla.cdusername,Tabla.cdterminal,Tabla.idcamion, "
	strSQL = strSQL & " CAST(Tabla.FECHACALADA as BIGINT)*1000000 + right('000000' + cast(Tabla.HORACALADA AS varchar(6)), 6) DTCALADA "
    strSQL = strSQL & " from ((SELECT '" & diaHoy & "' AS DTCONTABLE, A.sqauditoria, A.SQCALADA ,B.CDUSERNAME, B.CDTERMINAL,A.IDCAMION, "
    strSQL = strSQL & "         ((Year(A.dtcalada) * 10000) + (Month(A.dtcalada) * 100) + Day(A.dtcalada)) FECHACALADA, "
    strSQL = strSQL & "         ((DATEPART(HOUR, A.dtcalada) * 10000) + (DATEPART(MINUTE, A.dtcalada) * 100) + DATEPART(SECOND, A.dtcalada)) HORACALADA "
	strSQL = strSQL & " FROM  dbo.audcaladadecamiones A INNER JOIN dbo.auditoriacamiones B "
	strSQL = strSQL & " ON A.idcamion = B.idcamion "
	strSQL = strSQL & " AND A.sqauditoria = B.sqauditoria"
	strSQL = strSQL & " AND B.cdtransaccion = " & VISTEO_CALADA & ")"
	strSQL = strSQL & " UNION "
	strSQL = strSQL & "(SELECT (YEAR(A.DTCONTABLE)*10000 + Month(A.DTCONTABLE)*100 + DAY(A.DTCONTABLE)) DTCONTABLE, A.sqauditoria, A.SQCALADA ,B.CDUSERNAME, B.CDTERMINAL,A.IDCAMION, "
    strSQL = strSQL & "     ((Year(A.dtcalada) * 10000) + (Month(A.dtcalada) * 100) + Day(A.dtcalada)) FECHACALADA, "
    strSQL = strSQL & "     ((DATEPART(HOUR, A.dtcalada) * 10000) + (DATEPART(MINUTE, A.dtcalada) * 100) + DATEPART(SECOND, A.dtcalada)) HORACALADA "
	strSQL = strSQL & " FROM  dbo.haudcaladadecamiones A INNER JOIN dbo.hauditoriacamiones B "
	strSQL = strSQL & " ON A.idcamion = B.idcamion"
	strSQL = strSQL & " AND A.sqauditoria = B.sqauditoria "
	strSQL = strSQL & " AND B.cdtransaccion = " & VISTEO_CALADA
	strSQL = strSQL & " AND A.dtcontable = B.dtcontable)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY DTCONTABLE, IDCAMION, SQCALADA"		
	    
	Call GF_BD_Puertos(pto, rs, "OPEN",strSQL)
	set leerVisteosCamionCabecera =	rs
End function
'----------------------------------------------------------------------------------------
Function leerMuestrasCamionDetalle(psqcalada, pidcamion,pdtcontable,ppto) 		
	dim rs,myWhere,strSQL,diaHoy
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	Call mkWhere(myWhere, "IDCAMION", pidcamion, "=", 0)		
	Call mkWhere(myWhere, "SQCALADA", psqcalada, "=", 1)
	Call mkWhere(myWhere, "DTCONTABLE", pdtcontable, "=", 0)		
	strSQL = "Select * from "
	strSQL = strSQL & " ((select '" & diaHoy & "' AS DTCONTABLE , IDCAMION, SQCALADA,	SQMUESTRA,	VLHUMEDAD,	VLTEMPERATURA,	VLPESO"	
	strSQL = strSQL & " FROM dbo.MuestrasHumedCamiones A) "
	strSQL = strSQL & " UNION "
	strSQL = strSQL & " (select (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE, IDCAMION, SQCALADA,	SQMUESTRA,	VLHUMEDAD,	VLTEMPERATURA,	VLPESO"
	strSQL = strSQL & " FROM dbo.HMuestrasHumedCamiones A)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY SQMUESTRA"
    
	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
Set leerMuestrasCamionDetalle = rs
end function
'----------------------------------------------------------------------------------------
Function imprimirDatosCamion(pIdCamion,pCodProducto,pDsProducto,pNPorte,pChapa,pAcoplado,pCodCli, pDsCli,pdtcont,pCdCorredor,pDsCorredor,pCdEntregador,pDsEntregador,pCdVendedor,pDsVendedor, pKilosNetos, pMerma)	
	auxPRO = Trim(pCodProducto)&" - "&Trim(pDsProducto)	
	auxCli = Trim(pCodCli)&" - "&Trim(pDsCli)		
	auxCorredor = Trim(pCdCorredor)&" - "&Trim(pDsCorredor)
	auxVendedor = Trim(pCdVendedor)&" - "&Trim(pDsVendedor)	
	auxEntregador = Trim(pCdEntregador)&" - "&Trim(pDsEntregador)
	
	imprimirDatosCamion = arrTitulosCamion(0) & VALUE_TOKEN & GF_FN2DTE(pdtcont)
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(1) & VALUE_TOKEN & Trim(pIdCamion)
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(2) & VALUE_TOKEN & GF_EDIT_CTAPTE(pNPorte)
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(3) & VALUE_TOKEN & auxPRO
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(4) & VALUE_TOKEN & Trim(pChapa)
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(5) & VALUE_TOKEN & Trim(pAcoplado)
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(6) & VALUE_TOKEN & auxCli
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(7) & VALUE_TOKEN & auxCorredor
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(8) & VALUE_TOKEN & auxVendedor
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(9) & VALUE_TOKEN & auxEntregador
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(10) & VALUE_TOKEN & pKilosNetos
	imprimirDatosCamion = imprimirDatosCamion & FIELD_TOKEN & arrTitulosCamion(11) & VALUE_TOKEN & pMerma
End Function
'-----------------------------------------------------------------------------------------
Function imprimirVisteosCamionCabecera(pdtcontable,psqcalada,pcdusername,pcdterminal,pto)
	Dim auxValeRelacionado, dsPartSector, proveedor

	imprimirVisteosCamionCabecera = arrTitulosVisteo(0) & VALUE_TOKEN & GF_FN2DTE(pdtcontable)	
	imprimirVisteosCamionCabecera = imprimirVisteosCamionCabecera & FIELD_TOKEN & arrTitulosVisteo(1) & VALUE_TOKEN & getNombreApellidoCalada(Trim(pcdusername))
	imprimirVisteosCamionCabecera = imprimirVisteosCamionCabecera & FIELD_TOKEN & arrTitulosVisteo(2) & VALUE_TOKEN & Trim(pcdterminal)
		
End function
'-----------------------------------------------------------------------------------------
Function imprimirVisteosCamionDetalle(pcdrubro,pvl)	
	
	imprimirVisteosCamionDetalle = Trim(pcdrubro)
	imprimirVisteosCamionDetalle = imprimirVisteosCamionDetalle & VALUE_TOKEN & Trim(pvl)	
	
End Function
'-----------------------------------------------------------------------------------------
Function imprimirMuestrasCamionDetalle(cdMuestra, vlHumedad, vlTemperatura, vlPeso)
	imprimirMuestrasCamionDetalle = Trim(cdMuestra)
	imprimirMuestrasCamionDetalle = imprimirMuestrasCamionDetalle & VALUE_TOKEN & Trim(vlHumedad)	
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'*****************************************************************************************
'/****************************************************************************************
'* ETAPA 1 - GENERACION DEL ARCHIVO DE TEXTO TEMPORAL
'*
'*	Se procesaran los datos entre las fecha de inicio y fin, trabajando de a un día por 
'*	vez. Cada día será considerado un segmento de la información que será almacenado en un 
'*	archivo individual hasta que se completen todos los segmentos y los mismos sean
'*	unificados.
'*****************************************************************************************
fecContableD = GF_PARAMETROS7("fecContableDS", "", 6)
fecContableM = GF_PARAMETROS7("fecContableMS", "", 6)
fecContableA = GF_PARAMETROS7("fecContableAS", "", 6)
Call GF_STANDARIZAR_FECHA(fecContableD, fecContableM, fecContableA)

fecContableDH = fecContableD
fecContableMH = fecContableM
fecContableAH = fecContableA

'Si existe la borro
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arch = fs.OpenTextFile(strPath, 8, true)
Set fs = nothing

fechaD = fecContableA & fecContableM & fecContableD
fechaH = fecContableAH & fecContableMH & fecContableDH
Set rsDatos = leerDatosCamiones(g_idCamion, g_Puerto, g_cPorte, fechaD, fechaH, g_Producto,g_Vendedor,g_Corredor,g_Cliente,g_Entregador, g_Estado)
while not rsDatos.eof	
	'Comienzo el armado de los datos del camion.	
	stringCamion = imprimirDatosCamion(rsDatos("IDCAMION"),rsDatos("CDPRODUCTO"), rsDatos("DSPRODUCTO"), rsDatos("NUCARTAPORTE") & rsDatos("NUCTAPTEDIG"),rsDatos("CDCHAPACAMION"),rsDatos("CDCHAPAACOPLADO"),rsDatos("CDCLIENTE"), rsDatos("DSCLIENTE"),rsDatos("DTCONTABLE"), rsDatos("CDCORREDOR"),rsDatos("DSCORREDOR"),rsDatos("CDENTREGADOR"),rsDatos("DSENTREGADOR"),rsDatos("CDVENDEDOR"),rsDatos("DSVENDEDOR"), CLng(rsDatos("BRUTO"))-CLng(rsDatos("TARA")), rsDatos("Merma"))
	'LE DOY EL FORMATO AAAA-MM-DD PARA PODER TRABAJAR CON LA BASE DE DATOS	
	'myFormatFechaVieja = split(rsDatos("DTCONTABLE"),"/")	
	'myFormatFechaNueva = myFormatFechaVieja(2) & "-" & GF_nDigits(myFormatFechaVieja(0), 2) & "-" & GF_nDigits(myFormatFechaVieja(1), 2)
	Set rsCab = leerVisteosCamionCabecera(rsDatos("IDCAMION"), g_Puerto, rsDatos("DTCONTABLE")) 
	while (not rsCab.eof)		
		if (g_MaxSQCalada < CInt(rsCab("SQCALADA"))) then g_MaxSQCalada = CInt(rsCab("SQCALADA"))		
		'myFormatFechaVieja = split(rsCab("DTCALADA"),"/")
		'myfecha = myFormatFechaVieja(2) & "-" & GF_nDigits(myFormatFechaVieja(0), 2) & "-" & GF_nDigits(myFormatFechaVieja(1), 2)
		stringCamion = stringCamion & FIELD_TOKEN & imprimirVisteosCamionCabecera(rsCab("DTCALADA"),rsCab("SQCALADA"),rsCab("CDUSERNAME"),rsCab("CDTERMINAL"),g_Puerto)
		'Leo el detale de los rubros
		'myFormatFechaVieja = split(rsCab("DTCONTABLE"),"/")		
		'myfecha = myFormatFechaVieja(2) & "-" & GF_nDigits(myFormatFechaVieja(0), 2) & "-" & GF_nDigits(myFormatFechaVieja(1), 2)
		Set rsDet = leerVisteosCamionDetalle(rsCab("SQCALADA"), rsDatos("IDCAMION"),rsCab("DTCONTABLE"), g_Puerto) 		
		'Se agregan los rubros del camión.
		while (not rsDet.eof)		    
			cdRubro = Trim(rsDET("cdrubro"))
			if (not dicRubros.Exists(cdRubro)) then 
				txtRubro = Trim(getDsRubro(cdRubro))
				dicRubros.Add cdRubro, txtRubro
			end if						
			stringCamion = stringCamion & FIELD_TOKEN & imprimirVisteosCamionDetalle(cdRubro, rsDet("vlbonrebaja"))
			cdRubro = cdrubro & "_M"
			if (not dicRubros.Exists(cdRubro)) then 
				txtRubro = "MERMA_" & txtRubro
				dicRubros.Add cdRubro, txtRubro
			end if
			stringCamion = stringCamion & FIELD_TOKEN & imprimirVisteosCamionDetalle(cdRubro, rsDet("vlmerma"))
			rsDet.MoveNext()
		wend
		'Se agregan los valores de las muestras.		
		Set rsMuestras = leerMuestrasCamionDetalle(rsCab("SQCALADA"), rsDatos("IDCAMION"),rsCab("DTCONTABLE"), g_Puerto) 
		while (not rsMuestras.eof)	
			cdMuestra = PREFIX_MUESTRA & rsMuestras("SQMUESTRA")
			if (not dicRubros.Exists(cdMuestra)) then dicRubros.Add cdMuestra, "MUESTRA_HUMEDAD_" & rsMuestras("SQMUESTRA")
			stringCamion = stringCamion & FIELD_TOKEN & imprimirMuestrasCamionDetalle(cdMuestra, rsMuestras("VLHUMEDAD"), rsMuestras("VLTEMPERATURA"), rsMuestras("VLPESO"))
			rsMuestras.MoveNext()
		wend
		rsCab.MoveNext()
	wend
	'Imprimo la linea generada al archivo.	
	arch.WriteLine(stringCamion)
	rsDatos.MoveNext()
wend
arch.close()
Set arch = Nothing

Call saveDatosAdministrativos(dicRubros, strPathAdm)
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript">
	parent.generateSegment_callback();
</script>
</HEAD>
<BODY>
</BODY>
</HTML>
