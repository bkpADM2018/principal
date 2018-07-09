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
Function leerDatosVagones(pCdOperativo, pCdVagon, pto, pnuCartaPorte, pemision, phasta,pcdProducto,pcdVendedor,pcdCorredor,pcdCliente,pcdEntregador, pcdEstado)
	dim strSQL, rs, myWhere, diaHoy	, filtroProd, filtroEstado
	'Analizo los filtros.
	if (pnuCartaPorte <> "") 	then 
	    Call mkWhere(myWhere, "A.NUCARTAPORTESERIE", Left(pnuCartaPorte, 4), "LIKE", 3)
	    Call mkWhere(myWhere, "A.NUCARTAPORTE", Right(pnuCartaPorte, 8), "LIKE", 3)
	end if
	if (pCdOperativo > 0)			then Call mkWhere(myWhere, "A.CDOPERATIVO", pCdOperativo, "=", 0)
	if (pCdVagon > 0)			then Call mkWhere(myWhere, "A.CDVAGON", pCdVagon, "=", 0)
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
	strSQL = strSQL & "	A.NUCARTAPORTESERIE + A.NUCARTAPORTE NUCARTAPORTE,"
	strSQL = strSQL & " A.CDCLIENTE,"
	strSQL = strSQL & " E.DSCLIENTE,"	
	strSQL = strSQL & " A.CDPRODUCTO,"
	strSQL = strSQL & " C.DSPRODUCTO,"
	strSQL = strSQL & "	A.CDOPERATIVO,"	
	strSQL = strSQL & "	A.CDVAGON,"	
	strSQL = strSQL & "	A.DTCONTABLE,"	
	strSQL = strSQL & " D.DSCORREDOR,A.CDCORREDOR," 
	strSQL = strSQL & "	F.DSVENDEDOR,A.CDVENDEDOR,"
	strSQL = strSQL & " G.DSENTREGADOR,A.CDENTREGADOR,"
	strSQL = strSQL & "	CASE WHEN A.BRUTO is Null THEN 0 else A.BRUTO END BRUTO,"	
	strSQL = strSQL & "	CASE WHEN A.TARA is Null THEN 0 else A.TARA END TARA,"		
	strSQL = strSQL & "	A.MERMA"
	strSQL = strSQL & " from"
	
	strSQL = strSQL & "((SELECT '" & diaHoy & "' AS DTCONTABLE," 
	strSQL = strSQL & "	A.NUCARTAPORTE,"
	strSQL = strSQL & "	A.NUCARTAPORTESERIE,"
	strSQL = strSQL & " B.CDCLIENTE,"
	strSQL = strSQL & " A.CDPRODUCTO,"
	strSQL = strSQL & "	A.CDOPERATIVO,"	
	strSQL = strSQL & "	A.CDVAGON,"	
	strSQL = strSQL & "	B.CDCORREDOR,"
	strSQL = strSQL & "	B.CDVENDEDOR,"
	strSQL = strSQL & " B.CDENTREGADOR,"
	strSQL = strSQL & " (Select TOP 1 VLPESADA PESO from PESADASVAGON PC where CDPESADA=1 and PC.CDOPERATIVO=A.CDOPERATIVO and PC.CDVAGON=A.CDVAGON  order by PC.CDOPERATIVO, PC.CDVAGON, SQPESADA DESC ) BRUTO," 
    strSQL = strSQL & " (Select TOP 1 VLPESADA PESO from PESADASVAGON PC where CDPESADA=2 and PC.CDOPERATIVO=A.CDOPERATIVO and PC.CDVAGON=A.CDVAGON  order by PC.CDOPERATIVO, PC.CDVAGON, SQPESADA DESC ) TARA,"    
    strSQL = strSQL & " (Select TOP 1 VLMERMAKILOS MERMA from MERMASVAGONES MC where MC.CDOPERATIVO=A.CDOPERATIVO and MC.CDVAGON=A.CDVAGON order by MC.CDOPERATIVO, MC.CDVAGON, MC.SQPESADA DESC ) MERMA"
	strSQL = strSQL & " FROM VAGONES A "
	strSQL = strSQL & " INNER JOIN OPERATIVOS B on B.CDOPERATIVO=A.CDOPERATIVO"
	strSQL = strSQL & " WHERE A.CDESTADO in (" & filtroEstado & ")" &  filtroProd
	strSQL = strSQL & ")"
	strSQL = strSQL & " UNION " 	
	strSQL = strSQL & " (SELECT (YEAR(A.DTCONTABLE)*10000 + Month(A.DTCONTABLE)*100 + DAY(A.DTCONTABLE)) DTCONTABLE,"
	strSQL = strSQL & "	A.NUCARTAPORTE,A.NUCARTAPORTESERIE,"
	strSQL = strSQL & " B.CDCLIENTE,"
	strSQL = strSQL & "	A.CDPRODUCTO,"
	strSQL = strSQL & "	A.CDOPERATIVO,"	
	strSQL = strSQL & "	A.CDVAGON,"	
	strSQL = strSQL & "	B.CDCORREDOR,"
	strSQL = strSQL & "	B.CDVENDEDOR,"
	strSQL = strSQL & " B.CDENTREGADOR,"	
	strSQL = strSQL & " (Select TOP 1 VLPESADA PESO from HPESADASVAGON HPC where CDPESADA=1 and HPC.CDOPERATIVO=A.CDOPERATIVO and HPC.CDVAGON=A.CDVAGON and HPC.DTCONTABLE=A.DTCONTABLE order by HPC.DTCONTABLE, HPC.CDOPERATIVO, HPC.CDVAGON, SQPESADA DESC ) BRUTO,"
	strSQL = strSQL & " (Select TOP 1 VLPESADA PESO from HPESADASVAGON HPC where CDPESADA=2 and HPC.CDOPERATIVO=A.CDOPERATIVO and HPC.CDVAGON=A.CDVAGON and HPC.DTCONTABLE=A.DTCONTABLE order by HPC.DTCONTABLE, HPC.CDOPERATIVO, HPC.CDVAGON, SQPESADA DESC ) TARA,"
	strSQL = strSQL & " (Select TOP 1 VLMERMAKILOS MERMA from HMERMASVAGONES HMC where HMC.DTCONTABLE=A.DTCONTABLE and HMC.CDOPERATIVO=A.CDOPERATIVO and HMC.CDVAGON=A.CDVAGON  order by HMC.DTCONTABLE, HMC.CDOPERATIVO, HMC.CDVAGON, HMC.SQPESADA DESC ) MERMA"
	strSQL = strSQL & " FROM HVAGONES A "
	strSQL = strSQL & " INNER JOIN HOPERATIVOS B on B.DTCONTABLE=A.DTCONTABLE and B.CDOPERATIVO=A.CDOPERATIVO"
	strSQL = strSQL & " WHERE A.CDESTADO in (" & filtroEstado & ")" & filtroProd
	strSQL = strSQL & ")) A"		
	strSQL = strSQL & "	INNER JOIN dbo.PRODUCTOS C ON A.CDPRODUCTO = C.CDPRODUCTO"			
	strSQL = strSQL & " INNER JOIN dbo.CORREDORES D ON A.CDCORREDOR = D.CDCORREDOR"
	strSQL = strSQL & "	INNER JOIN dbo.CLIENTES E ON A.CDCLIENTE = E.CDCLIENTE"
	strSQL = strSQL & "	INNER JOIN dbo.VENDEDORES F ON A.CDVENDEDOR = F.CDVENDEDOR"
	strSQL = strSQL & "	LEFT JOIN dbo.ENTREGADORES G ON A.CDENTREGADOR= G.CDENTREGADOR"
	
	strSQL = strSQL & myWhere & " ORDER BY DTCONTABLE, CDOPERATIVO, CDVAGON"
	
	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	Set leerDatosVagones = rs
end function
'-----------------------------------------------------------------------------------------
Function leerVisteosVagonDetalle(psqcalada, pcdOperativo, pcdVagon,pdtcontable,ppto) 		
	dim rs,myWhere,strSQL,diaHoy
	diaHoy = GF_nDigits(Year(Now()), 4) & GF_nDigits(Month(Now()), 2)  & GF_nDigits(Day(Now()), 2) 
	Call mkWhere(myWhere, "CDOPERATIVO", pcdOperativo, "=", 0)	
	Call mkWhere(myWhere, "CDVAGON", pcdVagon, "=", 0)	
	Call mkWhere(myWhere, "SQCALADA", psqcalada, "=", 1)
	Call mkWhere(myWhere, "DTCONTABLE", pdtcontable, "=", 0)		
	strSQL = "Select * from "
	strSQL = strSQL & " ((select '" & diaHoy & "' AS DTCONTABLE ,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro,cdOperativo, cdVagon,sqcalada "
	strSQL = strSQL & " FROM dbo.rubrosvisteovagones) "
	strSQL = strSQL & " UNION "
	strSQL = strSQL & " (select (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE,cdrubro, vlbonrebaja, vlmerma, vlpesorubro, pcpesorubro, cdOperativo, cdVagon,sqcalada "
	strSQL = strSQL & " FROM dbo.hrubrosvisteovagones)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY cdrubro ASC"
	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
Set leerVisteosVagonDetalle = rs
end function
'-----------------------------------------------------------------------------------------
Function leerVisteosVagonCabecera(p_cdOperativo, p_cdVagon,pto,fecha)
	dim rs,myWhere,strSQL	
	
	if(p_cdOperativo > 0)then Call mkWhere(myWhere, "CDOPERATIVO", p_cdOperativo, "=", 0)	
	if(p_cdVagon > 0)then Call mkWhere(myWhere, "CDVAGON", p_cdVagon, "=", 0)	
	if (fecha <> "") then Call mkWhere(myWhere, "DTCONTABLE", fecha, "=", 0)		 
	'Armo la fecha del dia la SQL diaria.
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	strSQL = "Select * from "
	strSQL = strSQL & "((SELECT '" & diaHoy & "' AS DTCONTABLE, SQCALADA ,CDUSERNAME,CDOPERATIVO, CDVAGON "
	strSQL = strSQL & " FROM  CALADADEVAGONES)"
	strSQL = strSQL & " UNION "
	strSQL = strSQL & "(SELECT (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE, SQCALADA ,CDUSERNAME,CDOPERATIVO, CDVAGON "
	strSQL = strSQL & " FROM  HCALADADEVAGONES)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY DTCONTABLE, CDOPERATIVO, CDVAGON, SQCALADA"		
	
	Call GF_BD_Puertos(pto, rs, "OPEN",strSQL)
	set leerVisteosVagonCabecera =	rs
End function
'----------------------------------------------------------------------------------------
Function leerMuestrasVagonDetalle(psqcalada, pcdOperativo, pcdVagon, pdtcontable,ppto) 		
	dim rs,myWhere,strSQL,diaHoy
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()), 2) & GF_nDigits(Day(Now()), 2)
	Call mkWhere(myWhere, "CDOPERATIVO", pcdOperativo, "=", 0)
	Call mkWhere(myWhere, "CDVAGON", pcdVagon, "=", 0)
	Call mkWhere(myWhere, "SQCALADA", psqcalada, "=", 1)
	Call mkWhere(myWhere, "DTCONTABLE", pdtcontable, "=", 0)		
	strSQL = "Select * from "
	strSQL = strSQL & " ((select '" & diaHoy & "' AS DTCONTABLE , CDOPERATIVO, CDVAGON, SQCALADA,	SQMUESTRA,	VLHUMEDAD,	VLTEMPERATURA,	VLPESO"	
	strSQL = strSQL & " FROM dbo.MuestrasHumedVagones) "
	strSQL = strSQL & " UNION "
	strSQL = strSQL & " (select (YEAR(DTCONTABLE)*10000 + Month(DTCONTABLE)*100 + DAY(DTCONTABLE)) DTCONTABLE, CDOPERATIVO, CDVAGON, SQCALADA,	SQMUESTRA,	VLHUMEDAD,	VLTEMPERATURA,	VLPESO"
	strSQL = strSQL & " FROM dbo.HMuestrasHumedVagones)) as Tabla "
	strSQL = strSQL & myWhere	
	strSQL = strSQL & " ORDER BY SQMUESTRA"
	Call GF_BD_Puertos(ppto, rs, "OPEN", strSQL)
Set leerMuestrasVagonDetalle = rs
end function
'----------------------------------------------------------------------------------------
Function imprimirDatosVagon(pCdOperativo, pcdVagon,pCodProducto,pDsProducto,pNPorte,pCodCli, pDsCli,pdtcont,pCdCorredor,pDsCorredor,pCdEntregador,pDsEntregador,pCdVendedor,pDsVendedor, pKilosNetos, pMerma)	
	auxPRO = Trim(pCodProducto)&" - "&Trim(pDsProducto)	
	auxCli = Trim(pCodCli)&" - "&Trim(pDsCli)		
	auxCorredor = Trim(pCdCorredor)&" - "&Trim(pDsCorredor)
	auxVendedor = Trim(pCdVendedor)&" - "&Trim(pDsVendedor)	
	auxEntregador = Trim(pCdEntregador)&" - "&Trim(pDsEntregador)
	
	imprimirDatosVagon = arrTitulosVagon(0) & VALUE_TOKEN & GF_FN2DTE(pdtcont)
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(1) & VALUE_TOKEN & Trim(pCdOperativo)
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(2) & VALUE_TOKEN & Trim(pcdVagon)
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(3) & VALUE_TOKEN & GF_EDIT_CTAPTE(pNPorte)
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(4) & VALUE_TOKEN & auxPRO
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(5) & VALUE_TOKEN & auxCli
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(6) & VALUE_TOKEN & auxCorredor
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(7) & VALUE_TOKEN & auxVendedor
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(8) & VALUE_TOKEN & auxEntregador
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(9) & VALUE_TOKEN & pKilosNetos
	imprimirDatosVagon = imprimirDatosVagon & FIELD_TOKEN & arrTitulosVagon(10) & VALUE_TOKEN & pMerma
End Function
'-----------------------------------------------------------------------------------------
Function imprimirVisteosVagonCabecera(pdtcontable,psqcalada,pcdusername,pcdterminal,pto)
	Dim auxValeRelacionado, dsPartSector, proveedor

	imprimirVisteosVagonCabecera = arrTitulosVisteo(0) & VALUE_TOKEN & GF_FN2DTE(pdtcontable)	
	imprimirVisteosVagonCabecera = imprimirVisteosVagonCabecera & FIELD_TOKEN & arrTitulosVisteo(1) & VALUE_TOKEN & getNombreApellidoCalada(Trim(pcdusername))
	imprimirVisteosVagonCabecera = imprimirVisteosVagonCabecera & FIELD_TOKEN & arrTitulosVisteo(2) & VALUE_TOKEN & Trim(pcdterminal)
		
End function
'-----------------------------------------------------------------------------------------
Function imprimirVisteosVagonDetalle(pcdrubro,pvl)	
	
	imprimirVisteosVagonDetalle = Trim(pcdrubro)
	imprimirVisteosVagonDetalle = imprimirVisteosVagonDetalle & VALUE_TOKEN & Trim(pvl)	
	
End Function
'-----------------------------------------------------------------------------------------
Function imprimirMuestrasVagonDetalle(cdMuestra, vlHumedad, vlTemperatura, vlPeso)
	imprimirMuestrasVagonDetalle = Trim(cdMuestra)
	imprimirMuestrasVagonDetalle = imprimirMuestrasVagonDetalle & VALUE_TOKEN & Trim(vlHumedad)	
End Function
'*****************************************************************************************
'	COMIENZO DE PAGINA
'*****************************************************************************************
'/****************************************************************************************
'* ETAPA 1 - GENERACION DEL ARCHIVO DE TEXTO TEMPORAL - VAGONES
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

Set rsDatos = leerDatosVagones(g_cdOperativo, g_cdVagon, g_Puerto, g_cPorte, fechaD, fechaH, g_Producto,g_Vendedor,g_Corredor,g_Cliente,g_Entregador, g_Estado)
while not rsDatos.eof	
	'Comienzo el armado de los datos del camion.	
	stringVagon = imprimirDatosVagon(rsDatos("CDOPERATIVO"), rsDatos("CDVAGON"),rsDatos("CDPRODUCTO"), rsDatos("DSPRODUCTO"), rsDatos("NUCARTAPORTE"),rsDatos("CDCLIENTE"), rsDatos("DSCLIENTE"),rsDatos("DTCONTABLE"), rsDatos("CDCORREDOR"),rsDatos("DSCORREDOR"),rsDatos("CDENTREGADOR"),rsDatos("DSENTREGADOR"),rsDatos("CDVENDEDOR"),rsDatos("DSVENDEDOR"), CLng(rsDatos("BRUTO"))-CLng(rsDatos("TARA")), rsDatos("Merma"))
	Set rsCab = leerVisteosVagonCabecera(rsDatos("CDOPERATIVO"), rsDatos("CDVAGON"), g_Puerto, rsDatos("DTCONTABLE")) 
	while (not rsCab.eof)		
		if (g_MaxSQCalada < CInt(rsCab("SQCALADA"))) then g_MaxSQCalada = CInt(rsCab("SQCALADA"))
		stringVagon = stringVagon & FIELD_TOKEN & imprimirVisteosVagonCabecera(rsCab("DTCONTABLE"),rsCab("SQCALADA"),rsCab("CDUSERNAME"),"",g_Puerto)
		Set rsDet = leerVisteosVagonDetalle(rsCab("SQCALADA"), rsDatos("CDOPERATIVO"), rsDatos("CDVAGON"), rsCab("DTCONTABLE"), g_Puerto) 
		'Se agregan los rubros del camión.
		while (not rsDet.eof)					
			cdRubro = Trim(rsDET("cdrubro"))
			if (not dicRubros.Exists(cdRubro)) then 
				txtRubro = Trim(getDsRubro(cdRubro))
				dicRubros.Add cdRubro, txtRubro
			end if
			stringVagon = stringVagon & FIELD_TOKEN & imprimirVisteosVagonDetalle(cdRubro, rsDet("vlbonrebaja"))
			cdRubro = cdrubro & "_M"
			if (not dicRubros.Exists(cdRubro)) then 
				txtRubro = "MERMA_" & txtRubro
				dicRubros.Add cdRubro, txtRubro
			end if
			stringVagon = stringVagon & FIELD_TOKEN & imprimirVisteosVagonDetalle(cdRubro, rsDet("vlmerma"))
			rsDet.MoveNext()
		wend
		'Se agregan los valores de las muestras.
		Set rsMuestras = leerMuestrasVagonDetalle(rsCab("SQCALADA"), rsDatos("CDOPERATIVO"), rsDatos("CDVAGON"),rsCab("DTCONTABLE"), g_Puerto) 
		while (not rsMuestras.eof)	
			cdMuestra = PREFIX_MUESTRA & rsMuestras("SQMUESTRA")
			if (not dicRubros.Exists(cdMuestra)) then dicRubros.Add cdMuestra, "MUESTRA_HUMEDAD_" & rsMuestras("SQMUESTRA")
			stringVagon = stringVagon & FIELD_TOKEN & imprimirMuestrasVagonDetalle(cdMuestra, rsMuestras("VLHUMEDAD"), rsMuestras("VLTEMPERATURA"), rsMuestras("VLPESO"))
			rsMuestras.MoveNext()
		wend
		rsCab.MoveNext()
	wend
	'Imprimo la linea generada al archivo.	
	arch.WriteLine(stringVagon)
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

<P>&nbsp;</P>

</BODY>
</HTML>
