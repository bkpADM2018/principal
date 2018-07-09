<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosuser.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="reporteCaladorCommon.asp"-->
<%
'--------------PARAMETROS PUERTOS-------------
CONST PARAM_REPORTE_CALADOR  = "REPCALADOR"
Const DESC_PESO_HECTOLITRICO = "PESO HECTOLITRICO"
Const DESC_TEMPERATURA		 = "TEMPERATURA"
Const DESC_PORC_HUMEDAD		 = "PORC. HUMEDAD"
'------------------------------------------------------------------------------------------------------------
Function armarSQLReporteCaladorCamiones()
Dim strSQL
strSQL = "EXEC RMD.dbo.HCALADADECAMIONES_GET_REPORTE_CALADOR '" & g_FechaDesde &"','" & g_cdUsuario &"', "& g_cdCoordinado &","& g_cdCorredor &","& g_cdVendedor &","& g_cdRubro &","& g_cdAceptacion &","& g_cdProducto &","& TIPO_TRANSPORTE_CAMION
Call GF_BD_Puertos(g_Pto, rs, "OPEN", strSQL)
Set armarSQLReporteCaladorCamiones = rs	

End Function
'-------------------------------------------------------------------------------------------------------------
Function armarSQLReporteCaladorVagones()
Dim strSQL
strSQL = "EXEC RMD.dbo.HCALADADEVAGONES_GET_REPORTE_CALADOR '" & g_FechaDesde &"', '" & g_cdUsuario &"', "& g_cdCoordinado &","& g_cdCorredor &","& g_cdVendedor &","& g_cdRubro &","& g_cdAceptacion &","& g_cdProducto &","& TIPO_TRANSPORTE_VAGON
Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)
Set armarSQLReporteCaladorVagones = rs
    
End Function
'-------------------------------------------------------------------------------------------------------------
'Obtiene el promedio de un rubro, dependiendo del camion y de la fecha contable
Function promediarValorRubro(pCdRubro,pDtContable,pIdCamion,pDsRubro,pValor)
	Dim auxDtContable , auxVal
    auxDtContable = Left(pDtContable,4) &"-"& Mid(pDtContable,5,2)&"-"& Right(pDtContable,2)
    auxVal = pValor
    IF Trim(Ucase(pDsRubro)) = DESC_PESO_HECTOLITRICO or Trim(Ucase(pDsRubro)) = DESC_TEMPERATURA or Trim(Ucase(pDsRubro)) = DESC_PORC_HUMEDAD THEN
        Select Case Trim(Ucase(pDsRubro))
            Case DESC_PESO_HECTOLITRICO
                strSQL = "SELECT ROUND(AVG(VLPESO),2)valor FROM HMUESTRASHUMEDCAMIONES HMHC WHERE HMHC.IDCAMION='" & pIdCamion & "' AND HMHC.DTCONTABLE ='" & auxDtContable & "' group by dtcontable,idcamion"            
            Case DESC_TEMPERATURA
                strSQL = "SELECT AVG(VLTEMPERATURA)valor FROM HMUESTRASHUMEDCAMIONES HMHC WHERE HMHC.IDCAMION='" & pIdCamion & "' AND HMHC.DTCONTABLE ='" & auxDtContable & "' AND HMHC.sqcalada =(SELECT MAX(sqcalada) from hmuestrashumedcamiones where idcamion=HMHC.idcamion and dtcontable=HMHC.dtcontable) group by dtcontable,idcamion"
            Case DESC_PORC_HUMEDAD
                strSQL = "SELECT AVG(VLHUMEDAD)valor FROM HMUESTRASHUMEDCAMIONES HMHC WHERE HMHC.IDCAMION='" & pIdCamion & "' AND HMHC.DTCONTABLE ='" & auxDtContable & "' AND HMHC.sqcalada =(SELECT MAX(sqcalada) from hmuestrashumedcamiones where idcamion=HMHC.idcamion and dtcontable=HMHC.dtcontable) group by dtcontable,idcamion"
        End Select

        Call GF_BD_Puertos(g_strPuerto, rsVl, "OPEN", strSQL)
        if (not rsVl.Eof) Then auxVal = Round(rsVl("valor"), 2)
    END IF
	promediarValorRubro = auxVal
End Function
'--------------------------------------------------------------------------------------------------------------
Function imprimirDatosCamiones(pFecha,pIdCamion,pCPorte,pPatente,pCoordinado,pCorredor,pVendedor,pProcedencia,pProducto,pUser,pTerminal,pCalidad,pGrado,pNetoSMerma,pNetoCMerma,pMerma,pBruto,pTara)
	imprimirDatosCamiones = arrTitulosCamiones(0) & VALUE_TOKEN & GF_FN2DTE(pFecha)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(1) & VALUE_TOKEN & GF_nDigits(pIdCamion,10)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(2) & VALUE_TOKEN & GF_EDIT_CTAPTE(pCPorte)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(3) & VALUE_TOKEN & Trim(pPatente)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(4) & VALUE_TOKEN & Trim(pCoordinado) 
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(5) & VALUE_TOKEN & Trim(pCorredor)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(6) & VALUE_TOKEN & Trim(pVendedor)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(7) & VALUE_TOKEN & Trim(pProcedencia)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(8) & VALUE_TOKEN & Trim(pProducto)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(9) & VALUE_TOKEN & Trim(pUser)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(10) & VALUE_TOKEN & Trim(pTerminal)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(11) & VALUE_TOKEN & Trim(pCalidad)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(12) & VALUE_TOKEN & Trim(pGrado)
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(13) & VALUE_TOKEN & pNetoSMerma
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(14) & VALUE_TOKEN & pNetoCMerma
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(15) & VALUE_TOKEN & pMerma
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(16) & VALUE_TOKEN & pBruto
	imprimirDatosCamiones = imprimirDatosCamiones & FIELD_TOKEN & arrTitulosCamiones(17) & VALUE_TOKEN & pTara
End Function
'-------------------------------------------------------------------------------------------------------------
Function imprimirDatosRubros(pAbrRubro,pDsRubro,pRubro,pValor)
	imprimirDatosRubros = arrTitulosRubros(0) & VALUE_TOKEN & pRubro
	imprimirDatosRubros = imprimirDatosRubros & FIELD_TOKEN & arrTitulosRubros(1) & VALUE_TOKEN & Trim(pAbrRubro)
	imprimirDatosRubros = imprimirDatosRubros & FIELD_TOKEN & arrTitulosRubros(2) & VALUE_TOKEN & pValor
	imprimirDatosRubros = imprimirDatosRubros & FIELD_TOKEN & arrTitulosRubros(3) & VALUE_TOKEN & Trim(pDsRubro)
End function
'-------------------------------------------------------------------------------------------------------------
Function imprimirDatosVagones(pFecha,pCdOperativo,pCPorte,pCdVagon,pCoordinado,pCorredor,pVendedor,pProcedencia,pProducto,pUser,pTerminal,pCalidad,pGrado,pNetoSMerma,pNetoCMerma,pMerma,pBruto,pTara,pProteina)
	imprimirDatosVagones = arrTitulosVagones(0) & VALUE_TOKEN & GF_FN2DTE(pFecha)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(1) & VALUE_TOKEN & Trim(pCdOperativo)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(2) & VALUE_TOKEN & GF_EDIT_CTAPTE(pCPorte)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(3) & VALUE_TOKEN & Trim(pCdVagon)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(4) & VALUE_TOKEN & Trim(pCoordinado)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(5) & VALUE_TOKEN & Trim(pCorredor)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(6) & VALUE_TOKEN & Trim(pVendedor)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(7) & VALUE_TOKEN & Trim(pProcedencia)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(8) & VALUE_TOKEN & Trim(pProducto)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(9) & VALUE_TOKEN & Trim(pUser)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(10) & VALUE_TOKEN & Trim(pTerminal)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(11) & VALUE_TOKEN & Trim(pCalidad)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(12) & VALUE_TOKEN & Trim(pGrado)
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(13) & VALUE_TOKEN & pNetoSMerma
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(14) & VALUE_TOKEN & pNetoCMerma
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(15) & VALUE_TOKEN & pMerma
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(16) & VALUE_TOKEN & pBruto
	imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(17) & VALUE_TOKEN & pTara
    imprimirDatosVagones = imprimirDatosVagones & FIELD_TOKEN & arrTitulosVagones(18) & VALUE_TOKEN & pProteina
End Function
'----------------------------------------------------------------------------------------------------------------------
Function seguirProcesandoVagon(pOperativo,pVagon,pRs)
	Dim rtrn
	rtrn = false
	if (not pRs.Eof) then 
		if ((pOperativo = pRs("CDOPERATIVO")) and (pVagon = pRs("CDVAGON"))) then rtrn = true		
	end if
	seguirProcesandoVagon = rtrn
End Function
'----------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
g_fechaDesdeD = GF_PARAMETROS7("fecContableDS", "", 6)
g_fechaDesdeM = GF_PARAMETROS7("fecContableMS", "", 6)
g_fechaDesdeA = GF_PARAMETROS7("fecContableAS", "", 6)
Call GF_STANDARIZAR_FECHA(g_fechaDesdeD, g_fechaDesdeM, g_fechaDesdeA)
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arch = fs.OpenTextFile(strPath, 8, true)
Set fs = nothing

g_FechaDesde = g_fechaDesdeA & g_fechaDesdeM & g_fechaDesdeD

isFirstRegist = true
index = 0
     
'Controlo si esta chequeado la opcion de Camiones
if (g_chkCamiones = 1) then
	Set rsCam = armarSQLReporteCaladorCamiones()
	if not rsCam.Eof then		
		arch.WriteLine(REPORTE_CAMIONES)
		while not rsCam.Eof
            
			if ((idCamion_old <> rsCam("IDCAMION"))and(cPorte_old <> rsCam("CPORTE"))) Then
				if not isFirstRegist Then
					stringRubros = left(stringRubros,len(stringRubros)-3)
					arch.WriteLine(stringCamiones & stringRubros)
				end if
                
				stringCamiones = imprimirDatosCamiones(rsCam("fcontable"),rsCam("idcamion"),rsCam("cporte"),rsCam("patente"),rsCam("coordinado"),rsCam("corredor"),rsCam("vendedor"),rsCam("procedencia"),rsCam("producto"),rsCam("usr"),rsCam("term"),rsCam("calidad"),rsCam("grado"),rsCam("NetoSMer"),rsCam("NetoCMer"),rsCam("KgMerma"),rsCam("bruto"),rsCam("tara")) & SECTOR_TOKEN
				stringRubros = ""
			end if
			isFirstRegist = false
			idCamion_old = rsCam("IDCAMION")
			cPorte_old	 = rsCam("CPORTE")
			'Consulto si eligio la opcion promediar, verifico que el rubro est� habilitado para operar
			auxValor = rsCam("valor")
			if (g_chkPromediar = 1) Then
				auxValor = promediarValorRubro(rsCam("rubro"), rsCam("fcontable"), rsCam("idcamion"),rsCam("dsRubro"),rsCam("valor"))
			end if
            If (auxValor <> "") Then                
                if ((Not IsNull(rsCam("bruto")))and(Not IsNull(rsCam("tara"))))Then dNeto = Cdbl(rsCam("bruto")) - Cdbl(rsCam("tara"))
                auxTotales1 = auxTotales1 + dNeto
                auxTotales2 = auxTotales2 + dNeto * CDbl(auxValor)                
            end if
			stringRubros = stringRubros & imprimirDatosRubros(rsCam("rubroabr"),rsCam("dsRubro"),rsCam("rubro"),auxValor)& DETAIL_TOKEN			
            rsCam.MoveNext()
		wend		
		stringRubros = left(stringRubros,len(stringRubros)-3)		
		arch.WriteLine(stringCamiones & stringRubros)
	end if
end if
cPorte_old = ""
stringRubros = ""
isFirstRegist = true
if (g_chkVagones = 1) then
	Set rsVag = armarSQLReporteCaladorVagones()
	if not rsVag.Eof then
		arch.WriteLine(REPORTE_VAGONES)
		while not rsVag.Eof
			stringVagones = imprimirDatosVagones(rsVag("fcontable"),left(rsVag("operativo"), 12),left(rsVag("cartaporte"), 12),rsVag("cdvagon"),rsVag("coordinado"),rsVag("corredor"),rsVag("vendedor"),rsVag("procedencia"),rsVag("producto"),rsVag("usr"),rsVag("term"),rsVag("calidad"),rsVag("grado"),rsVag("NetoSMer"),rsVag("NetoCMer"),rsVag("KgMerma"),rsVag("bruto"),rsVag("tara"),rsVag("vlproteina")) & SECTOR_TOKEN
			cdOperativo_old = rsVag("CDOPERATIVO")			
			cdVagon_Old		= rsVag("CDVAGON")
			stringRubros = ""
			while(seguirProcesandoVagon(cdOperativo_old,cdVagon_Old,rsVag))
				stringRubros = stringRubros & imprimirDatosRubros(rsVag("rubroabr"),rsVag("dsRubro"),rsVag("rubro"),rsVag("valor")) & DETAIL_TOKEN
				rsVag.MoveNext()	
			wend
			if (stringRubros <> "") then stringRubros = left(stringRubros,len(stringRubros)-3)
			arch.WriteLine(stringVagones & stringRubros)
		wend
	end if
end if
arch.close()
Set arch = Nothing %>
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
