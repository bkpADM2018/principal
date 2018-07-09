<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<%
Const BOLSA_ARROYO = "90"
Const BOLSA_TRANSITO = "91"
Const BOLSA_PIEDRABUENA = "92"

Const RECORD_TOKEN = "||"
Const DATA_TOKEN = ","
Function getDescargas(pPuerto, pBolsa, pDtDesde, pDtHasta)
	Dim strSQL, rs, total311, totalOtros, rtrn		

	strSQL ="	        SELECT  FECHA, CPORR6 CPORTE, " & _
			"	                CASE WHEN SP.CPORC1 IS NULL THEN 0 ELSE 1 END AS MER582, " & _
			"	                CASE WHEN PC.CPORC1 IS NULL THEN 0 ELSE 1 END AS MER583, " & _
			"	                CASE WHEN PR.CPORCA IS NULL THEN 0 ELSE 1 END AS MER591 " & _
			"	                FROM " & _
			"           (SELECT MIN(FECDR6) AS FECHA, CPORR6 FROM MERFL.MER311F6 WHERE CDESR6=" & pPuerto & " AND FECDR6>=" & pDtDesde & " AND FECDR6<=" & pDtHasta & " AND CTRAR6 IN (" & TIPO_TRANSPORTE_CAMION & "," & TIPO_TRANSPORTE_VAGON & "," & TIPO_TRANSPORTE_CAMVAG & ") GROUP BY CPORR6) AS TBL  " & _
			"	        LEFT JOIN " & _
			"                   (SELECT FANAC1, CPORC1, COBEC1 FROM MERFL.MER582F1 GROUP BY FANAC1, CPORC1, COBEC1) AS SP " & _
			"						ON SP.CPORC1=TBL.CPORR6 AND SP.COBEC1=" & pBolsa & _
			"	        LEFT JOIN " & _
			"                   (SELECT FANAC1, CPORC1, COBEC1 FROM MERFL.MER583F1 GROUP BY FANAC1, CPORC1, COBEC1) AS PC " & _
			"						ON PC.CPORC1=TBL.CPORR6 AND PC.COBEC1=" & pBolsa & _
			"	        LEFT JOIN " & _
			"                   (SELECT FANACA, CPORCA, COBECA FROM MERFL.MER591CA GROUP BY FANACA, CPORCA, COBECA) AS PR " & _
			"						ON PR.CPORCA=TBL.CPORR6 AND PR.COBECA=" & pBolsa
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)
	Set getDescargas = rs
End Function
'------------------------------------------------------------
Function getTotalPorGrado(pPuerto, pBolsa, pDtDesde, pDtHasta)
	Dim strSQL, rs, total311, totalOtros, rtrn, myPuerto, myTexto, myProdAnt, myTotalGrados

	myPuerto = getNumeroPuerto(pPuerto)	
	strSQL = "	SELECT CPROCA, GRASCA, COUNT(*) AS TOTAL FROM " & _
			 "		        (SELECT MIN(FECDR6) AS FECHA, CPORR6 FROM MERFL.MER311F6 WHERE CDESR6=" & myPuerto & " AND CTRAR6 IN (" & TIPO_TRANSPORTE_CAMION & "," & TIPO_TRANSPORTE_VAGON & "," & TIPO_TRANSPORTE_CAMVAG & ") GROUP BY CPORR6) AS TBL " & _
			 "		INNER JOIN " & _
			 "		        MERFL.MER591CA AS PR ON PR.CPORCA=TBL.CPORR6 AND TBL.FECHA=PR.FANACA AND PR.COBECA=" & pBolsa & _
			 "		WHERE FECHA>=" & pDtDesde & " AND  FECHA<=" & pDtHasta & " GROUP BY CPROCA, GRASCA ORDER BY CPROCA, GRASCA" 
	Call GF_BD_COMPRAS(rs, conn, "OPEN", strSQL)

    myTexto = "Resumen de calidades (Grado)" & vbCrLf
    myTotalGrados = 0
    if (not rs.eof) then
        myProdAnt = 0        
	    While not rs.eof
	        if (CInt(rs("CPROCA")) <> myProdAnt) then
		        myProdAnt = CInt(rs("CPROCA"))
		        myTexto = myTexto & vbTab & "Producto " & rs("CPROCA") & ":" & vbCrLf 
	        end if	
	        myTexto = myTexto & vbTab & vbTab & "Grado " & rs("GRASCA") & ": " & rs("TOTAL") & " descargas." & vbCrLf
	        myTotalGrados = myTotalGrados + CDbl(rs("TOTAL"))
		    rs.MoveNext()
	    Wend		         
	end if
    myTexto = myTexto & vbTab & "----------------------------------------"  & vbCrLf
    myTexto = myTexto & vbTab & "TOTAL GRAL GRADOS: " & myTotalGrados & " descargas" & vbCrLf & vbCrLf        			
	getTotalPorGrado = myTexto	
End Function
'------------------------------------------------------------
Function procesarPuerto(pPuerto, pBolsa, pDtDesde, pDtHasta)
    Dim myMensaje, myPuerto, flagFound
    Dim totalSinProcasar, totalProcesado, totalCamara, totalAplicadas
    Dim listaSinProcasar, listaProcesado, listaCamara, listaAplicadas, listaSinAnalisis
    
    myMensaje = myMensaje & vbCrLf & "--------------------------------------------------------------------------------------------------" 
    myMensaje = myMensaje & vbCrLf & pPuerto & " - Del " & GF_FN2DTE(dtDesde) & " al " & GF_FN2DTE(dtHasta) & " - Bolsa:" & pBolsa    
    myMensaje = myMensaje & vbCrLf & "--------------------------------------------------------------------------------------------------" & vbCrLf  
    
    myPuerto = getNumeroPuerto(pPuerto)
        
    Set rsDescargas = getDescargas(myPuerto, pBolsa, pDtDesde, pDtHasta)
    
    'Inicializo los valores de contadores    
    totalSinProcasar = 0
    totalProcesado = 0
    totalCamara = 0
    totalAplicadas = 0
    totalSinAnalisis = 0
    
    'Se procesan todas las descargas aplicadas una por una.
    if (not rsDescargas.eof) then        
        while (not rsDescargas.eof)  
            flagFound = false          
            'Sumo la descarga en la lista que corresonde según su estado.
            listaAplicadas= listaAplicadas & rsDescargas("FECHA") & DATA_TOKEN & rsDescargas("CPORTE") & RECORD_TOKEN
            totalAplicadas = totalAplicadas + 1            

            if (CInt(rsDescargas("MER582")) = 1) then
                listaSinProcasar= listaSinProcasar & rsDescargas("FECHA") & DATA_TOKEN & rsDescargas("CPORTE") & RECORD_TOKEN
                totalSinProcasar = totalSinProcasar + 1
                flagFound = true
            end if
            
            if (CInt(rsDescargas("MER583")) = 1) then            
                listaCamara= listaCamara & rsDescargas("FECHA") & DATA_TOKEN & rsDescargas("CPORTE") & RECORD_TOKEN
                totalCamara = totalCamara + 1
                flagFound = true
            end if
            
            if (CInt(rsDescargas("MER591")) = 1) then            
                listaProcesado= listaProcesado & rsDescargas("FECHA") & DATA_TOKEN & rsDescargas("CPORTE") & RECORD_TOKEN
                totalProcesado = totalProcesado + 1
                flagFound = true
            end if
            
            if (not flagFound = true) then            
                listaSinAnalisis= listaSinAnalisis & rsDescargas("FECHA") & DATA_TOKEN & rsDescargas("CPORTE") & RECORD_TOKEN
                totalSinAnalisis = totalSinAnalisis + 1
            end if
            
            rsDescargas.MoveNext()
        wend
        'Se imprime el detalle del cuerpo del mail.
        myMensaje = myMensaje & getTotales2Mail(totalAplicadas, totalSinProcasar, totalCamara, totalProcesado)
        myMensaje = myMensaje & getTotalPorGrado(pPuerto, pBolsa, pDtDesde, pDtHasta)
        myMensaje = myMensaje & getLista2Mail("Descargas sin Análisis:", listaSinAnalisis)
        myMensaje = myMensaje & getLista2Mail("Descargas no Procesadas:", listaSinProcasar)
        myMensaje = myMensaje & getLista2Mail("Descargas con Análisis:", listaProcesado)
    else
        myMensaje = myMensaje & "No hay datos de descargas aplicadas para analizar."
    end if    
    
    procesarPuerto = myMensaje
    
End Function
'------------------------------------------------------------
Function getTotales2Mail(pTotalAplicaciones, pTotalSinProcesar, pTotalCamara, pTotalProcesadas)
         
    Dim myDescomposicion, myCantDescargas, myCantNoProcesados, myCantCamara, myCantProcesados, myAux, myTotal


    myTexto = "Cantidad de Descagas: " & pTotalAplicaciones & vbCrLf &vbCrLf
    myTexto = myTexto & vbTab & "Sin Procesar : " & pTotalSinProcesar & vbCrLf
    myTexto = myTexto & vbTab & "Envio a Cámara : " & pTotalCamara & vbCrLf
    myTexto = myTexto & vbTab & "Análisis Aplicados : " & pTotalProcesadas & vbCrLf
    myTexto = myTexto & vbTab & "-----------------------"  & vbCrLf
    myTexto = myTexto & vbTab & "TOTAL" & vbTab & ":" & (pTotalSinProcesar + pTotalCamara + pTotalProcesadas) & vbCrLf & vbCrLf

    getTotales2Mail = myTexto
    
end function
'------------------------------------------------------------
Function getLista2Mail(pTitle, pList)
    Dim myTexto, myDescomposicion, index, myAux

    myTexto = ""    
    myDescomposicion = split(pList, RECORD_TOKEN)
    myTexto = pTitle & vbCrLf
    myTexto = myTexto & vbTab & vbTab & "    FECHA" & vbTab & vbTab & "CARTA DE PORTE"  & vbCrLf
    myTexto = myTexto & vbTab & vbTab & "---------------" & vbTab & vbTab & "---------------------"  & vbCrLf
    for index=0 to ubound(myDescomposicion)-1
	    myAux = split(myDescomposicion(index), DATA_TOKEN)
	    if ubound(myAux) = 1 then myTexto = myTexto & vbTab & vbTab & GF_FN2DTE(myAux(0)) & vbTab & vbTab & GF_EDIT_CBTE(myAux(1)) & vbCrLf
    next
    getLista2Mail = myTexto
End Function
'------------------------------------------------------------
Function enviarMail(pTitulo,pEmail,pMensaje) 
	Dim emailToepfer , rtrn
	rtrn=false
	emailToepfer = obtenerMail(CD_TOEPFER)	
		if (emailToepfer <> "") and (pEmail <> "") then					
			Call GP_ENVIAR_MAIL(pTitulo, pMensaje, emailToepfer, pEmail)		
			rtrn = true
		end if	
	enviarMail = rtrn
End Function
'**************************************************************
'********************	INICIO DE PAGINA   ********************
'**************************************************************
Dim dtDesde, dtHasta, myCuerpoMail

if (session("Usuario") = "") then session("Usuario") = "JAS"

dtDesde = GF_Parametros7("dtDesde", 0, 6)
'dtDesde = 20140630

if dtDesde = 0 then 
	dtDesde = dateadd("d",-1,date())
	dtDesde = GF_STANDARIZAR_FECHA_RTRN(dtDesde)
	dtDesde = GF_DTE2FN(dtDesde)
end if	

dtHasta = GF_Parametros7("dtHasta", 0, 6)
if dtHasta = 0 then dtHasta = dtDesde

myCuerpoMail = procesarPuerto(TERMINAL_ARROYO, BOLSA_ARROYO, dtDesde, dtHasta)
myCuerpoMail = myCuerpoMail & procesarPuerto(TERMINAL_TRANSITO, BOLSA_TRANSITO, dtDesde, dtHasta)
myCuerpoMail = myCuerpoMail & procesarPuerto(TERMINAL_PIEDRABUENA, BOLSA_PIEDRABUENA, dtDesde, dtHasta)

Call enviarMail("DESCARGAS - CONTROL DE INTEGRIDAD EN ANALISIS", MAILTO_INTEGRIDAD, myCuerpoMail)
%>

