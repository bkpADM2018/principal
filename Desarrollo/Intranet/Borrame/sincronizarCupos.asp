<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<html>
<head>
    <script type="text/javascript">
        function bodyOnLoad() {
            window.close();
        }
    </script>
    
</head>
<body onload="bodyOnLoad()">
<%
'-----------------------------------------------------------------------------------------------------------------
'dim connPortName
dim strPuerto, strSQL, myHoy, rs, oConn, ptoAS400, myMmtoProceso
dim rsPto, rsAS400, proxNroCupo, flagSeguir, numeroDePuerto, letraPuerto
dim codigoDesde, codigoHasta, codAbrProducto, myCodigoCupo, myProducto

Call GP_CONFIGURARMOMENTOS

flagSeguir = true
strPuerto = GF_Parametros7("Puerto","",6)
myProducto = GF_Parametros7("prod",0,6)
myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))
'myHoy = "20150601"
myMmtoProceso = session("MmtoSistema")
Set logMig = new classLog
call startLog(HND_VIEW+HND_FILE, MSG_INF_LOG)
logMig.fileName = "CUPOS-SYNC-" & strPuerto & "-" & myHoy

on error resume next

'Eliminar del AS400-MER517F1 todos los cupos en los que la cantidad asignada sea 0
strSQL = "DELETE FROM MERFL.MER517F1 WHERE CUCCCP=0 AND CUFCCP > " & myHoy
logMig.info("INICIO DE PROCESO - ETAPA 1 ") 
logMig.info("BORRAR CUPOS SIN CAMIONES ASIGNADOS DEL AS400 - " & strSQL) 
Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)

'Buscar codigo de puerto CUCDES correcto
numeroDePuerto = getNumeroPuerto(strPuerto)

if numeroDePuerto = -1 then 
	logMig.info("Nombre de puerto incorrecto! - Valores esperados: TRANSITO, PIEDRABUENA, ARROYO")
	Response.End 
end if	 
'Conectar al puerto
if connect(strPuerto) then

	logMig.info("INICIO DE PROCESO - ETAPA 2 - CODIGOS DE CUPOS")
			
	'Lectura de codigos de cupos asignados
	strSQL = " SELECT CUFCCP, CUCPRO, CASE WHEN CUCOPE=4 then NRODOC else CUCOOR end CUCOOR, CASE WHEN CUCOPE=4 then 0 else case when CUCCOR = 5454 then CCORRH else CUCCOR end end CUCCOR, CASE WHEN CUCOPE=4 then 0 else CUCVEN end CUCVEN, C5DSDE, C5HSTA, DESCPR, A8BGTX, CUCDES FROM MERFL.MER517F1 F1 " & _
			 " INNER JOIN MERFL.MER517F5 F5 ON F1.CUCODI=F5.C5CODI " & _
			 " inner join MERFL.TCB6A1F1 PRO on PRO.NROPRO=F1.CUCVEN " & _
			 " left join MERFL.MER311FH FH on F1.CUCPRO=FH.CPRORH	and F1.CUCSUC=FH.CSUCRH	and F1.CUCOPE=FH.COPERH	and F1.CUNCTO=FH.NCTORH	and F1.CUACOS=FH.ACOSRH " &_
			 " LEFT JOIN MERFL.MER112F1 PROD ON F1.CUCPRO=PROD.CODIPR " & _
			 "		WHERE CUFCCP > " & myHoy  & " AND CUCDES in (" & numeroDePuerto & ")" & _			 			 
			 "      and CUUMMM>=20170612180000 " 
			 if (myProducto > 0) then strSQL= strSQL & " and CUCPRO= " & myProducto
			 strSQL= strSQL & " ORDER BY CUFCCP, CUCOOR, CUCPRO, CUCCOR, CUCVEN "
	Call executeQuery(rsAS400, "OPEN", strSQL)
	logMig.info("SELECCION DE RANGOS DE CUPOS ASIGNADOS AL CUPO - " & strSQL)			 
	while not rsAS400.eof	        
			'Response.Write rsAS400("C5HSTA")
			'Response.End
			letraPuerto = getLetraCupoSync(strPuerto) 
			codigoDesde = clng(rsAS400("C5DSDE"))
			codigoHasta = clng(rsAS400("C5HSTA"))
			codAbrProducto = trim(rsAS400("A8BGTX"))
			logMig.info("MIGRANDO CUPOS. RANGO " & codigoDesde & " A " & codigoHasta)			
			if codAbrProducto = "" then 
				codAbrProducto = "ER"
				logMig.info("FALTA LA DESCRIPCION ABREVIADA DEL PRODUCTO " & trim(rsAS400("CUCPRO")) & ". CUPOS NO ENVIADOS.")			
			else
				while codigoDesde <= codigoHasta
				     myCorredor = rsAS400("CUCCOR")
				     myVendedor = rsAS400("CUCVEN")
			         myKey = rsAS400("CUFCCP") & "_" & codigoDesde			         
			         if (diccNominaciones.Exists(myKey)) then
			            arrAux = Split(diccNominaciones(myKey), "|")
                        myCorredor = arrAux(0)
			            myVendedor = arrAux(1)			            
			         end if 				     
					 while len(codigoDesde) < 8 					    
						codigoDesde = "0" & codigoDesde 
					 wend	
					 if UCASE(strPuerto) = "PIEDRABUENA" then
						myCodigoCupo = LEFT(trim(rsAS400("DESCPR")),1) & (codigoDesde + 20000000)
					 else
						myCodigoCupo = letraPuerto & LEFT(trim(rsAS400("A8BGTX")),2) & codigoDesde
					 end if				
					 strSQL="Select * from CODIGOSCUPO where FECHACUPO=" & trim(rsAS400("CUFCCP")) & " and CODIGOCUPO='" & myCodigoCupo & "'"
					 logMig.info(strSQL)
					 Call GF_BD_Puertos(strPuerto, rsX, "OPEN", strSQL)
					 if (rsX.eof) then
					    strSQL="Insert into CODIGOSCUPO(FECHACUPO, CDPRODUCTO, CUITCLIENTE, CDCORREDOR, CDVENDEDOR, CODIGOCUPO, PATENTE, MOVIL, ESTADO, MMTO, " &_
                               "cuitOrigenWS, cuitIntermediarioWS, cuitRemComercialWS, cuitRepresentanteEntregadorWS, cuitTransportistaWS, cuitChoferWS, idCuitOrigenWS, " &_
                               "idCuitIntermediarioWS, idCuitRemComercialWS, idCuitCorredorVWS, idCuitRepresentanteEntregadorWS, idCuitTransportistaWS, idCuitChoferWS, ctgWS, " &_
		                       "fechaCTG_desdeWS, fechaCTG_HastaWS, cartaporteWS, fechaCP_cargaWS, fechaCP_VtoWS, codLocalidadOrigenWS, desvioWS, cosechaWS, nroEstablecimientoOrigenWS," &_
		                       "pesoNetoEstimadoWS, kmRecorrerWS, dominioWS ) " &_
					            " values (" & trim(rsAS400("CUFCCP")) & ", " &  trim(rsAS400("CUCPRO")) & ", '" & trim(rsAS400("CUCOOR")) & "', " & trim(myCorredor) & ", " & trim(myVendedor) & ",'" & myCodigoCupo & "', '', '', 1, " & myMmtoProceso &_
					            ", '', '', '', '', '', '', 0, 0, 0, 0, 0, 0, 0, '', '2000-01-01 00:00:00.000', '2000-01-01 00:00:00.000', '', '2000-01-01 00:00:00.000', '2000-01-01 00:00:00.000', 0, '', '', 0, 0, 0, '')"
					 else
					    campos = ""
					    if (CInt(rsX("ESTADO")) = 1) then campos = "CDCORREDOR= " & trim(myCorredor) & ", CDVENDEDOR=" & trim(myVendedor) & ", CDPRODUCTO=" & trim(rsAS400("CUCPRO")) & ", CUITCLIENTE='" & trim(rsAS400("CUCOOR")) & "', ESTADO=1, "					    
					    strSQL="Update CODIGOSCUPO Set " & campos & " MMTO=" & myMmtoProceso & " where FECHACUPO=" & trim(rsAS400("CUFCCP")) & " and CODIGOCUPO='" & myCodigoCupo & "'"
                     end if				
                     logMig.info(strSQL)		
                     Call GF_BD_Puertos(strPuerto, rsX, "EXEC", strSQL)                     	 					 					 
					 codigoDesde = codigoDesde + 1
				wend
			end if
		rsAS400.movenext
	wend

    'Quitar los codigos de cupo no existentes en Bs As (Son los que no se actualizaron.    	
	strSQL = " UPDATE CODIGOSCUPO set ESTADO=0 " & _
			 "		WHERE FECHACUPO > " & myHoy & " and MMTO < " & myMmtoProceso &_
			 "              AND (LEFT(CODIGOCUPO, 1) in ('J', 'K') or SUBSTRING(CODIGOCUPO, 2, 1)='2' or SUBSTRING(CODIGOCUPO, 2, 1)='3')"    
			 if (myProducto > 0) then strSQL= strSQL & " and CDPRODUCTO= " & myProducto
    logMig.info("QUITAR CUPOS DE MA�ANA EN ADELANTE " & strSQL)				 
    Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
	
	
	
	Call GP_ENVIAR_MAIL ("Sincronizacion de Cupos - Control de Sicronizacion " & strPuerto, strMsg, SENDER_MERCADERIAS ,"ScalisiJ@toepfer.com")
else
	Call GP_ENVIAR_MAIL ("Sincronizacion de Cupos - Error al conectar con " & strPuerto, "Imposible conectar con el puerto '" & strPuerto & "'" & chr(13) & chr(13) & "Por favor, contactese con el administrador del sistema." & chr(13) & chr(13) & "Muchas gracias.", SENDER_MERCADERIAS ,"ScalisiJ@toepfer.com")
end if
'-----------------------------------------------------------------------------------------------------------------
function getNumeroPuerto(pName)
dim rtrn
rtrn = -1
select case ucase(pName) 
	case "TRANSITO"
		rtrn = "10, 54"
	case "PIEDRABUENA"
		rtrn = 91
	case "ARROYO"
		'Para ARROYO se suma el puerto 18 que se utiliza para difrenciar condiciones de producto.
		rtrn = "36, 18"
end select 	
getNumeroPuerto = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
function getLetraCupoSync(pCodigo)
dim rtrn
rtrn = "?"
    select case ucase(pCodigo) 
        case TERMINAL_TRANSITO
	        rtrn = "J"
        case TERMINAL_ARROYO
	        rtrn = "K"
    end select 	
getLetraCupoSync = rtrn
end function

%>
</body>
</html>