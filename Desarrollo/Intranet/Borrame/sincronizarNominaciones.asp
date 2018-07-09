<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<%
'dim connPortName
dim strPuerto, strSQL, myHoy, rs, oConn, ptoAS400
dim rsPto, rsAS400, proxNroCupo, flagSeguir, numeroDePuerto, letraPuerto
dim codigoDesde, codigoHasta, codAbrProducto, myCodigoCupo
flagSeguir = true
strPuerto = GF_Parametros7("Puerto","",6)
if (strPuerto = "") then strPuerto = TERMINAL_PIEDRABUENA
myHoy = GF_DTE2FN(day(date) & "/" & month(date) & "/" & year(date))
'myHoy = "20150601"
Set logPic = new classLog
call startLog(HND_VIEW+HND_FILE, MSG_INF_LOG)
logPic.fileName = "CUPOS-SYNC-" & strPuerto & "-" & myHoy

on error resume next
'Eliminar del AS400-MER517F1 todos los cupos en los que la cantidad asignada sea 0
strSQL = "DELETE FROM MERFL.MER517F1 WHERE CUCCCP=0 AND CUFCCP > " & myHoy
logPic.info("INICIO DE PROCESO - ETAPA 1 ") 
logPic.info("BORRAR CUPOS SIN CAMIONES ASIGNADOS DEL AS400 - " & strSQL) 
Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)
    
'Buscar codigo de puerto CUCDES correcto
numeroDePuerto = getNumeroPuerto(strPuerto)
if numeroDePuerto = -1 then 
	logPic.info("Nombre de puerto incorrecto! - Valores esperados: TRANSITO, PIEDRABUENA, ARROYO")
	Response.End 
end if	 
'Conectar al puerto
if connect(strPuerto) then
	'logPic.info  "CONECTO"
	'Leer cupos desde el AS400
	strSQLCab = "SELECT TF.CUFCCP, TF.CUCOOR, TF.CUCPRO, TF.CUCCOR, TF.CUCVEN,CASE WHEN TF.QTASIGNADOS > TF.QTNOMINADOS THEN TF.QTNOMINADOS ELSE QTASIGNADOS END AS CANTIDAD "&_
             "FROM ( "&_
             " SELECT     CUFCCP, NRODOC CUCOOR, CUCPRO, CUCCOR, CUCVEN, CUCCCP AS QTASIGNADOS ,CASE WHEN NOMINADOS IS NULL THEN 0 ELSE NOMINADOS END AS QTNOMINADOS " & _
			 " FROM       MERFL.MER517F1 F1 " &_
			 " INNER JOIN MERFL.TCB6A1F1 PRO ON PRO.NROPRO=F1.CUCVEN " & _
             " LEFT JOIN ( SELECT   IDCUPO, SUM(CANTIDAD) AS NOMINADOS  "&_
             "            FROM MERFL.TBLCUPOSNOMINADOS "&_
             "            GROUP BY IDCUPO) N  ON F1.CUCODI = N.IDCUPO "&_
             " WHERE CUFCCP > " & MYHOY  & " AND CUCDES=" & NUMERODEPUERTO & " AND CUCOPE="& OPERACION_PRESTAMO_DEVOLUCION  & _
             "  ) TF "&_
             " WHERE TF.QTNOMINADOS > 0 "&_
             "  ORDER BY  tf.CUFCCP, tf.CUCOOR, tf.CUCPRO,  tf.CUCCOR, tf.CUCVEN"
    Call GF_BD_COMPRAS(rsAS400, oConn, "OPEN", strSQLCab)
	logPic.info("SELECCION DE CUPOS DEL AS400 - " & strSQLCab)
	'Leer cupos desde DB2
	strSQL = " SELECT IDCUPO, DTCUPO, NUCUITCOORDINADO, CDPRODUCTO, CDCORREDOR, CDVENDEDOR, QTASIGNADOS FROM " & _
			 "		ASIGNACIONCUPOS WHERE DTCUPO > " & myHoy & _
			 "			ORDER BY DTCUPO, NUCUITCOORDINADO, CDPRODUCTO, CDCORREDOR, CDVENDEDOR "
    
	call GF_BD_Puertos (strPuerto, rsPto, "OPEN", strSql)
	logPic.info("SELECCION DE CUPOS DEL PUERTO - " & strSQL)
    
	while flagSeguir
		if not rsPto.eof and not rsAS400.eof then
            cuposPermitidos = clng(trim(rsAS400("CANTIDAD")))
			select case CompararClave(trim(rsAS400("CUFCCP")), trim(rsPto("DTCUPO")), trim(rsAS400("CUCOOR")), trim(rsPto("NUCUITCOORDINADO")), trim(rsAS400("CUCPRO")), trim(rsPto("CDPRODUCTO")), trim(rsAS400("CUCCOR")), trim(rsPto("CDCORREDOR")), trim(rsAS400("CUCVEN")), trim(rsPto("CDVENDEDOR")))
				case 1
					logPic.info("CLAVE DEL AS400 MENOR A LA DEL PUERTO - AGREGAR CUPO DEL AS400.")
					call InsertarCupo(trim(rsAS400("CUFCCP")), trim(rsAS400("CUCOOR")), trim(rsAS400("CUCPRO")), trim(rsAS400("CUCCOR")), trim(rsAS400("CUCVEN")),trim(rsAS400("CUFCCP")), cuposPermitidos)
					rsAS400.movenext
				case 0
					logPic.info("CLAVE DEL AS400 IGUAL A LA DEL PUERTO - CUPOS AS400 " & cuposPermitidos & ", CUPOS PUERTO " & clng(trim(rsPto("QTASIGNADOS"))))
					if clng(cuposPermitidos) <> clng(trim(rsPto("QTASIGNADOS"))) then call ActualizarCupo(trim(rsPto("IDCUPO")), cuposPermitidos)	
					rsAS400.movenext
					rsPto.movenext
				case -1
					logPic.info("CLAVE DEL AS400 MAYOR A LA DEL PUERTO - QUITAR CUPO DEL PUERTO.")
					call QuitarCupo(trim(rsPto("IDCUPO"))) 
					rsPto.movenext
			end select	
		else
			logPic.info("QUITAR CUPOS ELIMINADOS DEL AS400 QUE AUN ESTEN EN PUERTO.")
			'Elimino los que quedaron en puerto
			while not rsPto.eof
					call QuitarCupo(trim(rsPto("IDCUPO"))) 
				rsPto.MoveNext
			wend	
			'Agrego los que quedaron en AS400
			logPic.info("AGREGAR NUEVOS CUPOS DEL AS400")
			while not rsAS400.eof
                    cuposPermitidos = clng(trim(rsAS400("CANTIDAD")))
					call InsertarCupo(trim(rsAS400("CUFCCP")), trim(rsAS400("CUCOOR")), trim(rsAS400("CUCPRO")), trim(rsAS400("CUCCOR")), trim(rsAS400("CUCVEN")), trim(rsAS400("CUFCCP")), cuposPermitidos)
				rsAS400.MoveNext
			wend			
			flagSeguir = false
		end if
	wend
	
    logPic.info("INICIO DE PROCESO - ETAPA 2 - CODIGOS DE CUPOS") 
	'Buscar letra de puerto 
    nroPuerto = getNumeroPuerto(strPuerto)
	letraPuerto = getLetraCupo(nroPuerto)
	'Quitar los codigos de cupo cargados
	strSQL = " DELETE FROM CODIGOSCUPO " & _
			 "		WHERE FECHACUPO > " & myHoy 
    Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
	logPic.info("QUITAR CUPOS DE MAÑANA EN ADELANTE " & strSQL)			 
	

	'Lectura de codigos de cupos asignados y nominados
	strSQL = " SELECT TF.CODIGODESDE,TF.CODIGOHASTA,TF.CUCODI,TF.CUFCCP, TF.CUCPRO, TF.CUCOOR,TF.IDCORREDOR, TF.IDVENDEDOR, TF.DESCPR, TF.A8BGTX,CASE WHEN TF.QTASIGNADOS > TF.QTNOMINADOS THEN TF.QTNOMINADOS ELSE QTASIGNADOS END AS CANTIDAD "&_
             " FROM(SELECT CUFCCP, CUCPRO, NRODOC as CUCOOR,IDCORREDOR, IDVENDEDOR, DESCPR, A8BGTX, "&_
             "            CUCCCP AS QTASIGNADOS, "&_
	         "            CASE WHEN NOMINADOS IS NULL THEN 0 ELSE NOMINADOS END AS QTNOMINADOS, "&_
	         "            NN.CODIGODESDE, NN.CODIGOHASTA,CUCODI "&_
             "      FROM MERFL.MER517F1 F1 " & _
			 "      INNER JOIN MERFL.TCB6A1F1 PRO ON PRO.NROPRO=F1.CUCVEN " & _
			 "      LEFT JOIN MERFL.MER112F1 PROD ON F1.CUCPRO=PROD.CODIPR " & _
             "      LEFT JOIN (SELECT IDCUPO, SUM(CANTIDAD) AS NOMINADOS "&_
		     "                  FROM MERFL.TBLCUPOSNOMINADOS "&_
             "                  GROUP BY IDCUPO) N ON F1.CUCODI = N.IDCUPO "&_
             "      LEFT JOIN MERFL.TBLCUPOSNOMINADOS NN ON NN.IDCUPO = F1.CUCODI "&_
			 "      WHERE CUFCCP > " & myHoy  & " AND CUCDES=" & numeroDePuerto &" AND F1.CUCOPE = "& OPERACION_PRESTAMO_DEVOLUCION &_
             " ) TF "&_
             " WHERE TF.QTNOMINADOS > 0 "&_
             " ORDER BY TF.CUFCCP, TF.CUCOOR, TF.CUCPRO, TF.IDCORREDOR, TF.IDVENDEDOR "
    Call executeQuery(rsAS400, "OPEN", strSQL)
	logPic.info("SELECCION DE RANGOS DE CUPOS ASIGNADOS AL CUPO - " & strSQL)			 
	while not rsAS400.eof
            cuposPermitidos = clng(trim(rsAS400("CANTIDAD")))
            
			codAbrProducto = trim(rsAS400("A8BGTX"))
			logPic.info("MIGRANDO CUPOS DEL CORREDOR " & rsAS400("IDCORREDOR") &", VENDEDOR " & rsAS400("IDVENDEDOR") ) 
			if codAbrProducto = "" then 
				codAbrProducto = "ER"
				logPic.info("FALTA LA DESCRIPCION ABREVIADA DEL PRODUCTO " & trim(rsAS400("CUCPRO")) & ". CUPOS NO ENVIADOS.")			
                rsAS400.movenext()
			else
                flagSeguir = true
                cuposGrabados = 0
                cdCupo = rsAS400("CUCODI")
                while (corteControlCuposNominados(rsAS400,cdCupo,flagSeguir))
                    codigoDesde = clng(rsAS400("CODIGODESDE"))
			        codigoHasta = clng(rsAS400("CODIGOHASTA"))
                    while ((codigoDesde <= codigoHasta)and(flagSeguir))
					     while len(codigoDesde) < 8 
						    codigoDesde = "0" & codigoDesde 
					     wend
					     myCodigoCupo = LEFT(trim(rsAS400("DESCPR")),1) & codigoDesde
					     call insertarCodigoDeCupo(trim(rsAS400("CUFCCP")), trim(rsAS400("CUCPRO")), trim(rsAS400("CUCOOR")),trim(rsAS400("IDCORREDOR")), trim(rsAS400("IDVENDEDOR")), myCodigoCupo)
					     codigoDesde = codigoDesde + 1
                         cuposGrabados = cuposGrabados + 1
                         if (CLng(cuposPermitidos) <= CLng(cuposGrabados)) then flagSeguir = false
				    wend
                    logPic.info("--->CODIGO DE CUPO DE " & rsAS400("CODIGODESDE") &" al "& codigoDesde)
                    rsAS400.movenext()
                wend
			end if
	wend
    
	logPic.info("INICIO DE PROCESO - ETAPA 3 - VALIDACION DE MIGRACION") 		
	
	strSQL="Select sum(cantidad) as Cantidad from ("& strSQLCab &") TT "
	Call executeQuery(rs, "OPEN", strSQL)
	myCant = 0
	if (not rs.eof) then myCant = rs("Cantidad")
	logPic.info("Cantidad de Cupos en Buenos Aires: " & myCant) 
	strMsg = strMsg & "Cantidad de Cupos en Buenos Aires: " & myCant & chr(13) & chr(13)
	'--
	strSQL="Select sum(QTASIGNADOS) Cantidad from ASIGNACIONCUPOS where DTCUPO > " & myHoy	
	Call GF_BD_Puertos(strPuerto, rs, "OPEN", strSQL)	
	myCant = 0
	if (not rs.eof) then myCant = rs("Cantidad")
	logPic.info("Cantidad de Cupos en Puerto: " & myCant) 
	strMsg = strMsg & "Cantidad de Cupos en Puerto: " & myCant & chr(13) & chr(13)
	'--
	strSQL="Select count(*) Cantidad from CODIGOSCUPO where FECHACUPO > " & myHoy
	Call GF_BD_Puertos(strPuerto, rs, "OPEN", strSQL)
	myCant = 0
	if (not rs.eof) then myCant = rs("Cantidad")
	logPic.info("Cantidad de Codigos en Puerto: " & myCant) 
	strMsg = strMsg & "Cantidad de Codigos en Puerto: " & myCant
	
	Call GP_ENVIAR_MAIL ("Sincronizacion de Cupos - Control de Sicronizacion " & strPuerto, strMsg, SENDER_MERCADERIAS ,"ScalisiJ@toepfer.com")
else
	Call GP_ENVIAR_MAIL ("Sincronizacion de Cupos - Error al conectar con " & strPuerto, "Imposible conectar con el puerto '" & strPuerto & "'" & chr(13) & chr(13) & "Por favor, contactese con el administrador del sistema." & chr(13) & chr(13) & "Muchas gracias.", SENDER_MERCADERIAS ,"ScalisiJ@toepfer.com")
end if
'-----------------------------------------------------------------------------------------------------------------
Function corteControlCuposNominados2(rsAS400,cdCorredor,cdVendedor,fechaCupo,cuitCoordinador,flagSeguir)
    corteControlCuposNominados = false
    if(not rsAS400.Eof)then
        if((CLng(rsAS400("CUFCCP"))=Clng(fechaCupo))and(Cdbl(rsAS400("CUCOOR"))=Cdbl(cuitCoordinador))and(CLng(rsAS400("IDVENDEDOR"))=Clng(cdVendedor))and(CLng(rsAS400("IDCORREDOR"))=Clng(cdCorredor)))then 
            if (flagSeguir)then corteControlCuposNominados = true
        end if
    end if
End Function 
'-----------------------------------------------------------------------------------------------------------------
Function corteControlCuposNominados(rsAS400,cdCupo,flagSeguir)
    corteControlCuposNominados = false
    if(not rsAS400.Eof)then
        if((CLng(rsAS400("CUCODI"))=Clng(cdCupo))AND(flagSeguir))then corteControlCuposNominados = true
    end if
End Function 
'-----------------------------------------------------------------------------------------------------------------
function CompararClave(pFechaCupoAS400, pFechaCupoPuerto, pCuitCoordinadoAS400, pCuitCoordinadoPuerto, pCdProductoAS400, pCdProductoPuerto, pCdCorredorAS400, pCdCorredorPuerto, pCdVendedorAS400, pCdVendedorPuerto)
dim rtrn
'rtrn puede tomar 3 valores
'	-1 = La clave del AS400 es mayor a la del Puerto - Quitar cupo del Pto y mover siguiente registro de Puerto
'	 0 = Las claves son iguales - Actualizar cupos asignados
'	+1 = La clave del AS400 es menor a la del Puerto - Insertar cupo del AS400 y mover siguiente registro en AS400
rtrn = 1
if cdbl(pFechaCupoAS400) = cdbl(pFechaCupoPuerto) then
	if cdbl(pCuitCoordinadoAS400) = cdbl(pCuitCoordinadoPuerto) then
		if cdbl(pCdProductoAS400) = cdbl(pCdProductoPuerto) then
			if cdbl(pCdCorredorAS400) = cdbl(pCdCorredorPuerto) then
				if cdbl(pCdVendedorAS400) = cdbl(pCdVendedorPuerto) then
					rtrn = 0
				else	
					if cdbl(pCdVendedorAS400) > cdbl(pCdVendedorPuerto) then rtrn = -1
				end if								
			else
				if cdbl(pCdCorredorAS400) > cdbl(pCdCorredorPuerto) then rtrn = -1
			end if						
		else	
			if cdbl(pCdProductoAS400) > cdbl(pCdProductoPuerto) then rtrn = -1
		end if						
	else	
		if cdbl(pCuitCoordinadoAS400) > cdbl(pCuitCoordinadoPuerto) then rtrn = -1
	end if						
else
	if cdbl(pFechaCupoAS400) > cdbl(pFechaCupoPuerto) then rtrn = -1
end if	
CompararClave = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
sub InsertarCupo(pFechaCupo, pCuitCoordinado, pCdProducto, pCdCorredor, pCdVendedor, pdtCupo, pQtAsinados)
'Se setea en la variable global cual sera el proximo nro de cupo
call setProximoNroCupo
'Insertar el cupo recibido por parametro
strSQL = "INSERT INTO ASIGNACIONCUPOS VALUES(" & proxNroCupo & ",'" & pCuitCoordinado & "'," & pCdCorredor & "," & pCdVendedor & "," & pCdProducto  & "," & pdtCupo & "," & pQtAsinados & ",0)"
logPic.info("Insercion de Cupo - " & strSQL)
Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------------------------------
sub QuitarCupo(pIdCupo)
'Quitar el cupo recibido por parametro
strSQL = " DELETE FROM ASIGNACIONCUPOS WHERE IDCUPO=" & pIdCupo
logPic.info("Eliminacion de Cupo - " & strSQL)
Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
'Quitar cupo especial
strSQL = " DELETE FROM CONTRATOSESPECIALES WHERE IDCUPO=" & pIdCupo
logPic.info("Eliminacion de Cupo Especial - " & strSQL)
Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------------------------------
sub ActualizarCupo(pIdCupo, pQtAsinados)
'Actualizar la cantidad de cupos asignados
strSQL = " UPDATE ASIGNACIONCUPOS SET QTASIGNADOS=" & pQtAsinados & " WHERE IDCUPO=" & pIdCupo
logPic.info("Actualizacion de Cupo - " & strSQL)
Call GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL)
end sub
'-----------------------------------------------------------------------------------------------------------------
function getLetraPuerto(pName)
dim rtrn
rtrn = ""
select case ucase(pName) 
	case "TRANSITO"
		rtrn = "T"
	case "PIEDRABUENA"
		rtrn = "P"
	case "ARROYO"
		rtrn = "A"
end select 	
getLetraPuerto = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
function getNumeroPuerto(pName)
dim rtrn
rtrn = -1
select case ucase(pName) 
	case "TRANSITO"
		rtrn = 10
	case "PIEDRABUENA"
		rtrn = 91
	case "ARROYO"
		rtrn = 36
end select 	
getNumeroPuerto = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
function setProximoNroCupo()
dim rsNextCupo
'Controlar si tengo el siguiente nro de cupo
if proxNroCupo = 0 then
	'Buscar el siguiente nro de cupo
	strSQL = " SELECT MAX(IDCUPO) AS ULTNROCUPO FROM ASIGNACIONCUPOS "
	Call GF_BD_Puertos (strPuerto, rsNextCupo, "OPEN", strSql)
	if IsNull(rsNextCupo("ULTNROCUPO")) then 
		proxNroCupo = 1
	else
		proxNroCupo = CDBL(rsNextCupo("ULTNROCUPO")) + 1
	end if	
else
	proxNroCupo = proxNroCupo + 1
end if
end function
'-----------------------------------------------------------------------------------------------------------------
function getNroCupo(pFechaCupo, pCuitCoordinado, pCdProducto, pCdCorredor, pCdVendedor)
dim rsCupo, rtrn
'Buscar el id cupo en puerto segun corresponda
strSQL = " SELECT IDCUPO FROM ASIGNACIONCUPOS WHERE DTCUPO=" & pFechaCupo & " AND NUCUITCOORDINADO='" & pCuitCoordinado & "' AND CDPRODUCTO=" & pCdProducto & " AND CDCORREDOR=" & pCdCorredor & " AND CDVENDEDOR=" & pCdVendedor 
logPic.info("NRO DE CUPO - " & strSQL)
Call GF_BD_Puertos (strPuerto, rsCupo, "OPEN", strSql)
if not rsCupo.eof then 
	rtrn = rsCupo("IDCUPO")
else
	rtrn = "ERROR"
end if	
getNroCupo = rtrn
end function
'-----------------------------------------------------------------------------------------------------------------
sub InsertarCodigodeCupo(pFechaCupo, pCdProducto, pCuitCoordinado, pCdCorredor, pCdVendedor, pCodigoCupo)
'Insertar el cupo recibido por parametro
strSQL = "INSERT INTO CODIGOSCUPO VALUES(" & pFechaCupo & "," & pCdProducto & ",'" & pCuitCoordinado & "'," & pCdCorredor & "," & pCdVendedor & ",'" & pCodigoCupo & "')"
if GF_BD_Puertos(strPuerto, rsPto, "EXEC", strSQL) then
	logPic.info("Insercion de Codigo de Cupo - " & strSQL)
else	
	logPic.info("Error en insercion de cupo - " & strSQL)
end if	
end sub
%>
