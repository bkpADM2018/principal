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
Function listarFaltantes(rs, tabla)
    Dim ret
    
    ret = ""
    if (not rs.eof) then    
        ret = "CAMIONES FALTANTES EN LA TABLA: " &  tabla & vbcrlf & vbcrlf    
        while (not rs.eof)
            ret = ret & rs("IDCAMION") & vbcrlf
            rs.MoveNext()
        wend
        logMig.info("Resultado: " & ret)
    else        
        logMig.info("Resultado: OK!")       
    end if        
    
    listarFaltantes = ret
    
End Function

dim strPuerto, strSQL, myHoy, rs, oConn, ptoAS400, myMmtoProceso
dim rsPto, rsAS400, proxNroCupo, flagSeguir, numeroDePuerto, letraPuerto
dim codigoDesde, codigoHasta, codAbrProducto, myCodigoCupo, myBody

Call GP_CONFIGURARMOMENTOS

strPuerto = GF_Parametros7("pto","",6)
g_strPuerto = strPuerto
myHoy = GF_Parametros7("fd","",6)
if (fd = "") then myHoy = Left(session("MmtoSistema"), 8)
myHoy = GF_FN2DTCONTABLE(myHoy)
'myHoy = "20150601"
Set logMig = new classLog
call startLog(HND_VIEW+HND_FILE, MSG_INF_LOG)
logMig.fileName = "HISTORICO-CTRL-" & strPuerto & "-" & myHoy

'on error resume next
'Comparo la tabla Camiones
strSQL = "Select D.* from CAMIONES D where CDESTADO in (6, 8, 7) and not EXISTS (Select * from HCAMIONES H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = listarFaltantes(rs, "HCAMIONES")
'Comparo la tabla CamionesDescarga
strSQL = "Select D.* from CAMIONESDESCARGA D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HCAMIONESDESCARGA H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and (D.NUCARTAPORTE=H.NUCARTAPORTE or D.CTG=H.CTG))"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HCAMIONESDESCARGA")
'Comparo la tabla CamionesCarga
strSQL = "Select D.* from CAMIONESCARGA D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HCAMIONESCARGA H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and D.NUCARTAPORTE=H.NUCARTAPORTE)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HCAMIONESCARGA")
'Comparo la tabla PesadasCamion
strSQL = "Select D.* from PESADASCAMION D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HPESADASCAMION H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and D.VLPESADA=H.VLPESADA)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HPESADASCAMION")
'Comparo la tabla CaladaDeCamiones
strSQL = "Select D.* from CALADADECAMIONES D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HCALADADECAMIONES H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and D.SQCALADA=H.SQCALADA and D.VLHUMEDAD=H.VLHUMEDAD)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HCALADADECAMIONES")
'Comparo la tabla RubroVisteoCalada
strSQL = "Select D.* from RUBROSVISTEOCAMIONES D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HRUBROSVISTEOCAMIONES H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and D.SQCALADA=H.SQCALADA and D.CDRUBRO=H.CDRUBRO and D.VLBONREBAJA=H.VLBONREBAJA)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HRUBROSVISTEOCAMIONES")
'Comparo la tabla MermaCAmiones
strSQL = "Select D.* from MERMASCAMIONES D where D.IDCAMION in (Select IDCAMION from CAMIONES where CDESTADO in (6, 8, 7)) and not EXISTS (Select * from HMERMASCAMIONES H where H.DTCONTABLE='" & myHoy & "' and D.IDCAMION=H.IDCAMION and D.SQPESADA=H.SQPESADA and D.VLMERMAKILOS=H.VLMERMAKILOS)"
logMig.info(strSQL)
Call GF_BD_PUERTOS(strPuerto, rs, "OPEN", strSQL)  
myBody = myBody & listarFaltantes(rs, "HMERMACAMIONES")

if (myBody = "") then myBody="TODO OK!"
logMig.info(myBody)
Call GP_ENVIAR_MAIL ("Poseidon - Control de Pasaje a Historico " & strPuerto, myBody, SENDER_MERCADERIAS ,"ScalisiJ@toepfer.com")

%>
