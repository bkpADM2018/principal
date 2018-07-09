<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<%

pPto = GF_Parametros7("pto","",6)

'Egresados
strSQL = "SELECT WS.CTG, WS.NUCARTAPORTE, CAST(DC.CDEXTERNO AS BIGINT) AS CDPRODUCTOAFIP, HC.DTCONTABLE, HC.CDPRODUCTO, HCD.CDCOSECHA, " & _
             "   (SELECT VLPESADA FROM HPESADASCAMION WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION AND CDPESADA=1 AND SQPESADA=(SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION AND CDPESADA=1)) - " & _
             "   (SELECT VLPESADA FROM HPESADASCAMION WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION AND CDPESADA=2 AND SQPESADA=(SELECT MAX(SQPESADA) FROM HPESADASCAMION WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION AND CDPESADA=2)) - " & _
             "   (SELECT VLMERMAKILOS FROM HMERMASCAMIONES WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION AND SQPESADA=(SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HC.DTCONTABLE AND IDCAMION=HC.IDCAMION)) AS KILOSNETOS " & _
             "    FROM HCAMIONESDESCARGA HCD " & _
             "                  INNER JOIN HCAMIONES HC" & _
             "                          ON HC.CDESTADO IN (6, 8) AND HCD.IDCAMION=HC.IDCAMION AND HCD.DTCONTABLE=HC.DTCONTABLE " & _
             "                  INNER JOIN (Select * from WSCTG_CAMIONES where (ESTADOCONFIRMACION = 0 or ESTADOCONFIRMACION is Null)) WS" & _
             "                          ON WS.NUCARTAPORTE=HCD.NUCARTAPORTE and WS.CTG=CAST(HCD.CTG AS INT) " & _
             "                  INNER JOIN DEVICES_CODE DC " & _
             "                          ON HC.CDPRODUCTO=DC.CDINTERNO AND DC.CDDEVICE= 3 " & _
             "    where HC.DTCONTABLE < '" & GF_FN2DTCONTABLE(GF_DTE2FN(now())) & "'" &_
             "           ORDER BY HC.DTCONTABLE "             
Call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)

while (not rs.eof)
    response.Write "DTCONTABLE: " & rs("DTCONTABLE") & " CARTADE PORTE: " & rs("NUCARTAPORTE") & " - CTG: " & rs("CTG") & "<br>"
    strSQL = "UPDATE dbo.WSCTG_CAMIONES SET MMTOCONFIRMACION=" & GF_DTE2FN(now()) & ", CDUSERCONFIRMACION='ASP', ESTADOCONFIRMACION=" & WSCTG_QUITADO & ", MMTODESVIO=" & GF_DTE2FN(now()) & ", CDUSERDESVIO='ASP', ESTADODESVIO=" & WSCTG_QUITADO & ", MMTORECHAZO=" & GF_DTE2FN(now()) & ", CDUSERRECHAZO='ASP', ESTADORECHAZO=" & WSCTG_QUITADO & " WHERE NUCARTAPORTE='" & rs("NUCARTAPORTE") & "' AND CTG=" & rs("CTG")
    'Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
    rs.MoveNext()
wend             

%>