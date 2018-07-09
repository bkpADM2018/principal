<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosformato.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosFacturacionCalidad.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<%
'****************************************************
'*****          COMIENZO DE LA PAGINA           *****
'***************************************************
Dim myHoy,logMig, myHasta, cliente, onScreen, transporte
Dim rs, strSQL, nreg, cartaPorte, strWhere

Call GP_CONFIGURARMOMENTOS()

'session("Usuario") = "SYNC"
'**** PARAMETROS OBLIGATORIOS ****
nreg = GF_PARAMETROS7("nreg", 0, 6)
cartaPorte = GF_PARAMETROS7("cartaPorte", "", 6)


'Se leen los registros de datos a re-migrar.
'strSQL="Select FECDR6 from MERFL.MER711F6 where FCRGR6 in (Select FCRGR7 from MERFL.MER711F7 where MAPLR7=0) and DVCDR6='" & getLetraPuerto(origen) & "' group by FECDR6"
'Call executeQuery(rs, "OPEN", strSQL)

Set logMig = new classLog
Call startLog(HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "REFACTURACION-SYNC-"& origen &"-"& left(session("MmtoDato"),8)	

strSQL="Select DISTINCT CPORR6, IDCAR6, DVCDR6, FCRGR6 from MERFL.MER711F6 " 
'response.write strSQL & "<br>"
if (cartaPorte <> "") then 
    Call mkWhere(strWhere, "CPORR6", cartaPorte, "=", 3)
else
    Call mkWhere(strWhere, "FCRGR6", nreg, "=", 1)
end if
strSQL = strSQL & strWhere
Call executeQuery(rs, "OPEN", strSQL)

if (not rs.eof) then
    if (CDbl(nreg) = 0) then nreg = rs("FCRGR6")
    'Se obtiene el puerto origen
    origen = getDsPuertoByLetra(rs("DVCDR6"))
    'Se agrega la orden de refacturacion.
    strSQL="Insert into MERFL.MER711F7 values(" & nreg & ", " & session("MmtoDato") & ", '" & session("Usuario") & "', 0, '', 0) "    
    'response.Write strSQL & "<br>"
    Call executeQuery(rs2, "EXEC", strSQL)
       
    While (not rs.eof)    
        'Se obtiene la fecha de descarga
        myCartaPorte = GF_nDigits(rs("CPORR6"), 12)
        mySerie = left(myCartaPorte, 4)
        myNro8 = right(myCartaPorte, 8)
        myTransporte = rs("IDCAR6")
        response.Write "Buscando Descarga: " & myCartaPorte & "..."
        'strSQL="Select YEAR(DTCONTABLE) ANIO, MONTH(DTCONTABLE) MES, DAY(DTCONTABLE) DIA from HCAMIONESDESCARGA where NUCARTAPORTE='" & GF_nDigits(rs("CPORR6"), 12) & "' order by DTCONTABLE DESC"
        strSQL="Select DISTINCT ANIO, MES, DIA from (" &_
            "Select YEAR(DTCONTABLE) ANIO, MONTH(DTCONTABLE) MES, DAY(DTCONTABLE) DIA from HCAMIONESDESCARGA where IDCAMION='" & myTransporte & "' and NUCARTAPORTE='" & myCartaPorte & "'" &_ 
            "union " &_
            "Select YEAR(DTCONTABLE) ANIO, MONTH(DTCONTABLE) MES, DAY(DTCONTABLE) DIA from HVAGONES where CDVAGON='" & myTransporte & "' and NUCARTAPORTE='" & myNro8 & "0000' and NUCARTAPORTESERIE='" & mySerie & "') A"
            'response.Write strSQL  
        Call GF_BD_Puertos(origen, rs2, "OPEN", strSQL)   
         if (not rs2.eof) then
            response.Write "encontrda!..."
            myHoy = GF_nDigits(rs2("ANIO"), 4) & GF_nDigits(rs2("MES"), 2) & GF_nDigits(rs2("DIA"), 2)
            Call migrarMermasAFacturar(origen, myHoy, CStr(rs("CPORR6")), myTransporte, TIPO_TRANSPORTE_CAMVAG, "", logMig)        
            response.Write "Migrada OK<br>"
        else            
            response.Write "no encontrda!...NO SE MIGRA.<br>"
        end if        
        rs.MoveNext()
    wend

    strSQL="Update MERFL.MER711F7 set MAPLR7=" & session("MmtoDato") & ", UAPLR7='" & session("Usuario") & "' where FCRGR7 =" & nreg
    'response.Write strSQL & "<br>"
    Call executeQuery(rs, "EXEC", strSQL)    
    
    response.Write " - Refacturación Finalizada - "
    
end if    
%>
