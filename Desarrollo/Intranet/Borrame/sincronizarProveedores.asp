<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<%
'dim connPortName
dim strPuerto, strSQL, myHoy, rs, oConn, ptoAS400, tipo
dim rsPto, rsAS400, proxNroCupo, flagSeguir, numeroDePuerto, letraPuerto
dim codigoDesde, codigoHasta, codAbrProducto, myCodigoCupo

Call GP_CONFIGURARMOMENTOS

strPuerto = GF_Parametros7("Puerto","",6)
Set logPic = new classLog
call startLog(HND_VIEW+HND_FILE, MSG_INF_LOG)
myHoy = Left(session("MmtoDato"), 8)
logPic.fileName = "PROVEEDORES-SYNC-" & strPuerto & "-" & myHoy

'on error resume next
'Buscar codigo de puerto CUCDES correcto
numeroDePuerto = getNumeroPuerto(strPuerto)
if numeroDePuerto = -1 then 
	logPic.info("Nombre de puerto incorrecto! - Valores esperados: TRANSITO, PIEDRABUENA, ARROYO")
	Response.End 
end if	 
'Conectar al puerto
if connect(strPuerto) then
    'Tomo los proveedores de Bs As
    strSQL="Select * from QS36F.""TG.6A1F1"" order by NROPRO"
    Call executeQuery(rs, "OPEN", strSQL)
    
    logPic.info("Sincronizando CORREDORES y VENDEDORES...")   
	'Tomo Corredores del Puerto
	strSQL = " SELECT * from CORREDORES  where CDCORREDOR<100000 order by CDCORREDOR"
	Call GF_BD_Puertos (strPuerto, rsPto, "OPEN", strSql)
	while ((not rs.eof) and (not rsPto.eof))
        idx = idx + 1
        logPic.info("Analizando CORREDOR: " & rs("NROPRO"))
        if (CLng(rs("NROPRO")) = CLng(rsPto("CDCORREDOR"))) then
            'Es el mismo proveedor, copio el estado.
            'if (rs("ESTADO") <> rsPto("CDESTADO")) then
                logPic.info("1.- Se actualiza el estado del Proveedor:" & rsPto("CDCORREDOR") & " a ESTADO='" & rs("ESTADO") & "'")                
                strSQL="Update CORREDORES set CDESTADO='" & rs("ESTADO") & "' where CDCORREDOR=" & rsPto("CDCORREDOR")
                Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)                
                strSQL="Update VENDEDORES set CDESTADO='" & rs("ESTADO") & "' where CDVENDEDOR=" & rsPto("CDCORREDOR")
                Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)                
            'end if            
            rs.MoveNext()
            rsPto.MoveNext()
        else 
            if (CLng(rs("NROPRO")) < CLng(rsPto("CDCORREDOR"))) then
                'El proveedores existe en bs pero no en el puerto!
                logPic.info("1.- Se da de alta el CORREDOR y VENDEDOR:" & rs("NROPRO") & " con  ESTADO='" & rs("ESTADO") & "'")                
                Call grabarProveedorPuertos(strPuerto, rs("NROPRO"), ucase(rs("RAZSOC")), ucase(rs("DOMICI")), getDsTipoDoc(rs("TIPDOC")), rs("NRODOC"), rs("ESTADO"))                
                rs.MoveNext()
             else  
                'El proveedor existe en el puerto pero no en bs as!!!!!!
                logPic.info("1.- ERROR - El CORREDOR/VENDEDOR:" & rsPto("CDCORREDOR") & " no existe en Bs AS !!! se da de baja.")  
                strSQL="Update CORREDORES set CDESTADO='" & ESTADO_DESHABILITADO & "' where CDCORREDOR=" & rsPto("CDCORREDOR")
                Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)        
                strSQL="Update VENDEDORES set CDESTADO='" & ESTADO_DESHABILITADO & "' where CDVENDEDOR=" & rsPto("CDCORREDOR")
                Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)                        
                rsPto.MoveNext()
             end if
        end if                              
    wend    
    'Analizo los que restaron en cada uno de los recordsets cuando el otro haya terminado.
    while (not rs.eof)        
        'El proveedores existe en bs pero no en el puerto!        
        logPic.info("2 - Se da de alta el CORREDOR/VENDEDOR:" & rs("NROPRO") & " con  ESTADO='" & rs("ESTADO") & "'")                 
        Call grabarProveedorPuertos(strPuerto, rs("NROPRO"), ucase(rs("RAZSOC")), ucase(rs("DOMICI")), getDsTipoDoc(rs("TIPDOC")), rs("NRODOC"), rs("ESTADO"))                        
        rs.MoveNext()        
    wend
    while (not rsPto.eof)         
        'El proveedor existe en el puerto pero no en bs as!!!!!!
        logPic.info("3.- ERROR - El CORREDOR/VENDEDOR:" & rsPto("CDCORREDOR") & " no existe en Bs AS !!! se da de baja.")  
        strSQL="Update CORREDORES set CDESTADO='" & ESTADO_DESHABILITADO & "' where CDCORREDOR=" & rsPto("CDCORREDOR")
        Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)
        strSQL="Update VENDEDORES set CDESTADO='" & ESTADO_DESHABILITADO & "' where CDVENDEDOR=" & rsPto("CDCORREDOR")
        Call GF_BD_Puertos (strPuerto, rsX, "OPEN", strSql)                
        rsPto.MoveNext()
    wend
    logPic.info("Finaliza Sincronizacion de CORREDORES y VENDEDORES...")           
    
end if    
%>
