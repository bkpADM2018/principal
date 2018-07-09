<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosBoletos.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<%

'-----------------------------------------------------------------------------------------------
Dim idProducto, g_strPuerto, nroPuerto

Call GP_CONFIGURARMOMENTOS


idProducto= GF_PARAMETROS7("producto", 0, 6)
g_strPuerto = GF_PARAMETROS7("pto","",6)
cdProveedor = GF_PARAMETROS7("prov",0,6)
nroPuerto   = getNumeroPuerto(g_strPuerto)

'------------------------------------------------------------------------------------------	
Function dibujarCampo(line, val, sz)

    dibujarCampo = line & GF_nChars(val,sz," ", CHR_AFT) & "&#09;"
    
End Function
'------------------------------------------------------------------------------------------	
' Dibuja los cupos que se tienen nominados para un determinado contrato entre un rango de fechas
' NOTA: 
'       Esta funcion trabaja solamente si el puerto es Bahia Blanca y la operacion es devolucion/prestamo (04),
'       al ver la informacion en la tabla que muestra el reporte por cada contrato se tomo la desicion de que esta funcion
'       trabaje correctamente con este puerto solamente. 
Function dibujarCuposNominados(p_cdCliente, p_fechaDesde, p_idProducto, p_NroPuerto)
    Dim rsNom,letraCupo,auxCodigoCupo,auxCodigo, linea
    Dim myFechaDesde, myFechaHasta
    
    myFechaDesde = GF_DTEADD(p_fechaDesde, 1, "D")
    myFechaHasta = GF_DTEADD(p_fechaDesde, 20, "D")
    
    Set sp_return = executeSP(rsNom, "MERFL.TBLCUPOSNOMINADOS_GET_BY_PARAMETERS", myFechaDesde &"||"& myFechaHasta &"||"& p_NroPuerto &"||"& p_idProducto &"||0||0||"& p_cdCliente)
        
    if (not rsNom.Eof) then        
        response.Write "Cupos otorgados a " & getDescripcionProveedor(rsNom("CLIENTE")) & "<br /><br />"
        
        linea = ""       
        linea = dibujarCampo(linea, "Fecha", 13)        
        linea = dibujarCampo(linea, "Cupo Asignado", 15)        
        linea = dibujarCampo(linea, "Producto", 15)
        linea = dibujarCampo(linea, "Corredor", 50)        
        linea = dibujarCampo(linea, "Vendedor", 50)
        response.write Replace(linea, " ", "&nbsp;") & "<br />"
        linea = ""       
        linea = dibujarCampo(linea, "-------------------------------------------------------------------------------------------------------------------------------------------------", 145)
        response.write Replace(linea, " ", "&nbsp;") & "<br />"
        while (not rsNom.Eof)             
            'Obtengo el codigo de cupo para Bahia blanca
            auxCodigo = GF_nDigits(rsNom("CODIGO"),8)
            auxCodigo = LEFT(Trim(rsNom("DSPRODUCTO")),1) & auxCodigo            
            linea = ""       
            linea = dibujarCampo(linea, GF_FN2DTE(rsNom("FECHACUPO")), 13)
            linea = dibujarCampo(linea, auxCodigo, 15)
            linea = dibujarCampo(linea, Left(rsNom("DSPRODUCTO"), 15), 15)            
            linea = dibujarCampo(linea, rsNom("IDCORREDOR") &"-"& Left(getDsCorredor(rsNom("IDCORREDOR")), 50), 50)
            linea = dibujarCampo(linea, rsNom("IDVENDEDOR") &"-"& Left(getDsVendedor(rsNom("IDVENDEDOR")), 50), 50)
            response.write Replace(linea, " ", "&nbsp;") & "<br />"
            rsNom.MoveNext()
        wend        
    end if    
End Function
'------------------------------------------------------------------------------------------	
%>

<html>
    <head>
    </head>
    <body style="font-family: Courier;">
        <h3>NOMINACIONES RELIZADAS</h3>        
<%        	
       Call dibujarCuposNominados(cdProveedor, Left(session("MmtoSistema"), 8), idProducto, nroPuerto)
	%>
	    <br /><b>Los cupos de lunes a viernes tendrán apertura el día anterior al cupo a las 15 hs y como cierre el día del mismo a las 17 hs.</b><br />
        <b>Para los días sábados y domingos será el día anterior a las 15 hs,  y el cierre el día del cupo a las 10 hs.</b>
    </body>
</html>