<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosuser.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosLog.asp"-->
<!--#include file="../../Includes/procedimientosLaboratorio.asp"-->
<!--#include file="exportarResultadosCommon.asp"-->
<%
'************************************************************************************************************************
'NOTA: 
'     Esta pagina solo trabaja si el puerto es Piedrabuena, la cual genera el reporte de los analisis a exportar.
'     
'************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------
Function drawTitleReport(pTipoTransporte, pTipoAnalisis)
    Dim strTransporte, strTitle
    if (Cint(pTipoTransporte) = TIPO_TRANSPORTE_CAMION) then
        strTransporte = "CAMIONES" 
    else
        strTransporte = "VAGONES" 
    end if
    archRep.writeline(strTransporte &" - "& pTipoAnalisis)
	strTitle = "Fecha" & string(arrLengthTitleReporte(1)-len("Fecha")," ") &_
               "Coordinador" & string(arrLengthTitleReporte(2)-len("Coordinador")," ") &_
               "Coordinado" & string(arrLengthTitleReporte(3)-len("Coordinado")," ") &_
               "Producto" & string(arrLengthTitleReporte(4)-len("Producto")," ") &_
               "Corredor" & string(arrLengthTitleReporte(5)-len("Corredor")," ") &_
               "Entregador" & string(arrLengthTitleReporte(6)-len("Entregador")," ") &_
               "Vendedor" & string(arrLengthTitleReporte(7)-len("Vendedor")," ") &_
               "Localidad" & string(arrLengthTitleReporte(8)-len("Localidad")," ") &_
               "Patente/Vagon" & string(arrLengthTitleReporte(9)-len("Patente/Vagon")," ") &_
               "Carta Porte" & string(arrLengthTitleReporte(10)-len("Carta Porte")," ") &_
               "Netos" & string(arrLengthTitleReporte(11)-len("Netos")," ") &_
               "Sticker" & string(arrLengthTitleReporte(12)-len("Sticker")," ")
    archRep.writeline(strTitle)
    archRep.writeline(string(240,"-"))
End Function
'------------------------------------------------------------------------------------------------------------------------
Function saveInfoReport(pDtContable,pCoordinador,pCoordinado,pProducto,pCorredor,pEntregador,pVendedor,pProcedencia,pIdTransporte,pCartaPorte,pNeto,pMuestra)
   Dim strDet, auxDescripcion
   strDet = GF_FN2DTE(pDtContable) & string(arrLengthTitleReporte(1)-len(GF_FN2DTE(pDtContable))," ")
   if (CInt(len(pCoordinador)) > CInt(arrLengthTitleReporte(2))) then 
       strDet = strDet & Left(pCoordinador,CInt(arrLengthTitleReporte(2))-2) & ".."
   else
       strDet = strDet & pCoordinador & string(arrLengthTitleReporte(2)-len(pCoordinador)," ")
   end if

   if (CInt(len(pCoordinado)) > CInt(arrLengthTitleReporte(3))) then 
       strDet = strDet & Left(pCoordinado,CInt(arrLengthTitleReporte(3))-2) & ".."
   else
       strDet = strDet & pCoordinado & string(arrLengthTitleReporte(3)-len(pCoordinado)," ")
   end if
   
   if (CInt(len(pProducto)) > CInt(arrLengthTitleReporte(4))) then 
       strDet = strDet & Left(pProducto,CInt(arrLengthTitleReporte(4))-2) & ".."
   else
       strDet = strDet & pProducto & string(arrLengthTitleReporte(4)-len(pProducto)," ")
   end if

   if (CInt(len(pCorredor)) > CInt(arrLengthTitleReporte(5))) then 
       strDet = strDet & Left(pCorredor,CInt(arrLengthTitleReporte(5))-2) & ".."
   else
       strDet = strDet & pCorredor & string(arrLengthTitleReporte(5)-len(pCorredor)," ")
   end if

   if (CInt(len(pEntregador)) > CInt(arrLengthTitleReporte(6))) then 
       strDet = strDet & Left(pEntregador,CInt(arrLengthTitleReporte(6))-2) & ".."
   else
       strDet = strDet & pEntregador & string(arrLengthTitleReporte(6)-len(pEntregador)," ")
   end if

   if (CInt(len(pVendedor)) > CInt(arrLengthTitleReporte(7))) then 
       strDet = strDet & Left(pVendedor,CInt(arrLengthTitleReporte(7))-2) & ".."
   else
       strDet = strDet & pVendedor & string(arrLengthTitleReporte(7)-len(pVendedor)," ")
   end if

   if (CInt(len(pProcedencia)) > CInt(arrLengthTitleReporte(8))) then 
       strDet = strDet & Left(pProcedencia,CInt(arrLengthTitleReporte(8))-2) & ".."
   else
       strDet = strDet & pProcedencia & string(arrLengthTitleReporte(8)-len(pProcedencia)," ")
   end if
   strDet = strDet & pIdTransporte & string(arrLengthTitleReporte(9)-len(pIdTransporte)," ")

   strDet = strDet & GF_EDIT_CTAPTE(pCartaPorte) & string(arrLengthTitleReporte(10)-len(GF_EDIT_CTAPTE(pCartaPorte))," ")
   
   strDet = strDet & GF_EDIT_DECIMALS(pNeto,0) & string(arrLengthTitleReporte(11)-len(GF_EDIT_DECIMALS(pNeto,0))," ")   

   strDet = strDet & pMuestra & string(arrLengthTitleReporte(12)-len(pMuestra)," ")
 
   archRep.writeline(strDet)

End Function
'------------------------------------------------------------------------------------------------------------------------
Function grabarAnalisisProteina(pPath, pTransporte)
    Dim txtLine,archTem,flagHayResultado, total
    Call drawTitleReport(pTransporte,"PROTEINA")
    Set archTem = fs.OpenTextFile(pPath, 1)
    logMig.info(string(20,"-")&" GENERANDO ANALISIS DE PROTEINA "&string(20,"-"))
    flagHayResultado = false
    total = 0
    while (not archTem.AtEndOfStream)        
	    txtLine = archTem.ReadLine()
        txtLine = split(txtLine, KEY_TOKEN)
        if ((CInt(txtLine(0)) = pTransporte)and(CInt(txtLine(1)) = ACEPTACION_REBAJA_CONVENIDA)and(CInt(txtLine(2)) = CInt(auxProductoProteina))) then
            flagHayResultado = true
            strDetalle = split(txtLine(3), FIELD_TOKEN)
            Call saveInfoReport(strDetalle(0),strDetalle(1),strDetalle(2),strDetalle(3),strDetalle(4),strDetalle(5),strDetalle(6),strDetalle(7),strDetalle(8),strDetalle(9),strDetalle(10),strDetalle(11))
            total = total + 1
        end if
	wend
    if (not flagHayResultado) then archRep.writeline(string(105," ") & "No se encontraron resultados" & string(105," "))
    archRep.writeline(string(240,"*"))
    archRep.writeline("TOTAL : " & total)
    archRep.writeline(string(240," "))
    archRep.writeline(string(240," "))
	archTem.Close
End Function
'------------------------------------------------------------------------------------------------------------------------
Function grabarAnalisisBiotecnologia(pPath, pTransporte)
    Dim txtLine,archTem,flagHayResultado, total
    Call drawTitleReport(pTransporte,"BIOTECNOLOGIA")
    Set archTem = fs.OpenTextFile(pPath, 1)
    logMig.info(string(20,"-")&" GENERANDO ANALISIS DE BIOTECNOLOGIA "&string(20,"-"))
    flagHayResultado = false
    total = 0
    while (not archTem.AtEndOfStream)        
	    txtLine = archTem.ReadLine()
        txtLine = split(txtLine, KEY_TOKEN)
        'if ((CInt(txtLine(0)) = pTransporte)and(CInt(txtLine(1)) = ACEPTACION_REBAJA_CONVENIDA)and(tieneBiotecnologia(txtLine(2), g_Pto))) then        
        'Call saveInfoReport(txtLine(2), tieneBiotecnologia(txtLine(2), g_Pto),"","","","","","","","","","")
        if ((CInt(txtLine(0)) = pTransporte) and (tieneBiotecnologia(txtLine(2), g_Pto))) then
            flagHayResultado = true
            strDetalle = split(txtLine(3), FIELD_TOKEN)
            Call saveInfoReport(strDetalle(0),strDetalle(1),strDetalle(2),strDetalle(3),strDetalle(4),strDetalle(5),strDetalle(6),strDetalle(7),strDetalle(8),strDetalle(9),strDetalle(10),strDetalle(11))
            total = total + 1
        end if
	wend
    if (not flagHayResultado) then archRep.writeline(string(105," ") & "No se encontraron resultados" & string(105," "))
    archRep.writeline(string(240,"*"))
    archRep.writeline("TOTAL : " & total)
    archRep.writeline(string(240," "))
    archRep.writeline(string(240," "))
	archTem.Close
End Function
'------------------------------------------------------------------------------------------------------------------------
Function grabarAnalisisTotal(pPath, pTransporte)
    Dim txtLine,archTem,flagHayResultado, total
    Call drawTitleReport(pTransporte,"TOTALES")
    Set archTem = fs.OpenTextFile(pPath, 1)
    logMig.info(string(20,"-") & " GENERANDO ANALISIS TOTALES " & string(20,"-"))
    flagHayResultado = false
    total = 0
    while (not archTem.AtEndOfStream)        
	    txtLine = archTem.ReadLine()
        txtLine = split(txtLine, KEY_TOKEN)
        if ( (CInt(txtLine(0)) = pTransporte) and (CInt(txtLine(1)) <> ACEPTACION_REBAJA_CONVENIDA)) then
            flagHayResultado = true
            strDetalle = split(txtLine(3), FIELD_TOKEN)
            Call saveInfoReport(strDetalle(0),strDetalle(1),strDetalle(2),strDetalle(3),strDetalle(4),strDetalle(5),strDetalle(6),strDetalle(7),strDetalle(8),strDetalle(9),strDetalle(10),strDetalle(11))
            total = total + 1
        end if
	wend
    if (not flagHayResultado) then archRep.writeline(string(105," ") & "No se encontraron resultados" & string(105," "))
    archRep.writeline(string(240,"*"))
    archRep.writeline("TOTAL : " & total)
    archRep.writeline(string(240," "))
    archRep.writeline(string(240," "))
	archTem.Close
End Function
'------------------------------------------------------------------------------------------------------------------------
'   Inicio de la Pagina
'------------------------------------------------------------------------------------------------------------------------
Dim archRep,arrLengthTitleReporte

arrLengthTitleReporte = Array(15,15,20,25,20,25,25,25,25,15,20,15,20)

Set fs = Server.CreateObject("Scripting.FileSystemObject")
if (fs.FileExists(pathTempExp)) then
    auxProductoProteina = getValueParametro(CAMARA_PARAMETER_PRODPROTEINA,g_Pto)    
    Set archRep = fs.OpenTextFile(strNamePathReporte, 2, true)
    archRep.writeline( string(240,"*"))
    archRep.writeline( string(5,"*") & string(100," ") & " REPORTE DE DATOS PARA C�MARA " & string(100," ") & string(5,"*"))
    archRep.writeline( string(240,"*"))
    Call grabarAnalisisTotal(pathTempExp, TIPO_TRANSPORTE_CAMION)
    Call grabarAnalisisTotal(pathTempExp, TIPO_TRANSPORTE_VAGON)
    Call grabarAnalisisBiotecnologia(pathTempExp, TIPO_TRANSPORTE_CAMION)
    Call grabarAnalisisBiotecnologia(pathTempExp, TIPO_TRANSPORTE_VAGON)
    Call grabarAnalisisProteina(pathTempExp, TIPO_TRANSPORTE_CAMION)
    Call grabarAnalisisProteina(pathTempExp, TIPO_TRANSPORTE_VAGON)
	fs.DeleteFile(pathTempExp)
end if

%>
<HTML>
    <HEAD>
        <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0" />
        <script type="text/javascript">
	        parent.generateReport_callback();
        </script>
    </HEAD>
    <BODY>
        <P>&nbsp;</P>
    </BODY>
</HTML>
