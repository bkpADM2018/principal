<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'-----------------------------------------------------------------------------------------
Function ActualizarControlStockDet(IdControl,idArticulo,StockFis)
	Dim strSQL,rs,oConn
	strSQL = " UPDATE TBLCSTKDETALLE SET STOCKFISICO = "&StockFis
	strSQL = strSQL & "	WHERE IDCONTROL = "& IdControl &" AND IDARTICULO = " &idArticulo
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End function 
'----------------------------------------------------------------------------------
Function ActualizarIdValeCab(pIdControl,pidVale)
	Dim strSQL,rs,oConn
	strSQL = " UPDATE TBLCSTKCABECERA SET IDRESULTADO = "&pidVale
	strSQL = strSQL & "	WHERE IDCONTROL = "& pIdControl
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
End function 
'-----------------------------------------------------------------------------------------
'Funcion que se asegura que al grabar el resultado de un control de stock no se haya grabado un resultado antes. Esto permite evitar que se repitan los vales.
Function puedeGrabarResultado(idControl)
    
    Dim strSQL, rs, ret
    
    ret= false
    strSQL="Select * from TBLCSTKCABECERA where IDCONTROL= " & idControl & " and IDRESULTADO=" & CTST_RESULTADO_PENDIENTE    
    Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
    if (not rs.eof) then 
        Call ActualizarIdValeCab(idControl, CTST_SIN_VALE)
        ret = true
    end if
    puedeGrabarResultado = ret
    
End Function
'******************************************************************************************
'***************************** COMIENZO DE LA PAGINA **************************************
'******************************************************************************************
Dim idVale, index,pIdArticulo,idControl_new,lastArticulos,vStockFis,cdResponsable, rsSS, idart,cantidadArticulos,pChkStock,precioMinimo
dim myIdAlmacen,v_index,esControlNuevo, flagGrabarVale,pSeleccion,vStockSist,i,callMsj,accion,IdControl,precioMaximo
'estos parametros son recibidos sin importar el tipo
myIdAlmacen = GF_PARAMETROS7("IdAlmacen",0,6)
pSeleccion = GF_PARAMETROS7("tipoReporte","",6)
lastArticulos = GF_PARAMETROS7("cantArticulos", 0, 6)
cdResponsable = GF_PARAMETROS7("CtSt_cdResponsable","",6)
'este parametro se recibe si es una carga de resultados
idControl_new = GF_PARAMETROS7("IdControl",0,6)
accion = GF_PARAMETROS7("accion","",6)
'se utiliza el parametro accion para poder grabar el Vale
cantidadArticulos = GF_PARAMETROS7("cantArticulostxt",0,6)
precioMinimo = GF_PARAMETROS7("precioMinimo",0,6)
precioMaximo = GF_PARAMETROS7("precioMaximo",0,6)
pChkStock = GF_PARAMETROS7("chkStock","",6)
if(pChkStock = "")then pChkStock = CTST_ARTICULO_SIN_STOCK

VS_idAlmacen = myIdAlmacen
IdControl = 0
callMsj = RESPUESTA_OK

Set vStockSist = Server.CreateObject("Scripting.Dictionary")
Set vStockFis = Server.CreateObject("Scripting.Dictionary")

esCargaResultado = false
if (idControl_new > 0) then esCargaResultado = true
	'/*************************************************************************************/
	'/**************************** RECIBE LOS PARAMETROS CON INDICE ***********************/
	'/*************************************************************************************/
	v_index = 0
	For i = 0 To lastArticulos-1			
		pIdArticulo = GF_PARAMETROS7("item" & i,"",6)
		if(pIdArticulo <> "")then 
			'Se agrega el artículo al diccionario y de paso se eliminan los duplicados ya que el artículo solo estará una vez en el diccionario.
			if (not vStockFis.Exists(pIdArticulo)) then vStockFis.Add pIdArticulo, 0
			if (esCargaResultado) then
				'Si es mayor quiere decir que se trata de una modificacion o una carga de resultados		
				vStockFis(pIdArticulo) = GF_PARAMETROS7("saldo" & i,"",6) 
			end if
			v_index = v_index + 1
		end if
	next
	'/*************************************************************************************/
	'/************************************ CONTROLA ***************************************/
	'/*************************************************************************************/
	if(v_index > 0)then
		For each idart in vStockFis
			if(not controlarArticulo(idart))then callMsj = ERROR_GRABAR_CTST
		next
	else
		callMsj = CANTIDAD_ATICULOS_CTST
	end if	
	'/*************************************************************************************/
	'/************************************** GRABA ****************************************/
	'/*************************************************************************************/
	if (esCargaResultado)then		
		'-------------------------CARGA DE RESULTADOS ----------------------
		'Controlo primero si ya no se grabó un resultado.
		if (puedeGrabarResultado(idControl_new)) then
		    'Tomo los stocks de sistema registrados al momento de imprimir por última vez el reporte.
		    Set rsSS = leerDetallesCtSt(idControl_new, myIdAlmacen, esCargaResultado,pSeleccion)		
		    While (not rsSS.eof) 
		        myKey = Trim(rsSS("IDARTICULO"))
			    myVal = rsSS("STOCKSISTEMA")
			    if (not vStockSist.Exists(rsSS("IDARTICULO"))) then vStockSist.Add myKey, myVal
			    rsSS.MoveNext()
		    Wend
		    'Se asume que ya no hay errores en los artículos dado que los mismos no pueden ser modificados.
		    'Graba una actualizacion del Control de Stock, se produce cuando se carga el Resultado 		
		    For each idart in vStockFis
			    Call ActualizarControlStockDet(idControl_new, idart,vStockFis(idart))				
			    if(CInt(vStockSist(idart)) <> CInt(vStockFis(idart)))then flagGrabarVale = true					
		    next
		     'En caso de que haya diferencia de los stock se genera el Vale de Ajuste			 		 
		     if(flagGrabarVale)then			 
			     Call initHeaderVale(idVale)
			     VS_cdSolicitante = session("Usuario")
			     VS_dsSolicitante = getUserDescription(VS_cdSolicitante)			 
			     VS_FechaSolicitud = GF_FN2DTE(Left(session("MmtoDato"),8))
			     VS_cdVale = CODIGO_VS_AJUSTE_STOCK
			     VS_ArticuloActual = 0			
			     Call grabarHeaderVale(idVale,0)			
			     Call grabarComentarioVale(idVale, "Control de Stock:" & idControl_new)			
			     For each idart in vStockFis			 
				    'Los stocks a considerar son los del sistema (no el actual sino el que habia a la ultima impresion del reporte) y el fisico contado. 
				    VS_idArticulo = idart
				    VS_saldo = vStockFis(idart)
				    VS_cantidad = vStockSist(idart)
				    if (VS_saldo >= 0) then
					    if(CInt(VS_saldo) <> CInt(VS_cantidad))then					
						     call grabarValeDetalle(idVale, 0)
						     call actualizarStock()
					     end if
				    end if
			     Next
			     Call grabarPreciosVigentesPorArticulo(idVale)
			     VS_cdSolicitante = cdResponsable
			     Call grabarFirmasValeAJS(idVale)
	         else
	            idVale = CTST_SIN_VALE
		     end if		
		     Call ActualizarIdValeCab(idControl_new,idVale)
		end if
		'------------------------------------------------------------------
	else 
		'-------------------------NUEVO CONTROL STOCK----------------------		
		'v_index : me indica la cantidad de registros que tiene el vector		
		if((v_index > 0)and(callMsj = RESPUESTA_OK))then
			idControl_new = AgregarControlStockCab(cdResponsable,myIdAlmacen,pSeleccion,pChkStock,precioMinimo,cantidadArticulos,precioMaximo)
			IdControl = idControl_new
			For each idart in vStockFis
				Call AgregarControlStockDet(idControl_new, idart, 0, 0, 0)
			next
		end if		
		'------------------------------------------------------------------	
	end if	
	if (callMsj <> RESPUESTA_OK) then callMsj = callMsj & "-" & errMessage(callMsj)
	'/*************************************************************************************/
	'/*************************************************************************************/
	'/*************************************************************************************/
%> 
<HTML>
<HEAD>
<script type="text/javascript">		
	parent.resultadoCarga_callback('<% =callMsj %>', <%=IdControl%>);	
</script>
</HEAD>
<BODY>
<P>&nbsp;</P>
</BODY>
</HTML>