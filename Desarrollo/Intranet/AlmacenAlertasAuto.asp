<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientossql.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosPM.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->

<%

	const DIAS_AVISO_SUPERVISOR = 2
	Dim dsAlmacen,emailMsg,rs0,rs1,rs2,oDiccCantidadesPedidas,stockEnPics
	Set oDiccCantidadesPedidas  = createObject("Scripting.Dictionary")
'-------------------------------------------------------------------------------------------------'
Function getCantidadPedida(pIdArticulo)
	Dim rtrn

	rtrn = 0
	if (oDiccCantidadesPedidas.Exists(cdbl(pIdArticulo))) then
		rtrn = oDiccCantidadesPedidas.Item(cdbl(pIdArticulo))
	end if

	getCantidadPedida = rtrn
End Function
'-------------------------------------------------------------------------------------------------'

	Function controlarFechasArticulo(byref pMsg, pFecha,pStockActual,pStockMinimo,pIdArticulo,pDsArticulo,pAlmacen)
		if (len(pFecha)>8) then pFecha = left(pFecha,8)
		if (GF_DTEDIFF(pFecha,session("mmtosistema"),"D") > DIAS_AVISO_SUPERVISOR) then
				'guardo el mensaje para enviarlo por email'
				pMsg = pMsg & pIdArticulo & " - " & pDsArticulo & " | Stock Actual: " & pStockActual & " - Pedidos: " & getCantidadPedida(pIdArticulo) & " - Stock Minimo: " & pStockMinimo & vbnewline&vbnewline
		end if
	End Function
'-------------------------------------------------------------------------------------------------'
Function iniciar()
	Dim lastDivision,finalizar
	strSQL = "select * from TBLALMACENES"
	call executeQueryDb(DBSITE_SQL_INTRA, rs0, "OPEN", strSQL)
	
	GP_ConfigurarMomentos
	emailMsg = ""
	
	lastDivision = 0
	while not rs0.EoF
		dsAlmacen = rs0("dsAlmacen")
		emailMsg = ""

		'Esto se hace para que no procese 2 veces la misma division'
		actualDivision = getDivisionAlmacen(rs0("idalmacen"))

		if (lastDivision <> actualDivision) then 
			lastDivision = actualDivision
			Set oDiccCantidadesPedidas = cargarCantidadesPedidas(rs0("idalmacen"),lastDivision)
		end if

		'Obtengo los articulos bajos de stock'
		strSQL = 		  "SELECT a.*,( existencia + sobrante ) stock,art.dsarticulo "
		strSQL = strSQL & "FROM   tblarticulosdatos a"
		strSQL = strSQL & " INNER JOIN tblarticulos art on a.idarticulo = art.idarticulo"
		strSQL = strSQL & " WHERE  idalmacen = " & rs0("idalmacen")
		strSQL = strSQL & " AND ( existencia + sobrante ) < stockminimo "
		strSQL = strSQL & "AND    stockminimo <> 0 order by idarticulo"
		call executeQueryDb(DBSITE_SQL_INTRA, rs1, "OPEN", strSQL)
		
		while not rs1.EoF

			stockEnPics = getCantidadPedida(rs1("idArticulo"))

			if (stockEnPics+cdbl(rs1("stock")) < cdbl(rs1("stockminimo")) ) then
				'el articulo tiene insuficiencia de stock aun contando los pics hechos'
				strSQL = 		  "SELECT   fecha, "
				strSQL = strSQL & "         cantidad "
				strSQL = strSQL & "FROM     tblvalescabecera c "
				strSQL = strSQL & "         INNER JOIN tblvalesdetalle d "
				strSQL = strSQL & "         ON       d.idvale = c.idvale "
				strSQL = strSQL & "WHERE    d.idarticulo      = " & rs1("idArticulo")
				strSQL = strSQL & " AND      idalmacen         = " & rs1("idalmacen")
				strSQL = strSQL & " ORDER BY fecha DESC"
				call executeQueryDb(DBSITE_SQL_INTRA, rs2, "OPEN", strSQL)
				
				stockArticulo = cdbl(rs1("stock"))
				if (rs2.EoF) then
					strSQL = "select momento from tblarticulos where idarticulo = " & rs1("idArticulo")
					call executeQueryDb(DBSITE_SQL_INTRA, rs3, "OPEN", strSQL)
					Call controlarFechasArticulo(emailMsg,rs3("momento"),stockArticulo,rs1("stockminimo"),rs1("idarticulo"),rs1("dsarticulo"),dsAlmacen)
				end if
				
				finalizar = false
				while not rs2.EoF and finalizar = true
					if ((stockArticulo + cdbl(rs2("cantidad"))) >= cdbl(rs1("stockminimo"))) then
						'en esta fecha se realizo el vale que produjo la insuficiencia'
						if( idLastArticulo <> rs1("idarticulo")) then
							idLastArticulo = rs1("idarticulo")
							
							Call controlarFechasArticulo(emailMsg,rs2("fecha"),stockArticulo,rs1("stockminimo"),rs1("idarticulo"),rs1("dsarticulo"),dsAlmacen)
							finalizar = true
						end if
					else
						stockArticulo = stockArticulo + cdbl(rs2("cantidad"))
					end if

					rs2.MoveNext
				wend
			else
				'El articulo tiene actualmente faltante de stock pero se solucionara cuando los pics que contienen dicho articulo sean antregados'				
			end if
			rs1.MoveNext
		wend

		'busco a quienes le corresponde recibir las alertas'
		response.write emailMsg
		response.end
		if (emailMsg <> "") then
			strSQL = "select * from TBLMAILSALERTASALMACENES where idalmacen = " &rs0("idalmacen")
			call executeQueryDb(DBSITE_SQL_INTRA, rs4, "OPEN", strSQL)
			auxEmail = ""
			while not rs4.EoF
				if (trim(cstr(rs4("email"))) <> "") then auxEmail = auxEmail & rs4("email") & ";"
				rs4.MoveNext
			wend
			Call GP_ENVIAR_MAIL("Falta Stock almacen "&dsAlmacen,emailMsg,"scalisij@toepfer.com",auxEmail)
			
		end if

		rs0.MoveNext
	wend
End Function

Call iniciar()
%>