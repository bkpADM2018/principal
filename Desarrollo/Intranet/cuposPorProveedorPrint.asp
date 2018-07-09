<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosBoletos.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<%

'-----------------------------------------------------------------------------------------------
         locale=session.lcid
		 session.lcid=2057	'Formato dd/mm/aaaa

dim rs, conn, strSQL, oPDF
dim idProveedor, nroContrato, fecha, periodo
Dim currentY, nroPagina
Dim cupos
Dim fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor
Dim esContratoNuevo
dim strPathAttachment, strCuerpoMail, strDestinosMail
dim chkEnviados, usr, mostrarConfirmacion,g_strPuerto

SEPARATION = 11
MARGIN = 0
PAGE_HEIGHT_SIZE = 800
PAGE_TOP_INIT = 82
SIN_MAIL = 1
nroPagina = 1

idProveedor = GF_Parametros7("id", 0, 6)
nroContrato = GF_Parametros7("nroContrato", 0, 6)
cdProducto = GF_Parametros7("cdProducto", 0, 6)
cdSucursal = GF_Parametros7("cdSucursal", 0, 6)
cdOperacion = GF_Parametros7("cdOperacion", 0, 6)
anioCosecha = GF_Parametros7("anioCosecha", 0, 6)
fecha = GF_Parametros7("fecha", 0, 6)
pdf_accion = GF_Parametros7("pdf_accion", 0, 6)
periodo = GF_Parametros7("periodo", 0, 6)
if periodo = 0 then periodo = 9

usr = session("Usuario")

'filtros
fltrCdProducto = GF_PARAMETROS7("fltrCdProducto","",6)
fltrCdSucursal = GF_PARAMETROS7("fltrCdSucursal","",6)
fltrCdOperacion = GF_PARAMETROS7("fltrCdOperacion","",6)
fltrNroContrato = GF_PARAMETROS7("fltrNroContrato","",6)
fltrAnioCosecha = GF_PARAMETROS7("fltrAnioCosecha","",6)
fltrPuerto = GF_PARAMETROS7("fltrPuerto",0,6)
fltrCorredor = GF_PARAMETROS7("fltrCorredor",0,6)
fltrVendedor = GF_PARAMETROS7("fltrVendedor",0,6)

chkEnviados = GF_PARAMETROS7("chkEnviados",0,6)
mostrarConfirmacion = GF_PARAMETROS7("mostrarConfirmacion","",6)'si se procese a enviar los mails en batch desde la pagina cuposPorProveedor.asp, viene con valor 'N'

'Response.Write "pdf_accion es " & pdf_accion

strDestinosMail =  getStringMailsProveedor(idProveedor) 
if pdf_accion = PDF_FILE_MODE and strDestinosMail = "" and mostrarConfirmacion = "N" then
	'si se envian los mail en batch desde cuposPorProveedor.asp y el proveedor no tiene mail, corta y devuelve por AJAX el aviso
	 Response.Write SIN_MAIL
	 Response.end
end if

filename = getFilename()
Set oPDF = GF_createPDF(Server.MapPath("temp/" & filename))
strPathAttachment = Server.mapPath("temp/" & filename)
Call GF_setPDFMODE(pdf_accion)
call armadoPDF()
Call GF_closePDF(oPDF)
if pdf_accion = PDF_FILE_MODE then
'enviar mail con attachment de pdf generado
	strCuerpoMail = GF_Parametros7("descripcionMail", "", 6)
	if strCuerpoMail = "" then strCuerpoMail = "Ver cupos asignados en el archivo adjunto"
	strDestinosMail = getUserMail(usr) & ";" & strDestinosMail '& ";scalisij@toepfer.com"
	strCuerpoMail = strCuerpoMail
	Call GP_ENVIAR_MAIL_ATTACHMENT("Alfred Toepfer S.R.L - Cupos Asignados", strCuerpoMail,determinarSender(cdSucursal), strDestinosMail, strPathAttachment)
	'borrar el archivo pdf creado
	set fs = Server.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPathAttachment) Then  call fs.deleteFile(strPathAttachment, true)
    if mostrarConfirmacion <> "N" then
    %>
      <html>
		<head>
			<script type="text/javascript" src="scripts/iwin.js"></script>
			<script type="text/javascript">
			var refPopUpEnviarMail;

			function bodyOnLoad() {
				refPopUpEnviarMail = startIWin('popupEnviarMail');
				refPopUpEnviarMail.hide();

			}
			</script>
		</head>
		<body onLoad="bodyOnLoad()">
		</body>
	</html>
    
<%
	end if
end if
 session.lcid=locale 'volver al formato de fecha original del servidor
 
function determinarSender(cdSucursal)
'la funcion determina el sender para el mail. Si el codigo de operacion es 1, entonces sender es Trading de Rosario. Otherwise es Mercaderias Buenos Aires
'Response.Write "codigo sucursal " & cdSucursal
'Response.end
	if CInt(cdSucursal) = 1 then
		determinarSender = SENDER_CUPOS_RO
	else
		determinarSender = SENDER_CUPOS_BA
	end if
end function
'-----------------------------------------------------------------------------------------------
function getContratos (idProveedor, fecha, periodo, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor)
	Dim strSQL, rs, conn
	Dim fechaInicio, fechaFin
	fechaInicio = fecha
	fechaFin = GF_DTE2FN(dateadd("d",periodo,GF_FN2DTE(fecha)))

	strSQL = "select cuncto as nroContrato, cucpro as cdProducto, cucsuc as cdSucursal, "
	strSQL = strSQL & " cucope as cdOperacion, cuacos as anioCosecha,cucdes as nropuerto from merfl.mer517f1 "
	strSQL = strSQL & "where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and (cuccor = " & idProveedor & " or cucven = " & idProveedor & ") "
	
	'aplicarFiltros
	if fltrCdProducto <> "" then strSQL = strSQL & " and cucpro = " & fltrCdProducto
	if fltrCdSucursal <> "" then strSQL = strSQL & " and cucsuc = " & fltrCdSucursal
	if fltrCdOperacion <> "" then strSQL = strSQL & " and cucope = " & fltrCdOperacion
	if fltrNroContrato <> "" then strSQL = strSQL & " and cuncto = " & fltrNroContrato
	if fltrAnioCosecha <> "" then strSQL = strSQL & " and cuacos = " & fltrAnioCosecha
	if fltrPuerto <> 0 then strSQL = strSQL & " and cucdes = " & fltrPuerto	
	if fltrCorredor <> 0 then strSQL = strSQL & " and (cuccor = " & fltrCorredor & " or (cucven = " & fltrCorredor & " and cuccor = " & SIN_CORREDOR & "))"
	if fltrVendedor <> 0 then strSQL = strSQL & " and cucven = " & fltrVendedor
	
	strSQL = strSQL & " group by cuncto, cucpro, cucsuc, cucope, cuacos,cucdes"
	strSQL = strSQL & " order by cuncto, cucope "
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	Set getContratos = rs
end function
'-----------------------------------------------------------------------------------------------
function getContratoVendedor (cdProducto, cdSucursal, cdOperacion, nroContrato, anioCosecha)
	Dim strSQL, rs, conn

	strSQL = "select concr1 as numContratoVendedor from merfl.mer311f1 "
	strSQL = strSQL & "where nctor1 = " & nroContrato 
	strSQL = strSQL & " and cpror1=" & cdProducto &  " and csucr1 =" & cdSucursal & " and coper1=" & cdOperacion & " and acosr1=" & anioCosecha
	
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	if not rs.eof then
		getContratoVendedor = rs("numContratoVendedor")
	else
		getContratoVendedor = ""
	end if	
end function
'------------------------------------------------------------------------------------------
sub grabarCuposInformados(cupos, ultimoNumeroSecuencia)
	Dim strSQL, rs, conn
	Dim sqlUpdate, sqlInsert

	strSQL = "select * from toepferdb.tblcuposinformados "
	strSQL = strSQL & " where codigocupo = " & cupos("cdCupo")
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	if rs.eof then
		'grabar
		sqlInsert = "insert into toepferdb.tblcuposinformados values("
		sqlInsert = sqlInsert & cupos("cdProducto") & ", " & cupos("cdSucursal") & ", " & cupos("cdOperacion")
		sqlInsert = sqlInsert & ", " & cupos("nroContrato") & ", " & cupos("anioCosecha") & ", " & cupos("fechaCupo") & ", " & cupos("cdPuerto")
		sqlInsert = sqlInsert & ", " & cupos("nroCamiones") & ", " & cupos("cdCupo") & ", " & ultimoNumeroSecuencia & ")"	
		Call GF_BD_AS400_2(rs, conn, "EXEC", sqlInsert)
	else
		'actualizar
		sqlUpdate = "update toepferdb.tblcuposinformados set cupos = (cupos + " & cupos("nroCamiones") & ") "
		sqlUpdate = sqlUpdate & ", ultimasecuencia = " & ultimoNumeroSecuencia
		sqlUpdate = sqlUpdate & " where codigocupo = " & cupos("cdCupo")
		Call GF_BD_AS400_2(rs, conn, "EXEC", sqlUpdate)
	end if		
end sub
'------------------------------------------------------------------------------------------
Function armadoPDF()
dim totalReg, i, cambioPagina
	cambioPagina = false
	dibujarTitulo(GF_TRADUCIR("CUPOS ASIGNADOS"))
	currentY = 80
	'dibujar Proveedor
	Call GF_squareBox(oPDF, 20	, currentY, 550	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 22, currentY+2 , "Proveedor: " & getDescripcionProveedor(idProveedor), 550, PDF_ALIGN_LEFT)
	avanzar()
	avanzar()
	
	if Cdbl(nroContrato) = 0 then
	'dibujar varios contratos
		set contratos = getContratos(idProveedor, fecha, periodo, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor)
		dim contadorContratos ' esta variable es necesaria para determinar el codigo de operacion del primer contrato, para saber el sender
		if not contratos.eof then
			contadorContratos = 0
			while not contratos.eof
				if contadorContratos = 0 then cdSucursal = contratos("cdSucursal")
				esContratoNuevo = true
				Call dibujarContrato (contratos("nroContrato"),contratos("cdProducto"),contratos("cdSucursal"),contratos("cdOperacion"),contratos("anioCosecha"), fecha)
				avanzar()
				avanzar()
				contadorContratos = contadorContratos + 1
				contratos.movenext
			wend
		end if
	else
		Call dibujarContrato (nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, fecha)
	end if	
End function
'------------------------------------------------------------------------------------------
function getStringSQLCupos(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)
dim fechaInicio, fechaFin
'arma la sentencia sql para traer los cupos de un numero de contrato dado.
'si la variable flag chkEnviados esta activa, entonces compara contra la tabla toepferdb.tblcuposinformados	
	fechaInicio = fecha
	fechaFin = GF_DTE2FN(dateadd("d",periodo,GF_FN2DTE(fecha)))
    strSQL = "select * from ("
    strSQL = strSQL & "select cucodi as cdCupo, cufccp as fechaCupo, cuzinf as cdDestino, cucdes as cdPuerto, cuncto as nroContrato, cucpro as cdProducto, "
	strSQL = strSQL & " cucsuc as cdSucursal, cucope as cdOperacion, cuacos as anioCosecha, MCPDRJ PProd, "	 
	if (chkEnviados = MOSTRAR_NO_ENVIADOS) then strSQL= strSQL & " sum (cucccp - case when cupos is null then 0 else cupos end) "
	if (chkEnviados = MOSTRAR_PUBLICADOS) then strSQL= strSQL & " sum (case when cupos is null then 0 else cupos end) "	
	if (chkEnviados = MOSTRAR_ENVIADOS) then strSQL = strSQL & " sum(cucccp) "	
	strSQL = strSQL & " as nroCamiones, cuccor as cdCorredor, cucven as cdVendedor "
	strSQL = strSQL & " from merfl.mer517f1 " 
	strSQL = strSQL & " left join MERFL.MER311FJ on CPRORJ=CUCPRO and CSUCRJ=CUCSUC and COPERJ=CUCOPE and NCTORJ=CUNCTO and ACOSRJ=CUACOS "
	if (chkEnviados = MOSTRAR_NO_ENVIADOS) or (chkEnviados = MOSTRAR_PUBLICADOS) then strSQL = strSQL & " left join toepferdb.tblcuposinformados  on CODIGOCUPO=CUCODI "
	strSQL = strSQL & "where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and cuncto = " & nroContrato 
	strSQL = strSQL & " and cucpro=" & cdProducto &  " and cucsuc =" & cdSucursal & " and cucope=" & cdOperacion & " and cuacos=" & anioCosecha	
	strSQL = strSQL & " group by cucodi, cufccp, cuzinf, cucdes, cuncto, cucpro, cucsuc, cucope, cuacos, cuccor, cucven, MCPDRJ) " 
	strSQL = strSQL & " as tablaGral "	
	strSQL = strSQL & " where nroCamiones <> 0 order by fechaCupo"
	getStringSQLCupos = strSQL
end function
'------------------------------------------------------------------------------------------
Sub dibujarContrato (nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)
	Dim strSQL, conn
	Dim fechaInicio, fechaFin, isMsjBahia ,MsjBahia
	dim iContColor, color
	Dim cantLineasAVolver 'esta variable sirve para volver (cuando se dibujan codigos de cupos) a primer linea
							 'en caso de que haya mas que una condicion especial.
	iContColor = 0
	isMsjBahia = false
	strSQL = getStringSQLCupos(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)
	
	Call GF_BD_AS400_2(cupos, conn, "OPEN", strSQL)
	if not cupos.eof then
	Call dibujarCabeceraContrato (cupos)
	avanzar()
	dibujarCorredorVendedor(cupos)
	avanzar()


		Call dibujarTituloCupos()
		avanzar()
		while not cupos.eof
			'se dibujan los cupos
			iContColor = iContColor + 1
			if (iContColor mod 2)  then 
				color = "#CECEF6"		
			else				
				color = "#FFFFFF"	
			end if	
			Call GF_squareBox(oPDF, 20, currentY, 550, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
			Call GF_writeTextAlign(oPDF,20, currentY, GF_FN2DTE(cupos("fechaCupo")), 100, PDF_ALIGN_CENTER)	
			Call GF_writeTextAlign(oPDF,120, currentY, cupos("nroCamiones"), 84, PDF_ALIGN_CENTER)
			Call GF_writeTextAlign(oPDF,204, currentY, getDsPort(CInt(cupos("cdPuerto"))), 183, PDF_ALIGN_CENTER)	
			if (((CInt(cupos("CDPUERTO")) = PUERTO_PIEDRABUENA)or(CInt(cupos("CDPUERTO")) = MUELLE_BAHIABLANCA))and(CInt(cupos("cdOperacion"))= OPERACION_PRESTAMO_DEVOLUCION)) then
                Call GF_writeTextAlign(oPDF,387, currentY, GF_TRADUCIR("Los codigo se obtendran luego de nominar"), 183, PDF_ALIGN_LEFT)
            else
                cantLineasAVolver = 0
			    Call dibujarCodigosCupos(cupos("cdCupo"), cupos("cdProducto"), CInt(cupos("cdPuerto")), cantLineasAVolver, color,ultimoNumeroSecuencia)
            end if
			if (pdf_accion = PDF_FILE_MODE) and (chkEnviados = MOSTRAR_NO_ENVIADOS) then
			'grabar en la tabla historica toepferdb.tblcuposinformados los cupos que se van a mandar para esta fecha
				Call grabarCuposInformados(cupos, ultimoNumeroSecuencia)
			end if
            if ((CInt(cupos("cdPuerto")) = PUERTO_PIEDRABUENA) OR (CInt(cupos("cdPuerto")) = MUELLE_BAHIABLANCA)) then isMsjBahia = true
			cupos.movenext
			avanzar()
		wend		
        if (isMsjBahia) then
            if (CInt(cdOperacion)= OPERACION_PRESTAMO_DEVOLUCION) then Call dibujarCuposNominados(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)
            MsjBahia="Los cupos de lunes a viernes tendrán apertura el día anterior al cupo a las 15 hs y como cierre el día del mismo a las 17 hs."
            Call GF_squareBox(oPDF, 20, currentY, 550, 10, 0, "#FFF000", "#FFF000", 1, PDF_SQUARE_NORMAL)
		    Call GF_writeTextAlign(oPDF,20, currentY, MsjBahia , 387, PDF_ALIGN_LEFT)
            currentY = currentY + SEPARATION
            MsjBahia="Para los días sábados y domingos será el día anterior a las 15 hs, y el cierre el día del cupo a las 10 hs."
            Call GF_squareBox(oPDF, 20, currentY, 550, 10, 0, "#FFF000", "#FFF000", 1, PDF_SQUARE_NORMAL)
            Call GF_writeTextAlign(oPDF,20, currentY, MsjBahia , 387, PDF_ALIGN_LEFT)	
        end if
	end if
End sub
'------------------------------------------------------------------------------------------	
' Dibuja los cupos que se tienen nominados para un determinado contrato entre un rango de fechas
' NOTA: 
'       Esta funcion trabaja solamente si el puerto es Bahia Blanca y la operacion es devolucion/prestamo (04),
'       al ver la informacion en la tabla que muestra el reporte por cada contrato se tomo la desicion de que esta funcion
'       trabaje correctamente con este puerto solamente. 
Function dibujarCuposNominados(p_nroContrato, p_cdProducto, p_cdSucursal, p_cdOperacion, p_anioCosecha, p_fecha)
    Dim rsNom,letraCupo,auxCodigoCupo,auxCodigoDesde,auxCodigoHasta
    
    Set rsNom = obtenerNominacionesCupos(p_nroContrato, p_cdProducto, p_cdSucursal, p_cdOperacion, p_anioCosecha, p_fecha)
    if (not rsNom.Eof) then
        Call GF_squareBox(oPDF, 20, currentY, 550, 10, 0, "#FF0000", "#FF0000", 1, PDF_SQUARE_NORMAL)
        Call GF_verticalLine(oPDF, 20, currentY, 10)
        Call GF_verticalLine(oPDF, 570, currentY, 10)
        Call GF_horizontalLine(oPDF, 20, currentY, 550)
        Call GF_setFontColor("FFFFFF")
        Call GF_writeTextAlign(oPDF,20, currentY +2, "Nominaciones realizadas" , 550, PDF_ALIGN_CENTER)
        Call GF_setFontColor("000000")
        Call GF_squareBox(oPDF, 20, currentY, 60 , 10, 0, "#CECEF6", "#000000", 1, PDF_SQUARE_NORMAL)
        Call GF_squareBox(oPDF, 80, currentY, 110 , 10, 0, "#CECEF6", "#000000", 1, PDF_SQUARE_NORMAL)
        Call GF_squareBox(oPDF, 190, currentY, 190, 10, 0, "#CECEF6", "#000000", 1, PDF_SQUARE_NORMAL)
        Call GF_squareBox(oPDF, 380, currentY, 190, 10, 0, "#CECEF6", "#000000", 1, PDF_SQUARE_NORMAL)
        Call GF_writeTextAlign(oPDF,20, currentY +2, "Fecha" , 60, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,80, currentY +2, "Cupos Asignados" , 110, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,190, currentY +2, "Corredor" , 190, PDF_ALIGN_CENTER)
        Call GF_writeTextAlign(oPDF,380, currentY +2, "Vendedor" , 190, PDF_ALIGN_CENTER)
        avanzar()
        while (not rsNom.Eof)
            g_strPuerto = getDsPuertoByNro(rsNom("PUERTO"))
            'Obtengo el codigo de cupo para Bahia blanca
            auxCodigoDesde = GF_nDigits(rsNom("CODIGODESDE"),8)
            auxCodigoDesde = LEFT(Trim(rsNom("DSPRODUCTO")),1) & auxCodigoDesde
            auxCodigoHasta = GF_nDigits(rsNom("CODIGOHASTA"),8)
			auxCodigoHasta = LEFT(Trim(rsNom("DSPRODUCTO")),1) & auxCodigoHasta
            Call GF_squareBox(oPDF, 20, currentY, 60, 10, 0, "#E0F8F7", "#E0F8F7", 1, PDF_SQUARE_NORMAL)
            Call GF_squareBox(oPDF, 80, currentY, 110 , 10, 0, "#E0F8F7", "#E0F8F7", 1, PDF_SQUARE_NORMAL)
            Call GF_squareBox(oPDF, 190, currentY, 190, 10, 0, "#E0F8F7", "#E0F8F7", 1, PDF_SQUARE_NORMAL)
            Call GF_squareBox(oPDF, 380, currentY, 190, 10, 0, "#E0F8F7", "#E0F8F7", 1, PDF_SQUARE_NORMAL)
            Call GF_writeTextAlign(oPDF,20, currentY +2, GF_FN2DTE(rsNom("FECHACUPO")) , 60, PDF_ALIGN_CENTER)
            Call GF_writeTextAlign(oPDF,80, currentY +2, auxCodigoDesde &" al "& auxCodigoHasta , 110, PDF_ALIGN_CENTER)
            auxCorredor = Trim(rsNom("IDCORREDOR")) &"-"& Trim(getDsCorredor(rsNom("IDCORREDOR")))
            if (Len(auxCorredor) > 39) then auxCorredor = Left(auxCorredor,37) & ".."
            Call GF_writeTextAlign(oPDF,190, currentY +2, auxCorredor , 190, PDF_ALIGN_LEFT)
            auxVendedor = Trim(rsNom("IDVENDEDOR")) &"-"& Trim(getDsVendedor(rsNom("IDVENDEDOR")))
            if (Len(auxVendedor) > 39) then auxVendedor= Left(auxVendedor,37) & ".."
            Call GF_writeTextAlign(oPDF,380, currentY +2, auxVendedor, 190, PDF_ALIGN_LEFT)
            avanzar()
            rsNom.MoveNext()
        wend
    end if
End Function
'------------------------------------------------------------------------------------------	
Function obtenerNominacionesCupos(p_nroContrato, p_cdProducto, p_cdSucursal, p_cdOperacion, p_anioCosecha, p_fecha)
    Dim strSQL,fechaInicio,fechaFin,rs
    fechaInicio = p_fecha
	fechaFin = GF_DTE2FN(dateadd("d",periodo,GF_FN2DTE(p_fecha)))
    strSQL = "SELECT A.FECHACUPO, A.IDCORREDOR, A.IDVENDEDOR, A.CODIGODESDE, A.CODIGOHASTA, A.PUERTO, B.DESCPR AS DSPRODUCTO "&_
             "FROM MERFL.TBLCUPOSNOMINADOS A "&_
             "  LEFT JOIN MERFL.MER112F1 B ON A.IDPRODUCTO = B.CODIPR "&_
             "WHERE A.IDPRODUCTO="& p_cdProducto &" AND A.IDSUCURSAL="& p_cdSucursal &" AND A.IDOPERACION="& p_cdOperacion &_
             "  AND A.NUMERO="& p_nroContrato &" AND A.COSECHA="& p_anioCosecha &" AND A.FECHACUPO >= "&fechaInicio&" AND A.FECHACUPO <= "&fechaFin &_
             " ORDER BY A.FECHACUPO, A.IDCORREDOR, A.IDVENDEDOR, A.CODIGODESDE, A.CODIGOHASTA"
    Call executeQuery(rs, "OPEN", strSQL)
    Set obtenerNominacionesCupos = rs
End Function
'------------------------------------------------------------------------------------------	
Sub	dibujarCodigosCupos(cdCupo, cdProducto, cdPuerto, cantLineasAVolver, color, ByRef ultimoNumeroSecuencia)
	Dim strSQL, codigosCupos, conn, ultimoCodigoCupo
	Dim strLineaCodigosCupos, primerLetraProducto
	dim esPrimerRegistro
	dim cantLineasAVolverAUX
	dim cantLineasCodigosCupos, letraPuerto
	Dim ultimoNumSecuenciaCodigosCuposInformado
	
	strLineaCodigosCupos = ""
	cantLineasAVolverAUX = cantLineasAVolver
	cantLineasCodigosCupos = 0
	ultimoNumeroSecuencia = 0
		
	primerLetraProducto = getDsAbrProductoParaCodigoCupo(cdProducto, cdPuerto)
	letraPuerto = getLetraCupo(cdPuerto)
	
	if (letraPuerto <> "?") then
		if (chkEnviados = MOSTRAR_NO_ENVIADOS) then
			'tomar ultimo numero de secuencia de codigos de cupos informados desde la tabla toepferdb.tblcuposinformados
			strSQL = "select case when ultimasecuencia is null then 0 else ultimasecuencia end as numSecuencia"
			strSQL = strSQL & " from toepferdb.tblcuposinformados "		
			strSQL = strSQl & " where codigocupo = " & cdCupo	
			Call GF_BD_AS400_2(ultimoCodigoCupo, conn, "OPEN", strSQL)
			if not ultimoCodigoCupo.eof then
				ultimoNumSecuenciaCodigosCuposInformado = ultimoCodigoCupo("numSecuencia")
			else
				ultimoNumSecuenciaCodigosCuposInformado = 0
			end if	
		end if
	
		strSQL = "SELECT C5DSDE AS CUPOSDESDE, C5HSTA AS CUPOSHASTA, C5ASNU FROM MERFL.MER517F5 F5 "
		strSQL = strSQL & " INNER JOIN MERFL.MER517F1 F1 ON F5.C5CODI=F1.CUCODI "
		strSQL = strSQL & "WHERE C5CODI = " & cdCupo
	
		if ultimoNumSecuenciaCodigosCuposInformado > 0 and (chkEnviados = MOSTRAR_NO_ENVIADOS) then
		'trae los codigos de cupos que todavia no se han informado
			strSQL = strSQL & " and c5asnu > " & ultimoNumSecuenciaCodigosCuposInformado
		end if	
	
		Call GF_BD_AS400_2(codigosCupos, conn, "OPEN", strSQL)
		if not codigosCupos.eof then
			if cantLineasAVolverAUX > 0 then
				'hubo mas de una condicion especial, hay que volver a primer linea
				while cantLineasAVolverAUX = 0 
					volver()
					cantLineasAVolverAUX = cantLineasAVolverAUX - 1
				wend
			end if
			esPrimerRegistro = true
			while not codigosCupos.eof		
				codigoDesde = trim(codigosCupos("cuposDesde"))
				codigoHasta = trim(codigosCupos("cuposHasta"))	
				while len(codigoDesde) < 8 
					codigoDesde = "0" & codigoDesde 
				wend	
				while len(codigoHasta) < 8 
					codigoHasta = "0" & codigoHasta 
				wend	
				strLineaCodigosCupos = letraPuerto & primerLetraProducto & codigoDesde & " al " & letraPuerto & primerLetraProducto & codigoHasta
				cantLineasCodigosCupos = cantLineasCodigosCupos + 1
				if esPrimerRegistro or cantLineasCodigosCupos <= cantLineasAVolver then
					Call GF_writeTextAlign(oPDF, 387, currentY+2	, strLineaCodigosCupos	, 183	, PDF_ALIGN_CENTER)	
					esPrimerRegistro = false
				else
					Call GF_squareBox(oPDF, 20, currentY, 550, 10, 0, color, color, 1, PDF_SQUARE_NORMAL)
					Call GF_writeTextAlign(oPDF, 387, currentY+2	, strLineaCodigosCupos	, 183	, PDF_ALIGN_CENTER)	
				end if
				'toma el ultimo numero de secuencia de cupos informados para actualizar el valor en la tabla toepferdb.tblcuposinformados
				ultimoNumeroSecuencia = codigosCupos("c5asnu")			
				codigosCupos.movenext				
			wend
		end if		
	end if
End sub
'------------------------------------------------------------------------------------------	
Sub	dibujarCabeceraContrato(cupos)
	Call GF_squareBox(oPDF, 20	, currentY, 184	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 204	, currentY, 183	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 387	, currentY, 183	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")	
	Call GF_writeTextAlign(oPDF, 22, currentY+2	, GF_TRADUCIR("Contrato") & ": " & GF_EDIT_CONTRATO(cupos("cdProducto"), cupos("cdSucursal"),cupos("cdOperacion"), cupos("nroContrato"), cupos("anioCosecha"))	, 184	, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 206, currentY+2	, GF_TRADUCIR("Contrato") & " " & GF_Traducir("Vendedor") & ": " & getContratoVendedor(cupos("cdProducto"), cupos("cdSucursal"),cupos("cdOperacion"), cupos("nroContrato"), cupos("anioCosecha"))	, 183	, PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF, 389, currentY+2	, GF_TRADUCIR("Producto") & ": " & getDsProducto(CInt(cupos("cdProducto"))), 183	, PDF_ALIGN_LEFT)	
	call GF_setFontColor("000000")	
End sub
'------------------------------------------------------------------------------------------	
Sub	dibujarCorredorVendedor(cupos)
    if (((CInt(cupos("CDPUERTO")) = PUERTO_PIEDRABUENA)or(CInt(cupos("CDPUERTO")) = MUELLE_BAHIABLANCA))and(CInt(cupos("cdOperacion"))= OPERACION_PRESTAMO_DEVOLUCION)) then    
        Call GF_squareBox(oPDF, 20	, currentY, 550, 13, 0, "#CECEF6", "#000000", 1, PDF_SQUARE_NORMAL)
        Call GF_writeTextAlign(oPDF, 22, currentY+2	, GF_TRADUCIR("Cupos otorgados al proveedor ") & getDescripcionProveedor(idProveedor), 550, PDF_ALIGN_LEFT)
    else
        Call GF_squareBox(oPDF, 20	, currentY, 275	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	    Call GF_squareBox(oPDF, 290	, currentY, 280	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	    call GF_setFont(oPDF,"ARIAL",8,8)
	    call GF_setFontColor("FFFFFF")
	    Call GF_writeTextAlign(oPDF, 22, currentY+2	, GF_TRADUCIR("Corredor") & ": " & 	getDescripcionProveedor(CInt(cupos("cdCorredor"))), 275	, PDF_ALIGN_LEFT)
	    Call GF_writeTextAlign(oPDF, 292, currentY+2	, GF_TRADUCIR("Vendedor") & ": " & 	getDescripcionProveedor(CInt(cupos("cdVendedor"))), 280	, PDF_ALIGN_LEFT)		
	    call GF_setFontColor("000000")
    end if
End sub
'------------------------------------------------------------------------------------------
function dibujarTituloCupos()
	Call GF_squareBox(oPDF, 20	, currentY, 100	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 120	, currentY, 84	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)	
	Call GF_squareBox(oPDF, 204	, currentY, 183	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, 387	, currentY, 183	, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	call GF_setFont(oPDF,"ARIAL",8,8)
	call GF_setFontColor("FFFFFF")
	Call GF_writeTextAlign(oPDF, 20, currentY+2	, GF_TRADUCIR("FECHA")			, 100	, PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, 120, currentY+2	, GF_TRADUCIR("CUPOS")			, 84	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 204, currentY+2	, GF_TRADUCIR("PUERTO")	, 183	, PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF, 387, currentY+2	, GF_TRADUCIR("CODIGOS CUPOS")	, 183	, PDF_ALIGN_CENTER)	
	call GF_setFontColor("000000")
end function
'------------------------------------------------------------------------------------------	
function dibujarTitulo(pTitulo)
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,2, 25, GF_TRADUCIR(pTitulo), 590, PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)		
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)	
end function
'------------------------------------------------------------------------------------------	
'Obtiene el nombre del archivo a generar.
Function getFilename()
         Randomize()
         getFilename = "cupos_asignados" & "-" & Int(100 * Rnd()) & ".pdf"
End Function
'------------------------------------------------------------------------------------------	
Sub	avanzar()
	currentY = currentY + SEPARATION
	if currentY >= PAGE_HEIGHT_SIZE - 55 then
		nuevaPagina()
		if not cupos.eof then
			Call dibujarCabeceraContrato (cupos)
			avanzar()
			Call dibujarTituloCupos()
			avanzar()
		end if
	end if
End sub
'------------------------------------------------------------------------------------------	
Sub	volver()
	currentY = currentY - SEPARATION	
End sub
'------------------------------------------------------------------------------------------
function nuevaPagina()
	Call GF_newPage(oPDF)
	currentY = 80
	nroPagina = nroPagina + 1
	call dibujarTitulo(GF_TRADUCIR("CUPOS"))
end function
%>