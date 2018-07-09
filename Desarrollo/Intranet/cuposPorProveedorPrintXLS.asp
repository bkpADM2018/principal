<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
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
xls_accion = GF_Parametros7("xls_accion", 0, 6)
periodo = GF_Parametros7("periodo", 0, 6)
if periodo = 0 then periodo = 9

usr = GF_PARAMETROS7("usr","",6)

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

strDestinosMail =  getStringMailsProveedor(idProveedor) 
if xls_accion = XLS_FILE_MODE and strDestinosMail = "" and mostrarConfirmacion = "N" then
	'si se envian los mail en batch desde cuposPorProveedor.asp y el proveedor no tiene mail, corta y devuelve por AJAX el aviso
	 Response.Write SIN_MAIL
	 Response.end
end if

filename = getFileName()
strPathAttachment = Server.mapPath("temp/" & filename & ".xls")
Call GF_setXLSMode(xls_accion)
Call GF_createXLS(filename)
Call armadoXLS()
Call closeXLS()

if xls_accion = XLS_FILE_MODE then
'enviar mail con attachment de xls generado
	strCuerpoMail = GF_Parametros7("descripcionMail", "", 6)
	if strCuerpoMail = "" then strCuerpoMail = "Ver cupos asignados en el archivo adjunto"
	myMail = getUserMail(usr)
	if (myMail <> "") then strDestinosMail = getUserMail(usr) & ";" & strDestinosMail
	Call GP_ENVIAR_MAIL_ATTACHMENT("Alfred Toepfer S.R.L - Cupos Asignados", strCuerpoMail,SENDER_CUPOS_BA, strDestinosMail, strPathAttachment)
	'borrar el archivo xls creado
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
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
'-----------------------------------------------------------------------------------------------
function getContratos (idProveedor, fecha, periodo, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor)
	Dim strSQL, rs, conn
	Dim fechaInicio, fechaFin
	fechaInicio = fecha
	fechaFin = GF_DTE2FN(dateadd("d",periodo,GF_FN2DTE(fecha)))

	strSQL = "select cuncto as nroContrato, cucpro as cdProducto, cucsuc as cdSucursal, "
	strSQL = strSQL & " cucope as cdOperacion, cuacos as anioCosecha,cucdes as nropuerto, concr1 as numContratoVendedor from merfl.mer517f1 "
	strSQL = strSQL & " inner join MERFL.MER311F1 on CUCPRO=CPROR1 and CUCSUC=CSUCR1 and CUCOPE=COPER1 and CUNCTO = NCTOR1 and CUACOS = ACOSR1 "
	strSQL = strSQL & "where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and (cuccor = " & idProveedor & " or cucven = " & idProveedor & ") "
	
	'aplicarFiltros
	if fltrCdProducto <> "" then strSQL = strSQL & " and cucpro = " & fltrCdProducto
	if fltrCdSucursal <> "" then strSQL = strSQL & " and cucsuc = " & fltrCdSucursal
	if fltrCdOperacion <> "" then strSQL = strSQL & " and cucope = " & fltrCdOperacion
	if fltrNroContrato <> "" then strSQL = strSQL & " and cuncto = " & fltrNroContrato
	if fltrAnioCosecha <> "" then strSQL = strSQL & " and cuacos = " & fltrAnioCosecha
	if fltrPuerto <> 0 then strSQL = strSQL & " and cucdes = " & fltrPuerto
	if fltrCorredor <> 0 then strSQL = strSQL & " and cuccor = " & fltrCorredor
	if fltrVendedor <> 0 then strSQL = strSQL & " and cucven = " & fltrVendedor
	
	strSQL = strSQL & " group by cuncto, cucpro, cucsuc, cucope, cuacos,cucdes, concr1"
	strSQL = strSQL & " order by cuncto, cucope "
	
	Call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
	Set getContratos = rs
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
Function armadoXLS()

writeXLS("<html><head><style type='text/css'>")
writeXLS(".titCupos{background-color:#CECEF6;FONT-WEIGHT: bold; border-color:#666666; border-style:solid; border-width:thin;}")
writeXLS(".titCupos2{background-color:#5577CC;FONT-WEIGHT: bold; border-color:#666666; color:#FFFFFF; text-align: center; border-style:solid; border-width:thin;}")
writeXLS(".cupos1{background-color:#E3F6CE;FONT-WEIGHT: bold;border-color:#666666; border-style:solid; border-width:thin;}")
writeXLS(".cupos2{background-color:#FFFACD;FONT-WEIGHT: bold; border-color:#666666; border-style:solid; border-width:thin;}")
writeXLS("</style></head><body>")
	if Cdbl(nroContrato) = 0 then
	'dibujar varios contratos		
		dim contadorContratos ' esta variable es necesaria para determinar el codigo de operacion del primer contrato, para saber el sender
		set contratos = getContratos(idProveedor, fecha, periodo, fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor)
		if not contratos.eof then	
			contadorContratos = 0
			while not contratos.eof
				if contadorContratos = 0 then cdSucursal = contratos("cdSucursal")
				esContratoNuevo = true					
				Call dibujarContrato(contratos("nroContrato"),contratos("cdProducto"),contratos("cdSucursal"),contratos("cdOperacion"),contratos("anioCosecha"), fecha)				                
				contadorContratos = contadorContratos + 1
				contratos.movenext
			wend			
		end if
	else		    
    	Call dibujarContrato (nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)		
	end if	
writeXLS("</body></html>")
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
	if (chkEnviados = MOSTRAR_NO_ENVIADOS) then 
	    strSQL= strSQL & " sum (cucccp - case when cupos is null then 0 else cupos end) "
	else
	    strSQL = strSQL & " sum(cucccp) "	
    end if	    
	strSQL = strSQL & " as nroCamiones, cuccor as cdCorredor, cucven as cdVendedor, C.RAZSOC dsCorredor, V.RAZSOC dsVendedor"
	strSQL = strSQL & " from merfl.mer517f1" 
	strSQL = strSQL & " inner join merfl.TCB6A1F1 C on cuccor=C.NROPRO " 
	strSQL = strSQL & " inner join merfl.TCB6A1F1 V on cucven=V.NROPRO " 
	strSQL = strSQL & " left join MERFL.MER311FJ on CPRORJ=CUCPRO and CSUCRJ=CUCSUC and COPERJ=CUCOPE and NCTORJ=CUNCTO and ACOSRJ=CUACOS "
	if (chkEnviados = MOSTRAR_NO_ENVIADOS) then strSQL = strSQL & " left join toepferdb.tblcuposinformados  on CODIGOCUPO=CUCODI "
	strSQL = strSQL & "where cufccp >= " & fechaInicio & " and cufccp <= " & fechaFin & " and cuncto = " & nroContrato 
	strSQL = strSQL & " and cucpro=" & cdProducto &  " and cucsuc =" & cdSucursal & " and cucope=" & cdOperacion & " and cuacos=" & anioCosecha	
	strSQL = strSQL & " group by cucodi, cufccp, cuzinf, cucdes, cuncto, cucpro, cucsuc, cucope, cuacos, cuccor, cucven, C.RAZSOC, V.RAZSOC, MCPDRJ) " 
	strSQL = strSQL & " as tablaGral "	
	strSQL = strSQL & " where nroCamiones <> 0 order by fechaCupo"
	getStringSQLCupos = strSQL
end function
'------------------------------------------------------------------------------------------
Sub dibujarContrato (nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, fecha)
	Dim strSQL, conn , isMsjBahia, strCodigosCupo
	Dim fechaInicio, fechaFin
	isMsjBahia = false
	strSQL = getStringSQLCupos(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)	
	Call GF_BD_AS400_2(cupos, conn, "OPEN", strSQL)
	
	if not cupos.eof then
	writeXLS("<table width='1000px'>")				
	Call dibujarCabeceraContrato (cupos)
	dibujarCorredorVendedor(cupos)
	Call dibujarTituloCupos()
		while not cupos.eof
			'se dibujan los cupos
			writeXLS("<tr>")
				writeXLS("<td class='cupos2' align='center'>" & GF_FN2DTE(cupos("fechaCupo")) & "</td>")
				writeXLS("<td class='cupos2' align='center'>" & cupos("nroCamiones") & "</td>")
				writeXLS("<td class='cupos2' align='center'>" & getDsPort(CInt(cupos("cdPuerto"))) & "</td>")				
				strCodigosCupo = dibujarCodigosCupos(cupos("cdCupo"), cupos("cdProducto"), CInt(cupos("cdPuerto")), ultimoNumeroSecuencia)
                if (((CInt(cupos("CDPUERTO")) = PUERTO_PIEDRABUENA)or(CInt(cupos("CDPUERTO")) = MUELLE_BAHIABLANCA))and(CInt(cupos("cdOperacion"))= OPERACION_PRESTAMO_DEVOLUCION)) then
                    writeXLS("<td class='cupos2' align='center'>"& GF_TRADUCIR("Los codigo de cupo se obtendran luego de nominar") &"</td>")                    
                else                    
				    writeXLS("<td class='cupos2' align='center'>" & strCodigosCupo & "</td>")
                end if
			writeXLS("</tr>")	
			if (xls_accion = XLS_FILE_MODE) and (chkEnviados = MOSTRAR_NO_ENVIADOS) then
				'grabar en la tabla historica toepferdb.tblcuposinformados los cupos que se van a mandar para esta fecha
				Call grabarCuposInformados(cupos, ultimoNumeroSecuencia)
			end if
            if ((CInt(cupos("cdPuerto")) = PUERTO_PIEDRABUENA) OR (CInt(cupos("cdPuerto")) = MUELLE_BAHIABLANCA)) then isMsjBahia = true
            cupos.movenext
		wend		
            'Mensaje que solo se agrega debajo de cada contrato hacia Bahia Blanca
        if (isMsjBahia) then		                
            writeXLS("<tr><td colspan='4' style='color:red; font-weight:900; background:yellow;'>Los cupos de lunes a viernes tendrán apertura el día anterior al cupo a las 15 hs y como cierre el día del mismo a las 17 hs.</td></tr>")
            writeXLS("<tr><td colspan='4' style='color:red; font-weight:900; background:yellow;'>Para los días sábados y domingos será el día anterior a las 15 hs,  y el cierre el día del cupo a las 10 hs.</td></tr>")
        end if
    writeXLS("</table><br>")		
	end if
End sub
'------------------------------------------------------------------------------------------	
function dibujarCodigosCupos(cdCupo, cdProducto, cdPuerto, ByRef ultimoNumeroSecuencia)
	Dim strSQL, codigosCupos, conn, ultimoCodigoCupo
	Dim strLineaCodigosCupos, primerLetraProducto
	Dim ultimoNumSecuenciaCodigosCuposInformado, letraPuerto
	
	strLineaCodigosCupos = ""
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
		strLineaCodigosCupos = ""
			while not codigosCupos.eof			
				codigoDesde = trim(codigosCupos("cuposDesde"))
				codigoHasta = trim(codigosCupos("cuposHasta"))	
				while len(codigoDesde) < 8 
					codigoDesde = "0" & codigoDesde 
				wend	
				while len(codigoHasta) < 8 
					codigoHasta = "0" & codigoHasta 
				wend				
				strLineaCodigosCupos = strLineaCodigosCupos & letraPuerto & primerLetraProducto & codigoDesde & " al " & letraPuerto & primerLetraProducto & codigoHasta & "<br>"
				'toma el ultimo numero de secuencia de cupos informados para actualizar el valor en la tabla toepferdb.tblcuposinformados
				ultimoNumeroSecuencia = codigosCupos("c5asnu")			
				codigosCupos.movenext			
			wend
		end if			
	end if
	dibujarCodigosCupos = strLineaCodigosCupos	
End function
'------------------------------------------------------------------------------------------	
Sub	dibujarCabeceraContrato(cupos)
	writeXLS("<tr width='1000px'>")
		writeXLS("<td class='titCupos' colspan='2'>"& GF_TRADUCIR("Contrato") & ": " & GF_EDIT_CONTRATO(cupos("cdProducto"), cupos("cdSucursal"),cupos("cdOperacion"), cupos("nroContrato"), cupos("anioCosecha")) & "</td>")
		writeXLS("<td class='titCupos' >" & GF_TRADUCIR("Contrato") & " " & GF_Traducir("Vendedor") & ": " & getContratoVendedor(cupos("cdProducto"), cupos("cdSucursal"),cupos("cdOperacion"), cupos("nroContrato"), cupos("anioCosecha")) & "</td>")
		writeXLS("<td class='titCupos' >" & GF_TRADUCIR("Producto") & ": " & getDsProducto(CInt(cupos("cdProducto"))) & "</td>")
	writeXLS("</tr>")
	if (cupos("PProd") = "V") then dibujarDeclaracionPP()
End sub
'------------------------------------------------------------------------------------------	
Sub	dibujarDeclaracionPP()
	writeXLS("<tr width='1000px'>")
		writeXLS("<td class='titCupos2' colspan='4'>"& GF_TRADUCIR("Contrato declarado como de ppia. producción") & "</td>")		
	writeXLS("</tr>")
End sub
'------------------------------------------------------------------------------------------	
Sub	dibujarCorredorVendedor(cupos)
	writeXLS("<tr>")
        if (((CInt(cupos("CDPUERTO")) = PUERTO_PIEDRABUENA)or(CInt(cupos("CDPUERTO")) = MUELLE_BAHIABLANCA))and(CInt(cupos("cdOperacion"))= OPERACION_PRESTAMO_DEVOLUCION)) then
            writeXLS("<td class='titCupos' colspan='4' >" & GF_TRADUCIR("Cupos otorgados al proveedor ") & cupos("dsVendedor") & "</td>")
        else
		    writeXLS("<td class='titCupos' colspan='2' >" & GF_TRADUCIR("Corredor") & ": " & cupos("dsCorredor") & "</td>")
		    writeXLS("<td class='titCupos' colspan='2' >" & GF_TRADUCIR("Vendedor") & ": " & cupos("dsVendedor") & "</td>")
        end if
	writeXLS("</tr>")
End sub
'------------------------------------------------------------------------------------------
function dibujarTituloCupos()
	writeXLS("<tr>")
		writeXLS("<td class='cupos1' align='center'>" & GF_TRADUCIR("FECHA") & "</td>")
		writeXLS("<td class='cupos1' align='center'>" & GF_TRADUCIR("CUPOS") & "</td>")
		writeXLS("<td class='cupos1' align='center'>" & GF_TRADUCIR("PUERTO") & "</td>")
		writeXLS("<td class='cupos1' align='center'>" & GF_TRADUCIR("CODIGOS CUPOS") & "</td>")
	writeXLS("</tr>	")
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
         getFilename = "cupos_asignados" & "-" & Int(100 * Rnd())
End Function
'------------------------------------------------------------------------------------------
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

%>