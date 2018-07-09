<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosBoletos.asp"-->
<!--#include file="Includes/procedimientosmail.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<%
'-----------------------------------------------------------------------------------------------
         locale=session.lcid
		 session.lcid=2057	'Formato dd/mm/aaaa

Const CONV_KEY_PUERTO = "PUERTO"
Const CONV_KEY_PRODUCTO = "PRODUCTO"

dim rs, conn, strSQL, oPDF
dim idProveedor, nroContrato, fecha, periodo
Dim currentY, nroPagina
Dim cupos
Dim fltrCdProducto, fltrCdSucursal, fltrCdOperacion,fltrNroContrato,fltrAnioCosecha,fltrPuerto,fltrCorredor,fltrVendedor
Dim esContratoNuevo, dicConv
dim strPathAttachment, strCuerpoMail, strDestinosMail
dim chkEnviados, usr, mostrarConfirmacion,g_strPuerto
Dim fs, myFile

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
if strDestinosMail = "" and mostrarConfirmacion = "N" then
	'si se envian los mail en batch desde cuposPorProveedor.asp y el proveedor no tiene mail, corta y devuelve por AJAX el aviso
	 Response.Write SIN_MAIL
	 Response.end
end if

filename = getFileName()
strPathAttachment = Server.mapPath("temp/" & filename & ".txt")
Set fs = Server.CreateObject("Scripting.FileSystemObject")
If fs.FileExists(strPathAttachment) Then  Call fs.deleteFile(strPathAttachment, true)
Set myFile = fs.OpenTextFile(strPathAttachment, 2, true)
Call armadoADM()
myFile.Close()    
Set myFile = Nothing
Set fs = Nothing

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
Function armadoADM()
    
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
	strSQL = strSQL & " cucsuc as cdSucursal, cucope as cdOperacion, cuacos as anioCosecha, MCPDRJ PProd, CUCCOO, V.NRODOC cuitVendedor, "	 
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
	strSQL = strSQL & " group by cucodi, cufccp, cuzinf, cucdes, cuncto, cucpro, cucsuc, cucope, cuacos, cuccor, cucven, CUCCOO, V.NRODOC, C.RAZSOC, V.RAZSOC, MCPDRJ) " 
	strSQL = strSQL & " as tablaGral "	
	strSQL = strSQL & " where nroCamiones <> 0 order by fechaCupo"
	getStringSQLCupos = strSQL
end function
'---------------------------------------------------------------------------------------------------------
Function cargarTablaConversion(pCuitCliente, pPto)
    
    Dim strSQL, rs, ret, auxkey, auxval
    
    
    Set dicConv = createObject("Scripting.Dictionary")    
    ret = false    
    strSQL="Select * from TBLCONVERSIONES where NUCUITCLIENTE='" & pCuitCliente& "'"
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)    
    while (not rs.eof)
        auxkey = rs("TIPODATO") & "_" & rs("CDPROPIO")
        auxval = rs("CDTERCERO")
        dicConv.Add auxkey, auxval
        ret = true
        rs.MoveNext()
    wend    
    cargarTablaConversion = ret
        
    
End Function
'---------------------------------------------------------------------------------------------------------
Function convertir(pTipo, pCodigoPropio)
    Dim rtrn
    
    rtrn=""
    if (dicConv.Exists(pTipo & "_" & pCodigoPropio)) then        
        rtrn = dicConv.Item(pTipo & "_" & pCodigoPropio)
    end if
    convertir = rtrn    
End Function
'------------------------------------------------------------------------------------------
Sub dibujarContrato (nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha, fecha)
	Dim strSQL, conn , isMsjBahia, cuitCoordinado
	Dim fechaInicio, fechaFin, registro
	
	isMsjBahia = false	
	strSQL = getStringSQLCupos(nroContrato, cdProducto, cdSucursal, cdOperacion, anioCosecha,fecha)	
	Call GF_BD_AS400_2(cupos, conn, "OPEN", strSQL)	
	if not cupos.eof then
	    cuitCoordinado = cupos("CUCCOO")
		if (CInt(cupos("cdOperacion")) = 4) then cuitCoordinado = cupos("cuitVendedor")
	    Call cargarTablaConversion(cuitCoordinado, getCxPuerto(cupos("cdPuerto")))
		while not cupos.eof				
			'se dibujan los cupos			
			Call dibujarCodigosCupos(cupos("fechaCupo"), cupos("cdCupo"), cupos("cdProducto"), CInt(cupos("cdPuerto")), ultimoNumeroSecuencia)            			
            if (chkEnviados = MOSTRAR_NO_ENVIADOS) then
				'grabar en la tabla historica toepferdb.tblcuposinformados los cupos que se van a mandar para esta fecha
				Call grabarCuposInformados(cupos, ultimoNumeroSecuencia)
			end if
            if ((CInt(cupos("cdPuerto")) = PUERTO_PIEDRABUENA) OR (CInt(cupos("cdPuerto")) = MUELLE_BAHIABLANCA)) then isMsjBahia = true
            cupos.movenext
		wend
	end if
End sub
'-----------------------------------------------------------------------------------------------------------------
function getCxPuerto(pNbr)
dim rtrn
rtrn = -1

select case Cint(pNbr) 
	case 10, 54
		rtrn = TERMINAL_TRANSITO
	case 91, 64
		rtrn = TERMINAL_PIEDRABUENA
	case 36, 18
		'Para ARROYO se suma el puerto 18 que se utiliza para difrenciar condiciones de producto.
		rtrn = TERMINAL_ARROYO
end select 	
getCxPuerto = rtrn
end function
'------------------------------------------------------------------------------------------	
function dibujarCodigosCupos(fechaCupo, cdCupo, cdProducto, cdPuerto, ByRef ultimoNumeroSecuencia)
	Dim strSQL, codigosCupos, conn, ultimoCodigoCupo
	Dim strLineaCodigosCupos, primerLetraProducto, myCodigo
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
				codigoDesde = CLng(codigosCupos("cuposDesde"))
				codigoHasta = CLng(codigosCupos("cuposHasta"))	
				for myCodigo = codigoDesde to codigoHasta
				    while len(myCodigo) < 8 
					    myCodigo = "0" & myCodigo 
				    wend	
				    '1.- Código de Cupo (Alfanumérico. máx 11 posiciones)
                    registro = letraPuerto & primerLetraProducto & myCodigo & "|"
                    '2.- Fecha del Cupo (Formato AAAAMMDD)
                    registro = registro & fechaCupo & "|"
                    '3.- Código de Producto (Según tabla ADM)
                    registro = registro & GF_nDigits(convertir(CONV_KEY_PRODUCTO, cdProducto), 3) & "|"
                    '4.- Código de Puerto (Según tabla ADM)         
                    registro = registro & GF_nDigits(convertir(CONV_KEY_PUERTO, cdPuerto), 2)            
                    myFile.WriteLine registro            
                next				
				'toma el ultimo numero de secuencia de cupos informados para actualizar el valor en la tabla toepferdb.tblcuposinformados
				ultimoNumeroSecuencia = codigosCupos("c5asnu")			
				codigosCupos.movenext			
			wend
		end if			
	end if
End function
'------------------------------------------------------------------------------------------	
'Obtiene el nombre del archivo a generar.
Function getFilename()
         Randomize()
         getFilename = "cupos_asignados" & "-" & Int(100 * Rnd())
End Function
'------------------------------------------------------------------------------------------

%>
