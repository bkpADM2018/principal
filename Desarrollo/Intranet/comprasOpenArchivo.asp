<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientos.asp"-->


<%
'/************************************************************************************
' * Funccion: 	getPCTDBFile
' * Descripcion: 	Obtiene un archivo binario de la base de datos de un PCT
' * Parametros: 	p_idPedido : id del pedido
' *			p_fileNo   : numero del archivo 
' *
' * Autor:		Guido Fonticelli
' * Fecha:		20/11/2009
' * Ultima Modificacion: 20/11/2009
'************************************************************************************/
Function getPCTDBFile(p_idPedido,p_fileNo,byRef myExt)
	dim strsql,rs
	
	strSQL = "Select * from TBLPCTBINARYFILES where idPedido=" & p_idPedido & " and FILENO=" & p_fileNo	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	myExt		= rs("ext")
	getPCTDBFile= rs("FILEBIN")
End Function
'--------------------------------------------------------------------------------------------------
Function getFilesPCT(pIdPedido,pFileNum)
	dim strSQL,rs,conn,rtrn,fileName
	
	strSQL = "Select filescan from TBLPCTBINARYFILES where idPedido=" & pIdPedido & " and FILENO=" & pFileNum
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	rtrn = 0
	fileName  = ""
	if (not rs.EoF) then rtrn = rs("cant")
	
	if (rtrn > 0) then fileName = buildFileName(pIdPedido, pFileNum, "")
	
		
	getFilesPCT = rtrn & "|" & fileName
	
End Function
'--------------------------------------------------------------------------------------------------
Function openPCTFile(idPedido,idContrato,fileno)

	if (idPedido > 0) then
		dbFile   = getPCTDBFile(idPedido,fileno,ext)
		fileName = buildFileName(idPedido, fileno, ext)
	elseif (idContrato > 0) then
		Call getCTCDBFile(idContrato, dbFile, fileName)
	else
		Response.Redirect "comprasAccesoDenegado.asp"
	end if

	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
End Function
'--------------------------------------------------------------------------------------------------
Function deleteFilePCT(pIdPedido,pFileNum)
	dim strSQL,rs,conn,rtrn
	
	if (pIdPedido > 0) then
		strSQL = "Delete from TBLPCTBINARYFILES where idPedido=" & pIdPedido & " and FILENO=" & pFileNum
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
'Function getCantFilesAFE(pIdAfe)
'	dim strSQL,rs,conn,rtrn
'	rtrn = 0
'	
'	strSQL = "Select count(*) cant from TOEPFERDB.TBLDATOSAFE where filescan is not null and idAfe= " & pIdAfe
'	Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)
'	
'	if (not rs.EoF) then rtrn = rs("cant")
'		
'	getCantFilesAFE = rtrn & "|" & "AFE-"&pIdAfe
'	
'End Function
'--------------------------------------------------------------------------------------------------
Function openAfeFile(pIdAfe)
	dim strSQL,rs,conn,rtrn,dbFile
	
	strSQL = "Select filescan,fileext from TBLDATOSAFE where idAfe=" & pIdAfe
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	if (not rs.EoF) then
		dbFile = rs("filescan")
		fileExt = rs("fileext")
		
		Response.AddHeader "Content-Disposition", "attachment; filename=" & "AFE-" & pIdAfe & "." & fileExt
		Response.CharSet     = "UTF-8"
		Response.ContentType = "application/octet-stream"

		Response.BinaryWrite dbFile
	end if
	
End Function
'--------------------------------------------------------------------------------------------------
Function deleteFilePCT(pIdAfe)
	dim strSQL,rs,conn,rtrn
	
	strSQL = "update tbldatosafe set filescan = null, fileext = null where idafe = " & pIdAfe
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
	
End Function
'--------------------------------------------------------------------------------------------------
Function openProveedorFile()
	Dim idProveedor,secuencia,fileName2
	
	idProveedor = GF_PARAMETROS7("idProveedor"	 ,0 ,6)
	secuencia = GF_PARAMETROS7("secuencia"	 ,0 ,6)

	if (idProveedor > 0 and secuencia <> 0) then
		dbFile   = getProveedorFileDB(idProveedor,secuencia,ext,fileName)
		fileName2 = fileName & "." & ext
	else
		Response.Redirect "comprasAccesoDenegado.asp"
	end if

	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName2
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
End Function
'-------------'--------------------------------------------------------------------------------------------------
Function openPicFile()
	Dim nroreq, secuencia,rs,fileName2
	idcotizacion = GF_PARAMETROS7("idcotizacion"	 ,0 ,6)
	secuencia = GF_PARAMETROS7("secuencia"	 ,0 ,6)

	'if (idcotizacion > 0 and secuencia <> 0) then
		strSQL = "select * from TBLCTZBINARYFILES where idcotizacion = " & idcotizacion & " and fileno = " & secuencia
		call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		dbFile = rs("filebin")
		fileName2 = rs("name") & "." & rs("ext")
	'else
	'	Response.Redirect "comprasAccesoDenegado.asp"
	'end if

	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName2
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
End Function
'-------------'--------------------------------------------------------------------------------------------------
Function openSmFile()
	Dim nroreq, secuencia, rs, fileName2
	id = GF_PARAMETROS7("id" ,0 ,6)
	typeO = GF_PARAMETROS7("typeO" ,0 ,6)
	secuencia = GF_PARAMETROS7("secuencia" ,0 ,6)

	strSQL = "SELECT * FROM TBLSMBINARYFILE WHERE ID = " & id & " AND TYPE=" & typeO & " AND FILENO = " & secuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	dbFile = rs("BINARYFILE")
	fileName2 = rs("NAME") & "." & rs("EXT")

	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName2
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
End Function
'-------------'--------------------------------------------------------------------------------------------------
Function openSmOtFile()
	Dim nroreq, secuencia, rs, fileName2
	id = GF_PARAMETROS7("id" ,0 ,6)
	secuencia = GF_PARAMETROS7("secuencia" ,0 ,6)

	strSQL = "SELECT * FROM TBLSMOTBINARYFILE WHERE ID = " & id & " AND FILENO = " & secuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	dbFile = rs("BINARYFILE")
	fileName2 = rs("NAME") & "." & rs("EXT")

	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName2
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
End Function
'--------------------------------------------------------------------------------------------------
Function getProveedorFileDB(pIdProveedor,pSecuencia,byRef myExt,byRef fileName)
	dim strsql,rs
	
	strSQL = "Select * from TBLPROVEEDORESARCHIVOS where idproveedor=" & pIdProveedor & " and secuencia=" & pSecuencia
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	myExt		= rs("ext")
	fileName 	= rs("filename")
	getProveedorFileDB= rs("file")
End Function
'--------------------------------------------------------------------------------------------------
Function openDraftSurveyFile()
	Dim idDraft ,secuencia, rs, fileName2	
	idDraft = GF_PARAMETROS7("iddraft",0 ,6)
	pto = GF_PARAMETROS7("pto","" ,6)
	strSQL = "SELECT NAMEFILE,EXTFILE,BINARYFILE FROM DB2ADMIN.TBLEMBARQUESDRAFTSURVEY WHERE IDDRAFT = " & idDraft 
	Call GF_BD_Puertos (pto, rs, "OPEN",strSQL)
	if not rs.Eof then
		dbFile = rs("BINARYFILE")
		fileName2 = rs("NAMEFILE") & "." & rs("EXTFILE")
	end if	
	Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName2
	Response.CharSet     = "UTF-8"
	Response.ContentType = "application/octet-stream"

	Response.BinaryWrite dbFile
	
	
End Function 
'***********************************************************************
'************		INICIO DE PAGINA 		******************
'***********************************************************************
Dim idPedido, fileno,ext, dbFile, idContrato, fileName,myType,myFileNum
Dim idAfe

idPedido 	= GF_PARAMETROS7("idPedido"	 ,0 ,6)
myId		= GF_PARAMETROS7("id"	 ,0 ,6)
idContrato 	= GF_PARAMETROS7("idContrato",0 ,6)
myType 		= GF_PARAMETROS7("type"		 ,"",6)
myFileNum 	= GF_PARAMETROS7("fileNum"	 ,0 ,6)
fileno   	= GF_PARAMETROS7("fileno"  ,0,6)

if (myType = "") then
	call openPCTFile(idPedido,idContrato,fileno)
else
	select case ucase(myType)
		case "PCT":
			response.write getFilesPCT(myId,myFileNum)
		case "PCT-OPEN":
			call openPCTFile(idPedido,idContrato,fileno)
		case "PCT-DELETE":
			Call deleteFilePCT(myId,myFileNum)
		case "AFE":
			response.write getCantFilesAFE(myId)
		case "AFE-OPEN":
			Call openAfeFile(myId)
		case "AFE-DELETE":
			Call deleteFilePCT(myId)
		case "PROV-OPEN"
			Call openProveedorFile()
		case "PIC-OPEN"
			call openPicFile()
		case "SM-OPEN"
			call openSmFile()
		case "SM-OT-OPEN"
			call openSmOtFile()
		case "DRAFT-OPEN"
			call openDraftSurveyFile()				
	end select
end if

Response.End



%>
