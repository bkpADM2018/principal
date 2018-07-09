<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosProveedores.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/json.asp"-->
<%

Function stringUnidades()
	
	Dim strSQL, rs, conn, ret	
	
	strSQL="Select * from TBLUNIDADES where ESTADO <> " & ESTADO_BAJA
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		ret = ret & rs("IDUNIDAD") & "|" & rs("DSUNIDAD")			
		rs.MoveNext()
	wend
	stringUnidades = ret		
End Function
'---------------------------------------------------------------------------------------------
Function stringCategorias()
	
	Dim strSQL, rs, conn, ret	

	strSQL="Select * from TBLARTCATEGORIAS where ESTADO <> " & ESTADO_BAJA
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		ret = ret & rs("IDCATEGORIA") & "|" & rs("CDCATEGORIA") & "|" & rs("DSCATEGORIA")	
		rs.MoveNext()		
	wend
	stringCategorias = ret
	
End Function
'---------------------------------------------------------------------------------------------
Function stringSolicitantes()
	
	Dim strSQL, rs, conn, ret
	
	strSQL= "Select * from VWPARAMSFIRMAUSUARIO order by Apellido, Nombre"	
	'Response.Write strSQL
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		
		ret = ret & UCase(rs("NombreUsuario")) & "-" & rs("APELLIDO") & ", " & rs("NOMBRE")
		rs.MoveNext()		
	wend
	ret = replace(ret,"'","")
	ret = replace(ret,"è","e")
	ret = replace(ret,"é","e")
	ret = replace(ret,"ò","o")
	ret = replace(ret,"ó","o")
	ret = replace(ret,"á","a")
	ret = replace(ret,"à","a")


	stringSolicitantes = ret
	
End Function
'---------------------------------------------------------------------------------------------
Function stringPersonas()
	
	Dim strSQL, rs, conn, ret
	
	strSQL= "Select CDUSUARIO, NOMBRE from VWUSUARIOS order by NOMBRE"	
	'Response.Write strSQL
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		
		ret = ret & UCase(rs("CDUSUARIO")) & "-" & rs("NOMBRE")
		rs.MoveNext()		
	wend
	ret = replace(ret,"'","")
	ret = replace(ret,"è","e")
	ret = replace(ret,"é","e")
	ret = replace(ret,"ò","o")
	ret = replace(ret,"ó","o")
	ret = replace(ret,"á","a")
	ret = replace(ret,"à","a")


	stringPersonas = ret
	
End Function
'---------------------------------------------------------------------------------------------
Function stringObras()
	
	Dim strSQL, rs, conn, ret, str

	str = GF_PARAMETROS7("divObras","",6)
	
	strSQL= "Select * from TBLDATOSOBRAS "
	'response.write strSQL
	if (str <> "") then strSQL= strSQL & " WHERE CDOBRA like '%" & str & "%' or DSOBRA like '%" & str & "%'"	
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		ret = ret & UCase(rs("CDOBRA")) & "-" & rs("DSOBRA")
		rs.MoveNext()		
	wend
	stringObras = ret
	
End Function
'---------------------------------------------------------------------------------------------
Function stringEmpresas()
	
	Dim strSQL, rs, conn, ret, str, myLinea, myWhere
	
	myLinea = GF_PARAMETROS7("linea",0,6)
	str = UCASE(GF_PARAMETROS7("companyName" & myLinea,"",6))
	if len(str) > 0 then
		if IsNumeric(str) then 			
			myWhere = " where (nroemp like '" & str & "%' or nrodoc like '" & str & "%')"
		else
			if len(str) = 1 then
				myWhere = " where nomemp like '" & str & "%'"
			else
				myWhere = " where nomemp like '%" & str & "%'"
			end if	
		end if	
	end if	
	
	'EAB VER
	'if len(myWhere) > 1 then myWhere = myWhere & " AND " 
	'myWhere = myWhere & " ESTADO <> '" & ESTADO_DESHABILITADO & "' or (ESTADO = '" & ESTADO_DESHABILITADO & "' and PROFOR = '" & PROV_PROFORMA & "')"
	
	
	strSQL= "Select * from [Database].[dbo].met001a " & myWhere	
	
	
	'Response.Write strSQL	
	Call executeQueryDB(DBSITE_SQL_MAGIC, rs, "OPEN", strSQL)
	'ret = "SQL-" & strSQL
	while (not rs.eof) 
		if (len(ret) > 0) then ret = ret & ";"			
		ret = ret & UCase(rs("nroemp")) & "-" & rs("nomemp") & "-" & rs("nrodoc")
		rs.MoveNext()		
	wend
	ret = Replace(ret, "'", "")	
	stringEmpresas = ret
	
End Function
'---------------------------------------------------------------------------------------------
Function stringArticulos()
	dim strSQL, rs, conn, ret, str, myLinea, myWhere, idAlmacen, auxCat, myWhereCat, stock
	dim tabla, seeAll, auxCat2
	myLinea = GF_PARAMETROS7("linea",0,6)
	seeAll = GF_PARAMETROS7("all",0,6)
	idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
	str = UCASE(GF_PARAMETROS7("articuloItem" & myLinea,"",6))
	'Se configuran los filtros			
	if (seeAll = 0) then 
		Call mkWhere(myWhere, "A.ESTADO", ESTADO_BAJA, "<>", 1)	
	else
		Call mkWhere(myWhere, "A.ESTADO", "(" & ESTADO_ACTIVO & ", " & ESTADO_BAJA & ")", "IN", 1)
	end if
	'No puede estar en una categoria que sea del tipo I(impuestos)
	auxCat = getCategoriasTipo(TIPO_CAT_IMPUESTOS)	
	if len(auxCat) > 0 then myWhereCat = " and A.IDCATEGORIA NOT IN (" & auxCat & ") "	
	strJoin = "" 
	strFields = ""
	if len(str) > 0 then
		if IsNumeric(str) then 			
			myWhere = myWhere & " and (A.IDARTICULO like '" & str & "%'"
		else
			if len(str) = 1 then
				myWhere = myWhere & " and (DSARTICULO like '" & str & "%'"
			else
				myWhere = myWhere & " and (DSARTICULO like '%" & str & "%'"
			end if	
		end if	
		'Se agrega el filtro por almacen, si se paso alguna
		if (idAlmacen > 0) then 
			myWhere = myWhere & " or D.CDINTERNO like '" & str & "%')"
			strJoin = " LEFT JOIN TBLARTICULOSDATOS D on A.IDARTICULO=D.IDARTICULO and D.IDALMACEN=" & idAlmacen
			strFields = ", (D.EXISTENCIA + D.SOBRANTE) STOCK, D.CDINTERNO"
		else
			myWhere = myWhere & ")"			
		end if
	end if	
	strSQL= "Select A.IDARTICULO, A.DSARTICULO" & strFields & ", '[' + U.ABREVIATURA + ']' ABREVIATURA from TBLARTICULOS A inner join TBLUNIDADES U on A.IDUNIDAD=U.IDUNIDAD" & myWhereCat & strJoin & myWhere
	Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof)
		if (len(ret) > 0) then ret = ret & ";"
		ret = ret & UCase(rs("IDARTICULO")) & "|" & rs("DSARTICULO")
		if (idAlmacen > 0) then 
			if (not isNull(rs("STOCK"))) then
				stock = rs("STOCK")
			else
				stock = 0
			end if
			ret = ret & "|" & rs("CDINTERNO") & "|" & stock 
		end if		
		ret = ret & rs("ABREVIATURA")
		rs.MoveNext()
	wend
	stringArticulos = ret
End Function
'---------------------------------------------------------------------------------------------
Function stringAreasBudget()
	dim strSQL, rs, conn, ret, str
			
	strSQL= "Select * from TBLBUDGETAREAS"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof)
		if (len(ret) > 0) then ret = ret & ";"
		ret = ret & rs("IDAREA") & "-" & rs("DSAREA")
		rs.MoveNext()
	wend
	stringAreasBudget = ret
End Function
'---------------------------------------------------------------------------------------------
Function stringDetalleBudget()
	dim strSQL, rs, conn, ret, str
	strSQL= "Select * from TBLBUDGETDETALLES where IDESTADO=1"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	while (not rs.eof)
		if (len(ret) > 0) then ret = ret & ";"
		ret = ret & rs("IDDETALLE") & "-" & rs("DSDETALLE")
		rs.MoveNext()
	wend
	stringDetalleBudget = ret
End Function
'---------------------------------------------------------------------------------------------
Function stringBudgetNivel1()
	Dim idObra, rs, ret
	idObra = GF_PARAMETROS7("idObra",0,6)
	Set rs = obtenerListaBudgetObra(idObra, 0, 0)	
	while (not rs.eof)
		if (CInt(rs("IDDETALLE")) = 0) then
			if (len(ret) > 0) then ret = ret & ";"
			ret = ret & rs("IDAREA") & "-" & rs("DSBUDGET")
		end if
		rs.MoveNext()
	wend
	stringBudgetNivel1 = ret
End Function
'---------------------------------------------------------------------------------------------
Function stringBudgetNivel2()
	Dim idObra, rs, ret, idArea
	idObra = GF_PARAMETROS7("idObra",0,6)
	idArea = GF_PARAMETROS7("idArea",0,6)
	Set rs = obtenerListaBudgetObra(idObra, idArea, 0)	
	while (not rs.eof)
		if (CInt(rs("IDDETALLE")) <> 0) then
			if (len(ret) > 0) then ret = ret & ";"
			ret = ret & rs("IDDETALLE") & "-" & rs("DSBUDGET")
		end if
		rs.MoveNext()
	wend
	stringBudgetNivel2 = ret
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQArticulos()
	Dim strSQL,articulo
	strSQL = ""
	
	articulo = GF_PARAMETROS7("term","",6)
	idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)
	
	strSQL =          "select uni.* ,cat.dscategoria, art.*,CASE WHEN D.EXISTENCIA is Null then 0 else (D.EXISTENCIA + D.SOBRANTE) end STOCK, D.CDINTERNO "
	strSQL = strSQL & " from tblarticulos art "
	strSQL = strSQL & " inner join tblunidades uni on art.idunidad = uni.idunidad "
	strSQL = strSQL & " inner join tblartcategorias cat on cat.idcategoria = art.idcategoria "
	strSQL = strSQL & " LEFT JOIN TBLARTICULOSDATOS D on art.IDARTICULO=D.IDARTICULO and D.IDALMACEN=" & idAlmacen
	strSQL = strSQL & " where art.estado = " & ESTADO_ACTIVO
	if (isNumeric(articulo)) then
			strSQL = strSQL & " and cast(art.idarticulo AS VARCHAR(10)) like '" & articulo & "%'"
	else
			strSQL = strSQL & " and UPPER(DSARTICULO) LIKE UPPER('%" & articulo & "%')"
	end if
	strSQL = strSQL & " order by art.idcategoria,art.dsarticulo"
	stringJQArticulos = strSQL
	
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQStock()
	almacen = GF_PARAMETROS7("almacen","",6)
	articulo = GF_PARAMETROS7("articulo","",6)
	
	response.write getCantidadArticulosEnAlmacen(almacen,articulo)
End Function 
'--------------------------------------------------------------------------------------------------
Function stringJQObras()
	Dim strSQL, rs, conn, ret, CdObra, dsObra,TipoGasto
	str = UCASE(GF_PARAMETROS7("term","",6))
	TipoGasto = GF_PARAMETROS7("TipoGasto","",6)
	if len(str) > 0 then
		myWhere = " WHERE "
		if len(str) = 1 then
		    myWhere = myWhere & " (CDOBRA like '" & str & "%' or DSOBRA like '" & str & "%') "
		else
			myWhere = myWhere & " (CDOBRA like '%" & str & "%' or DSOBRA like '%" & str & "%') "
		end if			
	end if		
	if(TipoGasto <> "")then myWhere = myWhere &" and TIPOGASTO = '"& TipoGasto &"' "	
	strSQL= "Select CDOBRA, DSOBRA from TBLDATOSOBRAS " & myWhere & " ORDER BY CDOBRA "	
	stringJQObras = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQUnidades()
	Dim strSQL,unidad
	strSQL = ""
	
	unidad = GF_PARAMETROS7("term","",6)
	
	if (isNumeric(unidad)) then
			strSQL = "select * from tblunidades where idunidad = " & unidad & " and ESTADO <> " & ESTADO_BAJA
	else
			strSQL = "select * from tblunidades where UPPER(dsunidad) LIKE UPPER('%" & unidad & "%') and ESTADO <> " & ESTADO_BAJA
	end if
	
	stringJQUnidades = strSQL
end Function
'--------------------------------------------------------------------------------------------------
Function stringJQEmpresas()
	
	Dim strSQL, rs, conn, ret, str, myLinea, myWhere
		
	str = UCASE(GF_PARAMETROS7("term","",6))
	myWhere = " where 1=1"
	if len(str) > 0 then
		if IsNumeric(str) then 			
			myWhere = myWhere & " and (nroemp like '" & str & "%' or nrodoc like '" & str & "%')"
		else
			if len(str) = 1 then
				myWhere = myWhere & " and nomemp like '" & str & "%'"
			else
				myWhere = myWhere & " and nomemp like '%" & str & "%'"
			end if	
		end if	
	end if		
	strSQL= "Select nroemp idempresa, nomemp dsempresa, nrodoc CUIT from [Database].[dbo].met001a " & myWhere & " order by nomemp"
	'Response.Write "<hr>" & strSQL & "<hr>"	
	stringJQEmpresas = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQPersonas()
	Dim strSQL,nombre
	
	strSQL = ""
	
	nombre = GF_PARAMETROS7("term","",6)
	
	strSQL= "Select CDUSUARIO, NOMBRE from VWUSUARIOS where CDUSUARIO + ' - ' + Nombre like '%"&nombre&"%' order by NOMBRE"
	stringJQPersonas = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQLocalidades()
	Dim str,strSQL,cdProvincia,myWhere
	
	str = UCASE(GF_PARAMETROS7("term","",6))
    cdProvincia = UCASE(GF_PARAMETROS7("cdProvincia","",6))

	if len(trim(str)) > 0 then        
        if(cdProvincia <> "")then myWhere = " and PROV.CODIPO = '"& cdProvincia &"'"
		if isNumeric(str) then
			strSQL = "Select PROC.CODIPC as id, PROC.DESCPC as label, PROC.AUXIPC as value, concat(concat('(',TRIM(PROV.DESCPO)), ')') as desc,PROV.CODIPO AS codprov from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE PROC.CODIPC LIKE '%" & str & "%' "& myWhere &" ORDER BY DESCPC"
		else
			strSQL = "Select PROC.CODIPC as id, PROC.DESCPC as label, PROC.AUXIPC as value, concat(concat('(',TRIM(PROV.DESCPO)), ')') as desc,PROV.CODIPO AS codprov from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE DESCPC LIKE '%" & str & "%' "& myWhere &" ORDER BY DESCPC"
		end if
	end if

	stringJQLocalidades = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQPaises()
	Dim str,strSQL

	str = UCASE(GF_PARAMETROS7("term","",6))
	if len(trim(str)) > 0 then
		if isNumeric(str) then
			strSQL = "select AIAFNB id, AIAJTX desc from EJIFL.ACAIREP where AIAFNB = " & str
		else
			strSQL = "select AIAFNB id, AIAJTX desc from EJIFL.ACAIREP where AIAJTX like '%"&str&"%' "
		end if
	end if

	stringJQPaises = strSQL
	
End Function
'--------------------------------------------------------------------------------------------------
Function stringJQProvincias()
	Dim strSQL,str
	
	str = UCASE(GF_PARAMETROS7("term","",6))
	
	strSQL = "select codipo id, descpo desc from MERFL.MER1K2F1 where descpo like '%"+str+"%'"
	
	stringJQProvincias = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function sqlGeneralProveedores(pCodigo,pDesc,pTabla)
	Dim strSQL,str
		
	strSQL = "select "&pCodigo&" CODIGO, "&pDesc&" DESC from "&pTabla&" order by " & pDesc
	
	sqlGeneralProveedores = strSQL
End Function 
'--------------------------------------------------------------------------------------------------
Function stringJQTiposProv()
	stringJQTiposProv = sqlGeneralProveedores("DJASST","DJGATX","provfl.acdjrep")
End Function 
'--------------------------------------------------------------------------------------------------
Function stringJQIssue()
	Dim strSQL,str,myWhere
	str = UCASE(GF_PARAMETROS7("term","",6))
	
	if isNumeric(str) then
		myWhere = "where issueid like '" & str & "%'"
		
	else
		myWhere = "where summary like '%" & str & "%'"
	end if
	
	strSQL =          "SELECT issue.issueid     id    , "
	strSQL = strSQL & "       issue.summary     dc    , "
	strSQL = strSQL & "       status.statusdesc status, "
	strSQL = strSQL & "       issue.issuestatusid statusId, "
	strSQL = strSQL & "		case "
	strSQL = strSQL & "			when users.firstname is null then 'Sin asignar' "
	strSQL = strSQL & "			else users.firstname +','+ users.surname "
	strSQL = strSQL & "		end 'nombre' "
	strSQL = strSQL & "FROM   ( SELECT issueid , "
	strSQL = strSQL & "               summary  , "
	strSQL = strSQL & "               issuestatusid "
	strSQL = strSQL & "       FROM    gemini_issues "
	strSQL = strSQL & "       " & myWhere
	strSQL = strSQL & "       ) "
	strSQL = strSQL & "       issue "
	strSQL = strSQL & "       LEFT JOIN gemini_issueresources b " 'uno con resource para saber a quien esta asignada
	strSQL = strSQL & "       ON     issue.issueid = b.issueid "
	strSQL = strSQL & "       LEFT JOIN gemini_issuestatus status " 'uno con status para traer la descripcion
	strSQL = strSQL & "       ON     status.statusid = issue.issuestatusid "
	strSQL = strSQL & "       LEFT JOIN gemini_users users " 'uno con users para traer la descripcion
	strSQL = strSQL & "       ON     users.userid = b.userid"
	
	stringJQIssue = strSQL
End Function 
'--------------------------------------------------------------------------------------------------
' esta funcion devolvera terminara devolviendo por ajax una respuesta json de la siguiente estructura:'
' [{"cantidad":"9"},{"cantidad":"1"}] siendo el primer objeto la cantidad de req que tienen la tarea,'
' y el segundo objeto indica si la tarea ya esta asignada al requerimiento cuando es mayor a 0 '
Function stringJQExistIssue()
	Dim rtrn,idTarea
	idTarea = UCASE(GF_PARAMETROS7("term","",6))
	idReq = UCASE(GF_PARAMETROS7("nroReq","",6))
	if (idTarea <> "" and idReq <> "") then
		'se utilizan los nros delante del count para que no una las 2 consultas en 1 solo registro '
		' cuando coincidan las cantidades'
		rtrn = "select 1,count(*) cantidad from tblsystareas where idtarea = " & idTarea
		rtrn = rtrn & " union "
		rtrn = rtrn & "select 2,count(*) cantidad from tblsystareas where idtarea = " & idTarea & " and idrequerimiento = " & idreq
		stringJQExistIssue = rtrn
	else
		response.end
	end if

End Function 
'--------------------------------------------------------------------------------------------------
Function stringJQListaCorreo()
	Dim strSQL, rs, conn, ret,DsLista	
	DsLista = GF_PARAMETROS7("term","",6)		
	if len(DsLista) = 1 then
		myWhere = myWhere & " DSLISTA like '" & Trim(Ucase(DsLista)) & "%' "
	else
		myWhere = myWhere & " DSLISTA like '%" & Trim(Ucase(DsLista)) & "%' "
	end if	
	strSQL= "SELECT DsLista FROM TBLMAILLSTCABECERA WHERE " & myWhere 	
	stringJQListaCorreo = strSQL
End function
'--------------------------------------------------------------------------------------------------
Function stringJQAseguradoras()
	Dim strSQL, dsAseguradora	
	dsAseguradora = GF_PARAMETROS7("term","",6)		
	if len(dsAseguradora) = 1 then
		myWhere = myWhere & " DSASEGURADORA like '" & Trim(Ucase(dsAseguradora)) & "%' "
	else
		myWhere = myWhere & " DSASEGURADORA like '%" & Trim(Ucase(dsAseguradora)) & "%' "
	end if	
	strSQL= "SELECT IDASEGURADORA AS id, DSASEGURADORA AS descr FROM TBLPDCASEGURADORAS WHERE " & myWhere	
	stringJQAseguradoras = strSQL
End Function
'--------------------------------------------------------------------------------------------------
Function QueryToJSON(strSQL)
        Dim rs, jsa, col
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)				
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null) (lcase(col.Name)) = trim(col.Value)
                Next
			rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function
'--------------------------------------------------------------------------------------------------
Function QueryToJSON2( strSQL)
        Dim rs, jsa, col
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null) (lcase(col.Name)) = trim(col.Value)
                Next
			rs.MoveNext
        Wend
        Set QueryToJSON2 = jsa
End Function
'--------------------------------------------------------------------------------------------------
Function QueryToJSON3( strSQL)
        Dim rs, jsa, col
        Call GF_BD_GEMINI(rs,"OPEN",strSQL)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null) (lcase(col.Name)) = trim(col.Value)
                Next
			rs.MoveNext
        Wend
        Set QueryToJSON3 = jsa
End Function
'*****************************************************
'***** 	COMIENZO DE LA PAGINA
'*****************************************************
Dim tipo, rtrn

tipo = GF_PARAMETROS7("tipo","",6)

rtrn=""
Select Case (tipo)
	Case "articulos":
		rtrn = stringArticulos()
	Case "categorias":
		rtrn = stringCategorias()
	Case "empresas":
		rtrn = stringEmpresas()
	Case "unidades":		
		rtrn = stringUnidades()
	Case "personas":
		rtrn = stringPersonas()
	Case "solicitantes":
		rtrn = stringSolicitantes()
	Case "obras":
		rtrn = stringObras()
	Case "aBudget":
		rtrn = stringAreasBudget()
	Case "dBudget":
		rtrn = stringDetalleBudget()
	Case "BudgetNivel1":
		rtrn = stringBudgetNivel1()
	Case "BudgetNivel2":
		rtrn = stringBudgetNivel2()
	Case "JQArticulos"	:
		rtrn = stringJQArticulos()
		QueryToJSON(rtrn).Flush
		response.end
	Case "JQUnidades":
		rtrn = stringJQUnidades()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQEmpresas":
		rtrn = stringJQEmpresas()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQPersonas":
		rtrn = stringJQPersonas()
		QueryToJSON2(rtrn).Flush
		response.end
	case "JQLocalidades":
		rtrn = stringJQLocalidades()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQpaises":
		rtrn = stringJQPaises()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQTipoProveedores":
		rtrn = stringJQTiposProv()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQStockArticulo"
		Call stringJQStock()
		response.end
	case "JQIssue":
		rtrn = stringJQIssue()
		QueryToJSON3(rtrn).Flush
		response.end
	case "JQProvincias":
		rtrn = stringJQProvincias()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQExisteIssue":
		rtrn = stringJQExistIssue()
		QueryToJSON(rtrn).Flush
		response.end
	case "JQObras":
		rtrn = stringJQObras()
		QueryToJSON(rtrn).Flush
		response.end	
	case "JQListaCorreo":
		rtrn = stringJQListaCorreo()
		QueryToJSON(rtrn).Flush
		response.end		
	case "JQAseguradoras":
		rtrn = stringJQAseguradoras()
		QueryToJSON(rtrn).Flush
		response.end			
End Select	
response.write rtrn
%>