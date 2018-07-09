<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
dim idDivision, idAlmacen, fechaCierre, fechaCierreAnterior, myPath, fs, fname, myAllFixes

idDivision = GF_Parametros7("idDivision", 0, 6)
idAlmacenes = GF_Parametros7("idAlmacen", "", 6)
fechaCierre = GF_Parametros7("fechaCierre", "", 6)
fechaCierreAnterior = GF_Parametros7("fechaCierreAnt", "", 6)

myAllFixes = "V"

set fs = Server.CreateObject("Scripting.FileSystemObject")
myName = "logs/CierresContables_PreInicializacion_" & Session("MmtoDato") & ".txt"
myPath = Server.MapPath(myName)
if fs.FileExists(myPath) then
    set fname = fs.OpenTextFile(myPath,8,true)
else
    set fname = fs.CreateTextFile(myPath,true)
end if

fname.WriteLine("******************************************************************************************************")
fname.WriteLine("                                  "& getDivisionDS(idDivision) &"                            ")
fname.WriteLine("******************************************************************************************************")


strSQL =	"SELECT VC.IDVALE, VC.NRVALE, VC.CDVALE, VC.FECHA, ART.IDARTICULO, ART.DSARTICULO, VD.CANTIDAD, UNI.ABREVIATURA, ALM.IDDIVISION FROM TBLVALESCABECERA VC " & _ 
			"	INNER JOIN TBLVALESDETALLE VD ON VC.IDVALE = VD.IDVALE " & _ 
			"	INNER JOIN TBLARTICULOS ART ON ART.IDARTICULO=VD.IDARTICULO " & _ 
			"	INNER JOIN TBLUNIDADES UNI ON ART.IDUNIDAD=UNI.IDUNIDAD " & _ 
			"	INNER JOIN TBLALMACENES ALM ON VC.IDALMACEN=ALM.IDALMACEN " & _ 
			"		WHERE VC.IDALMACEN IN (" & idAlmacenes & ") AND VC.FECHA LIKE '" & left(fechaCierre,6) & "%' AND VC.ESTADO=" & ESTADO_ACTIVO & _ 
			"			AND VC.CDVALE IN ('" & CODIGO_VS_SALIDA & "','" & CODIGO_VS_AJUSTE_VALE & "','" & CODIGO_VS_AJUSTE_TRANSFERENCIA & "','" & CODIGO_VS_AJUSTE_STOCK & "') AND VD.VLUPESOS=0 AND VD.EXISTENCIA<>0 " & _ 
			"ORDER BY VC.FECHA ASC "
'Response.Write strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
    fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine(" ITEMS NO VALORIZADOS")
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine("Mes de Cierre.: " & mid(fechaCierre,5,2) & "/" & left(fechaCierre,4))
	fname.WriteLine("Division......: " & getDivisionDS(idDivision))
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine("IDVALE" & vbTab & "NROVALE" & vbTab & vbTab & "CDVALE" & vbTab & "FECHA" & vbTab & vbTab & "IDART" & vbTab & "DSARTICULO" & string(30," ") & vbTab & "CANT" & vbTab & "PRECIO")

    while not rs.eof
		    myPrecioPesos = getUltimoPrecio(rs("IDDIVISION"), rs("IDARTICULO"), MONEDA_PESOS, rs("FECHA"))
		    if myPrecioPesos <> 0 then 			
			    myTipoCambio = getTipoCambioCV(MONEDA_DOLAR, rs("FECHA"), T_CAMBIO_COMPRADOR)
			    myPrecioDolares = round(clng(myPrecioPesos) / cdbl(myTipoCambio),0)
			    call setPreciosVigentesPorArticulo(rs("IDVALE"), trim(rs("IDARTICULO")), myPrecioPesos, myPrecioDolares)		
			    myLabel = myPrecioPesos & " - FIX" 
		    else	
			    descripcion = trim(rs("DSARTICULO"))
			    if len(descripcion)<40 then
				    descripcion = descripcion & string(40-len(descripcion)," ")
			    else
				    descripcion = left(descripcion,40)
			    end if	
			    myLabel = "0"
			    myAllFixes = "F"
		    end if	
		    fname.WriteLine(rs("IDVALE") & vbTab & rs("NRVALE") & vbTab & rs("CDVALE") & vbTab & rs("FECHA") & vbTab & rs("IDARTICULO") & vbTab & descripcion & vbTab & rs("CANTIDAD") & rs("ABREVIATURA") & vbTab & myLabel)
	    rs.movenext
    wend
    fname.WriteLine("")
end if
	

strSQL =	"SELECT * FROM " & _
			"	( " & _
			"		SELECT REM.IDREMITO, REM.NROREMITO, REM.IDALMACEN, REM.FECHA, REL.IDPIC, PIC.IDCOTIZACION FROM TBLREMCABECERA REM " & _
			"			LEFT JOIN TBLREMPIC REL ON REM.IDREMITO=REL.IDREMITO " & _
			"			LEFT JOIN TBLCTZCABECERA PIC ON REL.IDPIC=PIC.IDCOTIZACION " & _
			"				WHERE REM.IDALMACEN IN (" & idAlmacenes & ") AND REM.FECHA LIKE '" & left(fechaCierre,6) & "%' AND REM.ESTADO=" & ESTADO_ACTIVO & _ 
			"	) ALLS " & _
			"	WHERE ALLS.IDPIC IS NULL OR ALLS.IDCOTIZACION IS NULL ORDER BY FECHA"
'Response.Write STRsql
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if not rs.eof then
	
	'if not isObject(fname) then
		'myName = "logs/CierresContables_" & Year(now) & Month(now) & Day(now) & Hour(now) & Minute(now) & Second(now) & ".txt"
		'myName = "logs/CierresContables_RemSinPic_(" & idDivision & ")_" & replace(GF_MOMENTOSISTEMA,"'","") & ".txt"
		'set fname = fs.CreateTextFile(myPath,true)
	'else
	'	fname.WriteLine("")
	'end if
    fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine(" REMITOS SIN PICS")
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
    fname.WriteLine("Mes de Cierre.: " & mid(fechaCierre,5,2) & "/" & left(fechaCierre,4))
	fname.WriteLine("Division......: " & getDivisionDS(idDivision))
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine("IDREM" & vbTab & "NROREM" & vbTab & vbTab & "IDALMACEN" & vbTab & "FECHA")
    while not rs.eof
			nroRemito = trim(rs("NROREMITO"))
			if len(nroRemito)<10 then
				nroRemito = nroRemito & string(10-len(nroRemito)," ")
			else
				nroRemito = left(nroRemito,10)
			end if	
			myAllFixes = "F"
		fname.WriteLine(rs("IDREMITO") & vbTab & nroRemito & vbTab & rs("IDALMACEN") & vbTab & vbTab & rs("FECHA"))
	    rs.movenext
    wend
    fname.WriteLine("")
end if



'CONTROL N° 3:
'Se crea un control de firmas, permitiendo validar los vales del mes que esten todos firmados, caso contrario se lo informa en el log
strSQL = "SELECT IDVALE, CDVALE, FECHA, ESTADO, NRVALE, IDALMACEN "&_
         "FROM TBLVALESCABECERA "&_
         "WHERE FECHA LIKE '"& left(fechaCierre,6) &"%' "&_
         "  AND IDALMACEN IN (" & idAlmacenes & ")" &_
         "  AND ESTADO=" & ESTADO_ACTIVO &_
	     "  AND IDVALE IN (SELECT IDVALE "&_ 
         "                   FROM TBLVALESFIRMAS "&_
         "                   WHERE (HKEY IS NULL OR HKEY = ''))"
'Response.Write strSQL
Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
if (not rs.Eof) then
    fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine(" VALES SIN AUTORIZAR")
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine("Mes de Cierre.: " & mid(fechaCierre,5,2) & "/" & left(fechaCierre,4))
	fname.WriteLine("Division......: " & getDivisionDS(idDivision))
	fname.WriteLine("------------------------------------------------------------------------------------------------------")
	fname.WriteLine("IDVALE" & vbTab & "NROVALE" & vbTab & vbTab & "TIPO" & vbTab & "IDALMACEN"& vbTab & "FECHA")
    while not rs.eof
        fname.WriteLine(rs("IDVALE") & vbTab & Trim(rs("NRVALE")) & vbTab & Trim(rs("CDVALE")) & vbTab & rs("IDALMACEN") & vbTab& vbTab & GF_FN2DTE(rs("FECHA")))
	    rs.MoveNext()
    wend
    myAllFixes = "F"
end if
fname.WriteLine("")
fname.Close
set fname = Nothing

Response.Write myAllFixes & "-" & myName
Response.End 



%>
