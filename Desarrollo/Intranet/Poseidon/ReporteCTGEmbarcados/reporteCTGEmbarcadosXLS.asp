<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosUser.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosExcel.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<%
'ProcedimientoControl "GVPERMISOS"
dim cdProducto, dsProducto, cdCliente, dsCliente, cdAviso, cdBuque, dsBuque, kilos, mySelected, pKilosSobrantes, msjEmbarcados
dim myTableHTML, dicErr, accion, mySaveText, puerto, cont, color, secuencia, tipo, myName
Set dicErr = Server.CreateObject("Scripting.Dictionary")

cdProducto = GF_Parametros7("cdProducto", 0, 6)
dsProducto = GF_Parametros7("dsProducto", "", 6)
cdCliente = GF_Parametros7("cdCliente", 0, 6)
cdAviso = GF_Parametros7("cdAviso", 0, 6)
secuencia = GF_Parametros7("secuencia", "", 6)
accion = GF_Parametros7("accion", "", 6) 
puerto = GF_Parametros7("Pto", "", 6) 
dsCliente = getDsCliente(cdCliente)
if not loadDatosBuque(cdAviso, cdBuque, dsBuque) then
	setError(AVISO_NO_EXISTE)
end if
kilos = GF_Parametros7("kilos", 0, 6)
if kilos = 0 then setError(KILOS_NO_CARGADOS)

myTableHTML =   myTableHTML &	" <tr style='font-family:courier;'> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("Camion")			& "</b></td> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("Acoplado")			& "</b></td> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("Carta Porte")		& "</b></td> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("CTG")				& "</b></td> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("Producto")			& "</b></td> " & _
								"	<td bgcolor='#ffeecd' align='center'><b>" & GF_Traducir("Kilos")			& "</b></td> " & _
								"</tr>	"	
tipo = GF_Parametros7("tipo", "", 6) 
'Tipos
'S = Secuencia
'C= Cliente
'P= Producto
'A= Aviso
select case tipo
	case "A"
		myName = "AVISO_" & cdAviso
	case "C"
		myName = dsCliente 
	case "P"
		myName = dsCliente & "_" & dsProducto  		
	case "S"
		myName = "AVISO_" & cdAviso & "_SEC_" & secuencia 
	case else
		myName = "NO_DEFINIDO"		
end select	
if accion = "HIS2" then
		clearError
		Call GF_createXLS(myName)
		Call CargarHistorico(cdAviso, secuencia)
else
	if not hayError() then
			Call GF_createXLS(myName)
			kilosSobrantes = 0
			myQuery = ArmarCtaCteCTG(cdBuque,cdProducto,cdCliente, kilos, "MOSTRAR", kilosSobrantes, cdAviso, cdCosecha) 
			msjEmbarcados = KilosEmbarcados(myQuery, kilos, kilosSobrantes)
			If msjEmbarcados = kilos Then
			    KilosAlcanzado = True
			Else
			    KilosAlcanzado = False
			    msjEmbarcados = clng(kilos) - clng(msjEmbarcados)
			End If	
	end if
end if

'----------------------------------------------------------------------------------------
Function KilosEmbarcados(pSQLQuery, pKilosDefault, pKilosSobrantes)
' pKilosSobrantes As Double = 0) As String
Dim rs, NColumnas, acum, dif
	call GF_BD_Puertos(puerto, rs, "OPEN",pSQLQuery)
    while not rs.EOF
        if(cdbl(pKilosSobrantes) <> 0) then
			acum = acum + cdbl(pKilosSobrantes)
		else
			acum = acum + cdbl(rs("KILOSNETOS"))
		end if	
		cont = cont + 1 
		if cont mod 2 then
			color = "#ffffff"
		else
			color =	"#dcdcdc"
		end if	
        If clng(acum) >= clng(pKilosDefault) Then
            Dif = (clng(rs("KILOSNETOS")) - (clng(acum) - clng(pKilosDefault)))
			myTableHTML =   myTableHTML & "<tr style='font-family:courier; font-size:10;'>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_PATENTE(rs("CDCHAPACAMION")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_PATENTE(rs("CDCHAPAACOPLADO")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_CTAPTE(rs("CARTAPORTE")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & rs("CTG") & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & rs("DSPRODUCTO") & "</td>" & _
							"	<td bgcolor=" & color & " align=right>" & GF_EDIT_DECIMALS(clng(Dif)*100,2) & "</td>" & _
							"</tr>"
            'Guarda los kilos embarcados y Sale de la funcion
            If(Dif <> 0) then 
				KilosEmbarcados = clng(Dif) + (clng(acum) - clng(rs("KILOSNETOS")))
			else
				KilosEmbarcados = acum
			end if	
            Exit Function
        Else
			if cdbl(pKilosSobrantes) <> 0 then 
				kilosFinal = pKilosSobrantes 
				pKilosSobrantes = 0
			else
				kilosFinal = rs("KILOSNETOS")
			end if	
			myTableHTML =   myTableHTML & "<tr style='font-family:courier; font-size:10;'>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_PATENTE(rs("CDCHAPACAMION")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_PATENTE(rs("CDCHAPAACOPLADO")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & GF_EDIT_CTAPTE(rs("CARTAPORTE")) & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & rs("CTG") & "</td>" & _
							"	<td bgcolor=" & color & " align=center>" & rs("DSPRODUCTO") & "</td>" & _
							"	<td bgcolor=" & color & " align=right>" & GF_EDIT_DECIMALS(cdbl(kilosFinal)*100,2) & "</td>" & _
							"</tr>"
            rs.MoveNext
            pKilosSobrantes = 0
        End If
    wend
    KilosEmbarcados = acum
    rs.Close
    Set rs = Nothing
End Function
'---------------------------------------------------------------------------------------
function getDsCliente(pCdCliente)
Dim rs, strSQL, rtrn
rtrn = "ERROR"
strSQL = "SELECT DSCLIENTE FROM CLIENTES WHERE CDCLIENTE=" & pCdCliente
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
if not rs.eof then 
	rtrn = trim(rs("DSCLIENTE"))
end if	
rtrn = replace(rtrn,",","")
rtrn = replace(rtrn,".","")
rtrn = replace(rtrn," ","")
getDsCliente = rtrn 
end function
'---------------------------------------------------------------------------------------
Sub CargarHistorico(cdAviso, secuencia)
Dim rs, pQuery, MsjEmbarcados, pKilosSobrantes, myKilosNetos, myAnd
myAnd = ""
if secuencia <> "" then 
	myAnd = " AND SECUENCIA IN (" & secuencia & ")"
end if	
pKilosSobrantes = 0
pQuery = "SELECT SUM(KILOSNETOS) KILOSNETOS FROM CTGEMBARCADOS CTG WHERE CDAVISO= " & cdAviso & myAnd 
call GF_BD_Puertos(puerto, rs, "OPEN",pQuery)
if not rs.EOF then myKilosNetos = rs("KILOSNETOS")
pQuery =	" SELECT DTCONTABLE, IDCAMION, CDCHAPACAMION, CDCHAPAACOPLADO, NUCARTAPORTE AS CARTAPORTE, CTG, C.CDCLIENTE, C.DSCLIENTE, P.CDPRODUCTO, P.DSPRODUCTO, KILOSNETOS " & _
			"	FROM CTGEMBARCADOS CTG " & _
			"		LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = CTG.CDPRODUCTO " & _
			"		LEFT JOIN CLIENTES C ON C.CDCLIENTE = CTG.CDCLIENTE " & _
			"	WHERE CDAVISO= " & cdAviso & myAnd & _
			"	ORDER BY DTCONTABLE ASC, CDPRODUCTO, CDCLIENTE"
'Response.Write pQuery
MsjEmbarcados = KilosEmbarcados(pQuery, myKilosNetos, pKilosSobrantes)
End Sub
'--------------------------------------------------------------------------------------------------------------------
Function ArmarCtaCteCTG(pCdBuque, pCdproducto, pCdCliente, pKilosACargar, pTipoQuery, byref pKilosSobrantes, pCdAviso, pCdCosecha)
Dim rs, strSQL, strSql2, strSqlSobrantes, fechaInicio, idCamionInicio, kilosCargados
dim auxWhere1, auxWhere2, auxSql1
If cdCamionesDe <> 0 Then 
	auxWhere1 = " AND CDCLIENTE=" & cdCamionesDe
	auxWhere2 = " AND HCD.CDCLIENTE=" & cdCamionesDe
end if	
If pCdCosecha <> 0 Then 
	auxWhere1 = auxWhere1 & " AND CDCOSECHA=" & pCdCosecha
	auxWhere2 = auxWhere2 & " AND HCD.CDCOSECHA=" & pCdCosecha
end if	
fechaInicio = "2010-03-01"
	'MOSTRAR
		strsql =	"SELECT * FROM (" & _
					"SELECT (YEAR(TG.DTCONTABLE)*10000 + Month(TG.DTCONTABLE)*100 + DAY(TG.DTCONTABLE)) DTCONTABLE, TG.IDCAMION, TG.CDCHAPACAMION, TG.CDCHAPAACOPLADO, TG.CDCOSECHA, " & _
					"	TG.CARTAPORTE, TG.CTG, TG.DSCLIENTE, TG.DSPRODUCTO, CASE WHEN TG.KILOSCARGADOS IS NULL THEN TG.KILOSNETOS ELSE TG.KILOSNETOS-TG.KILOSCARGADOS END AS KILOSNETOS " & _ 
					"	  FROM " & _
					"	( " & _
					"	    SELECT HCD.DTCONTABLE, HCD.IDCAMION, HC.CDCHAPACAMION, HC.CDCHAPAACOPLADO, HCD.CDCOSECHA, " & _ 
					"	        RTRIM(HCD.NUCARTAPORTE) + RTRIM(HCD.NUCTAPTEDIG) AS CARTAPORTE, HCD.CTG, C.DSCLIENTE, P.DSPRODUCTO, " & _ 
					"	        ( " & _ 
					"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) " & _
					"	            -  " & _
					"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2))  " & _
					"	            -  " & _
					"	            ( SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION)) " & _
					"	        ) KILOSNETOS , EMBARCADOS.KILOSCARGADOS  " & _
					"	    FROM HCAMIONESDESCARGA HCD " & _ 
					"	        LEFT JOIN  " & _
					"	            (SELECT IDCAMION, DTCONTABLE, SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS GROUP BY IDCAMION, DTCONTABLE) " & _
					"	                EMBARCADOS ON HCD.IDCAMION = EMBARCADOS.IDCAMION AND HCD.DTCONTABLE=EMBARCADOS.DTCONTABLE " & _
					"	        LEFT JOIN HCAMIONES HC ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE=HCD.DTCONTABLE  " & _
					"	        LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO  " & _
					"	        LEFT JOIN CLIENTES C ON C.CDCLIENTE = HCD.CDCLIENTE "
    strSql2 = strSQL &  " WHERE HCD.DTCONTABLE >='" & fechaInicio & "'"
   
	strSql2 = strSql2 & auxWhere2
	strSql2 = strSql2 & " AND HC.CDPRODUCTO = " & pCdproducto & _
						" AND HC.CDESTADO IN (6,8) " & _
						" ) TG)TA WHERE KILOSNETOS>0 ORDER BY DTCONTABLE ASC, IDCAMION ASC"
    ArmarCtaCteCTG = strSql2
End Function
'---------------------------------------------------------------------------------------
function loadDatosBuque(pCdAviso, byref pCdBuque, byref pDsBuque)
dim rs2, strSQL, rtrn
rtrn = false
strSQL =	" SELECT 1 AS ETAPA, BUQ.CDBUQUE, BUQ.DSBUQUE FROM EMBARQUES EMB INNER JOIN BUQUES BUQ ON EMB.CDBUQUE=BUQ.CDBUQUE WHERE EMB.CDAVISO=" & pCdAviso & _
			" UNION " & _
			" SELECT 2 AS ETAPA, BUQ.CDBUQUE, BUQ.DSBUQUE FROM HEMBARQUES EMB INNER JOIN BUQUES BUQ ON EMB.CDBUQUE=BUQ.CDBUQUE WHERE EMB.CDAVISO=" & pCdAviso & _
			" ORDER BY ETAPA ASC "
call GF_BD_Puertos(puerto, rs, "OPEN",strSQL)
if not rs.EOF then
	pCdBuque = rs("CDBUQUE")
	pDsBuque = rs("DSBUQUE")
	rtrn = true
end if	
loadDatosBuque = rtrn	
end function
'---------------------------------------------------------------------------------------
%>
<HTML>
<HEAD>
   <TITLE>Reporte de CTGs Embarcados</TITLE>
</HEAD>
<BODY>
   <table border="1" cellpadding="0" cellspacing="0" width="60%">
      <tr>
	    <td bgcolor="#517B4A" align="center" colspan="6">
			<b>
				<font color="white" style="font-family:courier;" size="5">
					Reporte embarcados
				</font>
			</b>
		</td>
      </tr>
      <tr>
	    <td align="left" colspan="6">
				<b>
					<font style="font-family:courier;" size="1">Fecha.............:</font>
				</b>
				<font style="font-family:courier;" size="1">
					<%=GF_EDIT_FECHA(date())%>
				</font>
			<br>
				<b>
					<font style="font-family:courier;" size="1">Hora..............:</font>
				</b>
				<font style="font-family:courier;" size="1">
					<%=Time()%>
				</font>
			<br>
				<b>
					<font style="font-family:courier;" size="1">Buque.............:</font>
				</b>
				<font style="font-family:courier;" size="1">
					<%=dsBuque%>
				</font>
			<br>
				<b>
					<font style="font-family:courier;" size="1">Kilos embarcados..:</font>
				</b>
				<font style="font-family:courier;" size="1">
					<%=GF_EDIT_DECIMALS(cdbl(kilos)*100,2)%>
				</font>
		</td>
      </tr>
		<%
		if not hayError() then 
			if myTableHTML = "" then
				if mySaveText = "" then
					Response.Write "<tr><td align=center>No se encontraron camiones</td></tr>"
				else
					Response.Write mySaveText
				end if	
			else	
				Response.Write myTableHTML
			end if	
		end if
		%>
      <tr>
	    <td bgcolor="#ffeecd" align="center" colspan="6">
			<b>
				<font style="font-family:courier;" size="1">
					Fin Reporte
				</font>
			</b>
		</td>
      </tr>		
   </table>
</form>
</body>
</html>
