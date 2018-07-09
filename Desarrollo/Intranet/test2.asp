<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosCupos.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
pPto="ARROYO"
pCdProducto = 19
pFecha = 20171213

		myKey = getLetraCupo2(pPto)
		strSQL="Select DSPRODUCTO from PRODUCTOS where CDPRODUCTO=" & pCdProducto        
		Call executeQueryDb(pPto, rs, "OPEN", strSQL)
		if (not rs.eof) then 
			myKey = myKey & Left(Trim(rs("DSPRODUCTO")), 1)
		else
			myKey = myKey & "X"
		end if
		diff = GF_DTEDIFF(20170101,pFecha,"D")	
		colName = "M" & Right("0" & CStr(((CLng(pFecha) mod 100) mod 12) + 1), 2)		
		strSQL ="UPDATE  CODIGOSCUPO SET CODIGOCUPO = concat('" & myKey & "', Right(concat('0000', '" & diff & "'), 4), Right(concat('0000', B." & colName & "), 4)) " &_
				" FROM    CODIGOSCUPO a " &_
				"	INNER JOIN CODIGOSCUPOMTX b " &_
				"	ON (A.IDCUPO % 10000) = B.IDX " &_
				" where CODIGOCUPO='PROVISORIO' and CDPRODUCTO=" & pCdProducto		
		response.write strSQL

response.end

Dim mtx(10000)

Randomize
'For M = 1 to 12	
	For i= 0 to 9999
		idx = round(rnd * 10000, 0)
		idxOriginal = idx
		flg = true
		ecode = 0
		while (flg)
			if (mtx(idx) <> "") then
				idx = idx + 1
				if (idx > 9999) then idx = 0
				if (idx = idxOriginal) then 
					flg = false
					ecode = 1
				end if
			else
				flg = false
			end if
		wend
		if (ecode = 0) then 
			mtx(idx) = i
			'Call executeQueryDb(DBSITE_ARROYO, rs, "EXEC", "Insert into CODIGOSCUPOMTX(IDX, M01) values (" & i & ", " & idx & ")")			
			Call executeQueryDb(DBSITE_ARROYO, rs, "EXEC", "Update CODIGOSCUPOMTX SET M04 = " & idx & " where IDX = " & i)
		else
			idx = "ERROR!"
		end if
		'response.write "NRO: " & i & " - Idx Proopuesto: " & idxOriginal & " - Idx Asignado: " & idx & "<br>"				
	Next
	response.write "--- FIN ---<br>"

%>

