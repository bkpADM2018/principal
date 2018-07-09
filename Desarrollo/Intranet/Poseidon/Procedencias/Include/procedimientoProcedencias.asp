<%
dim gCdProcedencia, gDsProcedencia
dim gCdProcedenciaCamara, gDsProcedenciaCamara, gCdProcedenciaCamaraAsoc, gCdLocalidadCamara, gDsLocalidadCamara, gCdProvinciaCamara
dim gCdProcedenciaOncca, gDsProcedenciaOncca
dim gCdProvincia, gDsProvincia

'-----------------------------------------------------------------------------------------------------------------------------
sub clearDatos()
		gCdProcedencia = ""
		gDsProcedencia = ""
		gCdSubProcedencia = ""
		gCdProcedenciaCamara = 0
		gDsProcedenciaCamara = ""
		gCdProcedenciaCamaraAsoc = 0
		gCdLocalidadCamara = ""
		gDsLocalidadCamara = "" 
		gCdProvinciaCamara = ""
		gDsProvinciaCamara = ""
		gCdProcedenciaOncca = ""
		gDsProcedenciaOncca = ""		
		gCdProvincia = 0
		gDsProvincia = ""		
end sub
'-----------------------------------------------------------------------------------------------------------------------------
Function getProcedenciaByCdProcedencia(pCdProcedencia,pPto)
	Dim strSQL,rs
	getProcedenciaByCdProcedencia = false
	strSQL = "SELECT PROCE.CDPROCEDENCIA, PROCE.CDSUBPROCEDENCIA, PROCE.DSPROCEDENCIA, PROCE.CDPROCEDENCIACAMARA, " & _
			 "		LOCC.CDLOCALIDADCAMARA, LOCC.DSLOCALIDAD, PCIA.CDPROVINCIA, PCIA.DSPROVINCIA, L.IDPROV, PROCE.IDLOCONCCA, L.DSLOC " & _
			 " FROM dbo.PROCEDENCIAS PROCE " & _
			 "	Left join LOCPROVPART L on L.IDLOC=PROCE.IDLOCONCCA " & _
			 "	LEFT JOIN dbo.PROVINCIAS PCIA ON PROCE.CDPROV=PCIA.CDPROVINCIA " & _			 
			 "	LEFT JOIN dbo.LOCALIDADESCAMARAS LOCC ON CAST(SUBSTRING(PROCE.CDPROCEDENCIACAMARA,1,4) AS INT)=LOCC.CDLOCALIDADCAMARA AND CAST(SUBSTRING(PROCE.CDPROCEDENCIACAMARA,4,3) AS INT)=LOCC.CDLOCALIDADSUBCAMARA" & _
			 " WHERE CDPROCEDENCIA = " & pCdProcedencia
    call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)		
	call clearDatos
	if not rs.eof then
		gCdProcedencia = trim(rs("CDPROCEDENCIA"))	
		gCdSubProcedencia = trim(rs("CDSUBPROCEDENCIA"))	
		gDsProcedencia = trim(rs("DSPROCEDENCIA"))
		gCdProcedenciaCamara = trim(rs("CDPROCEDENCIACAMARA"))
		gCdProcedenciaCamaraAsoc = trim(rs("CDPROCEDENCIACAMARA"))	
		gCdLocalidadCamara = trim(rs("CDLOCALIDADCAMARA"))	
		gDsLocalidadCamara = trim(rs("DSLOCALIDAD"))	
		gCdProvinciaCamara = ""		
		gCdProcedenciaOncca = rs("IDLOCONCCA")		
		gDsProcedenciaOncca = rs("DSLOC")		
		gCdProvincia = verNull(trim(rs("CDPROVINCIA")),"N")	
		gDsProvincia = trim(rs("DSPROVINCIA"))			
		getProcedenciaByCdProcedencia = true
	end if
End Function	
'-----------------------------------------------------------------------------------------------------------------------------
sub getProcedenciaByParams()
	call clearDatos
	gCdProcedencia			 = GF_Parametros7("gCdProcedencia", 0, 6)
	gCdSubProcedencia		 = GF_Parametros7("gCdSubProcedencia", 0, 6)
	gDsProcedencia			 = ucase(GF_Parametros7("gDsProcedencia", "", 6))
	gCdProcedenciaCamara	 = GF_Parametros7("gCdProcedenciaCamara", 0, 6)
	gDsProcedenciaCamara	 = GF_Parametros7("gDsProcedenciaCamara", "", 6)
	gCdProcedenciaCamaraAsoc = GF_Parametros7("gCdProcedenciaCamaraAsoc", 0, 6)
	gCdLocalidadCamara		 = GF_Parametros7("gCdLocalidadCamara", 0, 6)
	gDsLocalidadCamara		 = GF_Parametros7("gDsLocalidadCamara", "", 6)
	gCdProvinciaCamara		 = GF_Parametros7("gCdProvinciaCamara", "", 6)
	gDsProvinciaCamara		 = GF_Parametros7("gDsProvinciaCamara", "", 6)
	gCdProcedenciaOncca		 = GF_Parametros7("gCdProcedenciaOncca", "", 6)
	gDsProcedenciaOncca		 = GF_Parametros7("gDsProcedenciaOncca", "", 6)	
	gCdProvincia			 = GF_Parametros7("gCdProvincia",0, 6)
	gDsProvincia			 = GF_Parametros7("gDsProvincia", "", 6)	
End sub	
'----------------------------------------------------------------------------------------------------------------------
Function checkProcedencia()
	Dim rsPro	
	checkProcedencia = false
	if len(TRIM(gDsProcedencia)) < 1 then call setError(DSPROCEDENCIA_VACIO)
	if clng(gCdProvincia) = 0 then call setError(CDPROVINCIA_VACIO)
	if (CLng(gCdProcedenciaOncca) = 0) then setError(FALTA_PROCEDENCIA_ONCCA)
	if clng(gCdProcedenciaCamaraAsoc) = 0 then call setError(CDPROCEDENCIACAMARA_VACIO)
	if (not hayError) then checkProcedencia = true
End Function
'-------------------------------------------------------------------------------------------------------------------
Function addProcedencia()
	Dim strSQL, cdProcedenciaGen, cdSubProcedenciaGen
	strSQL = "SELECT MAX(CDPROCEDENCIA) as ULTIMO FROM PROCEDENCIAS " & _
			 "	WHERE CDPROCEDENCIA BETWEEN " & gCdLocalidadCamara & "000 AND " & gCdLocalidadCamara & "999 " 
	call GF_BD_Puertos (g_strPuerto, rs, "OPEN",strSQL)		
	if not rs.eof then		
		if (RS("ULTIMO") <> "") then 
			cdProcedenciaGen = cdbl(gCdLocalidadCamara) * 1000
			cdSubProcedenciaGen = "000"
		else
			cdProcedenciaGen = CDbl(RS("ULTIMO")) + 1		
			cdSubProcedenciaGen = right(cdProcedenciaGen,3)
		end if
	else		
		cdProcedenciaGen = cdbl(gCdLocalidadCamara) * 1000
		cdSubProcedenciaGen = "000"
	end if	
	gCdProcedencia = cdProcedenciaGen
	gCdSubProcedencia = cdSubProcedenciaGen
	strSQL = "INSERT INTO PROCEDENCIAS(CDPROCEDENCIA,DSPROCEDENCIA,CDSUBPROCEDENCIA,CDPROCEDENCIACAMARA,CDPROV,CDPART, IDLOCONCCA) "&_
			 "VALUES (" & gCdProcedencia & ", '" & gDsProcedencia & "', '" & gCdSubProcedencia & _
			 "', '" & gCdProcedenciaCamaraAsoc & "', " & gCdProvincia & ", 0, " & gCdProcedenciaOncca & ")"	
	call consolidarProcedencias(strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function updateProcedencia()
	Dim strSQL
	gCdSubProcedencia = right(gCdProcedencia,3)
	strSQL = "UPDATE PROCEDENCIAS SET DSPROCEDENCIA= '" & gDsProcedencia & "' ,CDSUBPROCEDENCIA='" & gCdSubProcedencia & "', " & _
			 "	CDPROCEDENCIACAMARA='" & gCdProcedenciaCamaraAsoc & "', IDLOCONCCA=" & gCdProcedenciaOncca & ", CDPROV=" & gCdProvincia & _
			 " WHERE CDPROCEDENCIA = " & gCdProcedencia
	call consolidarProcedencias(strSQL)
End Function
'-------------------------------------------------
function verNull(pValue, pTipo)
dim rtrn
rtrn = pValue
if isnull(rtrn) then 
	select case pTipo
		case "N" 
			rtrn = 0
		case "S"
			rtrn = ""
	end select
end if		
verNull = rtrn		
end function
'-----------------------------------------------------------------------------------
sub consolidarProcedencias(strsql)
'Response.Write strSQL & "-" & TERMINAL_ARROYO	
Call GF_BD_Puertos (TERMINAL_ARROYO, rs, "EXEC",strSQL)

'Response.Write strSQL & "-" & TERMINAL_TRANSITO	
Call GF_BD_Puertos (TERMINAL_TRANSITO, rs, "EXEC",strSQL)

'Response.Write strSQL & "-" & TERMINAL_PIEDRABUENA	
Call GF_BD_Puertos (TERMINAL_PIEDRABUENA, rs, "EXEC",strSQL)
end sub
%>