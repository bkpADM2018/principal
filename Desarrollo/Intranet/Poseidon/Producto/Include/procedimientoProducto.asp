<%
Const VALUE_ATRIBUTE_NEGATIVO = 0
Const VALUE_ATRIBUTE_AFIRMATIVO = 1
Const VALUE_ATRIBUTE_OPCIONAL = 2
Const BIOTEC_ACTIVA = "V"
Const BIOTEC_INACTIVA = "F"

Function getProductoByCdProducto(pCdProducto,pPto)
	Dim strSQL,rs
	strSQL = "Select CDPRODUCTO, " &_
             "       DSPRODUCTO, " &_
             "       CASE WHEN PRK1 IS NULL THEN 0 ELSE PRK1 END AS PRK1," &_
             "       CASE WHEN PRK2 IS NULL THEN 0 ELSE PRK2 END AS PRK2," &_
             "       CASE WHEN PRK3 IS NULL THEN 0 ELSE PRK3 END AS PRK3," &_
             "       CASE WHEN PRK4 IS NULL THEN 0 ELSE PRK4 END AS PRK4," &_
             "       CASE WHEN NUULTBOLETCAMARA IS NULL THEN 0 ELSE NUULTBOLETCAMARA END AS NUULTBOLETCAMARA," &_
             "       CASE WHEN ICSTATUS IS NULL THEN '' ELSE ICSTATUS END AS ICSTATUS," &_
             "       CASE WHEN VLBASETRIGO IS NULL THEN 0 ELSE VLBASETRIGO END AS VLBASETRIGO," &_
             "       CASE WHEN ICTIPOENVIO IS NULL THEN 0 ELSE ICTIPOENVIO END AS ICTIPOENVIO," &_
             "       CASE WHEN CDPRODUCTOCAMARA IS NULL THEN 0 ELSE CDPRODUCTOCAMARA END AS CDPRODUCTOCAMARA," &_
             "       CASE WHEN CDNUMERADORTURNO IS NULL THEN '' ELSE CDNUMERADORTURNO END AS CDNUMERADORTURNO," &_
             "       CASE WHEN ICESTANDARBASE IS NULL THEN 0 ELSE ICESTANDARBASE END AS ICESTANDARBASE," &_
             "       CASE WHEN ICHUMEDIMETRO IS NULL THEN 0 ELSE ICHUMEDIMETRO END AS ICHUMEDIMETRO," &_
             "       CASE WHEN DSPRODUCTOABR IS NULL THEN '' ELSE DSPRODUCTOABR END AS DSPRODUCTOABR " &_
             "from dbo.productos where cdproducto = " & pCdProducto
    call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	Set getProductoByCdProducto = rs
End Function	
'-------------------------------------------------------------------------------------------------------------------
Function getAtributteProducto(pCdProducto,pCdAceptacion,pPto)
	Dim strSQL 
	strSQL = "SELECT * FROM dbo.ATRIBUTOSDEPRODUCTO WHERE CDPRODUCTO =" &pCdProducto& " AND CDACEPTACION = " & pCdAceptacion	 
	call GF_BD_Puertos (pPto, rs, "OPEN",strSQL)
	Set getAtributteProducto = rs
End Function
'-------------------------------------------------------------------------------------------------------------------
Function addAtribute(pCdProducto,pCdAceptacion,pSticker,pCamara,pRechazo,pGrado,pMerma,pRubro,pBalde,pInterno,pSupervisor,pAcondicionamiento,pPto)
	Dim strSQL
	strSQL = "INSERT INTO ATRIBUTOSDEPRODUCTO (CDPRODUCTO,CDACEPTACION,ICSTICKER,ICSUPERVISOR,ICCAMARA,ICMOTIVORECHAZO,ICGRADO,ICMERMA,ICRUBRO,ICBALDE,ICACON,ICINFORMEINTERNO) "&_											  
			 "VALUES("&pCdProducto&","&pCdAceptacion&","&pSticker&","&pSupervisor&","&pCamara&","&pRechazo&","&pGrado&","&pMerma&","&pRubro&","&pBalde&","&pAcondicionamiento&","&pInterno&")"		
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function updateAtribute(pCdProducto,pCdAceptacion,pSticker,pCamara,pRechazo,pGrado,pMerma,pRubro,pBalde,pInterno,pSupervisor,pAcondicionamiento,pPto)
	Dim strSQL
	strSQL = "UPDATE ATRIBUTOSDEPRODUCTO SET ICSUPERVISOR="&pSupervisor&",ICACON="&pAcondicionamiento&",ICSTICKER="&pSticker&",ICCAMARA="&pCamara&",ICMOTIVORECHAZO="&pRechazo&",ICGRADO="&pGrado&",ICMERMA="&pMerma&",ICRUBRO="&pRubro&",ICBALDE="&pBalde&",ICINFORMEINTERNO="&pInterno&_
			 " WHERE CDPRODUCTO = " & pCdProducto & " AND CDACEPTACION = " &pCdAceptacion	
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function deleteAtributo(pCdProducto,pCdAceptacion,pPto)
	Dim strSQL, auxWhere
	if Cdbl(pCdAceptacion) > 0 then auxWhere = " AND CDACEPTACION = " & pCdAceptacion
	strSQL  = "DELETE FROM ATRIBUTOSDEPRODUCTO WHERE CDPRODUCTO = " & cdProducto & auxWhere
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function deleteProducto(pCdProducto,pPto)
	Dim strSQL 
	strSQL  = "DELETE FROM Productos WHERE CDPRODUCTO = " & cdProducto
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function addProducto(pcdProducto,pDescripcion,pHumedadRecep,pHumedadBase,pCoeficiente1,pCoeficiente2,pUltimaBoleta,pBaseTrigo,pTipoEnvio,pCodigoCamara,pUltimoTurno,pTipoProducto,pHumedimetro,pDescripcionAbr,pPto)
	Dim strSQL
	strSQL = "INSERT INTO PRODUCTOS(CDPRODUCTO,DSPRODUCTO,PRK1,PRK2,PRK3,PRK4,NUULTBOLETCAMARA,ICSTATUS,VLBASETRIGO,ICTIPOENVIO,CDPRODUCTOCAMARA,CDNUMERADORTURNO,ICESTANDARBASE,ICHUMEDIMETRO,DSPRODUCTOABR) "&_
			 "VALUES ("&pcdProducto&",'"&pDescripcion&"',"&pHumedadRecep&","&pHumedadBase&","&pCoeficiente1&","&pCoeficiente2&","&pUltimaBoleta&",'0',"&pBaseTrigo&","&pTipoEnvio&","&pCodigoCamara&",'"&pUltimoTurno&"',"&pTipoProducto&","&pHumedimetro&",'"&pDescripcionAbr&"')"	
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function updateProducto(pcdProducto,pDescripcion,pHumedadRecep,pHumedadBase,pCoeficiente1,pCoeficiente2,pUltimaBoleta,pBaseTrigo,pTipoEnvio,pCodigoCamara,pUltimoTurno,pTipoProducto,pHumedimetro,pDescripcionAbr,pPto)
	Dim strSQL
	strSQL = "UPDATE PRODUCTOS SET DSPRODUCTO='"&pDescripcion&"',PRK1="&pHumedadRecep&",PRK2="&pHumedadBase&",PRK3="&pCoeficiente1&",PRK4="&pCoeficiente2&_
			 ",NUULTBOLETCAMARA="&pUltimaBoleta&",VLBASETRIGO="&pBaseTrigo&",ICTIPOENVIO="&pTipoEnvio&",CDPRODUCTOCAMARA="&pCodigoCamara&",CDNUMERADORTURNO='"&pUltimoTurno&"',ICESTANDARBASE="&pTipoProducto&",ICHUMEDIMETRO="&pHumedimetro&",DSPRODUCTOABR='"&pDescripcionAbr&"'"&_
			 " WHERE CDPRODUCTO = "& pcdProducto
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function deleteCosecha(pcdProducto,pCdCosecha,pPto)
	Dim strSQL, auxWhere
	if Cdbl(pCdCosecha) > 0 then auxWhere = " AND CDCOSECHA = " & pCdCosecha
	strSQL  = "DELETE FROM COSECHAS WHERE CDPRODUCTO = " & pcdProducto & auxWhere
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function updateCosecha(pCdProducto,pCdCosecha,pChkHabilitado,pPto)
	Dim strSQL
	strSQL = "UPDATE COSECHAS SET COSDEF = '"& pChkHabilitado &"'"&_
			 " WHERE CDPRODUCTO = " & pCdProducto & " AND CDCOSECHA = " &pCdCosecha
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function addCosecha(pCdProducto,pCdCosecha,pChkHabilitado,pPto)
	Dim strSQL
	strSQL = "INSERT INTO COSECHAS (CDPRODUCTO,CDCOSECHA,COSDEF) "&_
			 "VALUES("&pCdProducto&","&pCdCosecha&",'"&pChkHabilitado&"')"
	Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL)
End Function
'-------------------------------------------------------------------------------------------------------------------
Function EliminarBiotecnologiaDeProducto(pIdBiotecnologia, pProducto, pPto)
    Dim strSQL, mywhere
    if (Cdbl(pIdBiotecnologia) > 0) then mywhere  = " AND IDBIOTECNOLOGIA = "& pIdBiotecnologia
    strSQL = "UPDATE TBLBIOTECNOLOGIAS SET HABILITADO = '"&p_estado&"' WHERE IDBIOTECNOLOGIA = " &p_idBiotecnologia
    Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL) 
End Function
'-------------------------------------------------------------------------------------------------------------------
Function GuardarBiotecnologia( p_idBiotecnologia, p_dsBiotecnologia, p_cdCliente, p_cdProducto, p_estado, p_nuSobre, pPto)    
    Dim strSQL 
    if (Cdbl(p_idBiotecnologia) <> 0) then
        strSQL = "UPDATE TBLBIOTECNOLOGIAS SET DSBIOTECNOLOGIA = '"&p_dsBiotecnologia&"',IDPROVEEDOR = "&p_cdCliente&",HABILITADO = '"&p_estado&"' ,NUSOBRES = "&p_nuSobre&_
                 " WHERE IDBIOTECNOLOGIA = " &p_idBiotecnologia
    else
        strSQL = "INSERT INTO TBLBIOTECNOLOGIAS (DSBIOTECNOLOGIA,IDPROVEEDOR,IDPRODUCTO,HABILITADO,NUSOBRES) "&_
                 "VALUES ('"&p_dsBiotecnologia&"', "&p_cdCliente&", "&p_cdProducto&", '"&p_estado&"', "&p_nuSobre&")"
    end if
    Call GF_BD_Puertos (pPto, rs, "EXEC",strSQL) 
End Function 
'-------------------------------------------------------------------------------------------------------------------
'Traduce el valor del Atributo a una descripcion
Function getDsResultAtribute(pValue)
	Dim rtrn 
	rtrn = ""
	if Not isNull(pValue) then 
		Select Case Cdbl(pValue)
			Case VALUE_ATRIBUTE_AFIRMATIVO
				rtrn = "Sí"
            Case VALUE_ATRIBUTE_NEGATIVO
                rtrn = "No"
            Case VALUE_ATRIBUTE_OPCIONAL
                rtrn = "Opcional"
		End Select
	end if
	getDsResultAtribute = rtrn
End Function 
%>