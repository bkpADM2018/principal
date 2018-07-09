<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<%
Function guardarCambiosCabecera(pDsLista, pIdLista,pcdLista, pIdDivision)
	Dim strSQL, rs, rsLista, myIdLista,oConn,rtrn
	if(pIdLista = 0)then
		'Ademas de guardar la cabecera devuelve el Id de la lista guardada
		myIdLista = 0
		strSQL=" Insert into TBLMAILLSTCABECERA(CDLISTA,DSLISTA,IDDIVISION,CDUSUARIO,MOMENTO) "
		strSQL= strSQL & " values('" & Trim(UCase(pcdLista)) & "','" & Trim(Ucase(pDsLista)) & "'," & pIdDivision & ",'" & session("Usuario") & "' ,'" & session("MmtoSistema") &"' )"
		'Response.Write " CABECERA: " & strSQL & "<br></br>"
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
		strSQL = ""
		strSQL=" Select MAX(IDLISTA) AS IDLISTA FROM TBLMAILLSTCABECERA "
        Call executeQueryDb(DBSITE_SQL_INTRA, rsLista, "OPEN", strSQL)	
		myIdLista = rsLista("IDLISTA")
		rtrn = myIdLista		
		'rtrn = 20
	else
		'en caso de que sea mayor a 0 se toma como una actualizacion
		strSQL = "Update TBLMAILLSTCABECERA set DSLISTA = '" & Trim(Ucase(pDsLista)) &"', CDLISTA = '" & Trim(UCase(pcdLista)) & "' WHERE IDLISTA = " &pIdLista
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)		
		rtrn = pIdLista
	end if
	'Response.Write " CABECERA: " & strSQL & "<br></br>"
	guardarCambiosCabecera = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function guardarCambiosDetalle(pIdLista, pUser, pDsMail, pIdEstado)
	Dim strSQL, rs
	if(pIdEstado = ESTADO_ACTIVO)then
		strSQL = "Insert into TBLMAILLSTSDETALLE(IDLISTA, CDUSER, EMAIL, CDUSUARIO, MOMENTO)"
		strSQL= strSQL & " values('" & pIdLista & "', '" & pUser & "', '" & pDsMail & "' ,'" & session("Usuario") &"','" & session("MmtoSistema") &"')"		
	end if
	if(pIdEstado = ESTADO_BAJA)then
		strSQL = "Delete from TBLMAILLSTSDETALLE where idLista = " & pIdLista &" and CDUSER = '" & pUser &"'"
	end if
	'Response.Write " DETALLE: " & strSQL & "<br></br>"
    Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)	
End Function
'--------------------------------------------------------------------------------------------------
'SE ENCARGA DE CONTROLAR QUE LOS REGISTROS (IDDIVISION Y CDLISTA) NO ESTE REPETIDO, DE ESTA MANERA PUEDO 
'ASEGURARME QUE UN CODIGO DE LISTA SOLO ESTE DISPONIBLE UNA SOLA VEZ PARA CADA DIVISION
Function validarCodigoListaExistente(pCdLista,pIdLista,pIdivision)
	Dim strSQL,rtrn,rsLista,cdLista_Old
	flagModifico = true
	rtrn = false
	if(pIdLista > 0)then		
		strSQL=" Select CDLISTA FROM TBLMAILLSTCABECERA WHERE IDLISTA = " & pIdLista	
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)		
		if(rs("CDLISTA") = Ucase(pCdLista))then flagModifico = false		
	end if
	if(flagModifico)then
		strSQL=" Select COUNT(CDLISTA) AS CANTIDAD FROM TBLMAILLSTCABECERA WHERE CDLISTA = '" & Trim(UCase(pCdLista)) & "' AND IDDIVISION = "&pIdivision		
        Call executeQueryDb(DBSITE_SQL_INTRA, rsLista, "OPEN", strSQL)	
		if(not rsLista.eof)then
			if(rsLista("CANTIDAD") > 0)then rtrn = true
		end if
	end if
	validarCodigoListaExistente = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function controlarListaCorreo(pDsLista, pcantUsuario,pCdLista ,pIdDivision,pIdLista, ByRef phaveUser)
	Dim MsgErr, cdUser, flagOk,cont,IdDivision, IdEstado
	cont = 0
	MsgErr = RESPUESTA_OK
	flagVacio = true
	flagOk = true
	if((Len(pCdLista) = 0)or(pCdLista = ""))then
		MsgErr = CODIGO_VACIO
	else
		if(validarCodigoListaExistente(pCdLista,pIdLista,pIdDivision))then
			MsgErr = CODIGO_EXISTE			
		else
			if(Len(pDsLista) = 0)then
				MsgErr = DESCRIPCION_VACIA		
			else
				if (Len(PDsLista) > 50) then 
					MsgErr = DESCRIPCION_DEMASIADO_LARGA
				else
					if(pIdDivision <= 0)then
						MsgErr = DIVISION_NO_EXISTE
					else
						for i = 0 to pcantUsuario - 1
							cdUser = GF_PARAMETROS7("cdUsuario" & i, "", 6)				
							IdEstado = GF_PARAMETROS7("IdEstado"&i, 0, 6)						
							if(Len(cdUser) > 0)then
								phaveUser = true
								for z = 0 to cont - 1
									if((cdUser = vCdUser(z))and(vEstado(z) <> ESTADO_BAJA))then
										flagOk = false
										MsgErr = RESPONSABLE_NO_EXISTE
									end	if
								next							
								if(flagOk)then
									Redim Preserve vCdUser(cont)
									Redim Preserve vEstado(cont)
									vCdUser(cont) = cdUser
									vEstado(cont) = IdEstado
									cont = cont + 1
								end if
							end if
						next
					end if
				end if
			end if
		end if
	end if	
	controlarListaCorreo = MsgErr	
End function
'--------------------------------------------------------------------------------------------------
Dim  idLista, DsLista,IdDivision, CdUser,cantUsuario,vCdUser(),vEstado(),dsMail,IdEstado
Dim DsUsuario, rtr,cdSolicitante,CdLista,haveUser
IdLista = GF_PARAMETROS7("IdLista", 0, 6)
DsLista = GF_PARAMETROS7("DsLista", "", 6)
cantUsuario = GF_PARAMETROS7("cantUsuario", 0, 6) 
IdDivision = GF_PARAMETROS7("idDivision", 0, 6)
CdLista = GF_PARAMETROS7("CdLista","",6)
haveUser = false
rtr = controlarListaCorreo(DsLista, cantUsuario,CdLista,IdDivision,IdLista,haveUser)
if(rtr = RESPUESTA_OK)Then
	idLista = guardarCambiosCabecera(DsLista, IdLista, CdLista, IdDivision)	
	if(haveUser)then
		for i = 0 to  UBound(vCdUser)			
			dsMail = getUserMail(vCdUser(i))
			Call guardarCambiosDetalle(idLista, vCdUser(i), dsMail, vEstado(i))			 
		next			
	end if	
end if

%>
<HTML>
<HEAD>
<script type="text/javascript">		
	parent.resultadoCarga_callback('<%=rtr%>');	
</script>
</HEAD>
<BODY>
<P>&nbsp;</P>
</BODY>
</HTML>