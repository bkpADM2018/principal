<%
'Codigos de Mensaje
Const RESPUESTA_OK 						= "OK"
Const RESPUESTA_ULTIMA_FIRMA			= "UF"
Const CODIGO_VACIO 						= "0001"
Const CODIGO_EXISTE 					= "0002"
Const ABRAVIATURA_VACIA 				= "0003"
Const DESCRIPCION_VACIA 				= "0004"
Const TIPOUNIDAD_NO_EXISTE 				= "0005"
Const LIMITES_STOCK 					= "0006"
Const LIMITES_COMPRA 					= "0007"
Const CATEGORIA_NO_EXISTE 				= "0008"
Const UNIDAD_NO_EXISTE 					= "0009"
Const SOLICITANTE_NO_EXISTE 			= "0010"
Const PROVEEDOR_NO_EXISTE 				= "0011"
Const ARTICULO_NO_EXISTE 				= "0012"
Const EMPRESA_EXISTE 					= "0013"
Const CUIT_ERRONEO 						= "0014"
Const EMAIL_ERRONEO 					= "0015"
Const DIVISION_NO_EXISTE 				= "0016"
Const FALTA_JUSTIFICACION 				= "0017"
Const PERIODO_ERRONEO 					= "0018"
Const PROVEEDOR_REPETIDO 				= "0019"
Const CANTIDAD_NO_EXISTE 				= "0020"
Const VALOR_NO_VALIDO 					= "0021"
Const FECHA_ENTREGA_INCORRECTA 			= "0022"
Const POCOS_ARTICULOS 					= "0023"
Const IMPORTE_NO_EXISTE 				= "0024"
Const IMPORTE_SUPERA_NORMA 				= "0025"
Const FECHA_INICIO_INCORRECTA 			= "0026"
Const FECHA_FIN_INCORRECTA 				= "0027"
Const SALDO_OBRA_INSUFICIENTE 			= "0028"
Const RESPONSABLE_NO_EXISTE 			= "0029"
Const IMPORTE_NO_COINCIDE 				= "0030"
Const STOCK_INSUFICIENTE 				= "0031"
Const CANTIDAD_MENOR_SALDO 				= "0032"
Const PM_REFERENCIA_NO_EXISTE 			= "0033"
Const APRESTAR_MAY_PEDIDOS 				= "0034"
Const ARECIBIR_MAY_PRESTADOS 			= "0035"
Const REMITO_NO_EXISTE 					= "0036"
Const OBRA_NO_EXISTE 					= "0037"
Const ALMACEN_NO_EXISTE 				= "0038"
Const ALMACENDEST_NO_EXISTE 			= "0039"
Const ALMACENDEST_IGUAL_ORIGEN 			= "0040"
Const USUARIO_NO_AUTORIZADO 			= "0041"
Const MAX_USUARIOS_X_SECTOR 			= "0042"
Const USUARIO_YA_FIRMO 					= "0043"
Const COMPRAS_EN_CURSO					= "0044"
Const FALTA_TITULO 						= "0045"
Const NRO_REMITO_REPETIDO 				= "0046"
Const FALTA_RESPONSABLE 				= "0047"
Const FALTA_MIEMBROS_ADJUDICACION		= "0048"
Const ERROR_AUTENTICACION 				= "0049"
Const LLAVE_NO_CORRESPONDE 				= "0049"
Const FALTA_PROV_ADJUDICADO 			= "0050"
Const ALMACEN_NO_GUARDAR 				= "0051"
Const OBRA_NO_SELECCIONADA 				= "0052"
Const BUDGET_NO_EXISTE 					= "0053"
Const ITEMS_SIN_BUDGET 					= "0054"
Const CANTIDAD_NO_NEGATIVA 				= "0055"
Const PM_REQUERIDO 						= "0056"
Const NUEVA_MAY_CUMPLIDOS 				= "0057"
Const NUEVA_MAY_ORIGINAL 				= "0058"
Const ADMINISTRADOR_NO_EXISTE 			= "0059"
Const FALTA_CATAGORIA_AFE 				= "0060"
Const FALTA_AFE_COMPLEMENTARIO 			= "0061"
Const FALTA_TEXTO_OTROS 				= "0062"
Const FALTA_TIPO_GASTO 					= "0063"
Const FALTA_TIPO_CUMPLIMIENTO 			= "0064"
Const DESCRIPCION_DEMASIADO_LARGA		= "0065"
Const FALTA_TITULO_AFE 					= "0066"
Const PM_REFERENCIA_NO_TIPO 			= "0067"
Const VALE_NO_EXISTE 					= "0068"
Const VALE_ESTA_ANULADO 				= "0069"
Const PM_REFERENCIA_NO_EXISTE_NO_TIPO	= "0070"
Const IMPORTE_SUPERA_DISPONIBLE			= "0071"
Const FALTA_DETALLE_DESTINO 			= "0072"
Const FALTA_ASIGNAR_OBRA_SECTOR			= "0073"
Const NUEVA_MAY_NO_DEVUELTO 			= "0074"
Const ARECIBIR_MAY_TRANSFERIDOS			= "0075"
Const ADEVOLVER_MAY_PRESTADOS 			= "0076"
Const TIPO_CAT_INVALIDO 				= "0077"
Const COMENTARIO_REQUERIDO 				= "0078"
Const DETALLE_DUPLICADO 				= "0079"
Const BONIFICACION_PENDIENTE 			= "0080"
Const REQUIERE_FIRMA_AUD_PUE 			= "0081"
Const SALDO_OBRA_MENOR_10_PORC 			= "0082"
Const DESC_BUDGET_OBLIGATORIA 			= "0083"
Const PERSONA_DUPLICADA 				= "0084"
Const PIC_FIRMA_SUPERIOR_NECESARIA		= "0085"
Const DIVISION_PCT_DIFF_OBRA 			= "0086"
Const AVISO_BGT_EXCEDIDO 				= "0087"
Const STOCK_ACTUAL_NO_CUBRE 			= "0088"
Const CANTIDAD_NO_MODIFICADA 			= "0089"
Const CIERRE_NO_EXISTE 					= "0090"
Const PROVEEDOR_NO_RECOMENDADO			= "0091"
Const PRECIO_DIFIERE_ULTIMO_REGISTRO	= "0092"
Const PIC_NECESITA_AFE					= "0093"
Const FALTA_MIEMBROS_AFE				= "0094"
Const DIVISION_PM_VS_DIFF_OBRA 			= "0095"
Const MERGE_UNIDADES_NO_COINCIDEN		= "0096"
Const MERGE_FALTAN_ARTICULOS			= "0097"
Const MERGE_FALTA_DESC_ART_NUEVO		= "0098"
Const MERGE_ART_NUEVO_EXISTE			= "0099"
Const NUEVA_MAY_RECIBIDOS				= "0100"
Const AVISO_NO_EXISTE					= "0101"
Const KILOS_NO_CARGADOS					= "0102"
Const STOCK_EXISTENCIA_INSUFICIENTE		= "0103"
Const AVISO_REQUERIDO					= "0104"
Const PRODUCTO_REQUERIDO				= "0105"
Const FALTAN_DATOS_AVISO				= "0106"
Const KILOS_EXCEDEN						= "0107"
Const CARGA_COMPLETA					= "0108"
Const COMPLIANCE_FTP_FAIL				= "0109"
Const COMPLIANCE_OUT_OF_DATE			= "0113"
const ARTICULO_REQUIERE_PP				= "0114"
const ARTICULO_REQUIERE_SEC				= "0115"
const SIN_NOMBRE						= "0116"
const SIN_DOCUMENTO						= "0117"
const SIN_TIPO_DOCUMENTO				= "0118"
const SIN_DOMICILIO						= "0119"
const SIN_LOCALIDAD						= "0120"
const SIN_TELEFONO						= "0121"
const SIN_TIPO_PROVEEDOR				= "0122"
const SIN_SECTOR						= "0123"
const FAMILIAR_EXISTE					= "0124"
const SIN_TIPO_IGA          			= "0125"
const COD_IVA_INCORRECTO				= "0126"
const COD_IGA_INCORRECTO				= "0127"
const C_MULTI_SIN_CUIT					= "0128"
const PROV_REPETIDA						= "0129"
const COEF_SUPERADO						= "0130"
const PROV_SIN_COEF						= "0131"
const SIN_C_MULTI						= "0132"
const CTZ_AJU_TOTAL_BAJO				= "0133"
const CTZ_AJU_IMP_IGUALES				= "0134"
const CTZ_AJU_APROBADO					= "0135"
const CTZ_AJU_NO_APROBADO				= "0136"
const SIN_TITULO						= "0137"
const SIN_SOLICITANTE					= "0138"
const PP_PIC_DISTIN_PCT 				= "0139"
const ARTICULO_DUPLICADO				= "0140"
const FALTA_ARTICULO					= "0141"
const VERSION_MENOR_PRODUCCION			= "0142"
const VERSION_EXISTE					= "0143"
const SIN_PRODUCTO						= "0144"
const SIN_VERSION						= "0145"
const PRODUCTO_EXISTE 					= "0146"
const CTC_NO_EXISTE						= "0147"
const SIN_PROJECT_LEADER				= "0148"
const CTZ_NCR_MAYOR_SALDO				= "0149"
const PCT_LIMITE_CIERRE					= "0150"
Const SIN_PLATAFORMA					= "0151"
Const SOLICITANTE_DUPLICADO				= "0152"
Const CTZ_AJU_CANT_TOTAL_BAJO			= "0153"
Const NRO_CTA_PTE_INCOMPLETO			= "0154"
Const PAT_CHASIS_INCOMPLETO				= "0155"
Const PAT_ACOPLADO_INCOMPLETO			= "0156"
Const FEC_CONTABLE_INCOMPLETA			= "0157"
Const MANT_AREA_DET_OBLIGATORIO		    = "0158"
Const FEC_DESDE_MAYOR_HASTA 		    = "0159"
Const FEC_DESDE_REQUERIDA			    = "0160"
Const CANTIDAD_ATICULOS_CTST		    = "0161"
Const DIFERENTE_TIPO_GASTO		    	= "0162"
Const PCT_FECHA_EXTENDIDA		    	= "0163"
Const PCT_COTIZACIONES_ABIERTAS	    	= "0164"
Const PCT_FECHA_CIERRE_VENCIDA	    	= "0165"
Const REQUIERE_FIRMA_CEO		    	= "0166"
Const ERROR_GRABAR_CTST                 = "0167"
Const CTC_SALDO_INSUFICIENTE            = "0168"
Const PDC_MONTO_INCORRECTO		        = "0169"
Const PDC_ASEGURADORA_NO_EXISTE	        = "0170"
Const PDC_NRO_ASEGURADORA_NO_EXISTE	    = "0171"
Const TURNO_DESDEHASTA_INCORRECTO	    = "0172"
Const STICKER_DESDEHASTA_INCORRECTO	    = "0173"
Const PDC_PENDIENTE_INCORRECTA		    = "0174"
Const PDC_SALDO_PENDIENTE_INCORRECTO    = "0175"
Const SM_CODIGO_EQUIPO_INCORRECTO       = "0176"
Const SM_DESCRIPCION_EQUIPO_INCORRECTO  = "0177"
Const SM_CODIGO_EQUIPO_EXISTENTE		= "0178"
Const SM_COMPONENTE_YA_EXISTE			= "0179"
Const SM_COMPONENTE_EXISTE_EN_OT		= "0180"
Const SM_COMPONENTE_DESC_REQ			= "0181"
Const SM_EQUIPO_ACTIVO_YA_EXISTE		= "0182"
Const SM_EQUIPO_ACTIVO_EXISTE_EN_OT		= "0183"
Const FILE_MISSING						= "0184"
Const FILE_EXT_NOT_ALLOWED				= "0185"
Const USUARIO_PASS_INCORRECTO			= "0186"
Const USUARIO_BLOQUEADO					= "0187"
Const SM_TEMPLATE_TIENE_EQUIPO_ACTIVO	= "0188"
Const PRESTAMOS_EN_CURSO				= "0189"
Const SM_OT_FALTA_TITULO				= "0190"
Const SM_OT_FALTA_EQUIPO				= "0191"
Const SM_OT_FALTA_TIPO_MANT				= "0192"
Const SM_OT_FALTA_TIPO_ORDEN			= "0193"
Const SM_OT_FALTA_SOLICITANTE			= "0194"
Const SM_OT_FALTA_RESPONSABLE			= "0195"
Const SM_OT_FALTA_FECHA_PROG			= "0196"
Const SM_OT_FALTA_FECHA_PROG_VIEJA		= "0197"
Const ARTICULO_NO_ELEGIBLE				= "0198"
Const SM_OT_FALTA_OBRA					= "0199"
Const SM_OT_FALTA_BUDGET				= "0200"
Const FALTA_COORD_PUERTOS        		= "0201"
Const CTC_PARTIDA_YA_EXISTE				= "0202"
Const PCP_AREA_DET_OBLIGATORIO		    = "0203"
Const PCP_NECESITA_AFE					= "0204"
Const PROV_FALTA_NVA					= "0205"
Const PROV_CUIT_INHABILITADO			= "0206"
Const PROV_CUIT_REGISTRADO				= "0207"
Const SIN_PROVINCIA						= "0208"
Const PROV_EMPELA_INCORRECTO			= "0209"
Const PROV_SOCHEC_INCORRECTO			= "0210"
Const REM_ANULA_ARTICULO_PAGO			= "0211"
Const FALTA_CODIGO_ACEPTACION			= "0212"
Const FIRMANTE_DIR_NO_AUTORIZADO		= "0213"
Const BU_MEZCLA_ARTICULOS	        	= "0214"
Const BU_PARTIDA_INCORRECTA 		    = "0215"
Const PIC_FIRMA_DIRECCION_DUPLICADO     = "0216"
Const COSECHA_SIN_HABILITAR             = "0217"
Const ESTADO_AFE_NO_DETERMINADO         = "0218"
Const PROV_NO_COINCIDE_PCT              = "0219"
Const SM_ART_REPETIDOS				    = "0220"
Const SM_OBS_REQUERIDAS				    = "0221"
Const SM_OT_FALTA_FREQUENCIA		    = "0222"
Const DSPROCEDENCIA_VACIO			    = "0224"
Const CDPROVINCIA_VACIO					= "0225"
Const CDPARTIDO_VACIO					= "0226"
Const CDPROCEDENCIACAMARA_VACIO		    = "0227"
Const FECHA_CEC_INCORRECTA			    = "0228"
Const DIVISION_CTC_DIFF_OBRA 			= "0232"
Const AUTORIZANTE_NO_EXISTE 			= "0233"
Const PROVEEDOR_NO_HAB_CTC              = "0234"
Const ERROR_AUTORIZACION_PROVISIONES	= "0235"
Const BUDGET_REASIGNACION_EN_PROCESO    = "0236"
Const CTC_PARTIDA_NO_UNICA              = "0237"
Const CTZ_AJS_POS_NO_PERMITIDO          = "0238"
Const FALTA_PROCEDENCIA_ONCCA           = "0239"
'mensajes de info
Const MAIL_ENVIO_EXITOSO				= "1000"

Dim errMessages
Dim wrnMessages
Dim infMessages
Set errMessages = Server.CreateObject("Scripting.Dictionary")
Set wrnMessages = Server.CreateObject("Scripting.Dictionary")
Set infMessages = Server.CreateObject("Scripting.Dictionary")
'--------------------------------------------------------------------------------------
Function setError(e)
	Dim errValue
	errValue = "ERR|" & e
	if not errMessages.Exists(e) then errMessages.add e, errValue
End Function
'--------------------------------------------------------------------------------------
Function setWarning(e)
	Dim errValue
	errValue = "WRN|" & e
	if not wrnMessages.Exists(e) then wrnMessages.add e, errValue
end Function
'--------------------------------------------------------------------------------------
Function setInfo(e)
	Dim errValue
	errValue = "INF|" & e
	if not infMessages.Exists(e) then infMessages.add e, errValue
End Function
'--------------------------------------------------------------------------------------
Function hayError()
hayError = false
if errMessages.Count > 0 then hayError = true
end function
'--------------------------------------------------------------------------------------
Function clearError()
	errMessages.RemoveAll 
end function
'--------------------------------------------------------------------------------------
Function showErrors()		
	Dim index, Ky, It, aux
	if (isObject(errMessages)) then
		if ((errMessages.Count > 0) or (wrnMessages.Count > 0) or (infMessages.Count > 0)) then
			Set dicErr = Server.CreateObject("Scripting.Dictionary")
			Set dicWrn = Server.CreateObject("Scripting.Dictionary")
			Set dicInf = Server.CreateObject("Scripting.Dictionary")
			'Se separan los warnings de los errores.
			Ky = errMessages.Keys
			It = errMessages.Items
			for index = 0 To errMessages.Count -1			
				aux = Split(It(index), "|")
				dicErr.Add Ky(index), aux(1)
			next	
			
			Ky = wrnMessages.Keys
			It = wrnMessages.Items
			for index = 0 To wrnMessages.Count -1			
				aux = Split(It(index), "|")
				dicWrn.Add Ky(index), aux(1)
			next
			
			Ky = infMessages.Keys
			It = infMessages.Items
			for index = 0 To infMessages.Count -1			
				aux = Split(It(index), "|")
				dicInf.Add Ky(index), aux(1)
			next	
			
			'Se muestran los errores
			Call showErrorMessages(dicErr, "reg_Header_Error")
			'Se muestran los warnings
			Call showErrorMessages(dicWrn, "reg_Header_Warning")
			'Se muestran los infos
			Call showErrorMessages(dicInf, "reg_Header_success")
		end if
	end if
End Function
'--------------------------------------------------------------------------------------
function showMessages()
dim myErrors, myWarnings, myInfo
myErrors = getErrors()
if len(myErrors) > 0 then
	%>
	<div class="errormsj">
		<p> <%=myErrors%> </p>
	</div> 
	<%
end if	
myWarnings = getWarnings() 
if len(myWarnings) > 0 then
	%>
	<div class="alertmsj">
		<p> <%=myWarnings%> </p>
	</div> 
	<%
end if	
myInfo = getInfo()
if len(myInfo) > 0 then
	%>
	<div class="alertmsj">
		<p> <%=myInfo%> </p>
	</div> 
	<%
end if	

end function
'--------------------------------------------------------------------------------------
Function getErrors()		
	Dim index, Ky, It, aux, rtrn
	if (isObject(errMessages)) then
		if (errMessages.Count > 0) then
			for each element in errMessages
				'if not isEmpty(element) then
					if len(rtrn) > 0 then rtrn = rtrn & "<br>"
					rtrn = rtrn & element & " - " & UCASE(errMessage(element))
				'end if
			next
		end if
	end if
getErrors = rtrn
End Function
'--------------------------------------------------------------------------------------
Function getWarnings()		
	Dim index, Ky, It, aux, rtrn
	if (isObject(wrnMessages)) then
		if (wrnMessages.Count > 0) then
			for each element in wrnMessages
				if len(rtrn) > 0 then rtrn = rtrn & "<br>"
				rtrn = rtrn & element & " - " & UCASE(errMessage(element))
			next
		end if
	end if
getWarnings = rtrn
End Function
'--------------------------------------------------------------------------------------
Function getInfo()		
	Dim index, Ky, It, aux, rtrn
	if (isObject(infMessages)) then
		if (infMessages.Count > 0) then
			for each element in infMessages
				if len(rtrn) > 0 then rtrn = rtrn & "<br>"
				rtrn = rtrn & element & " - " & UCASE(errMessage(element))
			next
		end if
	end if
getInfo = rtrn
End Function

'--------------------------------------------------------------------------------------
Function showErrorMessages(pDic, pStyle)
	
	if (pDic.Count > 0) then
%>
	<table align="center" width="100%">
<%		Ky = pDic.Keys
		It = pDic.Items
		for index = 0 To pDic.Count -1			%>
		<tr>
			<td class="<% =pStyle %>"><% =Ky(index) %> - <% =GF_TRADUCIR(errMessage(It(index))) %></td>
		</tr>
<%		next	%>
	</table>	
<%	end if

End Function
'--------------------------------------------------------------------------------------
Function errMessage(errCode)
	Dim strSQL, rs
	errMessage = "Mensaje no definido!!"
	strSQL = "Select * from TBLMENSAJES where CDMENSAJE='" & errCode & "'"
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	if (not rs.eof) then errMessage = Trim(rs("DSMENSAJE"))
End Function


%>