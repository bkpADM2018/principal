<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosUser.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosMG.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<%
'--------------------------------------------------------------------------------------
Dim ppuerto,pppcosecha,pppkilos,ppproducto,ppcliente,dtContable_Old,idCamion_Old,pcontadorDiv

ppuerto = GF_Parametros7("pto", "", 6)
pppcosecha = GF_Parametros7("cosecha", "", 6)
ppproducto = GF_Parametros7("producto", "", 6)
pppkilos = GF_Parametros7("kilos", 0, 6)
ppcliente = GF_Parametros7("cliente", 0, 6)
dtContable_Old = GF_Parametros7("dtContable", "", 6)
idCamion_Old = GF_Parametros7("idCamion", "", 6)
pcontadorDiv = GF_Parametros7("Contador", "", 6)

call armarTablaCosechaXCamiones(ppcliente,ppproducto,pppcosecha,ppuerto,dtContable_Old,idCamion_Old,pcontadorDiv)
response.end
'-----------------------------------------------------------------------------------------
Function validarEscritura(ByRef pcontador,pidCamion_Old,pdtContable_Old)
Dim rtrn
rtrn = false
if((pcontador > 0)and(pcontador <= 100))then
	if((len(pdtContable_Old) > 0)and(len(pidCamion_Old) > 0))then'es el 2� registro 
		if(pcontador > 1)then 
			rtrn = true
		else
			pcontador = pcontador + 1			
		end if
	else'es el 1� registro
		rtrn = true		
	end if
end if	
validarEscritura = rtrn
End function
'-----------------------------------------------------------------------------------------
Function armarTablaCosechaXCamiones(ppcliente,ppproducto,pppcosecha,ppuerto,dtContable_Old,idCamion_Old,pcontadorDiv)
Dim rsCosCam,auxCliente,auxProducto,auxIdCam,auxDtCont,aaa,myFormatFechaVieja,myFormatFechaNueva,contador
Dim v_estadoFinalizado

Set rsCosCam = armarSQLCosechaXCamiones(ppcliente,ppproducto,pppcosecha,ppuerto,dtContable_Old,idCamion_Old)

Response.Write("<table cellspacing='0' align='center'  width='100%'>") 
Response.Write("<tr class='reg_Header_nav'>")
Response.Write("<td width='5%' align='center'>Fecha</td>")
Response.Write("<td width='6%' align='center'>IdCamion</td>")
Response.Write("<td width='4%' align='center'>Chapa</td>")
Response.Write("<td width='4%' align='center'>Acoplado</td>")
Response.Write("<td width='5%' align='center'>Carta Porte</td>")
Response.Write("<td width='4%' align='center'>CTG</td>")
Response.Write("<td width='5%' align='center'>Cliente</td>")
Response.Write("<td width='7%' align='center'>Producto</td>")
Response.Write("<td width='3%' align='center'>Kilos</td>")
Response.Write("</tr>")

if(rsCosCam.eof)then	
	Response.Write("<tr  class='reg_Header_navdos' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>")
	Response.Write("<td align='center' colspan='9'>No se encontraron mas camiones disponibles </td>")
	Response.Write("</tr>")		
	Response.end	
end if

auxIdCam = ""
auxDtCont = ""
v_estadoFinalizado = FINALIZAR_LISTADO
contador = 0
while (not rsCosCam.eof) and (contador < 100)
	
	Response.Write("<tr class='reg_Header_navdos' onMouseOver='javascript:lightOn(this)' onMouseOut='javascript:lightOff(this)'>")	
	Response.Write("<td align='center'>" & GF_FN2DTE(rsCosCam("DTCONTABLE")) & "</td>")
	Response.Write("<td align='center'>" & rsCosCam("IDCAMION") & "</td>") 
	Response.Write("<td align='left'>" & GF_EDIT_PATENTE(rsCosCam("CDCHAPACAMION")) & "</td>")
	Response.Write("<td align='left'>" & GF_EDIT_PATENTE(rsCosCam("CDCHAPAACOPLADO")) & "</td>")
	Response.Write("<td align='center'>" & GF_EDIT_CTAPTE(rsCosCam("NUCARTAPORTE")) & "</td>")
	Response.Write("<td align='center'>" & rsCosCam("CTG") & "</td>")				
	auxCliente = Trim(rsCosCam("cdcliente"))&"-"&Trim(rsCosCam("dscliente"))		
	if(len(auxCliente) > 32)then auxCliente = left(auxCliente,29) & "..."
	Response.Write("<td align='left'>" & Trim(auxCliente) & "</td>")		
		
	auxProducto = Trim(ppproducto)&"-"&Trim(rsCosCam("dsproducto"))				
	if(len(auxProducto) > 29)then auxProducto = left(auxProducto,26) & "..." 	
	Response.Write("<td  align='left'>" & Trim(auxProducto) & "</td>")
		
	Response.Write("<td  align='right'>" & GF_EDIT_DECIMALS(cdbl(rsCosCam("KILOSNETOS"))*100,2) & "</td>")	
	Response.Write("</tr>")						
	contador = contador + 1
	
	rsCosCam.movenext	
wend
'Si tengo un registro valido como proximo, seteo las variables de control.
if (not rsCosCam.eof) then
	v_estadoFinalizado = CONTINUAR_LISTADO
	auxDtCont = Left(rsCosCam("DTCONTABLE"),4) &"-"& Mid(rsCosCam("DTCONTABLE"),5,2) &"-"& Right(rsCosCam("DTCONTABLE"),2)
	auxIdCam = rsCosCam("IDCAMION")	
end if
Response.Write("</table>")
response.write("<input type='hidden' id='FinListado_"&pcontadorDiv&"' value="&v_estadoFinalizado&">")	
response.write("<input type='hidden' id='dtcot_old_"&pcontadorDiv&"' value="&auxDtCont&">")
response.write("<input type='hidden' id='IdCam_old_"&pcontadorDiv&"' value="&auxIdCam&">")

End function
'-----------------------------------------------------------------------------
Function armarSQLCosechaXCamiones(pcliente,pproducto,ppcosecha,pto,dtContable_Old,idCamion_Old)								 
Dim rs, strSQL  , mywherePro, mywhereCam
mywherePro = ""
mywhereCam = ""

'Si tengo variables de control de paginacion, las proceso
if (len(dtContable_Old) > 0) then
	mywherePaginacion = " and (HCD.DTCONTABLE > '" & dtContable_Old & "' or (HCD.DTCONTABLE = '" & dtContable_Old & "' and HCD.IDCAMION>= '" & idCamion_Old & "'))"
end if

if(pproducto > 0)then
	mywherePro = " AND P.CDPRODUCTO = " & pproducto
end if
if(pcliente > 0)then
	mywhereCam = " AND C.CDCLIENTE = " & pcliente
end if

fechaInicio = "2010-03-01"
strSQL = " select * from ( " 
strSQL = strSQL &	"SELECT (YEAR(TG.DTCONTABLE)*10000 + Month(TG.DTCONTABLE)*100 + DAY(TG.DTCONTABLE)) DTCONTABLE, TG.IDCAMION, TG.CARTAPORTE NUCARTAPORTE, TG.CDCHAPACAMION, TG.CDCHAPAACOPLADO, TG.CDCOSECHA, " 
strSQL = strSQL &	"	TG.CARTAPORTE, TG.CTG, TG.CDCLIENTE, TG.DSCLIENTE, TG.DSPRODUCTO, CASE WHEN EMBARCADOS.KILOSCARGADOS IS NULL THEN TG.KILOSNETOS ELSE TG.KILOSNETOS-EMBARCADOS.KILOSCARGADOS END AS KILOSNETOS " 
strSQL = strSQL &	"	  FROM " 
strSQL = strSQL &	"	 "
strSQL = strSQL &	"	    (SELECT HCD.DTCONTABLE, HCD.IDCAMION, HC.CDCHAPACAMION, HC.CDCHAPAACOPLADO, HCD.CDCOSECHA, " 
strSQL = strSQL &	"	        RTRIM(HCD.NUCARTAPORTE) + RTRIM(HCD.NUCTAPTEDIG) AS CARTAPORTE, HCD.CTG, C.CDCLIENTE, C.DSCLIENTE, P.DSPRODUCTO, " 
'Para SQL SERVER - strSQL = strSQL &	"	        RTRIM(RTRIM(HCD.NUCARTAPORTE)+''+RTRIM(HCD.NUCTAPTEDIG)) AS CARTAPORTE, HCD.CTG, C.DSCLIENTE, P.DSPRODUCTO, " 
strSQL = strSQL &	"	        ( "
strSQL = strSQL &	"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 1 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 1)) " 
strSQL = strSQL &	"	            -  "
strSQL = strSQL &	"	            ( SELECT PC.VLPESADA FROM dbo.HPESADASCAMION PC WHERE PC.DTCONTABLE = HCD.DTCONTABLE AND PC.IDCAMION = HCD.IDCAMION AND PC.CDPESADA = 2 AND PC.SQPESADA = (SELECT MAX(SQPESADA) FROM dbo.HPESADASCAMION WHERE PC.DTCONTABLE = DTCONTABLE AND PC.IDCAMION = IDCAMION AND CDPESADA = 2))  " 
strSQL = strSQL &	"	            -  "
strSQL = strSQL &	"	            ( SELECT CASE WHEN HMC.VLMERMAKILOS IS NULL THEN 0 ELSE HMC.VLMERMAKILOS END FROM HMERMASCAMIONES HMC WHERE HMC.DTCONTABLE=HCD.DTCONTABLE AND HMC.IDCAMION = HCD.IDCAMION AND HMC.SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES WHERE DTCONTABLE=HCD.DTCONTABLE AND IDCAMION = HCD.IDCAMION)) "
strSQL = strSQL &	"	        ) KILOSNETOS"
strSQL = strSQL &	"	    FROM HCAMIONESDESCARGA HCD"
strSQL = strSQL &	"	    LEFT JOIN HCAMIONES HC ON HC.IDCAMION = HCD.IDCAMION AND HC.DTCONTABLE=HCD.DTCONTABLE  "
strSQL = strSQL &	"	    LEFT JOIN PRODUCTOS P ON P.CDPRODUCTO = HC.CDPRODUCTO  "
strSQL = strSQL &	"	    LEFT JOIN CLIENTES C ON C.CDCLIENTE = HCD.CDCLIENTE "
strSQL = strSQL &	"		WHERE HCD.DTCONTABLE >='" & fechaInicio & "'"
strSQL = strSQL &	mywherePro
strSQL = strSQL &	mywhereCam
strSQL = strSQL &	mywherePaginacion
strSQL = strSQL &	"		AND HC.CDESTADO IN (6,8) "
strSQL = strSQL &   "		AND  HCD.cdcosecha =" & ppcosecha
strSQL = strSQL &	"		) TG "
strSQL = strSQL &	"	        LEFT JOIN  "
strSQL = strSQL &	"	            (SELECT IDCAMION, DTCONTABLE, SUM(KILOSNETOS) AS KILOSCARGADOS FROM CTGEMBARCADOS GROUP BY IDCAMION, DTCONTABLE) "
strSQL = strSQL &	"	                EMBARCADOS ON TG.IDCAMION = EMBARCADOS.IDCAMION AND TG.DTCONTABLE=EMBARCADOS.DTCONTABLE "
strSQL = strSQL &	") TABLA "
strSQL = strSQL &	"WHERE KILOSNETOS > 0 "
strSQL = strSQL &	"ORDER BY TABLA.DTCONTABLE, TABLA.IDCAMION"

call GF_BD_Puertos(pto, rs, "OPEN",strSQL)
set armarSQLCosechaXCamiones = rs  
End function
 
%>
