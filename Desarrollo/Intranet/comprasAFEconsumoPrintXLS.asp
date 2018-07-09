<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosExcel.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<% Response.Buffer = False
Const TOTAL_COLUMNAS = 18
'************************************************************************************************************
Function writeFilter() %>
	<table style="font-size:12; font-weight:bold; font-family:courier;">
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="right" style="font-weight:normal; font-size:10"><% =GF_FN2DTE(session("MmtoSistema")) %><br><% =session("usuario") %></td></tr>
		<tr><td colspan="<% =TOTAL_COLUMNAS %>" align="center" style="font-size:24"><% =GF_TRADUCIR("PLANILLA DE CONSUMOS ASOCIADOS A AFES") %></td></tr>
		<tr><td></td></tr>
	</table>
	<table style="font-size:16; font-weight:bold; font-family:courier">
		<tr><td></td></tr>
		<tr><td>Desde......:</td><td align="left"><% =GF_INT2MES(g_MesDesde) &" de "& g_AnioDesde %></td></tr>
		<tr><td>Hasta......:</td><td align="left"><% =GF_INT2MES(g_MesHasta) &" de "& g_AnioHasta %></td></tr>
		<tr><td>Importe....:</td><td align="left"><% =GF_EDIT_DECIMALS(Cdbl(g_Importe),2) & " u$s" %></td></tr>				
		<tr><td></td></tr>
	</table>
<%
End Function
'-------------------------------------------------------------------------------------------------------------
Function getCodeLocation(pCdDivision)	
	Dim rtrn
	Select Case pCdDivision
		case CODIGO_EXPORTACION
			rtrn = "VB6"
		case CODIGO_ARROYO
			rtrn = "VN5"
		case CODIGO_PIEDRABUENA
			rtrn = "VN7"
		case CODIGO_TRANSITO		
			rtrn = "VN4"
	End select 
	getCodeLocation = rtrn
End Function 
'-------------------------------------------------------------------------------------------------------------
Function getCategoryAfe(pIdCategoria)
	Dim rtrn
	Select Case pIdCategoria
		case AFE_CATEGORIA_CAPITAL
			rtrn = "CAP"
		case else 
			rtrn = "EXP"		
	End select 
	getCategoryAfe = rtrn
End Function 
'-------------------------------------------------------------------------------------------------------------
Function getStatusAfe(pEstado)
	Dim rtrn
	Select Case pEstado
		case AFE_APROBADO
			rtrn = "Approved"
		case AFE_ANULADO 
			rtrn = "Rejected"		
		case else
			rtrn = "Pending"	
	End select 
	getStatusAfe = rtrn
End Function 
'--------------------------------------------------------------------------------------------------------------------
' Función:	  
'				getStringAFEComplement
' Autor: 	  
'				CNA - Ajaya Nahuel
' Fecha: 	  
'				19/03/2014
' Objetivo:	  
'				Obtiene un string concatenado de todos los AFEs complementarios que tiene un AFE
' Parametros: 
'				pIdAfe		[integer]	-  Id del afe
' Devuelve:		
'				Codigo de los Afe separadas por , [string]
'--------------------------------------------------------------------------------------------------------------------
Function getStringAFEComplement(pIdAfe)
	Dim rs,str
	Set rs = listaAFESComplementarios(pIdAfe)
	while not rs.EoF
		str = str & rs("CDAFE") & ","
		rs.MoveNext()
	wend
	if (Len(str) > 0) Then str = left(str,len(str)-1)	
	getStringAFEComplement = str
End Function
'--------------------------------------------------------------------------------------------------------------------
' Función:	  
'				getProcessTypeAFE
' Autor: 	  
'				CNA - Ajaya Nahuel
' Fecha: 	  
'				19/03/2014
' Objetivo:	  
'				Procesa el/los tipo que puede tener el AFE y devuelve la descripcion de cada uno separada por coma(contactenada)  
' Parametros: 
'				pTipo		[string]	-  El/los tipo/s  que tiene el afe
'				pTipoOtros  [string]	-  Descripcion de tipo Otro (si el parametro pTipo es Otro)
'				pTipoCC		[string]	-  El tipo de cumplimiento del afe (si el parametro pTipo es un Cumplimiento)
' Devuelve:		
'				Tipo que tiene el AFE [string]
'--------------------------------------------------------------------------------------------------------------------
Function getProcessTypeAFE(pTipo,pTipoOtros,pTipoCC)
	Dim str,arrTipo,k,auxTipo
	str = ""	
	arrTipo = Split(pTipo, ",")
	for k=LBound(arrTipo) to UBound(arrTipo)
		auxTipo = Trim(arrTipo(k))
		'Si se elige un tipo de cumplimiento, se muestra el tipo directamente.
		if (auxTipo = AFE_TIPO_CUMPIMIENTO) then auxTipo = pTipoCC
		str = str & getDescripcionTipoAFE(auxTipo)
		'Si eligió otros, entonces se muestra la descripción.	
		if (auxTipo = AFE_TIPO_OTROS) then str = str & pTipoOtros
		str = Trim(str) & ", "
	Next
	if Len(str) > 0 then str = left(str, Len(str)-2)
	getProcessTypeAFE = str
End Function 
'------------------------------------------------------------------------------------------------------------- 
'NOTA:
	'El período de fechas tiene el siguiente formato:
	'Fecha inicio : 1AAMMDD-->Ultimos dos digitos de cada uno. El dia de inicio siempre se considera 01.
	'Fecha final  : 1AAMMDD-->Ultimos dos digitos de cada uno. El dia de fin siempre se considera el ultimo dia del mes en cuestión.	
Function drawDetalle()	
	Dim idAfe_old,rs,fechaDesde,fechaHasta,sumFacturado
	fechaDesde = g_AnioDesde & "-" & g_MesDesde & "-01"
    fechaHasta = g_AnioHasta & "-" & g_MesHasta & "-" & LastDayOfMonth(g_AnioHasta, g_MesHasta)	
	Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLDATOSAFE_GET_CONSUMOS_BY_PARAMETERS", fechaDesde & "||"& fechaHasta &"||"& g_Importe &"||"& detalle)
	if not rs.Eof then
		if detalle = RESPUESTA_OK Then
			'Dibuja el reporte detalladamente, esto implica mostrar cada Pic con su minuta/fecha facturada
			while not rs.Eof 
				idAfe_old = cdbl(rs("IDAFE"))
				idAfe = idAfe_old %>
				<tr style="font-size:10;" class="border">
					<td align="center"><% =rs("CDAFE") %></td>
					<td align="center"><% =rs("DSDIVISION") %></td>
					<td align="center"><% =getCodeLocation(rs("CDDIVISIONABR")) %></td>
					<td align="center"><% =getCategoryAfe(rs("CATEGORIA")) %></td>
					<td align="center"><% =getProcessTypeAFE(rs("TIPO"),rs("TIPOOTROS"),rs("TIPOCC")) %></td>
					<td align="center"><% Response.Write "USD" %></td>
					<td align="left"><% =rs("TITULO") %></td>
					<td class="fieldImportant" align="center"><% =getStatusAfe(rs("CONFIRMADO")) %></td>
					<td class="fieldImportant" align="center"><% =getStringAFEComplement(rs("IDAFE")) %></td>
					<td align="center"><% If(rs("CONFIRMADO") = AFE_APROBADO)then Response.Write GF_FN2DTE(Left(rs("FECHAPUBLICACION"),8))  %></td>				
					<td align="left"><% =rs("CDOBRA") &"-"& rs("DSOBRA")%></td>
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("AFEPESOS")),2)%></td>
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("TIPOCAMBIO")),3) %></td>
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("AFEDOLARES")),2)%></td>
					<td align="right" colspan="4"></td>
				</tr>
				<%
				sumFacturado = 0
				While ((not rs.eof) and (idAfe_old = idAfe)) 
					sumFacturado = sumFacturado + Cdbl(rs("DOLARESFACT")) %>
					<tr style="font-size:10;">
						<td colspan="<%=TOTAL_COLUMNAS-4%>"></td>
						<td class="border" align="center"><% =rs("IDCOTIZACION")%></td>
						<td class="border" align="center"><% =rs("MINUTA")%></td>
						<td class="border" align="center"><% ="20" & Mid(rs("FECHA"),1,2) & "/" & Mid(rs("FECHA"),3,2) &"/"& Right(rs("FECHA"),2) %></td>
						<td class="border" align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("DOLARESFACT")),2)%></td>
					</tr>
				<%	rs.MoveNext()
					if (not rs.eof) then idAfe = Cdbl(rs("IDAFE"))
				wend %>
				<tr style="font-size:10;">
					<td colspan="<%=TOTAL_COLUMNAS-1%>"></td>
					<td class="fieldImportant" align="right" ><% =GF_EDIT_DECIMALS(Cdbl(sumFacturado),2)%></td>
				</tr>
				<tr><td colspan="<%=TOTAL_COLUMNAS%>"></td></tr>
			<%wend
		else
			'Dibuja el reporte sumando todos los gastos del AFE (sin detalle)
			while not rs.Eof %>
				<tr style="font-size:10;" class="border">
					<td align="center"><% =rs("CDAFE") %></td>
					<td align="center"><% =rs("DSDIVISION") %></td>
					<td align="center"><% =getCodeLocation(rs("CDDIVISIONABR")) %></td>
					<td align="center"><% =getCategoryAfe(rs("CATEGORIA")) %></td>
					<td align="center"><% =getProcessTypeAFE(rs("TIPO"),rs("TIPOOTROS"),rs("TIPOCC")) %></td>
					<td align="center"><% Response.Write "USD" %></td>
					<td align="left"><% =rs("TITULO") %></td>
					<td class="fieldImportant" align="center"><% =getStatusAfe(rs("CONFIRMADO")) %></td>
					<td class="fieldImportant" align="center"><% =getStringAFEComplement(rs("IDAFE")) %></td>
					<td align="center"><% If(rs("CONFIRMADO") = AFE_APROBADO)then Response.Write GF_FN2DTE(Left(rs("FECHAPUBLICACION"),8))  %></td>				
					<td align="left"><% =rs("CDOBRA") &"-"& rs("DSOBRA")%></td>					
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("AFEPESOS")),2)%></td>
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("TIPOCAMBIO")),3) %></td>
					<td align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("AFEDOLARES")),2)%></td>
					<td class="fieldImportant" colspan="4" align="right"><% =GF_EDIT_DECIMALS(Cdbl(rs("DOLARESFACT")),2)%></td>				
				</tr>
			<%	rs.MoveNext()
			wend
		end if		 
	else %>
		<tr>
			<td colspan="<%=TOTAL_COLUMNAS%>" class="border" align="center"><% =GF_TRADUCIR("No se encontraron resultados") %></td>
		</tr>		
<%	end if
End Function
'-------------------------------------------------------------------------------------------------------------
Dim g_AnioHasta,g_MesHasta,g_AnioDesde,g_MesDesde,g_Importe,fname,g_verDetalle,detalle

g_AnioHasta = GF_PARAMETROS7("anioHasta","",6)
g_MesHasta = GF_PARAMETROS7("mesHasta","",6)
g_AnioDesde = GF_PARAMETROS7("anioDesde","",6)
g_MesDesde = GF_PARAMETROS7("mesDesde","",6)
g_Importe = GF_PARAMETROS7("importe",2,6)
g_verDetalle = GF_PARAMETROS7("verDetalle","",6)
if Ucase(g_verDetalle) = "ON" Then detalle = RESPUESTA_OK
fname = "REPORTE_AFE_CONSUMO_" & session("MmtoSistema")
Call GF_createXLS(fname)
 %>
<html>
<head>
	<style type="text/css">
		.border { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
		}

		.titulos {
			background-color:#87D5B0;
			font-weight:bold;
			border-style:solid; 
			border-width:thin;
		}
		.fieldImportant{
			background-color:#F4F461;
			font-weight:bold;
			border-style:solid; 
			border-width:thin;
		}
	</style>
</head>
<body>	
	<table class="border" style="background-color:#FFFACD; font-weight:bold">
		<tr>
			<td><% Call writeFilter() %></td>
		</tr>
	</table>
	<table >
		<tr style="font-size:12;" >
			<td class="titulos" align="center"><% =GF_TRADUCIR("AFE No.") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Location") %></td>			
			<td class="titulos" align="center"><% =GF_TRADUCIR("Location Code") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Capital or Expense") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Asset Categories") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Currency") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Proyect Name / Description") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Approved / Pending") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Supp. AFE") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Start date") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Job") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Capital approved [LC]") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("USD rate at approval") %></td>
			<td class="titulos" align="center"><% =GF_TRADUCIR("Capital approved [USD]") %></td>						
			<td class="titulos" align="center" colspan="4"><% =GF_TRADUCIR("2014 January Actuals") %></td>
		</tr>
		<tr style="font-size:12;" >
			<td class="titulos" colspan="<%=TOTAL_COLUMNAS-4%>"></td>
			<% If detalle = RESPUESTA_OK Then %> 
				<td class="titulos" ><%=GF_TRADUCIR("Pic")%></td>
				<td class="titulos" ><%=GF_TRADUCIR("Minuta")%></td>
				<td class="titulos" ><%=GF_TRADUCIR("Date")%></td>
				<td class="titulos" ><%=GF_TRADUCIR("Invoiced [USD]")%></td>
			<% else %>
				<td colspan="4" class="titulos" ></td>				
			<% end if %>
		</tr>
		<% Call drawDetalle() %>
	</table>	
</body>

</html>


	