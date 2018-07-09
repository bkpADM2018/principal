<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosUser.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->

<%
Const ORDER_BZA_TARA = 4

dim rs, conn, strSQL, rsMovimientos,fechaDesde, fechaHasta, patente, acoplado, estado, pto,fileName,v_haveBZA(5),v_difrenciaBZA(5) , tControl

fechaDesde = GF_PARAMETROS7("fechaDesde", "", 6)
fechaHasta = GF_PARAMETROS7("fechaHasta", "", 6)
pto		   = GF_PARAMETROS7("pto", "", 6)
patente	   = GF_PARAMETROS7("patente", "", 6)
acoplado   = GF_PARAMETROS7("acoplado", "", 6)
tControl   = GF_PARAMETROS7("tControl", "", 6)
estado	   = GF_PARAMETROS7("estado", 0, 6)	
 

%>
<html>
<head>
	<style type="text/css">
		.xls_border_left { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
		}
		.xls_align_center { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
			text-align: center;
		}
		.xls_align_right { 
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;
			text-align: right;
		}
		.xls_precioUC_tabla
		{
			BACKGROUND-COLOR: #ffff80;
			border-color:#666666; 
			border-style:solid; 
			border-width:thin;

		}
		</style>
	</head>
	<body>	
	
<%


fileName = "BZA_Camiones" & gv_pto & "_" & Session("MmtoSistema")
Call GF_createXLS(fileName)
Call armarExcel()	
Call closeXLS()
'-----------------------------------------------------------------------------------------------
'Devuelve el numero de Balanzas con que se realizo el control , se obtiene mediante el campo CANTIDADBZA de la tabla CTRLBZACAMIONES 
Function getCantidadBalanza()
	Dim strSQL, myWhere, cantBZA 
	Call buscarFiltrosControlBalanza(myWhere,fechaDesde,fechaHasta,patente,acoplado,estado, tControl)
	strSQL = " SELECT COUNT(BRUTO1) CANT1, COUNT(BRUTO2) CANT2, COUNT(BRUTO3) CANT3, COUNT(BRUTO4) CANT4, COUNT(TARA) CANT5 FROM CTRLBZACAMIONES " & myWhere
	GF_BD_Puertos pto, rs, "OPEN", strSQL	
	cantBZA = 0
	if not rs.Eof then		
		if(rs("CANT1") > 0)then 
			v_haveBZA(0) = true
			cantBZA = cantBZA + 1
		end if	
		if(rs("CANT2") > 0)then 
			v_haveBZA(1) = true
			cantBZA = cantBZA + 1
		end if	
		if(rs("CANT3") > 0)then 
			v_haveBZA(2) = true
			cantBZA = cantBZA + 1
		end if	
		if(rs("CANT4") > 0)then 
			v_haveBZA(3) = true
			cantBZA = cantBZA + 1
		end if	
		if(rs("CANT5") > 0)then 
			v_haveBZA(4) = true		
			cantBZA = cantBZA + 1
		end if	
	end if	
getCantidadBalanza = cantBZA
End Function
'----------------------------------------------------------------------------------------------------
Function armarExcel()
	Dim totalReg, i, cantBZA, auxDiferencia
	cantBZA = getCantidadBalanza()	
	
	Set rs = leerControlBalanza(pto,fechaDesde,fechaHasta,patente,acoplado,estado, tControl)
	Call armadoCabecera(cantBZA)
	totalReg = rs.RecordCount	
	if not rs.eof then	
		while not rs.EoF			
			writeXLS("	<TR style='font-size:12;' >")
			writeXLS("		<TD align='center'>" & GF_FN2DTE(rs("FECHA")) & "</TD>")
			writeXLS("		<TD align='center'>" & GF_FN2DTE(rs("TIPOCONTROL")) & "</TD>")
			writeXLS("		<TD align='center'>" & UCASE(Left(rs("CDCHAPACAMION"),3)) & "-" & UCASE(Right(rs("CDCHAPACAMION"),3)) & "</TD>")			
			writeXLS("		<TD align='center'>" & UCASE(Left(rs("CDCHAPAACOPLADO"),3)) & "-" & UCASE(Right(rs("CDCHAPAACOPLADO"),3)) & "</TD>")			
			For i = 0 to UBound(v_haveBZA) - 1				
				if(v_haveBZA(i))then
					if(i = ORDER_BZA_TARA)then
						myPesoBZA = rs("TARA")												
						if IsNull(myPesoBZA)then myPesoBZA = 0
					else
						myPesoBZA = rs("BRUTO" & i + 1)						
						if ((not IsNull(myPesoBZA)) and (not IsNull(rs("TARA")))) then v_difrenciaBZA(i) = CDbl(rs("TARA")) - Cdbl(myPesoBZA)
					end if
					if(myPesoBZA <> "")then myPesoBZA = GF_EDIT_DECIMALS(Cdbl(myPesoBZA)*100,2) & " KGS "
					writeXLS("<TD position='absolute' align='center' >" & myPesoBZA & " </TD>")
				end if
			next
			For i = 0 to UBound(v_haveBZA) - 1				
				if((v_haveBZA(i))and(i <> ORDER_BZA_TARA))then 
					auxDiferencia = 0
					if(v_difrenciaBZA(i) <> "" )then auxDiferencia =  v_difrenciaBZA(i)
					writeXLS("<TD bgcolor='#F7BE81' align='center'>" & GF_EDIT_DECIMALS(Cdbl(auxDiferencia)*100,2) & " KGS </TD>")
				end if
			next
			writeXLS("<TD position='absolute' align='center' >" & getUserDescription(rs("CDUSR")) & "</TD>")
			writeXLS("<TD position='absolute' align='center' ></TD>")
			writeXLS("<TD position='absolute' align='center' ></TD>")
			writeXLS("<TD position='absolute' align='center' >"& getDsEstadoBZA(rs("ESTADO")) &"</TD>")	
			writeXLS("</TR>")
			rs.MoveNext()
		wend 
	else		 
		 writeXLS("		<TR style='font-size:16;line-height:50%' >")		
		 writeXLS("			<TD  colspan='" & ((cantBZA * 2) - 1) + 9 & "' align='center' ><B>NO SE ENCONTRARON RESULTADOS</B></TD>")		 
		 writeXLS("		</TR>")			 
	end if
	
	writeXLS("</TABLE>")
end Function
'----------------------------------------------------------------------------------------------------
Function armadoCabecera(pCantBZA)		
	Dim strName, auxPatente, auxAcoplado, auxEstado, cspan
	
	cspan = ((pCantBza * 2) - 1) + 5
	
	writeXLS("<TABLE border='1'>")	
	writeXLS("		<TR style='font-size:16;line-height:50%' >")	
	writeXLS("			<TD align='center' colspan='" & ((pCantBza * 2) - 1) + 7 & "'><B><U>PLANILLA DE CONTROL DE BALANZAS</U></B></TD>")
	writeXLS("		</TR>")
	writeXLS("		<TR border='hidden' style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>FECHA DESDE:</B></TD>")			
	if(fechaDesde = "")then
		auxFechaDesde = "TODAS"
	else
		auxFechaDesde = GF_FN2DTE(fechaDesde)
	end if	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & auxFechaDesde & "</B></TD>")
	writeXLS("		</TR>")
	writeXLS("		<TR border='hidden' style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>FECHA HASTA:</B></TD>")	
	if(fechaHasta = "")then 
		auxFechaHasta = "TODAS"
	else
		auxFechaHasta = GF_FN2DTE(fechaHasta)	
	end if	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & auxFechaHasta & "</B></TD>")
	writeXLS("		</TR>")	
	writeXLS("		<TR border='hidden' style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>PATENTE:</B></TD>")
	
	if(patente = "")then 
		auxPatente = "TODAS"
	else
		auxPatente = Ucase(Left(patente,3)) & "-" & Ucase(Right(patente,3))
	end if	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & auxPatente & "</B></TD>")
	writeXLS("		</TR>")	
	writeXLS("		<TR border='hidden' style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>ACOPLADO:</B></TD>")
	
	if(acoplado = "")then 
		auxAcoplado = "TODAS"
	else
		auxAcoplado = UCASE(Left(acoplado,3)) & "-" & Ucase(Right(acoplado,3))		
	end if	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & auxAcoplado & "</B></TD>")
	writeXLS("		</TR>")		
	writeXLS("		<TR style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>ESTADO:</B></TD>")	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & getDsEstadoBZA(estado) & "</B></TD>")
	writeXLS("		</TR>")	
	writeXLS("		<TR style='font-size:12;line-height:50%' >")	
	writeXLS("			<TD align='left' colspan='2'><B>TIPO:</B></TD>")	
	writeXLS("			<TD align='left' colspan='" & cspan & "'><B>" & getDsTipoCtrlBZA(tControl) & "</B></TD>")
	writeXLS("		</TR>")	
	
	writeXLS("		<TR class='xls_border_center' style='font-size:12;line-height:50%' >")
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>FECHA</B></TD>")
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>TIPO</B></TD>")
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>PATENTE</B></TD>")	
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>ACOPLADO</B></TD>")	
	For i = 0 to UBound(v_haveBZA) - 1 
		if(v_haveBZA(i))then
			strName = "BRUTO " & i + 1
			if(i = ORDER_BZA_TARA)then 	strName = "TARA"
			writeXLS("		<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>" & strName & "</B></TD>")
		end if	
	next
	For i = 0 to UBound(v_haveBZA) - 1 
		if((v_haveBZA(i))and(i <> ORDER_BZA_TARA))then
			strName = "DIF. BZA " & i + 1			
			writeXLS("		<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>" & strName & "</B></TD>")
		end if	
	next
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>REALIZO</B></TD>")
	writeXLS("			<TD bgcolor='#BCF5A9' colspan='2' position='absolute' align='center' ><B>PRESINTOS</B></TD>")
	writeXLS("			<TD bgcolor='#BCF5A9' rowspan='2' position='absolute' align='center' ><B>ESTADO</B></TD>")	
	writeXLS("		</TR>")
	writeXLS("		<TR class='xls_border_center' style='font-size:12;line-height:50%'>")
	writeXLS("			<TD bgcolor='#BCF5A9' align='center' ><B> OK </B></TD>")
	writeXLS("			<TD bgcolor='#BCF5A9' align='center' ><B> NO OK </B></TD>")
	writeXLS("		</TR>")
End Function
'------------------------------------------------------------------------------------------	
Function getDsEstadoBZA(pIdEstado)
	Dim strDesc
	strDesc = "TODOS"
	select case pIdEstado
		case BZA_CAM_ESTADO_EN_CURSO
			strDesc = "EN CURSO"
		case BZA_CAM_ESTADO_FINALIZADO
			strDesc = "FINALIZADO"
		case BZA_CAM_ESTADO_CANCELADO
			strDesc = "CANCELADO"
	end select	
	getDsEstadoBZA = strDesc	
End Function
'------------------------------------------------------------------------------------------	
Function getDsTipoCtrlBZA(pIdEstado)
	Dim strDesc
	strDesc = "TODOS"
	select case pIdEstado
		case BZA_CAM_TIPO_CTRL_MANUAL
			strDesc = "MANUAL"
		case BZA_CAM_TIPO_CTRL_AUTOM
			strDesc = "AUTOMATICO"		
	end select	
	getDsTipoCtrlBZA = strDesc	
End Function
%>