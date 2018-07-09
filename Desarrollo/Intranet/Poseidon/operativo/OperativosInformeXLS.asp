<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->
<%
'******************************************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'-------------------------------------------------------------------------------------------
Function drawFiltroOperativo()
Dim auxOperativo
%>
<table width="100%" align="center" border="0">	
    <tr>
		<td colspan=9 class="border">
			<table style="font-size:12;">
				<tr>
					<td align="right" colspan="2">
						<b>Operativo:</b>
					</td>
					<%	auxOperativo = "Todos"
						if myOperativo <> "" then auxOperativo = myOperativo %>
					<td align="left" colspan="2"><%= auxOperativo%></td>		
					<td align="right">
						<b>Carta Porte:</b>
					</td>
					<%	auxCartaPorte = "Todas"
						if myCartaPorte <> "" then auxCartaPorte = myCartaPorte %>
					<td align="left" colspan="3">
						<%=auxCartaPorte%>
					</td>
				</tr>
				<tr>	
					<td align="right" colspan="2">
						<b>Turno:</b>
					</td>
					<%	auxTurno = "Todos"
						if myTurno <> "" then auxTurno = myTurno %>
					<td align="left" colspan="2"><%=auxTurno%></td>
					<td align="right">
						<b>Vagon:</b>
					</td>
					<%	auxVagon = "Todos"
						if myIdVagon <> "" then auxVagon = myIdVagon %>
					<td align="left" colspan="3"><%=auxVagon%></td>	
				</tr>
				<tr>
					<td align="right" colspan="2">
						<b>Fecha Inicio Desde:</b>
					</td>
					<td align="left" colspan="2">
					<%	if myFecContableDesde <> "" then %>
							<%=GF_FN2DTE(myFecContableDesde)%>
					<%	end if	%>
					</td>
					<td align="right">
						<b>Fecha Inicio Hasta:</b>
					</td>
					<td align="left" colspan="3">
						<%	if myFecContableHasta <> "" then %>
							<%=GF_FN2DTE(myFecContableHasta)%>
						<%	end if	%>
					</td>
				</tr>
				<tr>
					<td align="right" colspan="2">
						<b>Coordinado:</b>
					</td>
					<%	auxCoordinado = "Todos"
						if myCdCoordinado <> "" then auxCoordinado = myDsCoordinado &"-"& getDsCliente(myDsCoordinado) %>
					<td align="left" colspan="2">
						<%=auxCoordinado%>
					</td>
					<td align="right" >
						<b>Producto:</b>
					</td>
					<%	auxProducto = "Todos"
						if myCdProducto <> "" then auxProducto = myCdProducto &"-"& getDsProducto(myCdProducto) %>
					<td align="left" colspan="3">
						<%=auxProducto%>
					</td>
				</tr>
				<tr>
					<td align="right" colspan="2">
						<b>Corredor:</b>
					</td>
					<%	auxCorredor = "Todos"
						if myCdCorredor <> "" then auxCorredor = myCdCorredor &"-"& getDsCorredor(myCdCorredor) %>
					<td align="left" colspan="2">
						<%=auxCorredor%>
					</td>
					<td align="right" >
						<b>Vendedor:</b>
					</td>
					<%	auxVendedor = "Todos"
						if myCdVendedor <> "" then auxVendedor = myCdVendedor &"-"& getDsVendedor(myCdVendedor)%>
					<td align="left" colspan="3">
						<%=auxVendedor%>
					</td>
				</tr>
				<tr>	
					<td align="right" colspan="2">
						<b>Estado:</b>
					</td>
					<%	auxEstado = "Todos"
						if myEstado <> 0 then auxEstado = getDsEstadoOperativo(myEstado,pto)%>
					<td align="left" colspan="2">
						<%=auxEstado%>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</table>	
<%End function
'---------------------------------------------------------------------------------------------
Function drawTitulosOperativos() %>
	<TR>	
		<TD  class="reg_header_nav" align="center">	<%=GF_Traducir("Turno")%> </TD>		
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Operativo")%> </TD>
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Carta Porte")%> </TD>
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Fecha Inicio")%> </TD>
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Coordinado")%> </TD>
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Producto")%> </TD>
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Corredor")%> </TD>		
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Vendedor")%> </TD>				
		<TD class="reg_header_nav" align="center">	<%=GF_Traducir("Estado")%> </TD>
	</TR>
	<%
End Function 
'-------------------------------------------------------------------------------------------
Function drawCabeceraOperativos(pStr)
	Dim myRegistro,h
	myRegistro = Split(pStr, ";") %>
	<TR>
<%	For h = 0 To UBound(myRegistro) %>
		<TD class="reg_header_navdos" align="center"><%=myRegistro(h)%></TD>
<%	Next  %>
	</TR>
<%	
End Function


'********************************************************************
'					INICIO PAGINA
'********************************************************************
Call GP_CONFIGURARMOMENTOS()
Dim cont,index
index = 0
pto = GF_PARAMETROS7("pto", "", 6)
g_strPuerto = pto
call addParam("pto", pto, params)
pTipo = GF_PARAMETROS7("pTipo", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)
maxSegment = GF_PARAMETROS7("maxSegment", 0, 6)
totalVagones = 0
totalKilosNetos = 0
Call getParametros()
Call GF_createXLS("Operativos_"&session("MmtoSistema"))
Set fs = Server.CreateObject("Scripting.FileSystemObject")			
flagHayResultado = false %>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />   
    <style type="text/css">
        .reg_header
        {
            BORDER-BOTTOM: #f4b800 1px solid;
            BORDER-LEFT: #f4b800 1px solid;
            BACKGROUND-COLOR: #ffeecd;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            FONT-SIZE: 10px;
            BORDER-TOP: #f4b800 1px solid;
            BORDER-RIGHT: #f4b800 1px solid;
            TEXT-DECORATION: none;
            -moz-border-radius: 5px 5px 5px 5px
        }
        .reg_header_error
        {
            BORDER-BOTTOM: #f80800 1px solid;
            BORDER-LEFT: #f40800 1px solid;
            BACKGROUND-COLOR: #ffaa99;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            COLOR: #ffffff;
            FONT-SIZE: 10px;
            BORDER-TOP: #f40800 1px solid;
            FONT-WEIGHT: bold;
            BORDER-RIGHT: #f40800 1px solid;
            TEXT-DECORATION: none
        }
        .reg_header_nav
        {
            BACKGROUND-COLOR: #517b4a;
            COLOR: #ffffff;
            FONT-SIZE: 10px;
            FONT-WEIGHT: bold
        }
        .reg_header_navdos
        {
            BACKGROUND-COLOR: #dcdcdc;
            COLOR: #006400;
            FONT-SIZE: 10px;
            FONT-WEIGHT: bold
        }
        .titu_header
        {
            BORDER-BOTTOM: #006400 1px solid;
            BORDER-LEFT: #006400 1px solid;
            BACKGROUND-COLOR: #517b4a;
            FONT-FAMILY: verdana,arial,san-serif;
            HEIGHT: 19px;
            COLOR: white;
            FONT-SIZE: 12px;
            BORDER-TOP: #006400 1px solid;
            FONT-WEIGHT: bold;
            BORDER-RIGHT: #006400 1px solid;
            TEXT-DECORATION: none
        }
    </style>
</head>
<body>
<table  border="1" cellpadding="0" cellspacing="0" width="60%">
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td colspan="9" align="center" class="titu_header"><%=GF_Traducir("Informe de Operativos")%></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<%Call drawFiltroOperativo %>
		</td>
	</tr>
	<tr>
		<td>
		<table class="reg_header" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">
	<%
			while index <= maxSegment				
				pStrPath = Server.MapPath("../Temp/OPERATIVOS_" & session("Usuario") & "_" & index & ".txt")
				if (fs.FileExists(pStrPath)) then		
					Set fadm = fs.OpenTextFile(pStrPath, 1)	
					while (not fadm.AtEndOfStream)
						if not flagHayResultado then Call drawTitulosOperativos()
						txtLine = fadm.ReadLine()
						Call drawCabeceraOperativos(txtLine)
						flagHayResultado = true
					wend
					Set fadm = nothing
					fs.DeleteFile(pStrPath)
				end if
				index = index + 1
			wend 
			if not flagHayResultado then %>
			<TR>
				<TD  class="reg_header_nav" colspan="9" align="center">	<%=GF_Traducir("No se encontraron resultados")%> </TD>
			</TR>
		<%	end if
			%>
		</table>
	</tr></td>
</table>
</body>