<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/cor-IncludeCTO.asp"-->
<!--#include file="Includes/cor-IncludePC.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientospaginacion.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<%
Call ProcedimientoControl("CONFWEB")

Dim strErrorMsg, dia, mes, anio, rsContratos, sqlCampoOrden
Dim strTitulo,strLinkPagina,intMostrar, intIndex, unitDest
Dim cmbProducto, intProducto, aniosql, diasql, messql, retValue
Dim accion, flagCorredor, registroContrato, TDestado, TRlink
Dim rs, conn, campoOrden
dim chkVerHistoricos
'******************************************************************************************
  Function addParam(p_key,p_value,ByRef p_param)
           if (not isEmpty(p_value)) then
              if (isEmpty(p_param)) then
                 p_param = "?"
              else
                 p_param = p_param & "&"
              end if
              p_param = p_param & p_key & "=" & p_value
           end if
  End Function
'******************************************************************************************
'/**
' * Funcion: initHeaderAConfirmar
' * Descripcion: Esta funcion inicializa la lectura de la
' *              cabecera de los contratos para la opcion de CONFIRMA
' * Valor Devuelto: recordset con los datos leidos.
' * Autor: Javier A. Scalisi
' * Fecha 29/09/2010
' */
Function initHeaderAConfirmar(corredor, vendedor, producto, sucursal, operacion, numero, cosecha, fechaConcertacion, campoOrden)
    Dim strWhere, strORKC, oConn, strSQL, strOrder
    
    'Se preparan los filtros de información.
    if ((corredor = 0) and (vendedor = 0)) then
        strORKC = CInt(session("KCOrganizacion"))
        if (strORKC <> CDTOEPFER) then
            'Si no es el usuario de toepfer, q el usuario logueado sea el corredor o el vendedor
            strWhere = "where (CTO.CVENR1=" & strORKC & " or CTO.CCORR1=" & strORKC & ")"
			call mkWhere(strWhere, "CTO.CVENR1","1","<>",1)
			call mkWhere(strWhere, "CTO.CCORR1","1","<>",1)            
        else
            call mkWhere(strWhere, "CTO.CCORR1","5454","<>",3) 'Esto es para q no muestre los Boletos del Corredor 5454
        end if
    else
        if (corredor <> 0) then call mkWhere(strWhere, "CTO.CCORR1",corredor,"=",1)
        if (vendedor <> 0) then call mkWhere(strWhere, "CTO.CVENR1",vendedor,"=",1)
        call mkWhere(strWhere, "CTO.CVENR1","1","<>",1)
        call mkWhere(strWhere, "CTO.CCORR1","1","<>",1)
    end if    
    
    Call mkWhere(strWhere, "CTO.CSUCR1", 1,"<>",1)
    if (sucursal <> 0)	then Call mkWhere(strWhere, "CTO.CSUCR1", sucursal	,"=",1)
    if (operacion <> 0) then Call mkWhere(strWhere, "CTO.COPER1", operacion	,"=",1)
    if (numero <> 0)	then Call mkWhere(strWhere, "CTO.NCTOR1", numero	,"=",1)
    if (cosecha <> 0)	then Call mkWhere(strWhere, "CTO.ACOSR1", cosecha	,"=",1)
    
    if chkVerHistoricos = 0 then call mkWhere(strWhere, "CTO.CONFR1", "F", "=", 3)
    
    call mkWhere(strWhere, "BOL.PLRECI", "V", "<>", 3)
    
    'Se toman contratos de la cosecha 09 en adelante
    call mkWhere(strWhere, "CTO.ACOSR1", 10, ">=", 1)
    
    if (fechaConcertacion <> "") then strWhere = strWhere & GF_LIKE("CTO.FCCTR1", fechaConcertacion)
    
    if (producto <> 0) then Call mkWhere(strWhere, "CTO.CPROR1", producto,"=",1)
    
    'Los contratos de cebada solo deben mostrarse cuando sea para la consulta de la caratula
    'pero no para imprimir boleto o confirmar el contrato.
    'Caso especial solicitado por Ronchel
    'Si la Operacion es 9 se deben poder confirmar los contratos de Colza(09)
    strWhere = strWhere & " AND ((CTO.CPROR1 not in (9, 17, 25, 28))"  & _
                          " or (CTO.CPROR1 = 9 and CTO.COPER1 in (0, 9))" & _ 
                          " or (CTO.CPROR1 = 17 and CTO.COPER1 in (0, 1, 9) and CTO.ACOSR1 <= 17)) "
	'Se quita la operacion 04 (prestamos)
	strWhere = strWhere & " AND CTO.COPER1 <> 04 "   

    'Se arma la SQL.
	strSQL = "Select CTO.CPROR1 as Producto, CTO.CSUCR1 as Sucursal, CTO.COPER1 as Operacion, CTO.NCTOR1 as Numero, " & _
			 "   CTO.CONFR1, CTO.ACOSR1 as Cosecha, CTO.FCCTR1 as FechaConc, CTO.CCORR1 as KCCOR, CTO.CVENR1 as KCVEN," & _ 
			 "   CASE WHEN AAC.TOTAL IS NULL THEN CTO.KGCOR1 ELSE CTO.KGCOR1 + AAC.TOTAL END as Kilos, CF.PRODUCTO PRODCONF" & _
			 "   from MERFL.MER311F1 CTO" & _ 
			 "       LEFT JOIN MERFL.MER341F2 BOL on CTO.CPROR1=BOL.PLCPRO and CTO.CSUCR1=BOL.PLCSUC" & _ 
			 "           and CTO.COPER1=BOL.PLCOPE and CTO.NCTOR1=BOL.PLNCTO and CTO.ACOSR1=BOL.PLACOS" & _ 
			 "       LEFT JOIN" & _ 
			 "               (" & _ 
			 "               SELECT CPRORB,CSUCRB,COPERB,NCTORB,ACOSRB, SUM(KGCORB) AS TOTAL FROM MERFL.MER311FB GROUP BY CPRORB,CSUCRB,COPERB,NCTORB,ACOSRB" & _
			 "               )AAC" & _ 
			 "           on CTO.CPROR1=AAC.CPRORB and CTO.CSUCR1=AAC.CSUCRB" & _ 
			 "           and CTO.COPER1=AAC.COPERB and CTO.NCTOR1=AAC.NCTORB and CTO.ACOSR1=AAC.ACOSRB" & _ 
			 "       LEFT JOIN TOEPFERDB.TBLCONTRATOSCONF CF on CTO.CPROR1=CF.PRODUCTO" & _ 
			 "           and CTO.CSUCR1=CF.SUCURSAL and CTO.COPER1=CF.OPERACION and CTO.NCTOR1=CF.NUMERO" & _ 
			 "           and CTO.ACOSR1=CF.COSECHA " & strWhere
     
     
     
    if (campoOrden = "") then
        campoOrden = " CTO.FCCTR1 asc"
    else
        campoOrden = campoOrden
    end if    
    strSQL = strSQL & " AND (((AAC.TOTAL IS NULL) AND (CTO.KGCOR1>0)) OR ((NOT AAC.TOTAL IS NULL) AND (CTO.KGCOR1 + AAC.TOTAL<>0))) order by " & campoOrden

    
    'response.write "<br>la consulta dentro de include es " &  strSQL & "<BR>"
    Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    Set initHeaderAConfirmar = rs    
End Function
'******************************************************************************************
'*******	COMIENZO DE LA PAGINA
'******************************************************************************************

'Recupero los parametros.
cmbProducto= GF_Parametros7("cmbProducto",0,6)
intProducto= GF_Parametros7("txtProducto",0,6)
chkVerHistoricos= GF_Parametros7("chkVerHistoricos",0,6)
Call addParam("chkVerHistoricos",chkVerHistoricos,param)

producto = cmbProducto
if (producto = 0) then producto = intProducto
Call addParam("cmbProducto",producto,param)

sucursal= GF_Parametros7("txtSucursal",0,6)
Call addParam("txtSucursal", sucursal, param)

operacion= GF_Parametros7("txtOperacion",0,6)
Call addParam("txtOperacion", operacion, param)

numero= GF_Parametros7("txtNumero",0,6)
Call addParam("txtNumero", numero, param)

cosecha= GF_Parametros7("txtCosecha",0,6)
Call addParam("txtCosecha", cosecha, param)

dia = trim(GF_Parametros7("txtDia","",6))
Call addParam("txtDia",dia,param)

mes = trim(GF_Parametros7("txtMes","",6))
Call addParam("txtMes",mes,param)

anio = trim(GF_Parametros7("txtAnio","",6))
Call addParam("txtAnio",anio,param)

idCorredor= GF_Parametros7("txtCorredor",0,6)
call addParam("txtKCCOR", idCorredor, param)

idVendedor = trim(GF_Parametros7("txtVendedor",0,6))
call addParam("txtVendedor", idVendedor, param)

unitDest = GF_Parametros7("UnidadDestino",0,6)
if (unitDest = 0) then unitDest= UNIDAD_KILOS
Call addParam("UnidadDestino",unitDest,param)

campoOrden = GF_Parametros7("campoOrden","",6)
Call addParam("campoOrden",campoOrden,param)
select case campoOrden
    case "FechaAsc": sqlCampoOrden = "CTO.FCCTR1 asc"
    case "FechaDesc": sqlCampoOrden = "CTO.FCCTR1 desc"
    case "KilosAsc": sqlCampoOrden = "CTO.KGCOR1 Asc"
    case "KilosDesc": sqlCampoOrden = "CTO.KGCOR1 Desc"
    case "ContratoAsc": sqlCampoOrden = "CTO.CPROR1 asc, CTO.CSUCR1 asc, CTO.COPER1 asc, CTO.NCTOR1 asc, CTO.ACOSR1 asc"
    case "ContratoDesc": sqlCampoOrden = "CTO.CPROR1 desc, CTO.CSUCR1 desc, CTO.NCTOR1 desc, CTO.NCTOR1 desc, CTO.ACOSR1 desc"    
    case else: sqlCampoOrden = "CTO.FCCTR1 asc"
end select

accion = GF_Parametros7("accion","",6)

errorMsg= ""

fechaConsulta=""
if ((dia <> "") or (mes <> "") or (anio <> "")) then
	'Hay una fecha o parte de ella para filtrar los datos.
	diaSql  = "__"
	mesSql  = "__"
	anioSql = "____"	
	if (dia <> "") then diaSql = GF_nDigits(dia, 2)
    if (mes <> "") then mesSql = GF_nDigits(mes, 2)
    if (anio <> "") then anioSql = GF_nDigits(anio, 4)
	fechaConsulta = anioSql & mesSql & diaSql    
end if

'Se determina si es corredor o vendedor
if (GF_ES_CORREDOR(session("KCOrganizacion"))) then
	strTitulo="Vendedor"
	flagCorredor=true
else
	if session("KCOrganizacion") = KC_TOEPFER then
		strTitulo="Corredor/Vendedor"
	else
		strTitulo = "Corredor"	
	end if	
	flagCorredor=false
end if
'-------------------------------------------------------------------------------------
function tieneRegistro(producto, sucursal, operacion, numero, cosecha, byref registro)
dim rs, oConn, strSQL, registroContrato, rtrn
registro = ""
rtrn = false
strSQL = "SELECT * FROM TOEPFERDB.TBLCONTRATOSCONF WHERE PRODUCTO=" & producto & " AND SUCURSAL=" & sucursal & " AND OPERACION=" & operacion & " AND NUMERO=" & numero & " AND COSECHA=" & cosecha 
Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
if not rs.eof then
	if not isnull(rs("REGISTRO")) then
		registro = rs("REGISTRO")
		rtrn = true
	end if	
end if
tieneRegistro = rtrn
end function
%>
<html>
<head>
  <title>TOEPFER INTERNATIONAL - <%GF_Traducir("Contratos")%></title>
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
  <link rel="stylesheet" href="css/iwin.css" type="text/css">
  <script language="javascript" src="scripts/script_fechas.js"></script>
  <script language="javascript" src="scripts/scripts_ordenar.js"></script>
  <script language="javascript" src="scripts/script_checkboxes.js"></script>
  <script type="text/javascript" src="scripts/iwin.js"></script>	
  
  <script language="javascript">
		var childWindow
        function fcnCall(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha) {
			var params = "";
			params = "cmbProducto=" + p_intProducto + "&txtSucursal=" +  p_intSucursal;
			params= params + "&txtOperacion=" + p_intOperacion + "&txtNumero=" + p_intNumero;
			params= params + "&txtCosecha=" + p_intCosecha;
			childWindow = window.open('contratosDetalleConfirma.asp?' + params, "Detalle");
		    if (childWindow.opener == null) childWindow.opener = self;
        }
	function openRegistro(pProducto, pSucursal, pOperacion, pNumero, pCosecha) {	
		popUp = new PopUpWindow('popUp', 'contratosRegistro.asp?producto=' + pProducto + '&numero=' + pNumero + '&sucursal=' + pSucursal + '&operacion=' + pOperacion + '&cosecha=' + pCosecha, '700', '450', 'Registro de Contrato');
	}   
	function closePopUp(){
		self.close();
	}        
  </script>
</head>
<body>
<% GF_TITULO_2("Confirmación de Contratos") %>
<form method="POST" name="form1" action="contratosCabeceraConfirma.asp">
<table width="80%" align="center">
    <tr>
        <td align="center" width="100%">
			<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
            <input type="hidden" name="accion" id="accion" value="">
                <tr>
                    <td width="8"><img src="images/marco_r1_c1.gif"></td>
                    <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
                    <td width="8"><img src="images/marco_r1_c3.gif"></td>
                    <td width="73%"><td>
                    <td></td>
                </tr>
                <tr>
                    <td width="8"><img src="images/marco_r2_c1.gif"></td>
                    <td align=center valign="center"><font class="big" color="#517b4a"><% =GF_Traducir("Busqueda")%></font></td>
                    <td width="8"><img src="images/marco_r2_c3.gif"></td>
                    <td></td>
                    <td></td>
                </tr>
                <tr>
                    <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
                    <td></td>
                    <td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="6"></td>
                    <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
                    <td width="8"><img src="images/marco_r1_c3.gif"></td>
                </tr>
                <tr>
                    <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
                    <td colspan="3">
                        <table width="100%" align="center" border="0">
                            <tr>
                                <td align="right"><% =GF_Traducir("Fecha Conc.")%>:</td>
                                <td colspan="3">
                                    <input type="text" size="2" maxLength="2" value="<% =dia %>" name="txtDia" onBlur="javascript:ControlarDia(this);"> /
                                    <input type="text" size="2" maxLength="2" value="<% =mes %>" name="txtMes" onBlur="javascript:ControlarMes(this);"> /
                                    <input type="text" size="4" maxLength="4" value="<% =anio%>" name="txtAnio" onBlur="javascript:ControlarAnio(this);">
                                </td>

                                <td rowspan="5" valign="center" align="center"><input type="submit" value="<% =GF_Traducir("Buscar")%>..."></td>
                            </tr>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Producto")%>:</td>
                                <td colspan="3">
                                <% strSQL="Select * from MERFL.MER112F1 order by DESCPR asc"
                                   call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
                                %>
                                <select name="cmbProducto">
                                        <option SELECTED value="0">- <% =GF_TRADUCIR("Todos") %> -
                                <% while (not rs.eof)
                                        if (cmbProducto = CLng(rs("CODIPR"))) then %>
                                        <option SELECTED value="<% =rs("CODIPR") %>"><% =GF_TRADUCIR(rs("DESCPR")) %>
                                <%      else %>
                                        <option value="<% =rs("CODIPR") %>"><% =GF_TRADUCIR(rs("DESCPR")) %>
                                <%      end if %>
                                <%      rs.MoveNext
                                   wend
                                %>
                                </select>
                                </td>
                            </tr>
                            <%if (session("KCOrganizacion") = CDTOEPFER) then%>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Contrato")%>:</td>
                                <td colspan="3">
                                    <input type="text" size="2" maxLength="2" value="<% =producto %>" name="txtProducto"> -
                                    <input type="text" size="1" maxLength="1" value="<% =sucursal %>" name="txtSucursal"> -
                                    <input type="text" size="2" maxLength="2" value="<% =operacion %>" name="txtOperacion"> -
                                    <input type="text" size="5" maxLength="5" value="<% =numero %>" name="txtNumero"> /
                                    <input type="text" size="2" maxLength="2" value="<% =cosecha %>" name="txtCosecha">
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Cod. Corredor")%>:</td>
                                <td colspan="2">
                                    <input type="text" size="6" maxLength="6" value="<% =idCorredor %>" name="txtCorredor">
                                </td>
                                <td align="left"><% =GF_Traducir("Cod. Vendedor")%>:
									<input type="text" size="6" maxLength="6" value="<% =vendedor %>" name="txtVendedor">
								</td>
                            </tr>
							<tr>
								<td align="right">&nbsp;</td>
								<td colspan="3" align="left">	
									<input style="cursor:pointer;" type="checkBox" value="1" <%if chkVerHistoricos=1 then Response.Write "CHECKED" %> name="chkVerHistoricos">
									<% =GF_TRADUCIR("Incluir historicos") %>
                                </td>
                            </tr>                                
                            <%end if%>
                     </table>
                </td>
                <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
            </tr>
            <tr>
           <td width="8"><img src="images/marco_r3_c1.gif"></td>
           <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r3_c3.gif"></td>
            </tr>
    </table>
    </td>
    <td width="5" align="left"><div id="divAdobe" style="visibility:hidden;position:absolute;"><img src="images/get_adobe_reader.gif" onClick="javascript:window.open('http://www.adobe.com/products/acrobat/readstep2.html');" style="cursor:hand;"></div>
    </td>
    </tr>
</table>
<br>
<% 
	if (strErrorMsg = "") then
		Set rsContratos = initHeaderAConfirmar(idCorredor, idVendedor, producto, sucursal, operacion, numero, cosecha, fechaConsulta, sqlCampoOrden)
		if (not rs.eof) then
			strLinkPagina = "contratosCabeceraConfirma.asp" & param
			call GF_PAGINAR("N",strLinkPagina,intMostrar,50,rsContratos)%>
			<table class="reg_Header" border="0" cellspacing="1" cellpadding="2" width="100%">
				<tr>
				    <td colspan="3">
				    </td>
				    <td colspan="4" align="right"><% =GF_Traducir("Mostrar en")%>:
							<select name="UnidadDestino" id="UnidadDestino" onChange="javascript:form1.accion.value='unidad';form1.submit();">
								<option value="<% =UNIDAD_KILOS %>" <% if (unitDest = UNIDAD_KILOS) then response.write " selected" %>>Kilos
								<option value="<% =UNIDAD_TONELADAS %>" <% if (unitDest = UNIDAD_TONELADAS) then response.write " selected" %>>Toneladas
				        </select>
				    </td>
				</tr>
				<tr class="reg_Header_nav" align="center">
				    <td>
				        <IMG src="images/arrow_up.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','FechaAsc');">
				        <% =GF_TRADUCIR("Fecha Conc.") %>
				        <IMG src="images/arrow_down.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','FechaDesc');">
				    </td>
				    <%if (session("KCOrganizacion") = CDTOEPFER) then%>
				    <td>
				        <IMG src="images/arrow_up.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','ContratoAsc');">
				        <% =GF_TRADUCIR("Contrato") %>
				        <IMG src="images/arrow_down.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','ContratoDesc');">
				    </td>
				    <%end if%>
				    <td>
				        <% =GF_TRADUCIR(strTitulo) %>
				    </td>
				    <td>
				        <IMG src="images/arrow_up.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','KilosAsc');">
				        <% =GF_TRADUCIR("Kg Contratados") %>
				        <IMG src="images/arrow_down.gif" align=absMiddle style="cursor:pointer;" border=0 onClick="ordenar_onClick('contratosCabeceraConfirma.asp','<% =param%>','KilosDesc');">
				    </td>
				    <td>
				        <% =GF_TRADUCIR("Estado") %>
				    </td>
			        <% if (session("KCOrganizacion") = CDTOEPFER) then %>
							<td>
							    .
							</td>
				    <%end if%> 
				</tr>
				<%
				intIndex = 0
				while ((not rsContratos.eof) and (CInt(intIndex) < CInt(intMostrar))) 						
					if (trim(rsContratos("CONFR1"))) = "V" then 
						TDestado = "<td align='center'><font color='green'><b>" & GF_Traducir("Confirmado") & "</b></font></td>"
						TRlink = ""
					else
						TRlink = " style='cursor:pointer;' onclick='fcnCall(" & rsContratos("Producto") & "," & rsContratos("Sucursal") & "," & rsContratos("Operacion") & "," & rsContratos("Numero") & "," & rsContratos("Cosecha") & ")';"
						if (isNull(rsContratos("ProdConf"))) then 
							TDestado = "<td " & TRlink & " align='center'><font color='red'><b>" & GF_Traducir("Sin confirmar") & "</b></font></td>"
						else
							if (session("KCOrganizacion") <> CDTOEPFER) then TRlink = ""
							TDestado = "<td " & TRlink & " align='center'><font color=#ff9933><b>" & GF_Traducir("Pendiente") & "</b></font></td>"
						end if			
					end if
					%>
					<tr class="reg_Header_navdos">
						<td <%=TRlink%> align="center"><% =GF_FN2DTE(rsContratos("FechaConc")) %></td>
							<% if (session("KCOrganizacion") = CDTOEPFER) then %>
								<td <%=TRlink%> align="center">
									<%=GF_EDIT_CONTRATO(rsContratos("Producto"),rsContratos("Sucursal"),rsContratos("Operacion"),rsContratos("Numero"),rsContratos("Cosecha"))%>
								</td>
							<%end if%>
						<TD <%=TRlink%>>
							<%
							if clng(rsContratos("KCCOR")) = SIN_CORREDOR then
								vendedor = GetDSEnterprise2(rsContratos("KCVEN"))
							else	
								if (flagCorredor) then
									vendedor = GetDSEnterprise2(rsContratos("KCVEN"))
								else
								    vendedor = GetDSEnterprise2(rsContratos("KCCOR"))
								end if
							end if
							%>
							<span title="<%=vendedor%>">
								<%
								if (len(vendedor)>50)  then
								    response.write Left(vendedor,50) & "..."
								else
								    response.write vendedor
								end if
								%>
							</span>
						</TD>
						<%
						retValue = CDbl(rsContratos("Kilos"))*100
						if (unitDest = UNIDAD_TONELADAS) then retValue = retValue/1000             
						%>
						<td <%=TRlink%> align="right"><% =GF_EDIT_DECIMALS(retValue, 2)%></td>
						
						<%=TDestado%>
						
						<% if (session("KCOrganizacion") = CDTOEPFER) then %>
							<%if tieneRegistro(rsContratos("Producto"),rsContratos("Sucursal"),rsContratos("Operacion"),rsContratos("Numero"),rsContratos("Cosecha"), registroContrato) then%>
								<td align="center" style="cursor:pointer;" onclick=openRegistro(<%=rsContratos("Producto")%>,<%=rsContratos("Sucursal")%>,<%=rsContratos("Operacion")%>,<%=rsContratos("Numero")%>,<%=rsContratos("Cosecha")%>);>
								    <img src="images/contratos/Mail-16x16.png">
								</td>
							<%else%> 
								<td align="center">
								   .
								</td>
							<%end if%> 
						<%end if%> 						
					</tr>
					<%
				intIndex = intIndex + 1
				rsContratos.MoveNext()
			wend
			%>
		</table>
<%	
	else 'intHeader
%>
		<table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
			<tr>
				<td width="8"><img src="images/marco_r1_c1.gif"></td>
                <td width="100%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
                <td width="8"><img src="images/marco_r1_c3.gif"></td>
            </tr>
            <tr>
                <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
                <td align="center" class="TDTOTALES"><% =GF_TRADUCIR("NO SE ENCONTRARON CONTRATOS PARA MOSTRAR") %></td>
                <td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
            </tr>
            <tr>
                <td width="8"><img src="images/marco_r3_c1.gif"></td>
                <td width="100%"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
                <td width="8"><img src="images/marco_r3_c3.gif"></td>
            </tr>
		</table>
<% end if %>
<INPUT TYPE="HIDDEN" NAME="campoOrden" id="campoOrden" VALUE="<% =campoOrden %>">
</form>
<%else 'ERROR
%>
   <table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
          <tr>
              <td width="8"><img src="images/marco_r1_c1.gif"></td>
              <td width="100%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r1_c3.gif"></td>
          </tr>
          <tr>
              <td width="8" height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
              <td align="center" class="TDERROR"><% =GF_TRADUCIR(strErrorMsg) %></td>
              <td width="8" height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
          </tr>
          <tr>
              <td width="8"><img src="images/marco_r3_c1.gif"></td>
              <td width="100%"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
              <td width="8"><img src="images/marco_r3_c3.gif"></td>
          </tr>
   </table>
<%
end if
%>
</BODY>
</HTML>