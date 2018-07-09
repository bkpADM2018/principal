<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientospaginacion.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/cor-IncludeCTO.asp"-->
<!--#include file="Includes/cor-IncludePC.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<%
'---------------------------------------------------------------------------
Function setUnidadPeso(p_unit, p_kilos, ByRef p_value)
	if (p_unit = 1) then	'kilogramos	
		p_value = GF_EDIT_INTEGER(CDbl(p_kilos))
	else	'toneladas
		p_value = GF_EDIT_DECIMALS(p_kilos, 3)
	end if	
end Function
'---------------------------------------------------------------------------
'Call ProcedimientoControl("CONTCAB")

Dim strErrorMsg,p_intDia,p_intMes,p_intAnio,strDS
Dim strTitulo,strLinkPagina,intMostrar, intIndex, unitDest
Dim cmbProducto, intProducto, aniosql, diasql, messql, retValue
Dim accion,strConfirmado
Dim rs, conn, strCampoOrden

' Existen 3 tipos de operaciones que se realizan en esta pagina

' BOLETO -> Usuario que ingresa a ver los boletos pendientes de recepcion, contrato debe estar confirmado
' CONTRATO -> Usuario ingresa a ver Contratos

'Recupero los parametros.
strTipo=GF_Parametros7("strTipo","",6)
if (strTipo = "") then strTipo = "CONTRATO"
strTipo = ucase(strTipo)
g_tipo= strTipo
Call addParam("strTipo", strTipo, strParam)
cmbProducto= trim(GF_Parametros7("cmbProducto",0,6))
intProducto= trim(GF_Parametros7("txtProducto","",6))
g_intProducto = cmbProducto
if (g_intProducto = 0) then g_intProducto = intProducto
Call addParam("cmbProducto",g_intProducto,strParam)
g_intSucursal= trim(GF_Parametros7("txtSucursal","",6))
Call addParam("txtSucursal",g_intSucursal,strParam)
g_intOperacion= trim(GF_Parametros7("txtOperacion","",6))
Call addParam("txtOperacion",g_intOperacion,strParam)
g_intNumero= trim(GF_Parametros7("txtNumero","",6))
Call addParam("txtNumero",g_intNumero,strParam)
g_intCosecha= trim(GF_Parametros7("txtCosecha","",6))
Call addParam("txtCosecha",g_intCosecha,strParam)
p_intDia= trim(GF_Parametros7("txtDia","",6))
Call addParam("txtDia",p_intDia,strParam)
p_intMes= trim(GF_Parametros7("txtMes","",6))
Call addParam("txtMes",p_intMes,strParam)
p_intAnio= trim(GF_Parametros7("txtAnio","",6))
Call addParam("txtAnio",p_intAnio,strParam)
g_intKCCOR = trim(GF_Parametros7("txtCorredor","",6))
call addParam("txtCorredor", g_intKCCOR, strParam)
g_intKCVEN = trim(GF_Parametros7("txtVendedor","",6))
call addParam("txtVendedor", g_intKCVEN, strParam)
unitDest = GF_Parametros7("UnidadDestino","",6)
if unitDest = "" then unitDest= "1"
Call addParam("UnidadDestino",unitDest,strParam)
strCampoOrden = GF_Parametros7("campoOrden","",6)
select case strCampoOrden
    case "FechaAsc": g_strCampoOrden = "FechaConc asc"
    case "ContratoAsc": g_strCampoOrden = "Producto asc, Sucursal asc, Operacion asc, Numero asc, Cosecha asc"
    case "ContratoDesc": g_strCampoOrden = "Producto desc, Sucursal desc, Operacion desc, Numero desc, Cosecha desc"
    case "KilosAsc": g_strCampoOrden = "Kilos Asc"
    case "KilosDesc": g_strCampoOrden = "Kilos Desc"
    case "EntAsc": g_strCampoOrden = "SaldoEnt Asc"
    case "EntDesc": g_strCampoOrden = "SaldoEnt Desc"
    case else: g_strCampoOrden = "FechaConc desc"
end select
accion = GF_Parametros7("accion","",6)

strErrorMsg= ""
if (p_intDia = "") then
    diasql = "__"
else
    diasql=p_intDia
end if
if (p_intMes = "") then
   messql="__"
else
   messql=p_intMes
end if
if (p_intAnio = "") then
   aniosql="____"
else
   aniosql=p_intAnio
end if
g_intFechaConc = aniosql & messql & diasql


'Se determina si es corredor o vendedor
if (GF_ES_CORREDOR(Session("KCOrganizacion"))) then
    strTitulo = "Vendedor"
else
	strTitulo = "Corredor"
end if

%>
<html>
<head>
  <title>TOEPFER INTERNATIONAL - <%GF_Traducir("Contratos")%></title>
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
  <script language="javascript" src="scripts/script_fechas.js"></script>
   <script language="javascript" src="scripts/scripts_ordenar.js"></script>
   <script language="javascript" src="scripts/script_checkboxes.js"></script>
   <script language="javascript">
          function fcnCall(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha,p_unitDest,p_KCCOR)
          {
                   var params = "cmbProducto=" + p_intProducto + "&txtSucursal=" +  p_intSucursal;
                   params= params + "&txtOperacion=" + p_intOperacion + "&txtNumero=" + p_intNumero;
                   params= params + "&txtCosecha=" + p_intCosecha + "&UnidadDestino=" + p_unitDest;
                   params= params + "&txtCorredor=" + p_KCCOR + "&strTipo=<% =strTipo %>";
                   window.open("cor-DetalleCto.asp?" + params);
          }   
         function fcnCall2()
         {
            <%if (strTipo = "CONTRATO") then%>
            var respuesta

            respuesta = window.showModalDialog("cor-CtoGeneratorOpciones.asp","","dialogHeight:180px;dialogWidth:305px;status:no;scroll:no")
            //alert(respuesta);
            if (respuesta != 'nada')
            {
                if (respuesta == 'completos')
                    form1.accion.value = 'todos'
                else if (respuesta == 'descargas')
                    form1.accion.value = 'descargas'
                else
                    form1.accion.value = 'contratos';
                form1.action = 'cor-CtoGenerator.asp';
                fcnCall3();
            }
            <%else%>
                form1.action = 'cor-BolGenerator.asp';
                fcnCall3();
            <%end if%>
         }

         function fcnCall3() {
            form1.target = 'new';
            form1.submit();
            form1.action = 'cor-cabeceraCTO.asp';
            form1.target = '';
         }

		function checkAll() {
			var checks = document.getElementsByTagName('INPUT');
			var estado;
        
			if (document.getElementById('CHKALL').checked) {
				estado = true;
			} else {
				estado = false;
			}
			for (var k = 0; k < checks.length; k++) {
				checks[k].checked = estado;
			}
		}
    
		function uncheckAll(p_id) {
			if (!document.getElementById(p_id).checked)
				document.getElementById('CHKALL').checked = false;
		}
  </script>
</head>
<body>
<%if strTipo = "BOLETO" then
    GF_TITULO_2("Boletos de Compra/Venta Pendientes de Recepción")
else
    GF_TITULO_2("Contratos")
end if
%>
<form method="POST" name="form1" action="cor-cabeceraCto.asp">
<input type="hidden" name="strTipo" value="<%=strTipo%>">
<table width="60%" cellspacing="0" cellpadding="0" align="center" border="0">
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
           <td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="2"></td>
           <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
                     <table width="100%" align="center" border="0">
                            <tr>
                                <td align="right"><% =GF_Traducir("Fecha Conc.")%>:</td>
                                <td>
                                    <input type="text" size="2" maxLength="2" value="<% =p_intDia %>" name="txtDia" onBlur="javascript:ControlarDia(this);"> /
                                    <input type="text" size="2" maxLength="2" value="<% =p_intMes %>" name="txtMes" onBlur="javascript:ControlarMes(this);"> /
                                    <input type="text" size="4" maxLength="4" value="<% =p_intAnio%>" name="txtAnio" onBlur="javascript:ControlarAnio(this);"></td>
                                    <td rowspan=3 align="center"><input type="submit" value="<% =GF_Traducir("Buscar")%>..."></td>
                            </tr>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Producto")%>:</td>
                                <td>
                                <% strSQL="Select * from MERFL.MER112F1 order by DESCPR asc"
                                   call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
                                %>
                                <select name="cmbProducto">
                                        <option SELECTED value="0">- <% =GF_TRADUCIR("Todos") %> -
                                <% while (not rs.eof)
                                        if (cLng(cmbProducto) = cLng(rs("CODIPR"))) then %>
                                        <option SELECTED value="<% =cInt(rs("CODIPR")) %>"><% =GF_TRADUCIR(rs("DESCPR")) %>
                                <%      else %>
                                        <option value="<% =cInt(rs("CODIPR")) %>"><% =GF_TRADUCIR(rs("DESCPR")) %>
                                <%      end if %>
                                <%      rs.MoveNext
                                   wend
                                %>
                                </select>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Contrato")%>:</td>
                                <td>
                                    <input type="text" size="2" maxLength="2" value="<% =intProducto %>" name="txtProducto"> -
                                    <input type="text" size="1" maxLength="1" value="<% =g_intSucursal %>" name="txtSucursal"> -
                                    <input type="text" size="2" maxLength="2" value="<% =g_intOperacion %>" name="txtOperacion"> -
                                    <input type="text" size="5" maxLength="5" value="<% =g_intNumero %>" name="txtNumero"> /
                                    <input type="text" size="2" maxLength="2" value="<% =g_intCosecha %>" name="txtCosecha">
                                </td>
                            </tr>
                            <%if Session("KCOrganizacion") = 99999997 then%>
                            <tr>
                                <td align="right" width="20%"><% =GF_Traducir("Cod. Corredor")%>:</td>
                                <td>
                                    <table border=0 cellpadding=0 cellspacing=0 width="100%">
                                        <tr>
                                            <td><input type="text" size="6" maxLength="6" value="<% =g_intKCCOR %>" name="txtCorredor"></td>
                                            <td align="right"><% =GF_Traducir("Cod. Vendedor")%>:</td>
                                            <td><input type="text" size="6" maxLength="6" value="<% =g_intKCVEN %>" name="txtVendedor"></td>
                                        </tr>
                                    </table>
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
<div id="divAdobe" style="visibility:hidden;position:absolute;"><img src="images/get_adobe_reader.gif" onClick="javascript:window.open('http://www.adobe.com/products/acrobat/readstep2.html');" style="cursor:hand;"></div>
<br>
    <%if (strErrorMsg = "") then
      'Se asume modo CONTRATO
      'g_chrMrcConfirma = "V"
      g_chrMrcConfirma = ""
      g_chrMrcRecibido = ""
      'JAS - 20/01/2012 - Por pedido de Pablo Schmidt se deja de solicitar que los boletos esten recibidos para poder ver las aplicaciones realizadas a los contratos.
      'g_chrMrcRecibido = "V"
      'Modo en que un usuario comun ve los boletos
      'if (strTipo = "BOLETO") then g_chrMrcRecibido = "F"      
      
      if (initHeader()) then      
         strLinkPagina = "cor-CabeceraCto.asp" & strParam & "&campoOrden=" & strCampoOrden
         call GF_PAGINAR("N",strLinkPagina,intMostrar,50,g_rsContratos)%>
       <table class="reg_Header" border="0" cellspacing="1" cellpadding="2" width="100%">
       <tr>
           <td colspan="3">
               <IMG ALIGN="absMiddle" SRC="Images/print-16x16.png" ALT="<% =GF_TRADUCIR("Imprimir contratos seleccionados")%>" style="cursor:hand;" onClick="javascript:fcnCall2();"> <input type="button" value="<% =GF_TRaducir("Imprimir")%>" style="cursor:hand;" onClick="javascript:fcnCall2();">
           </td>
           <td colspan="4" align="right"><% =GF_Traducir("Mostrar en")%>:
               <select name="UnidadDestino" id="UnidadDestino" onChange="javascript:form1.accion.value='unidad';form1.submit();">
					<option value="1"<%if unitDest = 1 then response.write " selected"%>><% =GF_Traducir("Kilogramos")%></option>
                    <option value="2"<%if unitDest = 2 then response.write " selected"%>><% =GF_Traducir("Toneladas")%></option>
               </select>
           </td>
       </tr>
       <tr class="reg_Header_nav" align="center">
           <td align="center" width="5%">
                <input type="checkbox" class="NOBORDER" id="CHKALL" name="CHKALL" onClick="javascript:checkAll();"
                <%if accion="unidad" and GF_Parametros7("CHKALL","",6)="on" then response.write " checked"%>
                >
           </td>
           <td width="12%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','FechaAsc');">
               <% =GF_TRADUCIR("Fecha") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','FechaDesc');">
           </td>
           <td width="20%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','ContratoAsc');">
               <% =GF_TRADUCIR("Contrato") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','ContratoDesc');">
           </td>
           <td width="22%">
               <% =GF_TRADUCIR(strTitulo) %>
           </td>
           <td width="13%">
               <% =GF_TRADUCIR("Producto") %>
           </td>
           <td width="14%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','KilosAsc');">
               <% =GF_TRADUCIR("Contrat.") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','KilosDesc');">
           </td>
           <td width="14%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','EntAsc');">
               <% =GF_TRADUCIR("Entreg.") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-CabeceraCto.asp','<% =strParam%>','EntDesc');">
           </td>
       </tr>
<%     intIndex = 0
       while ((getNextHeader()) and (CInt(intIndex) < CInt(intMostrar))) %>
       <tr class="reg_Header_navdos">
           <td align="center" width="5%">
               <input type="checkbox" class="NOBORDER" id="CHK<%=intIndex%>" name="CHK<%=intIndex%>"
               <%if accion="unidad" and GF_Parametros7("CHK" & intIndex,"",6)="on" then response.write " checked"%>
                onClick="uncheckAll('CHK<%=intIndex%>');">
           </td>
           <td align="center">
            <%fecha = GF_FN2DTE(g_intFechaConc)
            response.write left(fecha,6) & mid(fecha,9,2)%>
           </td>
           <td align="center">
                <a href="javascript:fcnCall(<% =g_intProducto %>,<% =g_intSucursal %>,<% =g_intOperacion %>,<% =g_intNumero %>,<% =g_intCosecha %>,<% =unitDest %>,<% =g_intKCCOR%>)" title="<%=GF_Traducir("Ver Contrato")%>"><% =GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha)%></a>
           </td>
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Prod" NAME="P_CTO<%=intIndex %>Prod" VALUE="<%=g_intProducto%>">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Suc" NAME="P_CTO<%=intIndex %>Suc" VALUE="<%=g_intSucursal%>">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Oper" NAME="P_CTO<%=intIndex %>Oper" VALUE="<%=g_intOperacion%>">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Sec" NAME="P_CTO<%=intIndex %>Sec" VALUE="<%=g_intNumero%>">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Cos" NAME="P_CTO<%=intIndex %>Cos" VALUE="<%=g_intCosecha%>">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>CtoVen" NAME="P_CTO<%=intIndex %>ContVen" VALUE="">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>CPProcedencia" NAME="P_CTO<%=intIndex %>CPProcedenciaMerc" VALUE="">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>CAProcedencia" NAME="P_CTO<%=intIndex %>CAProcedenciaMerc" VALUE="">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>Clausula" NAME="P_CTO<%=intIndex %>Clausula" VALUE="">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>MercProdPropia" NAME="P_CTO<%=intIndex %>MercProdPropia" VALUE="">
           <INPUT TYPE="HIDDEN" ID="P_CTO<%=intIndex %>MercEnConsig" NAME="P_CTO<%=intIndex %>MercEnConsig" VALUE="">
<%if (GF_ES_CORREDOR(Session("KCOrganizacion"))) then
                vendedor = GetDSEnterprise2(cdbl(g_intKCVEN))
           else
                vendedor = GetDSEnterprise2(cdbl(g_intKCCOR))
           end if%>
           <TD style="cursor: default">
                <span title="<%=vendedor%>">
                <%if len(vendedor)>14  then
                    response.write left(vendedor,14) & "..."
                else
                    response.write vendedor
                end if%>
                </span>
           </TD>
           <td align="center" style="cursor: default">
                <% strSQL="Select DESCPR from MERFL.MER112F1 where CODIPR=" & g_intProducto
					Call GF_BD_AS400_2(rsProd,oConn,"OPEN",strSQL)
					strDs = GF_TRADUCIR(trim(rsProd("DESCPR")))%>
                <span title="<%=strDs%>" align="center">
                <%if len(strDs)>5 then
                    response.write left(strDs,5) & "..."
                else
                    response.write strDs
                end if%>
                </span>
           </td>
           <% Call setUnidadPeso(unitDest, g_intKilosNetos, retvalue) %>
           <td align="right"><% =retValue %></td>
           <% Call setUnidadPeso(unitDest, g_intKgEntregados, retvalue) %>
           <td align="right"><% =retvalue %></td>
       </tr>
<%          intIndex = intIndex + 1
       wend
%>
       </table>
       <INPUT TYPE="HIDDEN" ID="P_MAXCTO" NAME="P_MAXCTO" VALUE="<% =intIndex %>">
<%else 'intHeader%>
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
<% end if%>
<INPUT TYPE="HIDDEN" NAME="CampoOrden" id="CampoOrden" VALUE="<% =strCampoOrden %>">
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
<script language="javascript">check_qnt=<% =intIndex%>;</script>
<%
end if
%>

</BODY>
<script language="javascript">
var divAdobe = document.getElementById('divAdobe');
divAdobe.style.top = '160px';
divAdobe.style.left = window.screen.width - 293 + 'px';
divAdobe.style.visibility = 'visible';
</script>
</HTML>
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
'******************************************************************************************
%>
