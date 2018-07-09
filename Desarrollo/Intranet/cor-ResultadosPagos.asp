<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientospaginacion.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP"-->
<!--#include file="Includes/procedimientosAS400.asp"-->

<%
'***************************************************************************************
Function getTipoRetencion(pcdTipo)
	Dim rtrn, strSQL, conn, rs
	strSQL="Select DSCONCEPTO from TBLCONCEPTOPAGO where CDCONCEPTO='" & pcdTipo & "'"
	Call GF_BD_AS400(rs, conn, "OPEN", strSQL)
	if (not rs.eof) then
		rtrn = rs("DSCONCEPTO")
	end if
	getTipoRetencion = rtrn
End Function
'***************************************************************************************
dim fechaControl

'Recupero los parametros.
strTipo=GF_Parametros7("strTipo","",6)
Call addParam("strTipo", strTipo, strParam)
p_KCPRO = GF_Parametros7("KCPRO",0,6)
Call addParam("KCPRO", p_KCPRO, strParam)
p_Minuta = GF_Parametros7("minuta",0,6)
Call addParam("minuta", p_Minuta, strParam)
p_intDia= trim(GF_Parametros7("dia","",6))
Call addParam("dia",p_intDia,strParam)
p_intMes= trim(GF_Parametros7("mes","",6))
Call addParam("mes",p_intMes,strParam)
p_intAnio= trim(GF_Parametros7("anio","",6))
Call addParam("anio",p_intAnio,strParam)
p_TipoCbte = GF_Parametros7("tipoCbte","",6)
Call addParam("tipoCbte", p_TipoCbte, strParam)
p_nroRet = GF_Parametros7("nroRet",0,6)
Call addParam("nroRet", p_nroRet, strParam)
p_tipoRet = GF_Parametros7("tipoRet","",6)
Call addParam("tipoRet", p_tipoRet, strParam)
campoOrden = GF_Parametros7("campoOrden","",6)
select case campoOrden
    case "FechaAsc": strCampoOrden = "WDFPAG asc"
    case "NumeroAsc": strCampoOrden = "WDNRET asc"
    case "NumeroDesc": strCampoOrden = "WDNRET desc"
    case "MinutaAsc": strCampoOrden = "WDNING asc"
    case "MinutaDesc": strCampoOrden = "WDNING desc"
    case "tipoRetAsc": strCampoOrden = "WDCODE asc"
    case "tipoRetDesc": strCampoOrden = "WDCODE desc"
    case else: strCampoOrden = "WDFPAG desc"
end select

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
fechaSQL = aniosql & messql & diasql

fechaControl = GF_DTEADD(mid(session("MmtoSistema"),1,8),-1,"a")
fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,5,2)-1)*-1,"M")
fechaControl = GF_DTEADD(fechaControl,(mid(fechaControl,7,2)-1)*-1,"D")
%>
<html>
<head>
  <title>Resultados Pagos</title>
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
  <script language=javascript>

    function ordenar_onClick(p_campoOrden) {
       document.location.href = 'cor-ResultadosPagos.asp<%=strParam%>&campoOrden=' + p_campoOrden;
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
    
    function fcnPrint() {
        document.getElementById('frmRetenciones').action = "cor-RetGenerator.asp";
        document.getElementById('frmRetenciones').submit();
    }
  </script>
</head>

<body>
    <%
    strSQL="Select WDFPAG as FechaPago, WDTCBT as TipoCbte, WDNING as Minuta, WDCODE as KCDetalle, WDFOPG as KCPago, WDNRET as RetNro, WDPOPG as KCPRO, WDCODE as TipoRet "
	strSQL= strSQL & "from TESFL.TES960F2 where WDFPAG >='" & fechaControl & "' and WDFPAG like '" & fechaSQL & "' and WDCODE<>'W' "
	if (p_KCPRO > 0) then strSQL = strSQL & " and WDPOPG=" & p_KCPRO
    if (p_Minuta > 0) then strSQL = strSQL & " and WDNING=" & p_Minuta
    if (p_TipoCbte <> "") then strSQL = strSQL & " and WDTCBT='" & p_TipoCbte & "'"
    if (p_nroRet > 0) then strSQL = strSQL & " and WDNRET=" & p_nroRet
    if (p_tipoRet <> "0") then strSQL = strSQL & " and WDCODE='" & p_tipoRet & "'"
    strSQL = strSQL & " order by " & strCampoOrden
    call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
    intIndex = 0
    if (not rs.eof) then
       strLinkPagina = "cor-ResultadosPagos.asp" & strParam & "&campoOrden=" & campoOrden
       intMostrar = 10
       call GF_PAGINAR("N",strLinkPagina,intMostrar,50,rs)%>
       <table class="reg_Header" border="0" cellspacing="1" cellpadding="2" align="center" width=100%>
         <tr>
           <td colspan="6" align="left" width="5%">
               <IMG ALIGN="absMiddle" SRC="Images/mail.gif" ALT="<% =GF_TRADUCIR("Enviar retenciones seleccionadas")%>" style="cursor:hand;" onClick="document.getElementById('frmRetenciones').submit();"> <input type="button" value="<% =GF_TRaducir("Enviar")%>" style="cursor:hand;" onClick="document.getElementById('frmRetenciones').submit();">
               <IMG ALIGN="absMiddle" SRC="Images/Printer.gif" ALT="<% =GF_TRADUCIR("Imprimir retenciones seleccionadas")%>" style="cursor:hand;" onClick="javascript:fcnPrint();"> <input type="button" value="<% =GF_TRaducir("Imprimir")%>" onClick="javascript:fcnPrint()">
           </td>
         </tr>
         <tr class="reg_Header_nav" align="center">
           <td align="center">
                <input type="checkbox" class="NOBORDER" id="CHKALL" name="CHKALL" onClick="checkAll();">
           </td>
           <td width="10%">
               <table width=100% cellpadding=0 cellspacing=0 border=0 class="reg_Header_nav" align=center>
                <tr>
                    <td align=center><IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('FechaAsc');"></td>
                    <td align=center><% =GF_TRADUCIR("Fecha Pago") %></td>
                    <td align=center><IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('FechaDesc');"></td>
                </tr>
               </table>
           </td>
           <td width="10%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('NumeroAsc');">
               <% =GF_TRADUCIR("Retención") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('NumeroDesc');">
           </td>
           <td width="15%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('tipoRetAsc');">
               <% =GF_TRADUCIR("Tipo de Ret.") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('tipoRetDesc');">
           </td>
           <td width="30%">
               <% =GF_TRADUCIR("Proveedor") %>
           </td>
           <td width="20%">
               <% =GF_TRADUCIR("Tipo Cbte.") %>
           </td>
           <td width="10%">
               <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('MinutaAsc');">
               <% =GF_TRADUCIR("Minuta") %>
               <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('MinutaDesc');">
           </td>
       </tr>
<%	while ((not rs.eof) and (CInt(intIndex) < CInt(intMostrar))) %>
       <form id="frmRetenciones" action="cor-RetMailGenerator.asp" target="new" method="GET">
       <tr class="reg_Header_navdos">
           <td align="center">
                <input type="checkbox" class="NOBORDER" id="CHK<%=intIndex%>" name="CHK<%=intIndex%>" onClick="uncheckAll('CHK<%=intIndex%>');">
           </td>
           <INPUT TYPE="HIDDEN" ID="P_RET<%=intIndex %>Tipo" NAME="P_RET<%=intIndex %>Tipo" VALUE="<%=rs("KCDetalle")%>">
           <INPUT TYPE="HIDDEN" ID="P_RET<%=intIndex %>Nro" NAME="P_RET<%=intIndex %>Nro" VALUE="<%=rs("RetNro")%>">
           <td align="center">
            <%fecha = GF_FN2DTE(rs("FechaPago"))
            response.write left(fecha,6) & mid(fecha,9,2)%>
            <INPUT TYPE="HIDDEN" ID="P_RET<%=intIndex %>Fecha" NAME="P_RET<%=intIndex %>Fecha" VALUE="<%=rs("FechaPago")%>">
           </td>
           <td align="center">
                <%=GF_EDIT_CBTE(GF_nDigits(rs("RetNro"),12))%>
           </td>
           <td align="center">
                <%=getTipoRetencion(rs("TipoRet"))%>
           </td>
           <TD style="cursor: default; padding-left:10px;">
                <%kcProveedor=rs("KCPRO")
                dsProveedor = getDSEnterprise2(kcProveedor)%>
                <INPUT TYPE="HIDDEN" ID="P_RET<%=intIndex %>KCPRO" NAME="P_RET<%=intIndex %>KCPRO" VALUE="<%=kcProveedor%>">
                <span title="<%=dsProveedor%>">
                <%if len(dsProveedor)>35  then
                    response.write left(dsProveedor,30) & "..."
                else
                    response.write dsProveedor
                end if%>
                </span>
           </TD>
           <td align="center" style="cursor: default">
                <%call GF_MGC("SF", rs("TipoCbte"),"",strDS)
                strDs = GF_TRADUCIR(strDS)%>
                <span title="<%=strDs%>" align="center">
                <%if len(strDs)>20 then
                    response.write left(strDs,16) & "..."
                else
                    response.write strDs
                end if%>
                </span>
           </td>
           <td align="right"><%=rs("Minuta")%></td>
       </tr>
<%          intIndex = intIndex + 1
            rs.movenext
       wend%>
         <INPUT TYPE="HIDDEN" NAME="P_MAXRET" id="P_MAXRET" VALUE="<% =intIndex-1 %>">
       </form>
       </table>
       <INPUT TYPE="HIDDEN" ID="P_MAXCTO" NAME="P_MAXCTO" VALUE="<% =intIndex %>">
<%     else%>
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
<%     end if%>

<INPUT TYPE="HIDDEN" NAME="CampoOrden" id="CampoOrden" VALUE="<% =strCampoOrden %>">

</body>
<script language=javascript>
    if (<%=intIndex%>>0)
        parent.expandirIFrame((<%=intIndex%> * 25) + 130);
</script>
</html>
<%'******************************************************************************************
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
