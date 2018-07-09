<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/cor-IncludeCTO.asp"-->
<!--#include file="Includes/procedimientosexecute.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<%call ProcedimientoControl("DESCDET")
Dim intMostrar, auxDs, strErrorMsg, retValue, unitDest

'Levanto parametros
g_intProducto = GF_Parametros7("producto",0,6)
g_intSucursal = GF_Parametros7("sucursal",0,6)
g_intOperacion = GF_Parametros7("operacion",0,6)
g_intNumero = GF_Parametros7("numero",0,6)
g_intCosecha = GF_Parametros7("cosecha",0,6)
g_intPuerto = GF_Parametros7("puerto",0,6)
g_intPlanillaCos = GF_Parametros7("planillaCos","",6)
g_intPlanillaNro = GF_Parametros7("planillaNro","",6)
g_intFechaDescarga = GF_Parametros7("fechaDescarga",0,6)
g_intPlanillaSec = GF_Parametros7("planillaSec","",6)
g_intSolicitudNro = GF_Parametros7("solicitudNro","",6)
g_intCartaPorte = GF_Parametros7("cartaPorte","",6)

unitDest = GF_Parametros7("unitDest","",6)
if (unitDest = "") then unitDest = "1"
%>
<html>
<head>
  <title><% =GF_Traducir("TOEPFER INTERNATIONAL - Analisis de Descarga")%></title>
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
</head>

<body>
   <%strErrorMsg = ""
   GF_TITULO_2(GF_Traducir("Descarga")) %>
   <!-- INICIO DESCARGA -->
   <%if (initDescarga()=1) then
   call getNextDescarga()%>
   <TABLE WIDTH="95%" ALIGN="CENTER" CELLSPACING="0" CELLPADDING="0" BORDER="0">
          <TR>
              <TD WIDTH="8"><img src="images/marco_r1_c1.gif"></TD>
              <TD COLSPAN="3"><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
              <TD WIDTH="8"><img src="images/marco_r1_c3.gif"></TD>
          </TR>
          <TR>
              <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
              <TD COLSPAN="3">
                  <TABLE WIDTH="100%">
                     <tr>
                        <td width="23%" align="right"><b><% =GF_Traducir("Contrato")%></b></td>
                        <td width="1%"><b>:</b></td>
                        <td width="26%"><% =g_intProducto%>-<% =g_intSucursal%>-<% =g_intOperacion%>-<% =g_intNumero%>-<% =g_intCosecha%></td>
                        <td align="right" width="23%"><b><% =GF_Traducir("Fecha de Descarga")%></b></td>
                        <td width="1%"><b>:</b></td>
                        <td><% =GF_FN2DTE(g_intFechaDescarga)%></td>
                     </tr>
                     <tr>
                        <td align="right"><b><% =GF_Traducir("Contrato Corredor")%></b></td>
                        <td><b>:</b></td>
                        <td><% =g_strCtoCorredor%></td>
                        <td align="right"><b><% =GF_Traducir("Carta de Porte")%></b></td>
                        <td><b>:</b></td>
                        <td><% =g_intCartaPorte%></td>
                     </tr>
                  </TABLE
               </TD>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </TR>
            <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
            </TR>
            <tr>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <td colspan="3" align="left"><b><% =GF_Traducir("Partes Involucradas")%></b></td>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </tr>
            <tr>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <td colspan="3">
                  <br>
                  <table width="100%">
                     <tr>
                        <td WIDTH="18%" align="right"><% =GF_Traducir("Corredor")%></td>
                        <td width="1%">:</td>
                        <td><% =GetDSEnterprise2(g_intKcCor)%></td>
                     </tr>
                     <tr>
                        <td align="right"><% =GF_Traducir("Vendedor")%></td>
                        <td>:</td>
                        <td><% =GetDSEnterprise2(g_intKcVen)%></td>
                     </tr>
                  </table>
                  <br>
               </td>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </tr>
            <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
            </TR>
            <tr>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <td colspan="3" align="left"><b><% =GF_Traducir("Detalles de la Planilla")%></b></td>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </tr>
            <TR>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <TD COLSPAN="3">
                  <br>
                  <TABLE WIDTH="100%">
                     <tr>
                        <td width="30%" align="right"><% =GF_Traducir("Nro. de Recibo/Romaneo")%></td>
                        <td width="1%">:</td>
                        <td width="20%"><% =g_intReciboNro%></td>
                        <td width="25%" align="right"><% =GF_Traducir("Tipo Movimiento")%></td>
                        <td width="1%">:</td>
                        <td width="23%">
                            <%if (ucase(g_intCdeEs)="E") then
                                 response.write GF_Traducir("Salida")
                            elseif (ucase(g_intCdeEs)="I") then
                                   response.write GF_Traducir("Entrada")
                            end if%>
                        </td>
                     </tr>
                     <tr>
                        <td align="right"><% =GF_Traducir("Mercaderia Conforme")%></td>
                        <td>:</td>
                        <td>
                           <%if g_CHRMrcConforme="V" then
                              response.write GF_Traducir("Si") & "&nbsp;(Gdo:&nbsp;" & g_intAnalisisGdo & ")"
                           elseif g_CHRMrcConforme="F" then
                              response.write GF_Traducir("No")
                           end if%>
                        </td>
                        <td align="right"><% =GF_Traducir("Solicitud de Analisis")%></td>
                        <td>:</td>
                        <td><% =g_intSolicitudNro%></td>
                     </tr>
                     <tr>
                        <td align="right"><% =GF_Traducir("Cantidad Descargada")%></td>
                        <td>:</td>
                        <td colspan="4">
                            <%Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilosDescarga,"1",unitDest,"","",retValue,7,8,9,10)
                            response.write retValue%>&nbsp;<% =GF_DT1("READ","DSAB","","","MU",unitDest)%>
                        </td>

                     </tr>
                     <tr>
                        <td align="right"><% =GF_Traducir("Puerto")%></td>
                        <td>:</td>
                        <td colspan="4">
                           <%call GF_MGC("PU",g_intPuerto,0,auxDs)
                           response.write auxDs
                           'response.write g_intPuerto%>
                        </td>
                     </tr>
                  </table>
               </td>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </tr>

          <tr>
              <td><img src="images/marco_r3_c1.gif"></td>
              <td COLSPAN="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
              <td><img src="images/marco_r3_c3.gif"></td>
          </tr>
   </TABLE>
   <br>
   <br>
   <%call GF_Titulo_4(GF_Traducir("Analisis"))%>
   <!--INICIO CABECERA ANALISIS-->
   <%if (initHeaderAnalisis()) then
   call getNextAnalisis()%>
   <TABLE WIDTH="95%" ALIGN="CENTER" CELLSPACING="0" CELLPADDING="0" BORDER="0">
          <TR>
              <TD WIDTH="8"><img src="images/marco_r1_c1.gif"></TD>
              <TD COLSPAN="3"><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
              <TD WIDTH="8"><img src="images/marco_r1_c3.gif"></TD>
          </TR>
          <TR>
                 <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD COLSPAN="3">
                    <br>
                    <TABLE WIDTH="100%" border =0>
                        <TR>
                           <TD width="23%" align="right"><% =GF_Traducir("Producto")%></td>
                           <td width="1%">:</td>
                           <td width="20%" colspan="5">
                               <%call GF_MGC("AR",g_intProducto,0,auxDs)
                               response.write GF_Traducir(auxDs)%>
                           </TD>
                        </TR>
                        <TR>
                           <TD align="right"><% =GF_Traducir("Nro. de Analisis")%></td>
                           <td >:</td>
                           <td width="15%"><% =g_intNumeroAnalisis%></TD>
                           <TD width="15%" align="right"><% =GF_Traducir("Fecha de Analisis")%></td>
                           <td width="1%">:</td>
                           <td><% =GF_FN2DTE(g_intFechaAnalisis)%></TD>
                        </TR>
                        <TR>
                           <TD  align="right"><% =GF_Traducir("Cantidad Analizada")%></td>
                           <td >:</td>
                           <td >
                               <%Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilos,"1",unitDest,"","",retValue,7,8,9,10)
                               response.write retValue%>&nbsp;<% =GF_DT1("READ","DSAB","","","MU",unitDest)%>
                           </TD>
                           <TD  align="right"><% =GF_Traducir("Entidad")%></td>
                           <td >:</td>
                           <td ><%
                               Call GF_MGC("ME",g_intBolsa,0,strDS)
                               response.write strDS
                               %>
                           </TD>
                        </TR>
                        <TR>
                           <td  align="right"><% =GF_Traducir("Grado del Analisis")%></td>
                           <td >:</td>
                           <td ><% =g_intAnalisisGdo%></td>
                           <TD  align="right"><% =GF_Traducir("Costo del Analisis")%></td>
                           <td >:</td>
                           <td >$&nbsp;<% =Editar_Importe(g_intCosto)%></TD>
                        </TR>
                    </TABLE>
                 </TD>
                 <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
            </TR>
            <TR>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <TD colspan="3"><b><% =GF_Traducir("Detalles del Analisis")%></b></TD>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </TR>
            <TR>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <TD colspan="3">&nbsp;</TD>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </TR>
            <!--INICIO LOS DETALLES DEL ANALISIS-->

            <% Call initHeaderDetAnalisis %>
            <TR>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
               <TD COLSPAN="3" ALIGN="CENTER">
            <%if g_rsDetAnalisis.eof then%>
               <font size="3" color="red"><b><% =GF_Traducir("No se encontraron Detalles del Analisis para esta descarga")%></b></font>
            <%else %>
                  <TABLE class="reg_Header" border="0" cellspacing="1" cellpadding="2" width="95%">
                    <TR class="reg_Header_nav" align="center">
                        <TD><% =GF_Traducir("Concepto")%></TD>
                        <TD><% =GF_Traducir("Valor")%></TD>
                        <TD>&nbsp;<% =GF_Traducir("Rebaja")%></TD>
                        <TD>&nbsp;<% =GF_Traducir("Bonificacion")%></TD>
                     </TR>
                     <%while getNextDetAnalisis()%>
                        <TR class="reg_Header_navdos">
                           <TD ALIGN="CENTER"><% =g_intConcepto%>&nbsp;-&nbsp;<% =GF_Traducir(g_strConceptoDs)%></TD>
                           <TD ALIGN="CENTER"><% =Editar_Importe(g_intValor)%>&nbsp;%</TD>
                           <TD ALIGN="CENTER"><% =Editar_Importe(g_intRebaja)%>&nbsp;%</TD>
                           <TD ALIGN="CENTER"><% =Editar_Importe(g_intBonif)%>&nbsp;%</TD>
                        </TR>
                     <%wend%>
                  </TABLE>
            <%end if%>
               </TD>
               <TD WIDTH="8" HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
            </TR>
            <tr>
              <td><img src="images/marco_r3_c1.gif"></td>
              <td COLSPAN="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
              <td><img src="images/marco_r3_c3.gif"></td>
          </tr>
   </TABLE>
   <%else 'ERROR
          strErrorMsg = "Aún no se encuentra disponible esta información"%>
   <% end if%>
   <% else
      strErrorMsg = "No se encontraron datos de la descarga"
   end if %>
   <% if strErrorMsg<>"" then%>
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
   <% end if%>
</body>
</html>
