<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/cor-IncludeCTO.asp"-->
<!--#include file="Includes/procedimientosexecute.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#Include File="Includes/ExternalFunctions.ASP" -->
<!-- #Include File="Includes/procedimientospaginacion.ASP" -->

<%call ProcedimientoControl("CONTDET")

Function getListoFijaciones(intProducto,intSucursal,intOperacion,intNumero,intCosecha)
	Dim strSQL, conn, rs
	
	strSQL="Select * from MERFL.MER311F7 where CPROF7=" & intProducto & " and CSUCF7=" & intSucursal & " and COPEF7=" & intOperacion & " and NCTOF7=" & intNumero & " and ACOSF7=" & intCosecha
	Call GF_BD_AS400_2(rs, conn,"OPEN",strSQL)
	
	Set getListoFijaciones = rs
	
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim unitDest, intMostrar, cont, auxDs, strTexto
DIM unitDestino, retValue, strLink, strParam, strCampoOrden

g_intProducto= GF_Parametros7("cmbProducto","",6)
Call addParam("cmbProducto",g_intProducto,strParam)
g_intSucursal= GF_Parametros7("txtSucursal","",6)
Call addParam("txtSucursal",g_intSucursal,strParam)
g_intOperacion= GF_Parametros7("txtOperacion","",6)
Call addParam("txtOperacion",g_intOperacion,strParam)
g_intNumero= GF_Parametros7("txtNumero","",6)
Call addParam("txtNumero",g_intNumero,strParam)
g_intCosecha= GF_Parametros7("txtCosecha","",6)
Call addParam("txtCosecha",g_intCosecha,strParam)
g_intKCCOR= GF_Parametros7("txtCorredor","",6)
call addParam("txtKCCOR",g_intKCCOR,strParam)
unitDest = GF_Parametros7("UnidadDestino","",6)
if unitDest = "" then unitDest= "1"
Call addParam("UnidadDestino",unitDest,strParam)

Call initHeader()
Call getNextHeader()
%>
<HTML>
<HEAD>
      <TITLE>TOEPFER INTERNATIONAL - Contratos</TITLE>
      <LINK HREF="CSS/ActisaIntra-1.css" REL="stylesheet" TYPE="text/css">
      <script language="javascript" src="scripts/scripts_ordenar.js"></script>
      <script language="javascript">
        function fcnExpand(P_TBL,P_IMG)
        	{
        		if (P_TBL.style.visibility == 'visible')
        		{
        			P_TBL.style.visibility= 'hidden';
        			P_TBL.style.position= 'absolute';
        			P_IMG.src='images/Tplusik.gif';
        		}
        		else
        		{
        			P_TBL.style.visibility= 'visible';
        			P_TBL.style.position= 'relative';
        			P_IMG.src='images/TMinus.gif';
        		}
        	}
        	
        function fcnCall3(p_intProducto, p_intSucursal, p_intOperacion, p_intNumero, p_intCosecha, p_unitDest, p_puerto, p_fechaDescarga, p_planillaSec,p_planillaCos,p_planillaNro, p_solicitudNro, p_cartaPorte)
          {
                   var params = "Producto=" + p_intProducto + "&Sucursal=" +  p_intSucursal;
                   params= params + "&Operacion=" + p_intOperacion + "&Numero=" + p_intNumero;
                   params= params + "&Cosecha=" + p_intCosecha + "&UnidadDestino=" + p_unitDest;
                   params= params + "&Puerto=" + p_puerto + "&FechaDescarga=" + p_fechaDescarga;
                   params= params + "&planillaCos=" + p_planillaCos + "&planillaNro=" + p_planillaNro;
                   params= params + "&solicitudNro=" + p_solicitudNro + '&cartaPorte=' + p_cartaPorte;
                   if (p_planillaSec != '')
                        params= params + "&planillasec=" + p_planillaSec;
                   window.open("cor-infoDescarga.asp?" + params);
          }
      </script>
</HEAD>
<BODY>
<FORM NAME="form1" ID="form1" METHOD="POST" ACTION="cor-detalleCto.asp">
      <INPUT TYPE="HIDDEN" NAME="cmbProducto" ID="cmbProducto" VALUE="<% =g_intProducto%>">
      <INPUT TYPE="HIDDEN" NAME="txtSucursal" ID="txtSucursal" VALUE="<% =g_intSucursal%>">
      <INPUT TYPE="HIDDEN" NAME="txtOperacion" ID="txtOperacion" VALUE="<% =g_intOperacion%>">
      <INPUT TYPE="HIDDEN" NAME="txtNumero" ID="txtNumero" VALUE="<% =g_intNumero%>">
      <INPUT TYPE="HIDDEN" NAME="txtCosecha" ID="txtCosecha" VALUE="<% =g_intCosecha%>">
      <% GF_TITULO_2("Detalle de Contrato") %>
      <!-- INICIO CONTRATO -->
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
                            <% if ((g_chrMrcConfirma = "F") and (g_chrTipoCto <> "C")) then %>
                            <TR>
                                <TD COLSPAN="2" class="TDERROR"><% =GF_TRADUCIR("CONTRATO SIN CONFIRMAR")%></TD>
                            </TR>
                            <% end if %>
                            <TR>
                                <TD WIDTH="50%"><B><% =GF_TRADUCIR("Contrato") %>:</B> <% =GF_EDIT_CONTRATO(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha) %></TD>
                                <TD><B><% =GF_TRADUCIR("Fecha de Concertacion") %>:</B> <% =GF_FN2DTE(g_intFechaConc) %></TD>
                            </TR>
                            <TR>
                                <TD WIDTH="50%"><B><% =GF_TRADUCIR("Cto Corredor") %>:</B> <% =g_strCtoCorredor %></TD>
                                <TD><B><% =GF_TRADUCIR("Operacion") %>:</B>
                                <%
                                Call GF_MGC("MO",GF_nDigits(g_intOperacion,2),"",strDS)
                                response.write GF_Traducir(strDS)
                                %>
                                </TD>
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
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD COLSPAN="3">
                     <TABLE WIDTH="100%">
                           <TR>
                                <TD COLSPAN="3"><B><% =GF_TRADUCIR("Partes Involucradas") %></B></TD>
                            </TR>
                            <TR>
                                <TD WIDTH="10%"></TD>
                                <TD WIDTH="10%" ALIGN="RIGHT"><% =GF_TRADUCIR("Corredor") %>:</TD>
                                <TD><% =GetDSEnterprise2(g_intKCCOR) %></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Vendedor") %>:</TD>
                                <TD><% =GetDSEnterprise2(g_intKCVEN) %></TD>
                            </TR>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_t.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
             </TR>
             <TR>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD WIDTH="50%" VALIGN="TOP">
                     <TABLE WIDTH="100%">
                            <TR>
                                <TD COLSPAN="4"><B><% =GF_TRADUCIR("Mercaderia")%></B></TD>
                            </TR>
                            <TR>
                                <TD WIDTH="5%"></TD>
                                <TD ALIGN="RIGHT" WIDTH="35%"><% =GF_TRADUCIR("Producto")%>:</TD>
                                <TD colspan=2><% call GF_MGC("AR",g_intProducto,"",strDS)
                                       response.write GF_TRADUCIR(strDS) %></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Cosecha")%>:</TD>
                                <TD>
                                    <%if cint(g_intCosecha) > 95 then
                                        intCosecha = cint(g_intCosecha) + 1900
                                    else
                                        intCosecha = cint(g_intCosecha) + 2000
                                    end if
                                    response.write intCosecha%>
                                </TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Contratada")%>:</TD>
                                <%
                                  Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilos,"1",unitDest,"","",retValue,7,8,9,10)
                                %>
                                <TD>
                                    <%response.write retValue & "&nbsp;" & GF_DT1("READ","DSAB","","","MU",unitDest)%>
                                </TD>
                                <TD width="35%"></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Entregada")%>:</TD>
                                <%
                                  Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKgEntregados,"1",unitDest,"","",retValue,7,8,9,10)
                                %>
                                <TD>
                                    <%response.write retValue & "&nbsp;" & GF_DT1("READ","DSAB","","","MU",unitDest)%>
                                </TD>
                                <TD width="35%"></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Procedencia")%>:</TD>
                                <TD COLSPAN=2><% =g_strProcedencia %></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD></TD>
                                <TD></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <%
                                  if (cDbl(g_intAnulaciones) <> 0) and (cDbl(g_intAnulaciones) <> "") then
                                    intAnulaciones = cDbl(g_intAnulaciones)
                                    strTexTo = "Ampliaciones"
                                    if (intAnulaciones < 0) then
                                        strTexto = "Anulaciones"
                                        intAnulaciones = intAnulaciones * -1
                                    end if
                                %>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR(strTexto)%>:</TD>
                                <%
                                  Call GP_SERVEREXECUTE("GF_convertUnit.asp",intAnulaciones,"1",unitDest,"","",retValue,7,8,9,10)
                                %>
                                <TD><%response.write retValue & "&nbsp;" & GF_DT1("READ","DSAB","","","MU",unitDest)%></TD>
                                <% else %>
                                <TD></TD>
                                <TD></TD>
                                <% end if %>
                                <TD></TD>
                            </TR>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_c_v.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD VALIGN="TOP">
                     <TABLE WIDTH="100%">
                            <TR >
                                <TD COLSPAN="4"><B><% =GF_TRADUCIR("Fijacion") %></B></TD>
                            </TR>
                            <%if len(g_intFechaFijaDesde)=8 then%>
                               <TR>
                                   <TD WIDTH="5%"></TD>
                                   <TD ALIGN="RIGHT" WIDTH="30%"><% =GF_TRADUCIR("Desde")%>:</TD>
                                   <TD COLSPAN=2><% =GF_FN2DTE(g_intFechaFijaDesde) %></TD>
                               </TR>
                               <TR>
                                   <TD></TD>
                                   <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Hasta")%>:</TD>
                                   <TD COLSPAN=2><% =GF_FN2DTE(g_intFechaFijaHasta) %></TD>
                               </TR>
                               <TR>
                                   <TD></TD>
                                   <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Cant (Min)")%>:</TD>
                                   <%
                                     Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilosMin,"1",unitDest,"","",retValue,7,8,9,10)
                                   %>
                                   <TD align="right"><%response.write retValue & "&nbsp;" & GF_DT1("READ","DSAB","","","MU",unitDest)%></TD>
                                   <TD WIDTH="35%"></TD>
                               </TR>
                               <TR>
                                   <TD></TD>
                                   <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Cant (Max)")%>:</TD>
                                   <%
                                     Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilosMax,"1",unitDest,"","",retValue,7,8,9,10)
                                   %>
                                   <TD align="right"><%response.write retValue & "&nbsp;" & GF_DT1("READ","DSAB","","","MU",unitDest)%></TD>
                                   <TD></TD>
                               </TR>
                            <%else%>
                                 <tr valign="center" height="90%">
                                    <td align="center">
                                       <% =GF_Traducir("A este contrato no se le aplica fijacion")%>
                                    </td>
                                 </tr>
                            <%end if%>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_t_l.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_plus.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_r.gif"></TD>
             </TR>
             <TR>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD>
                     <TABLE WIDTH="100%">
                           <TR>
                                <TD COLSPAN="5"><B><% =GF_TRADUCIR("Entrega") %></B></TD>
                            </TR>
                            <TR>
                                <TD WIDTH="5%"></TD>
                                <TD WIDTH="35%" ALIGN="RIGHT"><% =GF_TRADUCIR("Desde") %>:</TD>
                                <TD><% =GF_FN2DTE(g_intFechaEntDesde) %></TD>
                            </TR>
                            <TR>
                                <TD></TD>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Hasta") %>:</TD>
                                <TD><% =GF_FN2DTE(g_intFechaEntHasta) %></TD>
                            </TR>
                            <%if cInt(g_intPuertoRecepcion)>0 then%>
                               <TR>
                                   <TD></TD>
                                   <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Puerto Entrega") %>:</TD>
                                   <TD>
                                       <%call GF_MGC("PU",g_intPuertoRecepcion,0,auxDs)
                                       response.write auxDs%>
                                   </TD>
                               </TR>
                            <%end if%>
                            <%if cInt(g_intPuertoDevolucion)>0 then%>
                               <TR>
                                   <TD></TD>
                                   <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Puerto Devol.") %>:</TD>
                                   <TD>
                                       <%call GF_MGC("PU",g_intPuertoDevolucion,0,auxDs)
                                       response.write auxDs%>
                                   </TD>
                               </TR>
                            <%end if%>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_c_v.gif" WIDTH="8" HEIGHT="100%"></TD>
                 <TD VALIGN="TOP">
                     <TABLE WIDTH="100%">
                           <TR>
                                <TD COLSPAN="4"><B><% =GF_TRADUCIR("Pago") %></B></TD>
                            </TR>
                            <TR>

                                    <%select case cInt(g_intOperacion)
                                        case 0,1,2,3,5:
                                            simboloMoneda = "$"
                                            precioTonelada = cDbl(g_intPrecioP)
                                        case 6,9,10,11,12:
                                            simboloMoneda = "U$S"
                                            precioTonelada = cDbl(g_intPrecioD)
                                    end select%>
                                <TD WIDTH="40%" ALIGN="RIGHT"><% =GF_TRADUCIR("Precio") %>:</TD>
                                <TD align="right" width="20%"><%=simboloMoneda%>&nbsp;<%=Editar_Importe(precioTonelada)%></TD>
                                <TD></TD>
                            </TR>
                            <TR>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Parcial") %>:</TD>
                                <TD align="right"><% =Editar_Importe(g_intPjeParcial) %></TD>
                                <TD>%</TD>
                            </TR>
                            <TR>
                                <TD ALIGN="RIGHT"><% =GF_TRADUCIR("Forma de Pago") %>:</TD>
                                <TD COLSPAN=2>
                                    <%if (g_strCodigoPago <> "X") then
                                        response.write GF_Traducir(g_strKCPago)
                                    else
                                        response.write replace(GF_Traducir(g_strKCPago),"X", g_intCamionesPactados)
                                    end if%>
                                </TD>
                            </TR>
                     </TABLE>
                 </TD>
                 <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif"  WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_r3_c1.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_t_b.gif"></TD>
                 <TD><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_r3_c3.gif"></TD>
             </TR>
      </TABLE>
      <!-- FIN CONTRATO -->
	  <!-- INICIO FIJACIONES -->
	  <BR>
<% 		Call GF_TITULO_4("Fijaciones") 
		if (CLng(g_intOperacion) = 1) or (CLng(g_intOperacion) = 3) or (CLng(g_intOperacion) = 10) or (CLng(g_intOperacion) = 12)then		
			Set rsFijaciones = getListoFijaciones(g_intProducto,g_intSucursal,g_intOperacion,g_intNumero,g_intCosecha)
%>
		<TABLE class="reg_Header" align="center" border="0" cellspacing="1" cellpadding="2" width="80%">
			<TR class="reg_header_nav">
				<TD align="center">Nro.</TD>
				<TD align="center">Fecha de Fijación</TD>				
				<TD align="center">Kilos Fijados</TD>
				<TD align="center">Precio</TD>
				<TD align="center">Puerto</TD>
			</TR>
<%			
			if (not rsFijaciones.eof) then
				while (not rsFijaciones.eof)
					call GF_MGC("PU",rsFijaciones("PTOFF7"),0,auxDs)
%>
					<TR class="reg_header_navdos">
						<TD align="center"><% =rsFijaciones("NROFF7") %></TD>
						<TD align="center"><% =GF_FN2DTE(rsFijaciones("FEFIF7")) %></TD>				
						<TD align="center"><% =rsFijaciones("KGFIF7") %> Kg</TD>
						<TD align="right"><% =simboloMoneda & " " & GF_EDIT_DECIMALS(CDbl(rsFijaciones("PREFF7"))*100, 2) %></TD>
						<TD align="center"><% =auxDs %></TD>
					</TR>
<%				
					rsFijaciones.MoveNext()
				wend
			else
%>
				<TR><TD class="TDNOHAY" colspan="5">Aún no se registran fijaciones.</TD></TR>
<%
			end if
%>
		</TABLE>
<%		end if			%>
	  <!-- FIN FIJACIONES -->
      <BR>
      <% GF_TITULO_4("Descargas") %>
      <!-- INICIO DESCARGAS -->
      <%'Lo leo aca para que no interfiera con las cabeceras que lei arriba
      strCampoOrden = GF_Parametros7("campoOrden","",6)
      select case strCampoOrden
            case "FechaDesc": g_strCampoOrden = "D.FECDR6 desc"
            case "CPAsc": g_strCampoOrden = "D.CPORR6 asc"
            case "CPDesc": g_strCampoOrden = "D.CPORR6 desc"
            case "KilosAsc": g_strCampoOrden = "KgDescarga asc"
            case "KilosDesc": g_strCampoOrden = "KgDescarga desc"
            case else: g_strCampoOrden = "D.FECDR6 asc"
      end select
      %>
      <TABLE WIDTH="95%" ALIGN="CENTER" CELLSPACING="0" CELLPADDING="0">
             <TR>
                 <TD WIDTH="8"><img src="images/marco_r1_c1.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_r1_c3.gif"></TD>
             </TR>
             <TR>
                  <TD HEIGHT="100%"><img src="images/marco_r2_c1.gif" WIDTH="8" HEIGHT="100%"></TD>
                  <TD ALIGN="CENTER" COLSPAN="3">
                     <%g_flgAgrupar = true
                     if initDescarga() = 0 then%>
                        <br>
                        <TABLE class="reg_Header" border="0" cellspacing="1" cellpadding="2" width="95%">
                           <TR>
                              <TD ALIGN="CENTER"><FONT COLOR="RED"><b><% =GF_Traducir("Este contrato no presenta descargas hasta el momento")%></b></FONT></TD>
                           </TR>
                     <%else
                        strLink = "cor-DetalleCto.asp" & strParam & "&CampoOrden=" & strCampoOrden
                        CALL GF_Paginar("N",strLink,intMostrar,50,g_rsDescargas)%>
                        <TABLE class="reg_Header" border="0" cellspacing="1" cellpadding="2" width="100%">
                           <TR>
                               <TD COLSPAN="6" ALIGN="RIGHT"><% =GF_Traducir("Mostrar en")%>:</TD>
                               <TD>
                                   <SELECT NAME="UnidadDestino" ID="UnidadDestino" onChange="javascript:form1.submit();">
										<option value="1"<%if unitDest = 1 then response.write " selected"%>><% =GF_Traducir("Kilogramos")%></option>
										<option value="2"<%if unitDest = 2 then response.write " selected"%>><% =GF_Traducir("Toneladas")%></option>
                                   </SELECT>
                               </TD>
                           </TR>
                           <TR class="reg_Header_nav" align="center">
                              <td width="2%">
                                    &nbsp;
                              </td>
                              <TD ALIGN="CENTER" width="10%">
                                  <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','FechaAsc');">
                                  <% =GF_TRADUCIR("Fecha") %>
                                  <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','FechaDesc');">
                              </TD>
                              <TD ALIGN="CENTER"><% =GF_Traducir("Puerto")%></TD>
                              <TD ALIGN="CENTER" width="15%">
                                  <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','CPAsc');">
                                  <% =GF_TRADUCIR("C. de Porte") %>
                                  <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','CPDesc');">
                              </TD>
                              <TD ALIGN="CENTER" width="13%">
                                  <% =GF_TRADUCIR("Analisis Gdo.") %>
                              </TD>
                              <TD ALIGN="CENTER" width="10%">
                                  <% =GF_TRADUCIR("Proteinas") %>
                              </TD>
                              <%unitDestino = 1%>
                              <TD ALIGN="CENTER" width="12%">
                                  <IMG src="images/arrow_up.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','KilosAsc');">
                                  <% =GF_TRADUCIR("Cant.") %>
                                  <IMG src="images/arrow_down.gif" align=absMiddle style="cursor=hand;" border=0 onClick="ordenar_onClick('cor-DetalleCto.asp','<% =strParam %>','KilosDesc');">
                              </TD>
                           </TR>
                         <%cont = 0
                         while (getNextDescarga() = 1) and (cint(cont) <= cint(intMostrar))
                            'Leo los datos del analisis
                            strSql = "Select NSANR6 as solicitudnro, sum(KGNER6) as kgDescarga, SECUR6 as planillaSec from MERFL.MER311F6 where CPROR6=" & g_intProducto & " and CSUCR6=" & g_intSucursal
                            strSQL = strSQL & " and COPER6=" & g_intOperacion & " and NCTOR6=" & g_intNumero
                            strSQL = strSQL & " and ACOSR6=" & g_intCosecha & " and CDESR6=" & g_intPuerto & " and ACOPR6="
                            strSQL = strSQL & g_intPlanillaCos & " and PLANR6=" & g_intPlanillaNro & " and CPORR6=" & g_intCartaPorte & " and FECDR6="
                            strSQL = strSQL & g_intFechaDescarga & " group by CPROR6, CSUCR6, COPER6, NCTOR6, ACOSR6, CDESR6, ACOPR6, PLANR6, FECDR6, NSANR6, SECUR6 order by NSANR6 asc"
                            'response.write strSql & "<br>"
                            call GF_BD_AS400_2(rs,conn,"OPEN", strSQL)%>
                           <TR class="reg_Header_navdos">
                              <td align="center">
                                    <%if (g_intCantidad = 1) or (rs.recordCount = 1) then%>
                                        <img src="images/ver.gif" style="cursor:hand;" onClick="javascript:fcnCall3(<% =g_intProducto %>,<% =g_intSucursal %>,<% =g_intOperacion %>,<% =g_intNumero %>,<% =g_intCosecha %>,<% =unitDest%>,<% =g_intPuerto %>,<% =g_intFechaDescarga %>,'',<% =g_intPlanillaCos %>,<% =g_intPlanillaNro %>,<% =g_intSolicitudNro %>,<% =g_intCartaPorte %>);" title="<% =GF_Traducir("Ver Detalle Descarga")%>">
                                    <%else%>
                                        <img id="img<% =Cont %>" src="images/Tplusik.gif" style="cursor:hand;" onClick="fcnExpand(tbl<% =Cont %>,img<% =Cont %>)">
                                    <%end if%>
                               </td>
                              <TD ALIGN="CENTER"><% =GF_FN2DTE(g_intFechaDescarga)%></TD>
                              <TD ALIGN="CENTER">
                                 <%call GF_MGC("PU",g_intPuerto,0,auxDs)
                                 response.write auxDs
                                 'response.write g_intPuerto%>
                              </TD>
                              <TD ALIGN="CENTER"><% =g_intCartaPorte %></TD>
                              <TD align="center">
                                    <%if (GF_DT1("READ","SHOWGRAD","","","AR",g_intProducto) <> "?") then
                                        if (ucase(g_chrMrcConforme) = "V") then
                                            if (cInt(g_intAnalisisGdo) <= 3) then
                                                response.write g_intAnalisisGdo
                                            else
                                                response.write "3"
                                            end if
                                        elseif initHeaderAnalisis() = 1 then
                                            call getNextAnalisis()
                                            if (cint(g_intAnalisisGdo) <= 3) then
                                                response.write g_intAnalisisGdo
                                            else
                                                response.write "3"
                                            end if
                                        else
                                            response.write "&nbsp;"
                                        end if
                                    else
                                        if (ucase(g_chrMrcConforme) = "V") then
                                            response.write GF_Traducir("Conf.")
                                        else
                                            response.write "&nbsp;"
                                        end if
                                    end if%>
                              </td>
                              <TD align="right">
                                    <%retValue = GF_DT1("READ","showprot","","","AR",g_intProducto)
                                    if lcase(retValue) = "true" then
                                        'Busco el valor de prot en el analisis
                                        if initHeaderAnalisis()="1" then
                                            call getNextAnalisis()
                                            g_intConcepto = "33" 'Proteinas
                                            if initHeaderDetAnalisis()="1" then
                                                call getNextDetAnalisis()
                                                response.write Editar_Importe(g_intValor) & "%"
                                            else
                                                response.write "&nbsp;"
                                            end if
                                        else
                                            response.write "&nbsp;"
                                        end if
                                    else
                                        response.write "&nbsp;"
                                    end if%>
                              </td>
                              <TD ALIGN="RIGHT">
                                  <%Call GP_SERVEREXECUTE("GF_convertUnit.asp",g_intKilosDescarga,"1",unitDest,"","",retValue,7,8,9,10)
                                  RESPONSE.WRITE retValue & " " & GF_DT1("READ","DSAB","","","MU",unitDest)%>
                              </TD>
                          </TR>
                          <%if (g_intCantidad <> 1) and (rs.Recordcount<>1) then%>
                               <tr id="tbl<% =cont %>" style="visibility=hidden;position=absolute;">
                                   <td width="5%">&nbsp</td>
                                   <td colspan="6">
                        		      <table class=reg_header align="left" cellSpacing=1 cellPadding=2 width="30%">
                        			     <tr class=reg_header_nav align="center">
                                             <td width="5%">&nbsp;</td>
                                             <td width="50%"><% =GF_TRADUCIR("Sol. Analisis")%></td>
                            			     <td><% =GF_TRADUCIR("Cantidad")%></td>
                                         </tr>
                                        <%while not rs.eof%>
                                            <tr class="reg_header_navdos" align="center">
                                                <td>
                                                    <img src="images/ver.gif" style="cursor:hand;" onClick="javascript:fcnCall3(<% =g_intProducto %>,<% =g_intSucursal %>,<% =g_intOperacion %>,<% =g_intNumero %>,<% =g_intCosecha %>,<% =unitDest %>,<% =g_intPuerto%>,<% =g_intFechaDescarga%>,'<% =rs("PlanillaSec")%>',<%=g_intPlanillaCos%>,<%=g_intPlanillaNro%>,<% =g_intSolicitudNro %>,<% =g_intCartaPorte %>);" title="<% =GF_Traducir("Ver Detalle Descarga")%>;">
                                                </td>
                                                <td><% =rs("SolicitudNro")%></td>
                                			     <td align="right">
                                                    <%Call GP_SERVEREXECUTE("GF_convertUnit.asp",rs("KgDescarga"),"1",unitDest,"","",retValue,7,8,9,10)
                                                    response.write retValue%>
                                                 </td>
                                             </tr>
                                            <%rs.movenext
                                        wend%>
                                       </table>
                                   </td>
                                </tr>
                            <%end if
                            cont = cont + 1
                         wend
                         g_intConcepto=""
                         call GF_reset_Descargas()
                     end if%>
                     </TABLE>
                  </TD>
                  <TD HEIGHT="100%"><img src="images/marco_r2_c3.gif" WIDTH="8" HEIGHT="100%"></TD>
             </TR>
             <TR>
                 <TD WIDTH="8"><img src="images/marco_r3_c1.gif"></TD>
                 <TD COLSPAN="3"><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></TD>
                 <TD WIDTH="8"><img src="images/marco_r3_c3.gif"></TD>
              </TR>
      </TABLE>
      <!-- FIN DESCARGAS -->
</FORM>
</BODY>
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
