<%
'                   ***   PROCEDIMIENTOS  RETENCIONES   ***
'/**
' * CONSTANTES MODIFICADORAS DE BODY
' */
Const BODY_NN = 0
Const BODY_IG = 1 'Impuesto a las Ganancias
Const BODY_IB = 2 'Ingresos Brutos
Const BODY_CP = 3 'Contribuciones Patronales

Const RET_RECLASIFICACION = "RCL"
'/**
' * CONSTANTES MODIFICADORAS DE BODY
' */
'Variables Globales de las Retenciones cor_PGCAB
dim Gbl_FechaPago, Gbl_TipoCbte, Gbl_Minuta, Gbl_NroCbteProv
dim Gbl_FechaCbte, Gbl_PagoCobro, Gbl_NroOrden, Gbl_KCCOR, Gbl_KCVEN
dim Gbl_ImpCbte, Gbl_ImpRet, Gbl_ImpMerc, Gbl_ImpIVA
dim Gbl_KCMERC, Gbl_KCIVA, Gbl_NroContrato
'Variables Globales de las Retenciones cor_PGDET
dim Gbl_KCDetalle, Gbl_ImpPesos, Gbl_ImpDolar, Gbl_DbtCdt, Gbl_KCPago
dim Gbl_NroLote, Gbl_MrcPagoIVA, Gbl_KCPRO, Gbl_KCBC, Gbl_BCSC
dim Gbl_NroChc, Gbl_NroRet, Gbl_MrcAnul
'Variables para el agente de retencion y sujeto retenido
dim Gbl_AgntDen, Gbl_AgntCUIT, Gbl_AgntDom, Gbl_AgntLoc, Gbl_Agntnib, Gbl_Agntconv, Gbl_Agntnar
dim Gbl_RspDen, Gbl_RspCUIT, Gbl_RspDom, Gbl_RspLoc, Gbl_Rspnib, Gbl_Rspconv, Gbl_Rspnar
'Variables especiales para datos extras de retenciones
dim Gbl_BaseImponible, Gbl_Alicuota, Gbl_NoSujetoARet, Gbl_SujetoARet
dim Gbl_Concepto, Gbl_AcumuladoMensual, Gbl_CondGanancias, Gbl_RetAcumulado
dim Gbl_RetAnteriores, Gbl_TotalFacturado, Gbl_TotalIVAFacturado
'Variables para la escritura del archivo PDF
dim Gbl_oPDF
'-----------------------------------------------------------------
function GF_ObtenerAgenteRetencion(p_kcAgnt, byref p_Denominacion, byref p_CUIT, byref p_Domicilio, byref p_Localidad, byref p_nib, byref p_convenio, byref p_nragret)
dim My_Sql, My_rs, My_cn, strDato
GF_ObtenerAgenteRetencion = false
  My_Sql = "Select NRODOC as CUIT, DOMICI as Domicilio, LOCALI as Localidad, NROIBR as NIBrutos, NROCML as Convenio	from MERFL.TCB6A1F1 where NROPRO =" & p_KCAgnt
  call GF_BD_AS400_2(My_rs, My_cn, "OPEN", My_Sql)
  if not My_Rs.eof then
     p_CUIT      = cDbl(My_rs("CUIT"))
     p_Domicilio = trim(My_rs("Domicilio"))
     p_Localidad = trim(My_rs("Localidad"))
     p_Denominacion = GetDSEnterprise2(p_kcAgnt)
     p_nib = cDbl(My_rs("NIBrutos"))
     p_convenio = cDbl(My_rs("Convenio"))
     p_nragret = getNrAgret(Gbl_KCDetalle)
     GF_ObtenerAgenteRetencion = true
  else
     p_CUIT = ""
     p_Domicilio = ""
     p_Localidad = ""
  end if
end function
'-----------------------------------------------------------------
'funcion que devuelve el numero de agente 
' * Funcion: getNrAgret
' * Descripcion: Devuelve el numero de agente de retención.
' * Parametros: p_kcDetalle.
' * Autor: Juan Pablo Santi.
' * Fecha: 29/03/2010
Function getNrAgret(p_kcDetalle)
	dim strSQL, conn, rsNrAgret, rntn
	strSQL = "select NRAGENTERETENCION from TBLCONCEPTOPAGO where CDCONCEPTO = '" & p_kcDetalle & "'"
	call GF_BD_AS400(rsNrAgret, conn, "OPEN", strSQL)
	if (not rsNrAgret.eof) then
		rntn = rsNrAgret("NRAGENTERETENCION")
	else
		rntn = "?"
	end if
	getNrAgret = rntn
end Function
'-----------------------------------------------------------------
function GF_ObtenerResponsableRetencion(p_kcAgnt, byref p_Denominacion, byref p_CUIT, byref p_Domicilio, byref p_Localidad,byref p_nib, byref p_convenio, byref p_nragret)
   call GF_ObtenerAgenteRetencion(p_kcAgnt, p_Denominacion,  p_CUIT, p_Domicilio, p_Localidad, p_nib, p_convenio, p_nragret)
end function
'-----------------------------------------------------------------
function GF_ObtenerValoresRetencionCAB(p_FechaPago, p_TipoCbte, p_Minuta, p_nroProv)
Dim My_Sql, My_Rs, My_Cn,intSalir
Dim intFechaPago,strTipoCbte,intMinuta
Dim strSQL,rs,conn
            GF_CargarValoresRetencionCAB = false
            inSalir = 0
            intFechaPago = p_FechaPago
            strTipoCbte = p_TipoCbte
            intMinuta = p_Minuta
            call GF_CAB_LEER(My_Rs,intFechaPago,strTipoCbte,intMinuta, p_nroProv)
            while (intSalir = 0) and (not My_Rs.eof)
                  if (trim(My_Rs("TipoCbte")) = "RCL") and (Gbl_MrcAnul <> "F") then
					 My_Sql = "Select MRFPAR as FechaPago_1, MRNOPR as Orden_1, MRFPAO as FechaPago_2, MRNOPO as Orden_2 from TESFL.TES134F4 where MRFPAR =" & My_Rs("FechaPago") & " and MRNOPR = " & My_Rs("Orden")
                     'response.Write "<BR>" & MY_Sql
                     call GF_BD_AS400_2(My_Rs,My_Cn,"OPEN",MY_Sql)
                     if (not My_Rs.eof) then
                        My_Sql = "Select WCFPAG as fechaPago, WCTCBT as tipoCbte, WCNING as minuta, WCNCPV as CbteProveedor, WCFCPV as FechaCbte, WCPGCB as PC, WCNOPC as orden, WCNPRO as KCCOR, WCPRET as KCVEN, WCIMBT as Importe, WCIMRE as ImporteRet, WCIMME as importeMerc, WCIMIV as ImporteIVA, WCPOME as KCMERC, WCPOIV as KCIVA, WCCONT as Contrato "
                        My_Sql = My_Sql & " from TESFL.TES960F1 where WCFPAG = " & My_Rs("FechaPago_2") & " and WCNOPC=" & My_Rs("Orden_2")
                        'response.Write "<BR>" & MY_Sql
                        call GF_BD_AS400_2(My_Rs,My_Cn,"OPEN",MY_Sql)
                     end if
                  else
				  		Gbl_NroOrden = My_Rs("Orden")
                        if ( My_Rs("TIPOCBTE") = RET_RECLASIFICACION ) then
							strSQL =          "SELECT WCFPAG AS fechaPago    , "
							strSQL = strSQL & "       WCTCBT AS tipoCbte     , "
							strSQL = strSQL & "       WCNING AS minuta       , "
							strSQL = strSQL & "       WCNCPV AS CbteProveedor, "
							strSQL = strSQL & "       WCFCPV AS FechaCbte    , "
							strSQL = strSQL & "       WCPGCB AS PC           , "
							strSQL = strSQL & "       WCNOPC AS orden        , "
							strSQL = strSQL & "       WCNPRO AS KCCOR        , "
							strSQL = strSQL & "       WCPRET AS KCVEN        , "
							strSQL = strSQL & "       WCIMBT AS Importe      , "
							strSQL = strSQL & "       WCIMRE AS ImporteRet   , "
							strSQL = strSQL & "       WCIMME AS importeMerc  , "
							strSQL = strSQL & "       WCIMIV AS ImporteIVA   , "
							strSQL = strSQL & "       WCPOME AS KCMERC       , "
							strSQL = strSQL & "       WCPOIV AS KCIVA        , "
							strSQL = strSQL & "       WCCONT AS Contrato "
							strSQL = strSQL & " FROM   TESFL.TES134F4 "
							strSQL = strSQL & "       INNER JOIN TESFL.TES960F1 "
							strSQL = strSQL & "       ON     WCFPAG = MRFPAO "
							strSQL = strSQL & "       AND    WCNOPC = MRNOPO "
							strSQL = strSQL & " WHERE  MRFPAR        = " & intFechaPago
							strSQL = strSQL & " AND    MRNOPR        = " & intMinuta
	
							call GF_BD_AS400_2(rs,conn,"OPEN",strSQL)
							
							
							if (not rs.EoF) then
								Gbl_NroCbteProv = rs("CbteProveedor")
								Gbl_FechaCbte   = rs("FechaCbte")
								Gbl_PagoCobro   = rs("PC")
								Gbl_KCCOR       = cdbl(rs("KCCOR"))
								Gbl_KCVEN       = cdbl(rs("KCVEN"))
								Gbl_ImpCbte     = Cdbl(rs("Importe"))
								Gbl_ImpRet      = Cdbl(rs("ImporteRet"))
								Gbl_ImpMerc     = Cdbl(rs("ImporteMerc"))
								Gbl_ImpIVA      = Cdbl(rs("ImporteIVA"))
								Gbl_KCMERC      = cdbl(rs("KCMERC"))
								Gbl_KCIVA       = cdbl(rs("KCIVA"))
								Gbl_NroContrato = GF_nDigits(rs("Contrato"), 12)
							end if
						else
							Gbl_NroCbteProv = My_Rs("CbteProveedor")
							Gbl_FechaCbte   = My_Rs("FechaCbte")
							Gbl_PagoCobro   = My_Rs("PC")
							Gbl_KCCOR       = cdbl(My_Rs("KCCOR"))
							Gbl_KCVEN       = cdbl(My_Rs("KCVEN"))
							Gbl_ImpCbte     = Cdbl(My_Rs("Importe"))
							Gbl_ImpRet      = Cdbl(My_Rs("ImporteRet"))
							Gbl_ImpMerc     = Cdbl(My_Rs("ImporteMerc"))
							Gbl_ImpIVA      = Cdbl(My_Rs("ImporteIVA"))
							Gbl_KCMERC      = cdbl(My_Rs("KCMERC"))
							Gbl_KCIVA       = cdbl(My_Rs("KCIVA"))
							Gbl_NroContrato = GF_nDigits(My_Rs("Contrato"), 12)
						end if							

                        GF_CargarValoresRetencionCAB = true
                        intSalir=1
                  end if
            wend
            if (intSalir = 0) then
                Gbl_NroCbteProv = ""
                Gbl_FechaCbte   = ""
                Gbl_PagoCobro   = ""
                Gbl_NroOrden    = ""
                Gbl_KCCOR       = ""
                Gbl_KCVEN       = ""
                Gbl_ImpCbte     = ""
                Gbl_ImpRet      = ""
                Gbl_ImpMerc     = ""
                Gbl_ImpIVA      = ""
                Gbl_KCMERC      = ""
                Gbl_KCIVA       = ""
                Gbl_NroContrato = ""
                GF_CargarValoresRetencionCAB = False
            end if
end function
'------------------------------------------------------------------
function GF_ObtenerValoresRetencionDET(p_KCDetalle, p_RetNro, p_fecha)
dim My_Sql, My_Rs, My_Cn
            GF_ObtenerValoresRetencionDET = false
            My_Sql = "Select WDFPAG as FechaPago, WDTCBT as TipoCbte, WDNING as Minuta, WDCODE as KCDetalle, WDIMPP as ImportePesos, WDIMPD as ImporteDolar, WDDBCR as DBCR, WDFOPG as KCPago, WDNRLT as Lote, WDPIVA as MRCPago, WDPOPG as KCPRO, WDCBCO as KCBC, WDSBCO as KCBCSC, WDNCHE as Cheque, WDNRET as RetNro, WDMBAJ as MarcaAnulacion "
            My_Sql = My_Sql & " from TESFL.TES960F2 where WDFPAG='" & p_fecha & "' and WDCODE='" & p_KCDetalle & "' and WDNRET='" & p_RetNro & "'"
            call GF_BD_AS400_2 (My_Rs, My_Cn, "OPEN", My_Sql)
            if not My_Rs.eof then
               Gbl_FechaPago  = My_Rs("FechaPago")
               Gbl_TipoCbte   = My_Rs("TipoCbte")
               Gbl_Minuta     = My_Rs("Minuta")
               Gbl_KCDetalle  = trim(My_Rs("KCDetalle"))
               Gbl_ImpPesos   = CDbl(My_Rs("ImportePesos"))
               Gbl_ImpDolar   = CDbl(My_Rs("ImporteDolar"))
               Gbl_DbtCdt     = My_Rs("DBCR")
               Gbl_KCPago     = My_Rs("KCPago")
               Gbl_NroLote    = My_Rs("Lote")
               Gbl_MrcPagoIVA = My_Rs("MRCPago")
               Gbl_KCPRO      = My_Rs("KCPRO")
               Gbl_KCBC       = My_Rs("KCBC")
               Gbl_BCSC       = My_Rs("KCBCSC")
               Gbl_NroChc     = My_Rs("Cheque")
               Gbl_NroRet     = GF_nDigits(My_Rs("RetNro"),12)               
               Gbl_MrcAnul    = My_Rs("MarcaAnulacion")
               GF_ObtenerValoresRetencionDET = true
            else
               Gbl_FechaPago  = ""
               Gbl_TipoCbte   = ""
               Gbl_Minuta     = ""
               Gbl_KCDetalle  = ""
               Gbl_ImpPesos   = ""
               Gbl_ImpDolar   = ""
               Gbl_DbtCdt     = ""
               Gbl_KCPago     = ""
               Gbl_NroLote    = ""
               Gbl_MrcPagoIVA = ""
               Gbl_KCPRO      = ""
               Gbl_KCBC       = ""
               Gbl_BCSC       = ""
               Gbl_NroChc     = ""
               Gbl_NroRet     = ""
               Gbl_MrcAnul    = ""
            end if
end function
'------------------------------------------------------------------
Function GF_ObtenerDatosExtras(p_intFechaPago,p_strRetNro)
         Dim strSQL, oConn, rs
         
         'Inicializo las variables
         Gbl_BaseImponible=0
         Gbl_Alicuota=0
         Gbl_Concepto=""
         Gbl_NoSujetoARet=0
         Gbl_SujetoARet=0
         Gbl_AcumuladoMensual=0
         Gbl_CondGanancias=""
         Gbl_RetAnteriores=0
         Gbl_RetAcumulado=0
         Gbl_TotalFacturado=0
         Gbl_TotalIVAFacturado=0
         'Obtengolos valores
         strSQL="Select DCFPAG as FechaPago, DCCODE as Tipo, DCNRET as RetNro, DCVACD as Codigo, DCVADA as Valor from TESFL.TES960F3 where DCNRET ='" & p_strRetNro
         strSQL= strSQL & "' and DCFPAG =" & p_intFechaPago           
         Call GF_BD_AS400_2 (rs, oConn, "OPEN", strSQL)

         While (not rs.eof)
               Select case rs("Codigo")
                      case "BI": Gbl_BaseImponible=CDbl(rs("Valor"))
                      case "PR": Gbl_Alicuota=CDbl(rs("Valor"))
                      case "CR": Gbl_Concepto=rs("Valor")
                      case "MN": Gbl_NoSujetoARet=rs("Valor")
                      case "AM": Gbl_AcumuladoMensual=CDbl(rs("Valor"))
                      case "CG": Gbl_CondGanancias=rs("Valor")
                      case "RA": Gbl_RetAnteriores=CDbl(rs("Valor"))
                      case "RT": Gbl_RetAcumulado=CDbl(rs("Valor"))
                      case "MS": Gbl_SujetoARet=rs("Valor")
                      case "TF": Gbl_TotalFacturado=CDbl(rs("Valor"))
                      case "TI": Gbl_TotalIVAFacturado=getImporte(rs("Valor"))
               End Select
               rs.MoveNext
         Wend         
End Function
'----------------------------------------------------------------
Function getImporte(pValor)
    Dim ret
    
    if (isNumeric(pValor)) then
        ret = CDbl(pValor)
    else
        'El valor el negativo y el iSeries lo devuelve con un caracter al final en lugar del signo!!        
        ret = CDbl(Left(Trim(pValor), Len(Trim(pValor))-1)) * -1
    end if        
    getImporte = ret
End Function
'----------------------------------------------------------------
Function GF_Print_Header(p_titulo, p_detalle)
%>
<table align=center width="95%" border=0 cellpadding=0 cellspacing=0>
   <tr>
   <td><img src="images/marco_r1_c1.gif" border=0 width=8 height=8></td>
      <td colspan="2"><img src="images/marco_r1_c2.gif" border=0 width=730 height=8></td>
      <td><img src="images/marco_r1_c3.gif" border=0 width=8 height=8></td>
   </tr>
   <tr>
      <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
      <td colspan=2 width="70%"><B><font class="BIG"><% =GF_Traducir(p_titulo) %></font></B></td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
   	  <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
 	  <td colspan="2">&nbsp;</td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
      <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
	  <td colspan=2 align="left" valign="top"><B><font class="BIG"><tt>
        <%if left(Gbl_FechaPago,8) > "20050809" then
            response.write GF_Traducir(p_AgntDen)
        else
            response.write "Alfred C. Toepfer International S.A."
        end if%>
      </tt></font></B></td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
  	  <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
	  <td colspan=1 align="left" valign="top"><B><font class="BIG"><tt><% =GF_Traducir(p_AgntDom) %></tt></font></B></td>
 	  <td rowspan=3 align=right><img src="Images/kogge256.gif" width=75 height=60 border=0></td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
  	  <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
	  <td  colspan=2 align="left" valign="top"><B><font class="BIG"><tt><% =GF_Traducir(p_AgntLoc) %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =GF_Traducir(p_detalle) %></tt></font></B></td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
      <td><img src="images/marco_r2_c1.gif" border=0 width=8 height=20></td>
 	  <td colspan="2">&nbsp;</td>
  	  <td><img src="images/marco_r2_c3.gif" border=0 width=8 height=20></td>
   </tr>
   <tr>
    <td><img src="images/marco_r3_c1.gif" border=0 width=8 height=8></td>
	<td colspan="2"><img src="images/marco_r3_c2.gif" border=0 width=730 height=8></td>
        <td><img src="images/marco_r3_c3.gif" border=0 width=8 height=8></td>
   </tr>
</table>
<%
End function
'----------------------------------------------------------------
Function GF_Print_Controls(p_name,p_other)

Dim p_strControls
'Recibo el parametro que indica si se deben o no dibujar estos controles.
p_strControls = GF_PARAMETROS7("P_CONTROLS","",6)
if (isEmpty(p_strControls)) then
%>
<table id="<% =p_name %>" border=0 align="center" cellspacing=0 cellpadding=0>
      <tr>
        <td width="8"><img src="images/marco_r1_c1.gif"></td>
	<td><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></td>
	<td width="8"><img src="images/marco_t_s.gif"></td>
	<td><img src="images/marco_r1_c2.gif" WIDTH="100%" HEIGHT="8"></td>
	<td width="8"><img src="images/marco_r1_c3.gif"></td>
      </tr>
      <tr>
        <td width="8"><img src="images/marco_r2_c1.gif"></td>
	<td align="center"><a href="javascript:fcnAction();"><% =GF_TRADUCIR("CERRAR") %></a></td>
	<td width="8"><img src="images/marco_c_v.gif"></td>
	<td align="center"><a href="javascript:`Imprimir(<% =p_name %>,<% =p_other%>);"><% =GF_TRADUCIR("IMPRIMIR") %></a></td>
	<td width="8"><img src="images/marco_r2_c3.gif"></td>
      </tr>
      <tr>
        <td width="8"><img src="images/marco_r3_c1.gif"></td>
	<td><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></td>
	<td width="8"><img src="images/marco_t_b.gif"></td>
	<td><img src="images/marco_r3_c2.gif" WIDTH="100%" HEIGHT="8"></td>
	<td width="8"><img src="images/marco_r3_c3.gif"></td>
      </tr>
</table>
<%
end if
End Function
'----------------------------------------------------------------
Function GF_ObtenerParametrosRetencion(ByRef p_tipo, ByRef p_RetNro, ByRef p_fecha)

'Bajo los parametros
p_tipo=GF_Parametros7("P_TIPO","",6)
p_RetNro=GF_Parametros7("P_RET","",6)
p_fecha=GF_Parametros7("P_FECHA","",6)
if isEmpty(p_tipo) or isEmpty(p_RetNro) then
   response.redirect("mgmsg.asp?P_MSG=PARAMETROS INCORRECTOS.")
end if

End Function
'-------------------------------------------------------------------------------
Function GF_Print_Std_Header(p_titulo, p_detalle)
         'Se dibuja el recuadro
         Call GF_squareBox(Gbl_oPDF, 5, 10, 585, 125, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
         'Se coloca el Logo
         Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\kogge256.gif"), 520, 20, 60, 60, 0)
         'Se escribe la informacion de la cabecera
         Call GF_setFont(Gbl_oPDF,"ARIAL", 14,8)
         Call GF_writeText(Gbl_oPDF,15, 23, GF_Traducir(p_titulo), 0)
         Call GF_setFont(Gbl_oPDF,"Courier", 12, 8)
         if left(Gbl_FechaPago,8) > "20050809" then
            Call GF_writeText(Gbl_oPDF,15, 60, GF_Traducir(Gbl_AgntDen), 0)
         else
            Call GF_writeText(Gbl_oPDF,15, 60, "Alfred C. Toepfer International S.A.", 0)
         end if
         Call GF_writeText(Gbl_oPDF,15, 75, GF_Traducir(Gbl_AgntDom), 0)
         Call GF_writeText(Gbl_oPDF,15, 90, GF_Traducir(Gbl_AgntLoc), 0)
         Call GF_writeText(Gbl_oPDF,240, 90, GF_Traducir(p_detalle), 0)
End Function
'-------------------------------------------------------------------------------
Function GF_Print_Std_Body(p_showSpecial)

         Dim strTexto, auxRspDen
         
         'Se dibujan el recuedro.
         Call GF_squareBox(Gbl_oPDF, 5, 155, 585, 680, 0, "#FFFFFF", "#b1bca7", 2, PDF_SQUARE_ROUND)
         'Se setea la fuente a utilizar.
         Call GF_setFont(Gbl_oPDF,"Courier", 10, 0)
         'Datos de la Retencion.
         Call GF_writeText(Gbl_oPDF,15,170,GF_Traducir("Lugar y Fecha") & ".....: " & Gbl_AgntLoc & " " & GF_FN2DTE(Gbl_FechaPago),0)
         Call GF_writeText(Gbl_oPDF,335,170,GF_Traducir("Retencion Nro") & "....: " & GF_EDIT_CBTE(Gbl_NroRet),0)
         Call GF_writeText(Gbl_oPDF,15,185,GF_Traducir("C.U.I.T.") & "..........: " & GF_STR2CUIT(Gbl_AgntCUIT),0)
         Select case (p_showSpecial)
                case BODY_IB:
                     Call GF_writeText(Gbl_oPDF,335,185,GF_Traducir("No Agente Retenc") & ".: " & Gbl_Agntnar,0)
                case else:
                     Call GF_writeText(Gbl_oPDF,335,185,GF_Traducir("Orden de Pago Nro") & ": " & Gbl_NroOrden,0)
         End Select
         Call GF_H_SEPARADOR(Gbl_oPDF, 5, 201, 585)
         'Datos del Agente Retenido.
         auxRspDen = Gbl_RspDen
         'if (len(auxRspDen) > 35) then auxRspDen = left(Gbl_RspDen,30) & "..."
         Call GF_writeText(Gbl_oPDF,15,208,GF_Traducir("Sujeto Retenido") & "...: ",0)
         Call GF_writeTextPlus(Gbl_oPDF,135,208,auxRspDen, 200, 10,0)
         Call GF_writeText(Gbl_oPDF,347,208,GF_Traducir("C.U.I.T.") & ".......: " & GF_STR2CUIT(Gbl_RspCuit),0)
         Call GF_writeText(Gbl_oPDF,15,233,GF_Traducir("Direccion") & ".........: " & Gbl_RspDom,0)
         Call GF_writeText(Gbl_oPDF,135,248,Gbl_RspLoc,0)
         Call GF_writeText(Gbl_oPDF,347,228,GF_Traducir("Proveedor Nro") & "..: " & Gbl_KCPRO,0)
         Select case (p_showSpecial)
                case BODY_IG:
                     Call GF_writeText(Gbl_oPDF,347,248,GF_Traducir("Condicion") & "......: " & Gbl_CondGanancias,0)
                case BODY_IB:
                     if (isEmpty(Gbl_Rspconv) or isNull(Gbl_Rspconv)) then
                         Call GF_writeText(Gbl_oPDF,347,248,GF_Traducir("Nro Ing.Brutos") & ".: " & p_Rspnib,0)
                     else
                         Call GF_writeText(Gbl_oPDF,347,248,GF_Traducir("Conv. Multilat.") & ": " & Gbl_Rspconv,0)
                     end if
         End Select
         Call GF_H_SEPARADOR(Gbl_oPDF, 5, 264, 585)
         'Datos de la factura.
         Call GF_writeText(Gbl_oPDF,15,281,GF_Traducir("Lugar y Fecha Fac.") & ": " & Gbl_RspLoc & " " & GF_FN2DTE(Gbl_FechaCbte),0)
         Call GF_writeText(Gbl_oPDF,385,281,GF_Traducir("Fac. Nro") & ".: " & GF_EDIT_CBTE(Gbl_NroCbteProv),0)
         Call GF_writeText(Gbl_oPDF,15,296,GF_Traducir("Concepto Retencion") & ": " & GF_Traducir(Gbl_Concepto),0)
         Call GF_H_SEPARADOR(Gbl_oPDF, 5, 317, 585)
         'Firma del apoderado.
         if left(Gbl_FechaPago,8) > "20070524" then
            Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\firma_elizabeth.jpg"), 370, 550, 170, 170, 0)
         else
            if left(Gbl_FechaPago,8) > "20050809" then
                Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\firma_Pedro5.jpg"), 370, 550, 150, 150, 0)
            else
                Call GF_writeImage(Gbl_oPDF, Server.MapPath("Images\firma_Pedro4.jpg"), 370, 550, 150, 150, 0)
            end if
         end if
End Function
'-------------------------------------------------------------------------------
Function GF_Cargar_Retencion(p_RetNro,p_tipo,p_fecha)

Dim ret
ret = False
if (GF_ObtenerValoresRetencionDET(p_tipo, p_RetNro, p_fecha)) then
   Call GF_ObtenerValoresRetencionCAB(Gbl_FechaPago, Gbl_TipoCbte, Gbl_Minuta, Gbl_KCPRO)
   Call GF_ObtenerAgenteRetencion("99999997", Gbl_AgntDen, Gbl_AgntCUIT, Gbl_AgntDom, Gbl_AgntLoc, Gbl_Agntnib, Gbl_Agntconv, Gbl_Agntnar)
   Call GF_ObtenerResponsableRetencion(Gbl_KCPRO, Gbl_RspDen, Gbl_RspCUIT, Gbl_RspDom, Gbl_RspLoc, Gbl_Rspnib, Gbl_Rspconv, Gbl_Rspnar)
   Call GF_ObtenerDatosExtras(Gbl_FechaPago,CLng(Gbl_NroRet))
   ret = True
end if
GF_Cargar_Retencion = ret

End Function
'-------------------------------------------------------------------------------
'Retencion IVA 615/99 - 2854/10
Function GF_Retencion_C(p_RetNro,p_tipo,p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion al Impuesto al Valor Agregado", "Resolucion General Nro 2854/10 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_NN)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Total Fac. 100%") & "...: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_ImpCbte*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("I.V.A. Facturado") & "..: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349, GF_EDIT_DECIMALS(Gbl_TotalIVAFacturado,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("I.V.A. Retenido") & "...: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   Call GF_writeText(Gbl_oPDF,15,405,GF_Traducir("Nota : La firma ") & Gbl_RspDen & GF_Traducir(" se encuentra comprendida en el"),0)
   Call GF_writeText(Gbl_oPDF,55,420,GF_Traducir(" regimen especial previsto en la Resolucion General Nro.2854/10 A.F.I.P"),0)
   GF_Retencion_C = True
end if

End Function
'-------------------------------------------------------------------------------
'Retencion IVA 1394/02 - IVA 2300/07
Function GF_Retencion_E(p_RetNro,p_tipo,p_fecha)
Dim strProducto, strLine, coord_Y

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion al Impuesto al Valor Agregado", "Resolucion General Nro 2300/07 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_NN)
   'Se escriben los campos particulares.
   'Obtengo la descripcion del producto
   strProducto="???"
   if (len(Gbl_NroContrato) > 2) then
      strProducto=getDsProduct(left(Gbl_NroContrato,2))
   end if
   'Armo el texto.
   strLine = GF_Traducir("     Por la presente se deja constancia de haberse efectuado la retencion en el ")
   strLine = strLine & GF_Traducir("impuesto al valor agregado por la suma de Pesos ") & GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2) & GF_Traducir(" por la ")
   strLine = strLine & GF_Traducir("operacion de compra de") & " " & strProducto & GF_Traducir(", segun factura o documento ")
   strLine = strLine & GF_Traducir("equivalente Nro.") & GF_EDIT_CBTE(Gbl_NroCbteProv) & GF_Traducir(" de fecha ") & GF_FN2DTE(Gbl_FechaCbte) & GF_Traducir(" por un importe total de ")
   strLine = strLine & GF_Traducir("Pesos ") & GF_EDIT_DECIMALS(Gbl_ImpCbte*100,2) & GF_Traducir(" de acuerdo con lo establecido por la ")
   strLine = strLine & GF_Traducir("Resolucion General 2300/07. A.F.I.P.<br>")
   'Imprimo el texto
   'coord_Y = GF_writeTextPlus(Gbl_oPDF, 35, 334, strLine, 395, 15, 4)
   'Armo el texto.
   strLine = strLine & GF_Traducir("Este comprobante se extiende en las condiciones previstas en el articulo 11vo ")
   strLine = strLine & GF_Traducir("de la resolucion general ya citada, a los fines de respaldar el computo de la ")
   strLine = strLine & GF_Traducir("retencion practicada, con arreglo a lo dispuesto en el articulo 13 de la ")
   strLine = strLine & GF_Traducir("mencionada norma.")
   'Imprimo el texto
   Call GF_writeTextPlus(Gbl_oPDF,35,334,strLine,395,15,4)
   'Call GF_writeTextPlus(Gbl_oPDF,35,coord_Y,strLine,500,15,4)
   'Imprimo los datos de terminacion.
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 726, 585)
   strLine = GF_Traducir("Negocio:") & GF_INT2CTO(Gbl_NroContrato) & "   " & GF_Traducir("Minuta:") & Gbl_Minuta & "   " & GF_Traducir("Corredor:") & GetDSEnterprise2(Gbl_KCCOR)
   coord_Y = GF_writeTextPlus(Gbl_oPDF,15,733,strLine, 395, 15, 0)
   strLine = GF_Traducir("Codigo de Regimen:") & CodigoRegimen_E()
   Call GF_writeText(Gbl_oPDF,15,coord_Y,strLine,0)
   GF_Retencion_E = True
end if

End Function
'------------------------------------------------------------------
function CodigoRegimen_E()
    if ((Gbl_TipoCbte = "BDB") and (Gbl_DbtCdt = 2))then
        CodigoRegimen_E = 787
    else
        CodigoRegimen_E = 784
    end if
End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_B(p_RetNro,p_tipo,p_fecha)

Dim strProducto, strLine, coord_Y, i, resol

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   resol = "Resolucion General Nro "
   if (esPagoMercaderia(Gbl_Minuta)) then
    resol = resol & "2218/06"
   else
    resol = resol & "830/00"
   end if
   resol = resol & " A.F.I.P."
   Call GF_Print_Std_Header("Constancia de Retencion al Impuesto a las Ganancias", resol)
   Call GF_Print_Std_Body(BODY_NN)
   'Se escriben los campos particulares.
   'Importes
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Importe de la Operacion") & "..........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,120,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Importe Acumulado Mensual") & "........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,120,349, GF_EDIT_DECIMALS(Gbl_AcumuladoMensual,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,350,349,GF_Traducir("Minimo:"), PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,410,349, GF_EDIT_DECIMALS(Gbl_NoSujetoARet,2), 58, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe de la Retencion") & "..........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,120,364, GF_EDIT_DECIMALS(Gbl_RetAcumulado,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,350,364,GF_Traducir("Porc..:"), PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,410,364,GF_EDIT_DECIMALS(Gbl_Alicuota,4), 58, PDF_ALIGN_RIGHT)
   'response.end
   Call GF_writeText(Gbl_oPDF,15,379,GF_Traducir("Importe de Retenciones Acumuladas") & ": ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,120,379, GF_EDIT_DECIMALS(Gbl_RetAnteriores,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,394,GF_Traducir("Importe Retenido") & ".................: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,120,394, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 415, 585)
   'La nota
   Call GF_writeText(Gbl_oPDF,15,423,GF_Traducir("Nota:"),0)
   strLine = GF_Traducir("La retencion se informara en las declaracion jurada del mes de ") & GF_TRADUCIR(GF_INT2MES(GF_DateGet("M",Gbl_FechaPago))) & "."
   coord_Y = GF_writeTextPlus(Gbl_oPDF,45,423,strLine,395,15,3)
   strLine = GF_Traducir("No se aceptaran  reclamos sobre este comprobante pasados tres dias de entrega, debido a requerimientos impositivos.")
   Call GF_writeTextPlus(Gbl_oPDF,45,coord_Y,strLine,395,15,3)
   GF_Retencion_B = True
end if

End Function

'-------------------------------------------------------------------------------
Function GF_Retencion_H(p_RetNro,p_tipo,p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Impuesto sobre los Ingresos Brutos Pcia. de Bs As.", "D.N. B 43/96")
   Call GF_Print_Std_Body(BODY_IB)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Base Imponible") & "...: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Alicuota") & ".........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349,GF_EDIT_DECIMALS(Gbl_Alicuota,4), 142, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & ".: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   GF_Retencion_H = True
end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_D(p_RetNro,p_tipo,p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Impuesto sobre los Ingresos Brutos Pcia. de Santa Fe.", "")
   Call GF_Print_Std_Body(BODY_IB)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Base Imponible") & "..........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_TotalFacturado,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Minimo no Imponible") & ".....: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349, GF_EDIT_DECIMALS(Gbl_NoSujetoARet,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Monto Sujeto a Retencion") & ": ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,379,GF_Traducir("Importe Retenido") & "........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,379, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,394,GF_Traducir("Alicuota") & "................: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,394,GF_EDIT_DECIMALS(Gbl_Alicuota,4), 185, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 415, 585)
   GF_Retencion_D = True
end if
End Function
'-------------------------------------------------------------------------------
Function GF_RegistroeInspeccion(p_title,p_param)

   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Derecho de Registro e Inspeccion " & p_title,p_param)
   Call GF_Print_Std_Body(BODY_IB)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Base Imponible") & "..........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_TotalFacturado,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Monto Sujeto a Retencion") & ": ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & "........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,379,GF_Traducir("Alicuota") & "................: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,379,GF_EDIT_DECIMALS(Gbl_Alicuota,4), 185, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 400, 585)

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_J(p_RetNro,p_tipo,p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then

   Call GF_RegistroeInspeccion("Arroyo Seco - Santa Fe.", "")
   GF_Retencion_J = True
   
end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_G(p_RetNro,p_tipo,p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then

   Call GF_RegistroeInspeccion("Gral San Martin - Santa Fe.", "")
   GF_Retencion_G = True

end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_I(p_RetNro, p_tipo, p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then	
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Impuesto sobre los Ingresos Brutos Ciudad de Bs As.", "")
   Call GF_Print_Std_Body(BODY_IB)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Base Imponible") & "...: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Alicuota") & ".........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,160,349,GF_EDIT_DECIMALS(Gbl_Alicuota,4), 142,PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & ".: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   GF_Retencion_I = True
end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_K(p_RetNro, p_tipo, p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Contribuciones Patronales S.U.S.S.", "Resolución General Nro.4052/95 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_CP)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Monto Sujeto a Reteneción") & "..: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Porc. Retenido") & ".............: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349,GF_EDIT_DECIMALS(Gbl_Alicuota,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & "...........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   Call GF_writeTextPlus(Gbl_oPDF,55,405,GF_Traducir("Nota : El Importe retenido será ingresado conforme lo establecido en el Art. 18 de la presente resolución."),395, 10, 3)
   Call GF_writeTextPlus(Gbl_oPDF,55,425,GF_Traducir("La firma ") & Gbl_RspDen & GF_Traducir(" se encuentra comprendida en el regimen especial previsto en la Resolucion General Nro. 1784/04 A.F.I.P"), 395, 10,3)
   GF_Retencion_K = True

end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_L(p_RetNro, p_tipo, p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Contribuciones Patronales S.U.S.S.", "Resolución General Nro.1784/05 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_CP)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Monto Sujeto a Reteneción") & "..: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190,PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Porc. Retenido") & ".............: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349, GF_EDIT_DECIMALS(Gbl_Alicuota,2), 190,PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & "...........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190,PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   Call GF_writeTextPlus(Gbl_oPDF,55,405,GF_Traducir("Nota : El Importe retenido será ingresado conforme lo establecido en el Art. 18 de la presente resolución."),395, 10, 3)
   Call GF_writeTextPlus(Gbl_oPDF,55,425,GF_Traducir("La firma ") & Gbl_RspDen & GF_Traducir(" se encuentra comprendida en el regimen especial previsto en la Resolucion General Nro. 1784/04 A.F.I.P. S.U.S.S. Reg. Gral."), 395, 10,3)
   GF_Retencion_L = True

end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_P(p_RetNro, p_tipo, p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Contribuciones Patronales S.U.S.S.", "Resolución General Nro.1769/04 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_CP)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Monto Sujeto a Reteneción") & "..: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Porc. Retenido") & ".............: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349,GF_EDIT_DECIMALS(Gbl_Alicuota,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & "...........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   Call GF_writeTextPlus(Gbl_oPDF,55,405,GF_Traducir("Nota : El Importe retenido será ingresado conforme lo establecido en el Art. 17 de la presente resolución."),395, 10, 3)
   Call GF_writeTextPlus(Gbl_oPDF,55,425,GF_Traducir("La firma ") & Gbl_RspDen & GF_Traducir(" se encuentra comprendida en el regimen especial previsto en la Resolucion General Nro. 1769/04 A.F.I.P. S.U.S.S. Reg. Gral."), 395, 10,3)
   GF_Retencion_P = True

end if

End Function
'-------------------------------------------------------------------------------
Function GF_Retencion_M(p_RetNro, p_tipo, p_fecha)

if (GF_Cargar_Retencion(p_RetNro, p_tipo, p_fecha)) then
   'Se escriben los campos comunes.
   Call GF_Print_Std_Header("Constancia de Retencion Contribuciones Patronales S.U.S.S.", "Resolución General Nro.1556/03 A.F.I.P.")
   Call GF_Print_Std_Body(BODY_CP)
   'Se escriben los campos particulares.
   Call GF_writeText(Gbl_oPDF,15,334,GF_Traducir("Monto Sujeto a Reteneción") & "..: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,334, GF_EDIT_DECIMALS(Gbl_BaseImponible,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,349,GF_Traducir("Porc. Retenido") & ".............: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,349,GF_EDIT_DECIMALS(Gbl_Alicuota,2), 190, PDF_ALIGN_RIGHT)
   Call GF_writeText(Gbl_oPDF,15,364,GF_Traducir("Importe Retenido") & "...........: ", PDF_ALIGN_LEFT)
   Call GF_writeTextAlign(Gbl_oPDF,100,364, GF_EDIT_DECIMALS(Gbl_ImpPesos*100,2), 190, PDF_ALIGN_RIGHT)
   Call GF_H_SEPARADOR(Gbl_oPDF, 5, 385, 585)
   Call GF_writeTextPlus(Gbl_oPDF,55,405,GF_Traducir("Nota : El Importe retenido será ingresado conforme lo establecido en el Art. 10 de la presente resolución."),395, 10, 3)
   Call GF_writeTextPlus(Gbl_oPDF,55,425,GF_Traducir("La firma ") & Gbl_RspDen & GF_Traducir(" se encuentra comprendida en el regimen especial previsto en la Resolucion General Nro. 1556/03 A.F.I.P. S.U.S.S. Reg. Gral."), 395, 10,3)
   GF_Retencion_M = True

end if

End Function
'*****************************************************************************************
sub enviarMailsRetenciones(byval p_KCPRO, BYVAL p_tipoRet, byval p_nroRet)
    dim strDestinatario, strAsunto, strPathAttachment, ORKC, cantEnvios, de
    dim conn, rs, strSQL
    dim vecMails(1)

    'completo los datos del mail
    strTituloRet = getTituloTipo(p_tipoRet)
    strAsunto = "Retención " & strTituloRet & " Nro. " & p_nroRet
    strPathAttachment = Server.mapPath("temp/") & "\" & getFileName(p_tipoRet, p_nroRet)
    strToepferDenomination = GetDsEnterprise2("99999997")
    de = strToepferDenomination & " <" & SENDER_IMPUESTOS & ">"
    strBody = "Se adjunta al presente mail la retención  de " & strTituloRet & " Nro: " & GF_EDIT_CBTE(p_nroRet) & chr(13) & chr(10) & chr(13) & chr(10)
    strBody = strBody & "                  " & strToepferDenomination
    'Busco los mails del destinatario y envio
    call obtenerMailRetenciones(p_KCPRO, vecMails)
    'vecMails(0) = "bacariniE@toepfer.com"
	if ((not isnull(vecMails(0))) and (vecMails(0) <> "")) then strDestinatario = vecMails(0) & "; "
	if ((not isnull(vecMails(1))) and (vecMails(1) <> "")) then strDestinatario = strDestinatario & vecMails(1) & ";"

	if (strDestinatario <> "") then
		call GP_ENVIAR_MAIL_ATTACHMENT(strAsunto, strBody,de,vecMails(0), strPathAttachment)
		strSQL = "update controlEnvioRetencion set MrcEnvioMail = 'V' where RetNro=" & p_nroRet
		%>
		<td align=left style="left-padding:10px;">
			ha sido enviada a <%=strDestinatario%>
		</td>
		<%
	else
		strSQL = "update controlEnvioRetencion set MrcEnvioMail = 'N' where RetNro=" & p_nroRet
		%>
		<td align=left style="left-padding:10px;">No se ha podido enviar la retencion, debido a que no se han establecido las direcciones de mail donde enviarlos.</td>
		<%
	end if
    call GF_BD_AS400(rs, conn, "EXEC", strSQL)
end sub
'*****************************************************************************************
'Funcion que permite saber si el pago es por mercaderia (comodities) comprada.
Function esPagoMercaderia(p_minuta)
    Dim oConn, rs, strSQL  
    
    strSQL="Select * from PROVFL.VWACDSREP where DSQFNB=" & p_minuta
    Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
    
    esPagoMercaderia = false
    if (not rs.eof) then
        if (rs("DSD1ST") = "MR") then esPagoMercaderia = true
    end if
End function
'*****************************************************************************************
'Funcion que devuelve el titulo para la pantalla
Function getDescripcionRetencion(p_tipo)
    Dim oConn, rs, strSQL  
    
    strSQL="Select * from tblconceptopago where CDCONCEPTO = '" & p_tipo & "'"
    Call GF_BD_AS400(rs,oConn,"OPEN",strSQL)
    
    getDescripcionRetencion = "ERROR"
    if (not rs.eof) then getDescripcionRetencion = rs("DSCONCEPTO")
    
End Function
'*****************************************************************************************
'Funcion que arma el titulo del tipo de retencion para nombrar el archivo fisico.
function getTituloTipo(byval p_tipo)
    select case ucase(p_tipo)
        Case "C": getTituloTipo = "IVA 2854-10"						
        Case "E": getTituloTipo = "IVA 2300-07"						
        Case "B": getTituloTipo = "Imp Ganancias"
        Case "H": getTituloTipo = "IIBB Pcia Bs As"
        Case "D": getTituloTipo = "IIBB Pcia Santa Fe"
        Case "G", "J": getTituloTipo = "Derecho de Reg e Insp"
        Case "I": getTituloTipo = "IIBB Cdad de Bs As"
        Case "K": getTituloTipo = "Cont Patr SUSS 4052-95"			
        Case "L": getTituloTipo = "Cont Patr SUSS 1784-05"			
        Case "M": getTituloTipo = "Cont Patr SUSS 1556-03"
        Case "P": getTituloTipo = "Cont Patr SUSS 1769-04"
    end select
end function
'-------------------------------------------------------------------------------
function getDsProduct(p_KC)
    dim rsMG, conn, strSQL
    
    getDsProduct = "#ERROR KC Producto#"
    'strSQL = "select * from mg where mg_km='AR' and mg_kc='" & p_KC & "'"
    strSQL = "Select * from MERFL.MER112F1 where CODIPR=" & p_KC
    call GF_BD_AS400_2(rsMG, conn, "OPEN", strSQL)    
    if not rsMG.eof then
        getDsProduct = Trim(rsMG("DESCPR"))        
    end if
end function
%>
