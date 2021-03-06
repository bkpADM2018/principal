<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<%
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Se encarga de enviar mail al proximo firmante que tiene el lote, en caso de ser el ultimo envia a los primero autorizantes informando que se aplico
Function sendMailNextSignatory(pIdReasignacion, pIdObra)
    Dim rsFir, mailMsg, mailOrigen, mailDestino, mailAsunto
    'El sotre procedure devuelve el/los usuarios que deberan ser notificados por la alerta de mail de provisiones
    Call executeProcedureDb(DBSITE_SQL_INTRA, rsFir, "TBLBUDGETREASIGNACIONFIRMAS_GET_NEXT_SIGNATORY_BY_IDREASIGNACION", pIdReasignacion)
    if (not rsFir.Eof) then
        mailOrigen = getTaskMailList(TASK_COM_AUTH_REASSIGNING_BUDGET, MAIL_TASK_SENDER)
        mailAsunto = "Sistema Compras - Alerta de firma"
        mailMsg = "Tiene pendiente para autorizar el siguiente Ajuste de Partida Presupuestaria: "& vbcrlf
        mailMsg = mailMsg & "Numero: "& pIdReasignacion & vbcrlf
        mailMsg = mailMsg & "Obra: "& getDescripcionObra(pIdObra) & vbcrlf
        while(not rsFir.Eof)
            mailDestino = getUserMail(Trim(rsFir("CDUSUARIO")))
            Call GP_ENVIAR_MAIL(mailAsunto, mailMsg, mailOrigen, mailDestino)
            rsFir.MoveNext()
        wend
    end if
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getObrasBudget(pIdObra,pIdArea,pIdDetalle)
    Dim strSQL,rsBud
    strSQL = "SELECT CASE WHEN C.DLBUDGET IS NULL THEN 0 ELSE C.DLBUDGET END AS IMPORTEDOLARDETALLE,"&_
             "       CASE WHEN B.DSBUDGET IS NULL THEN '' ELSE B.DSBUDGET END AS DSBUDGETAREA,"&_
             "       CASE WHEN C.DSBUDGET IS NULL THEN '' ELSE C.DSBUDGET END AS DSBUDGETDETALLE, "&_
             "       D.DSDIVISION "&_
             "FROM TBLDATOSOBRAS A "&_
             "  LEFT JOIN TBLBUDGETOBRAS B ON A.IDOBRA = B.IDOBRA AND B.IDAREA = "& pIdArea &" AND B.IDDETALLE = 0 "&_
             "  LEFT JOIN TBLBUDGETOBRAS C ON A.IDOBRA = C.IDOBRA AND C.IDAREA = "& pIdArea &" AND C.IDDETALLE = "& pIdDetalle &_
             "  INNER JOIN TBLDIVISIONES D ON D.IDDIVISION = A.IDDIVISION  "&_
             "WHERE A.IDOBRA = "& pIdObra
    Call executeQueryDb(DBSITE_SQL_INTRA, rsBud, "OPEN", strSQL)
    Set getObrasBudget = rsBud
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function controlarAjustePartidaPresupuestaria(vlBudgetOld,vlBudget,pIdObra,pIdArea,pIdDetalle)
    Dim ret, gasto, totalReasignacion
    ret = false
    if (Cdbl(vlBudget) <> Cdbl(vlBudgetOld)) then
        if (((Cdbl(vlBudgetOld) + Cdbl(vlBudget)) < 0)and(Cdbl(vlBudget) < 0)) then 
            Call setError(IMPORTE_NO_EXISTE)
        else 
            'Calculo cuanto seria el total de la reasignacion con el nuevo importe para verificar que cubra lo gastado que ya tiene
            gasto = calcularGastosObra(MONEDA_DOLAR, pIdObra,pIdArea,pIdDetalle, false)
            'Quito los decimales que vienen en los gastos
            gasto = Cdbl(gasto)/100
            totalReasignacion = Cdbl(vlBudgetOld) + Cdbl(vlBudget)
            if ( Cdbl(totalReasignacion) >= cdbl(gasto) ) then
                ret = true
            else
                Call setError(IMPORTE_SUPERA_DISPONIBLE)    
            end if
        end if
    else
        Call setError(CTZ_AJU_IMP_IGUALES) 
    end if
    controlarAjustePartidaPresupuestaria = ret
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function grabarAjustePartidaPresupuestaria(idObra,idArea,idDetalle,vlBudget,observaciones)
    Dim importePesos, tipoCambio, rsIns
    tipoCambio = getTipoCambio(MONEDA_DOLAR, "")
    importePesos = (Cdbl(vlBudget)*100)*Cdbl(tipoCambio)    
    Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsIns, "TBLBUDGETREASIGNACION_INS", idObra &"||"& Left(Session("MmtoDato"),8) &"||0||0||"& idArea &"||"& idDetalle &"||"& importePesos &"||"& Cdbl(vlBudget)*100 &"||"& Replace(tipoCambio,".",",") &"||"& Session("Usuario") &"||"& Session("MmtoDato") &"||"& editText4DB(observaciones) &"||0$$IDREASIGNACION")
    grabarAjustePartidaPresupuestaria = sp_ret("IDREASIGNACION")
End function
'***********************************************************************************
'*******	                     COMIENZO DE LA PAGINA                      ********
'***********************************************************************************
dim accion, idArea, idDetalle, idObra ,vlBudgetOld, observaciones, vlBudget, rsBudObr, idAsignacion

guardado = false
idObra = GF_Parametros7("idObra",0,6)
idArea = GF_Parametros7("idArea",0,6)
idDetalle = GF_Parametros7("idDetalle",0,6)
accion = GF_PARAMETROS7("accion","",6)
vlBudget = GF_PARAMETROS7("vlBudget",0,6)
    
Set rsBudObr = getObrasBudget(idObra,idArea,idDetalle)

flagGrabar = false
if (isFormSubmit()) then
    vlBudgetOld = GF_PARAMETROS7("vlBudgetOld",0,6)
    observaciones = GF_PARAMETROS7("observaciones","",6)
    flagControl = controlarAjustePartidaPresupuestaria(vlBudgetOld,vlBudget,idObra,idArea,idDetalle)
    if ((flagControl)and(accion = ACCION_GRABAR)) then 
        idAsignacion = grabarAjustePartidaPresupuestaria(idObra,idArea,idDetalle,vlBudget,observaciones)
        Call sendMailNextSignatory(idAsignacion,idObra)
        flagGrabar = true
    end if
end if    

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title><% =GF_TRADUCIR("Sistema de Compras - Ajuste Partida Presupuestaria") %></title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">

<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript">
	
    
	function bodyOnLoad(){
		var tb = new Toolbar('toolbar', 4, 'images/');
		<% if not flagGrabar then %>
		    tb.addButton("compras/save-16x16.png", "<%=GF_Traducir("Guardar")%>", "submitInfo('<% =ACCION_GRABAR %>')");
		    tb.addButton("checkmark-16.png", "<%=GF_Traducir("Controlar")%>", "submitInfo('<% =ACCION_CONTROLAR %>')");
		    <% else %>
                document.getElementById("msjGrabar").style.display = "block";
		        document.getElementById("msjGrabar").innerHTML = "Se grabo correctamente el ajuste";
		        document.getElementById("msjGrabar").className = "confirmsj";
	    <% end if %>
		
		tb.draw();
	}
	function submitInfo(acc){
		document.getElementById("accion").value = acc;
		document.getElementById("frmSel").submit();
	}


</script>
</head>
<BODY onload="bodyOnLoad()">
<DIV id="toolbar"></DIV>
<form name="frmSel" id="frmSel" method=post action="comprasAjusteBudgetPopUp.asp">					
    <div class="tableasidecontent"><% call showErrors() %></div>
	        
    <% if not rsBudObr.Eof then %>
        <div class="tableasidecontent">
            <div id="msjGrabar" style="display:none;width:100%;height:15px;margin-bottom:10px"></div>
            
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Obra:") %> </div>        
            <div class="col46">
                <% Response.Write getDescripcionObra(idObra) %>
		    </div>
            <div class="col26 reg_header_navdos"  style="height:52px;margin-bottom:5px"> <% =GF_TRADUCIR("Detalle:") %> </div>        
            <div class="col46">
                <%= idArea &"-"& UCase(Trim(rsBudObr("DSBUDGETAREA"))) %>
		    </div>
            <div class="col46" style="float:right;">
                <%= idDetalle &"-"& UCase(Trim(rsBudObr("DSBUDGETDETALLE"))) %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Division:") %> </div>        
            <div class="col46">
                <%= rsBudObr("DSDIVISION") %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Gastos Vales:") %> </div>        
            <div class="col26">
                <% Set dicVales = obtenerTotalValesObraPorPPArea(idObra,idArea,MONEDA_DOLAR,session("MmtoSistema"))
                   gastoVale = 0
                   if dicVales.Exists(clng(idDetalle)) then gastoVale = dicVales(clng(idDetalle))/100
		           Response.Write TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(gastoVale)/100,0) %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Gastos Pedido:") %> </div>        
            <div class="col26">
                <% gastoPedido = calcularGastosObra(MONEDA_DOLAR, idObra,idArea,idDetalle, false)
                   Response.Write TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(gastoPedido)/100,0) %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Gastos Facturado:") %> </div>        
            <div class="col26">
                <% gastoFacturado = calcularGastosFacturados(idObra,idArea,idDetalle, "", "", MONEDA_DOLAR)
                   Response.Write TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(gastoFacturado)/100,0) %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Budget "& TIPO_MONEDA_DOLAR &":") %> </div>        
            <div class="col26">
                <% = TIPO_MONEDA_DOLAR &" "& GF_EDIT_DECIMALS(Cdbl(rsBudObr("IMPORTEDOLARDETALLE"))/100,0) %>
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Ajuste "& TIPO_MONEDA_DOLAR &":") %> </div>        
            <div class="col46">
                <input type="text" id="vlBudget" name="vlBudget" value="<%=vlBudget %>" onkeypress="return controlIngreso(this,event,'I');" style="text-align:right;"/>
                <input type="hidden" id="vlBudgetOld" name="vlBudgetOld" value="<%=cdbl(rsBudObr("IMPORTEDOLARDETALLE"))/100%>" />
		    </div>
            <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Observaciones:") %> </div>        
            <div class="col46">
                <textarea name="observaciones" id="observaciones" maxlength="1000" cols="83"><%=observaciones%></textarea>				
		    </div>
            <div class="col56"> </div>
            
        </div>
    <% end if %>
    <input type="hidden" name="accion" id="accion">
    <input type="hidden" name="idObra" id="idObra" value="<%= idObra %>">
    <input type="hidden" name="idArea" id="idArea" value="<% =idArea %>">
    <input type="hidden" name="idDetalle" id="idDetalle" value="<% =idDetalle %>">
</form>
</body>
</html>