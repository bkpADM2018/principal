<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<%

CONST SIN_DESTINO = -1

'--------------------------------------------------------------------------------------
Function obtenerRSCombo(pIdObra,pIdArea,pIdDetalle)
	Dim rsCombo,conn,strSQL
	
	strSQL = 		  " SELECT *"
	strSQL = strSQL & " FROM tblbudgetobras"
	strSQL = strSQL & " WHERE idobra   = " & pIdObra
	if (not flagCoordinador) then 
		strSQL = strSQL & " AND ((idArea =  "& pIdArea &")AND(idDetalle <> "& pIdDetalle &"))"
	else
		strSQL = strSQL & " AND ((idArea <> "& pIdArea &")OR (idDetalle <> "& pIdDetalle &"))"
	end if
	strSQL = strSQL & " ORDER BY idarea,iddetalle"
    Call executeQueryDb(DBSITE_SQL_INTRA, rsCombo, "OPEN", strSQL)
	
	if rsCombo.EoF then puedeReasignar = FALSE
	
	Set ObtenerRSCombo = rsCombo
End Function
'--------------------------------------------------------------------------------------
Function cargarCombo(pTrimestre,pIdArea,pRS)
	Dim seleccionado,antArea,seleccionadoArea
	pRS.MoveFirst
		
	if pRS.Eof then
		rtrn = false
	else
		rtrn = true
	end if
	seleccionado = SIN_DESTINO
	if (rtrn) then
		 %>
		<select name="<%=pTrimestre%>" id="combo<%=pTrimestre %>" onChange="seleccionarCombo(this)" style="width:200px">
		<option value="<%=SIN_DESTINO%>">-Seleccione-</option>
		<% 
			if (not reasignacionOK and accion = ACCION_GRABAR) then
				seleccionado = GF_Parametros7("combovalue"& pTrimestre,0,6)
				seleccionadoArea = GF_Parametros7("idArea_"& pTrimestre,0,6) 
			end if			
			while not pRS.EOF 
				if (CLng(pRS("IDAREA")) <> antArea) then
					antArea = CLng(pRS("IDAREA"))	%>
					<optgroup label="<%=pRS("IDAREA")%>-<%=pRS("DSBUDGET")%>"></optgroup>
			<%	else %>
					<option  value ="<%=pRS("IDDETALLE")%>"	alt="<%=pRS("IDAREA")%>" <% if ((Cdbl(seleccionado) = Cdbl(pRS("IDDETALLE")))and(Cdbl(seleccionadoArea) = Cdbl(pRS("IDAREA")))) then response.write "selected='true'" end if%> >
						&nbsp;&nbsp;<%=pRS("IDDETALLE") & " - " &  pRS("DSBUDGET")%>
					</option>
			<%	end if
				pRS.MoveNext
			wend%>
		</select>
		<input type="hidden" name="combovalue<%=pTrimestre%>" id="combovalue<%=pTrimestre%>" value="<%=seleccionado%>">
		<input type="hidden" name="idArea_<%=pTrimestre%>" id="idArea_<%=pTrimestre%>">		
	<%end if
	cargarCombo = rtrn
End Function
'---------------------------------------------------------------------------------------
Function controlarReasignacion(pIdObra, ByRef pRs)
		Dim rtrn,valor,cmbPartida ,strSQL,conn,gasto

		rtrn = true
		while ((not pRs.EOF)and(rtrn))
            valor = 0
            'Controlo solo aquellas que tienen AREA-DETALLE
			if (pRs("IDDETALLE") <> 0) then
		        valor = GF_Parametros7("valor" & pRs("periodo"), 0, 6)		
                valor = Cdbl(valor) * 100
                'Al no generar reasignaciones trimestrales se debera si o si ingresar un monto en la reasignacion que esta haciendo
                if (cdbl(valor) > 0) then
                    cmbPartida = GF_Parametros7("combovalue" & pRs("periodo"), "", 6)
					'Controlo que haya elegido una partida origen para reasignar el importe
                    if (cmbPartida <> cstr(SIN_DESTINO)) then
                        'Controlo que el nuevo importe a reasignar no sobrepase al importe que tenia antes la partida
                        if (Cdbl(valor) <= cdbl(pRs("DLBUDGET"))) then
						    'Obtengo el gasto que tiene la partida y controlo que el nuevo importe a reasignar sea inferior o igual al saldo de lo pagado en la partida
                            gasto = calcularGastosObra(MONEDA_DOLAR, pIdObra,pRs("IDAREA"),pRs("IDDETALLE"), false)
                            'el nuevo saldo se obtendra del importe que tenia la partida menos el gasto de la partida hasta el momento, el resultado tiene que ser menor o igual al nuevo importe
                            sadloDolares = cdbl(pRs("DLBUDGET")) - Cdbl(gasto)							
                            if ( cdbl(valor) > Cdbl(sadloDolares) ) then Call setError(IMPORTE_SUPERA_DISPONIBLE)
                        else
                            Call setError(FALTA_DETALLE_DESTINO)
                        end if
					else
                        Call setError(FALTA_DETALLE_DESTINO)
					end if
				else
                    Call setError(IMPORTE_NO_EXISTE)
                end if
			end if
            if (hayError()) then rtrn = false
			pRs.movenext()
		wend
		controlarReasignacion = rtrn
End Function
'---------------------------------------------------------------------------------------
	Function grabarReasignacion(pIdObra,ByRef pRs)
		Dim rtrn
		Dim strSQL,rs2,conn
		Dim importePesos,importeDolares,detalleDestino
		Dim nuevoTotalDolares
		
		pRs.MoveFirst
		
		while not pRs.EOF

			if (pRs("IDDETALLE") <> 0 AND GF_Parametros7("valor"&pRs("PERIODO"),0,6) <> 0) then
				importeDolares = GF_Parametros7("valor"&pRs("PERIODO") ,0 ,6)*100
				importePesos   = cdbl(importeDolares)*cdbl(pRs("TIPOCAMBIO"))
				motivo		   = GF_Parametros7("motivo"&pRs("PERIODO"),"",6)
				detalleDestino = GF_Parametros7("combovalue"&pRs("PERIODO"),"",6)
				areaDestino	   = GF_Parametros7("idArea_"&pRs("PERIODO"),"",6)
				tipocambio     = Replace(Cstr(pRs("TIPOCAMBIO")),".",",")

                Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rsIns, "TBLBUDGETREASIGNACION_INS", pIdObra &"||"& left(Session("MmtoSistema"),8) &"||"& pRs("IDAREA") &"||"& pRs("IDDETALLE") &"||"& areaDestino &"||"& detalleDestino &"||"& importePesos &"||"& importeDolares &"||"& tipocambio &"||"& Session("Usuario") &"||"& Session("MmtoDato") &"||"& motivo &"||"& pRs("PERIODO") &"$$IDREASIGNACION")
                auxIdReasignacion = sp_ret("IDREASIGNACION")
        
                Call sendMailNextSignatory(auxIdReasignacion, pIdObra)

            end if                    
			pRs.MoveNext()
		wend
		pRs.MoveFirst
	End Function
	
'---------------------------------------------------------------------------------------
'Function isCoordinadorPuerto : verifica si el usuario tien rol de Coordinador de Puerto. 
'					True  = Si
'					False = No
Function isCoordinadorPuerto()
	Dim strSQL 
	isCoordinadorPuerto =  false
	idRol = getRolFirma(Session("Usuario"), SEC_SYS_COMPRAS)
	if (idRol = FIRMA_ROL_SUP_PUERTO) then isCoordinadorPuerto = true
End Function
'---------------------------------------------------------------------------------------
'Se encarga de enviar mail al proximo firmante que tiene el lote, en caso de ser el ultimo envia a los primero autorizantes informando que se aplico
Function sendMailNextSignatory(pIdReasignacion, pIdObra)
    Dim rsFir, mailMsg, mailOrigen, mailDestino, mailAsunto
    'El sotre procedure devuelve el/los usuarios que deberan ser notificados por la alerta de mail de provisiones
    Call executeProcedureDb(DBSITE_SQL_INTRA, rsFir, "TBLBUDGETREASIGNACIONFIRMAS_GET_NEXT_SIGNATORY_BY_IDREASIGNACION", pIdReasignacion)
    if (not rsFir.Eof) then
        mailOrigen = getTaskMailList(TASK_COM_AUTH_REASSIGNING_BUDGET, MAIL_TASK_SENDER)
        mailAsunto = "Sistema Compras - Alerta de firma"
        mailMsg = "Tiene pendiente para autorizar la siguiente Reasignacion de Partida Presupuestaria: "& vbcrlf
        mailMsg = mailMsg & "Numero: "& pIdReasignacion & vbcrlf
        mailMsg = mailMsg & "Obra: "& getDescripcionObra(pIdObra) & vbcrlf
        while(not rsFir.Eof)
            mailDestino = getUserMail(Trim(rsFir("CDUSUARIO")))
            Call GP_ENVIAR_MAIL(mailAsunto, mailMsg, mailOrigen, mailDestino)
            rsFir.MoveNext()
        wend
    end if
End Function
'**************************************************************************************************
'*                                                                                                *
'*                                   INICIO DE PAGINA                                             *
'*                                                                                                *
'**************************************************************************************************
Dim idArea,idDetalle,idObra,puedeReasignar,rs2,totalreg
Dim strSQL, rsGral, conn,rsCombo,importeTrim(4),motivoTrim(4),reasignacinOK, flagCoordinador

idObra 		= GF_Parametros7("idObra",0 ,6)
idArea 		= GF_Parametros7("idarea",0 ,6)
idDetalle 	= GF_Parametros7("iddetalle",0 ,6)
accion 		= GF_Parametros7("accion","" ,6)

importeTrim(0) = GF_Parametros7("valor0",0 ,6)
importeTrim(1) = GF_Parametros7("valor1",0 ,6)
importeTrim(2) = GF_Parametros7("valor2",0 ,6)
importeTrim(3) = GF_Parametros7("valor3",0 ,6)

motivoTrim(0) = GF_Parametros7("motivo0","" ,6)
motivoTrim(1) = GF_Parametros7("motivo1","" ,6)
motivoTrim(2) = GF_Parametros7("motivo2","" ,6)
motivoTrim(3) = GF_Parametros7("motivo3","" ,6)

flagCoordinador = isCoordinadorPuerto()

strSQL = "SELECT A.IDOBRA, "&_
	     "       A.IDAREA, "&_
	     "       A.IDDETALLE, "&_
	     "       A.DSBUDGET, "&_
         "       CASE WHEN B.TIPOCAMBIO IS NULL THEN A.TIPOCAMBIO ELSE B.TIPOCAMBIO END AS TIPOCAMBIO, "&_
	     "       CASE WHEN B.PERIODO IS NULL THEN 0 ELSE B.PERIODO END AS PERIODO, "&_
	     "       CASE WHEN B.PERIODO IS NULL THEN A.DLBUDGET ELSE B.DLBUDGET END AS DLBUDGET "&_
         "FROM TBLBUDGETOBRAS A "&_
         "LEFT JOIN TBLBUDGETOBRASDETALLE B ON A.IDOBRA = B.IDOBRA AND A.IDAREA = B.IDAREA AND A.IDDETALLE = B.IDDETALLE "&_
         "WHERE A.IDOBRA = "& idobra &" AND A.IDAREA = "& idarea &" AND A.IDDETALLE = "& iddetalle 
Call executeQueryDb(DBSITE_SQL_INTRA, rsGral, "OPEN", strSQL)

reasignacionOK = false
if (accion=ACCION_GRABAR) then
	controlOK = controlarReasignacion(idObra, rsGral)
	rsGral.MoveFirst
	if (controlOK) then
		Call grabarReasignacion(IdObra,rsGral)
		reasignacionOK = TRUE
	end if
end if


puedeReasignar = TRUE
Set rsCombo = obtenerRSCombo(idObra,idArea,idDetalle)

%>

<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript">
			var popUpReasignar;
			function bodyOnLoad(){
				popUpReasignar = getObjPopUp('popUpReasignaciones');
				<% if (reasignacionOK) then %>
					popUpReasignar.hide();
				<% end if %>
				if (document.getElementById("valor0")) {
					document.getElementById("valor0").focus();
				}
			}
			function seleccionarCombo(me){
				var nombre = 'combovalue'+me.name;
				document.getElementById(nombre).value = me.value;
				document.getElementById("idArea_" + me.name).value = $("#combo"+ me.name +" option:selected").attr("alt");				
			}
		</script>
	</head>
	<body onLoad="bodyOnLoad();">
		<% if (puedeReasignar) then %>
		<form id="myForm" method="GET" action="comprasBudgetPopUpReasignacion.asp">
			<input type="hidden" id="idobra" name="idobra" value="<%=idobra%>">
			<input type="hidden" id="idarea" name="idarea" value="<%=idarea%>">
			<input type="hidden" id="iddetalle" name="iddetalle" value="<%=iddetalle%>">
			<input type="hidden" name="accion" id="accion" value="<%=ACCION_GRABAR%>">
			
			<table width="620" align="center" class="reg_header">
			<tr>
					<td width="612"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
					  <tr>
						<td colspan="3" class="reg_header_nav round_border_top" align="center" style="font-size:15px" ><%=getDescripcionObra(idObra)%><hr></td>
					  </tr>

					  <tr>
						<td class="reg_header_nav round_border_bottom_left" align="right" width="45%"><%=idArea%></td>
						<td class="reg_header_nav round" width="10%" align="center">-</td>
						<td class="reg_header_nav round_border_bottom_right" align="left" width="45%"><%=idDetalle%></td>
					  </tr>
					</table></td>
			</tr>
			<tr>
				<td>
					<% 	if (hayError()) then Call showErrors() 
						if (hayError() = false AND accion = ACCION_GRABAR) then 
							reasignacionOK = true
						else
							response.write "&nbsp;"
						end if
					%>
				</td>
			</tr>
			<tr>
				<td>
				  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="3" class="reg_header">
					<tr>
					  <td width="98"  align="center" class="reg_header_nav round_border_top_left">Periodo</td>
					  <td width="122" align="center" class="reg_header_nav">Disponible</td>
					  <td width="253" align="center" class="reg_header_nav">Detalle</td>
					  <td width="122" align="center" class="reg_header_nav round_border_top_right">Importe (USD)</td>
					</tr>
					<% 
					i = 0
					while not rsGral.EoF %>
							<tr>
							  <td width="98" class="reg_header_nav" align="center">Trimestre <%=cdbl(rsGral("periodo"))+1%></td>
							  <%
							  gasto = calcularGastosObra(MONEDA_DOLAR, rsGral("IDOBRA"),rsGral("IDAREA"),rsGral("IDDETALLE"), false)
								'el nuevo saldo se obtendra del importe que tenia la partida menos el gasto de la partida hasta el momento, el resultado tiene que ser menor o igual al nuevo importe
								saldoDolares = cdbl(rsGral("DLBUDGET")) - Cdbl(gasto)	
								%>
							  <td width="122" align="center"><input type="text" readonly style="text-align:right;background-color:#DDD" value="<%= GF_EDIT_DECIMALS(saldoDolares,2)%>"></td>
							  <td width="253" align="center"><% Call cargarCombo(i,rsGral("IDAREA"), rsCombo) %></td>
							  <td width="122" align="center"><input type="text" style="text-align:right;" id="valor<%=i%>" name="valor<%=i%>" value="<%=importeTrim(i)%>"></td>
							</tr>
						<%if (puedeReasignar) then%>
							<tr>
							  <td>&nbsp;</td>
							  <td colspan="3" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
								  <td width="15%" align="center" class="reg_header_navdos">Motivo</td>
								  <td width="85%" align="center"><input type="text" id="motivo<%=i%>" name="motivo<%=i%>" size="75" value="<%=motivoTrim(i)%>"></td>
								</tr>
							  </table></td>
							</tr>
						
						<%end if%>
							<tr><td colspan=4>&nbsp;</td></tr>
					<% 
						i =  i +1
						rsGral.MoveNext
					wend 
					if (i = 0) then %>
						<tr>
							<td width="100%" align="center" colspan="4">
								No hay presupuestos Asignados
							</td>
						</tr>
					<%end if%>
				  </table>			      
				</td>
			</tr>
			<tr>
				<td colspan=4 align="right">
				<% if (i > 0) then %>
					<input class="round_border_bottom_right" type="submit" value="Aceptar">
				<%end if%>
				</td>
			</tr>
			</table>
		</form>
		<%else%>
			<table width="620" align="center" class="reg_header">
				<tr>
					<td width="100%" align="center">
						Solo existe 1 detalle, por lo que no es posible realizar reasignaciones
					</td>
				</tr>
			</table>
		<%end if%>
</body>
</html>