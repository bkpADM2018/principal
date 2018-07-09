<!--#include file="../Includes/procedimientosMG.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<%
Const MUESTRAS_AUDITORIA_ONLY = 1

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
'------------------------------------------------------
Function crearTabla(p_mostrar, p_accion)
    Dim reg, pesoNeto, cuitTitular, dsTitular, cdIntermediario, dsIntermediario, cdRteCOmercial, dsRteComercial
%>
<table class="datagrid" width="100%" align="center">	
	<thead>
		<tr>			
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DTCONTABLE DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Fecha")%>
			    <% if (p_accion = "") then %>			    
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DTCONTABLE ASC')">
			    <% end if %>
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUCARTAPORTE DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Carta Porte")%>		
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUCARTAPORTE ASC')">
			    <% end if %>
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSPRODUCTO DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Producto")%>		
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSPRODUCTO ASC')">
			    <% end if %>
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSCLIENTE DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Cliente")%>			
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSCLIENTE ASC')">
			    <% end if %>
			    </td>	
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSCORREDOR DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Corredor")%>		
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSCORREDOR ASC')">
			    <% end if %>
			    </th>		
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSVENDEDOR DESC')">		
			    <% end if %>
			    <%=GF_Traducir("Vendedor")%>		
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSVENDEDOR ASC')">
			    <% end if %>
			    </th>
			<th>
			    <%=GF_Traducir("Kg Netos")%>					    
			    </th>			    
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY CDCHAPACAMION DESC')">	
			    <% end if %>
			    <%=GF_Traducir("Chasis")%>			
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY CDCHAPACAMION ASC')">
			    <% end if %>
			    </th>
			<th>
			    <%=GF_Traducir("Nro. Sticker")%>		
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSESTADO DESC')">			
			    <% end if %>
			    <%=GF_Traducir("Estado")%>			
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY DSESTADO ASC')">
			    <% end if %>
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUCUPO DESC')">			
			    <% end if %>
			    <%=GF_Traducir("Cod. Cupo")%>		
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUCUPO ASC')">
			    <% end if %>
			    </th>
			<th>
			   <% if (p_accion = "") then %>
			    <img src="images/arrow_down_12x12.gif" title="<%=GF_Traducir("Descendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUAUTSALIDA DESC')">		
			    <% end if %>
			    <%=GF_Traducir("N.Aceptacion")%>	
			    <% if (p_accion = "") then %>
			    <img src="images/arrow_up_12x12.gif" title="<%=GF_Traducir("Ascendiente")%>" style="cursor:pointer;" onclick="setSortBy('ORDER BY NUAUTSALIDA ASC')">
			    <% end if %>
			    </th>			
			    <% if (p_accion = "") then %>
					<th>.</th>
					<th>.</th>
					<th>.</th>
			   <% else %> 
					<th>Titular</th>
					<th>Intermediario</th>
					<th>Rte. Comercial</th>
					<th>Deposito</th>
			   <% end if %>  
		</tr>
	</thead>
	<tbody>
		<% if rsLista.eof then %>
			<tr>
				<td colspan="15" align="center"><%=GF_Traducir("No se encotraron resultados")%></td>
			</tr>	
		<% else 
				while not rsLista.eof  and (reg < p_mostrar) 
					reg = reg + 1
			        dtContable = rsLista("DTCONTABLE")	
			        camionSalio = isEstadoTerminal(rsLista("CIRCUITO"), rsLista("CDESTADO"))							        
			        pesoNeto = 0
			        if (not isNull(rsLista("PESO"))) then pesoNeto = rsLista("PESO")
					%>
					<tr>						
						<td align="center"><%=GF_FN2DTE(rsLista("DTCONTABLE")) %></td>
						<td  align="center"><%=GF_EDIT_CBTE(rsLista("NUCARTAPORTE"))%></td>
						<td  align="center"><%=rsLista("CDPRODUCTO") & " - " & rsLista("DSPRODUCTO")%></td>		
						<td  ><%=rsLista("CDCLIENTE") & "-" & rsLista("DSCLIENTE")%></td>
						<td  ><%=rsLista("CDCORREDOR") & "-" & rsLista("DSCORREDOR")%></td>
						<td  ><%=rsLista("CDVENDEDOR") & "-" & rsLista("DSVENDEDOR")%></td>
						<td  align="center"><% =pesoNeto %></td>
						<td  align="center"><%=left(rsLista("CDCHAPACAMION"),3) & "-" & right(trim(rsLista("CDCHAPACAMION")),3)%></td>
						<td  align="center"><%=rsLista("NUBARRAS")%></td>
						<td  align="center"><%=rsLista("DSESTADO")%></td>
						<td  align="center"><%=rsLista("NUCUPO")%></td>
						<td  align="center"><%=rsLista("NUAUTSALIDA")%></td>						
						<%if (p_accion = "") then   %>
						<td align="center" >
						<%    if (not isNull(rsLista("SQCALADA"))) then		
						    %>
    						
						        <img style="width:20;height:20;cursor:pointer" src="../images/analisis-16.png" onclick="abrirInfoAnalisis('<%=rsLista("idCamion")%>', '<% =dtContable %>', '<% =rsLista("NUCARTAPORTE") %>');" title="Ver analisis del Camion"/>
						    <%else%>
						        <img style="width:20;height:20" src="../images/analisis-16b.png" title="Sin analisis"/>
						    <%end if    %>
						</td>												
						<td  align="center">
						<%  'Si es un estado terminal (dependiendo el circuito del camion) y tiene kilos netos se supone que termino OK el ciruito
                            if ((camionSalio)and(Cdbl(pesoNeto) > 0)) then %>
							<img src="../images/pdf-16.png" title="Ver Nota de recepci�n" onclick="abrirNotaRecepcion('<%=Trim(rsLista("IDCAMION"))%>','<%=Trim(rsLista("NUCARTAPORTE"))%>','<%=dtContable%>',<%=rsLista("CIRCUITO") %>)" style="cursor:pointer;">
					    <%  end if %>
						</td>
                        <td  align="center">
                        <% if (isToepfer(session("KCOrganizacion"))) then %>
                            <%  if ((camionSalio)and(Cdbl(pesoNeto) > 0)and(Cdbl(rsLista("CIRCUITO")) = CIRCUITO_CAMION_DESCARGA)) then %>
                                <img src="../images/edit-2-16.png" title="Editar Carta de Porte" onclick="editCtaPte('<%=Trim(rsLista("NUCARTAPORTE"))%>','<%=dtContable%>','<% =rsLista("IDCAMION") %>')" style="cursor:pointer;">
                            <%  end if %>
                        <%  end if %>
                        </td>
                        <% else 
                            Call leerPartesIntervinientes(rsLista("DTCONTABLE"), rsLista("IDCAMION"), cuitTitular, dsTitular, cdIntermediario, dsIntermediario, cdRteCOmercial, dsRteComercial)                            
                        %>                            
                            <td  ><% =cuitTitular & dsTitular %></td>
						    <td  ><% =cdIntermediario & " - " & dsIntermediario %></td>
						    <td  ><% =cdRteComercial & " - " & dsRteComercial%></td>
						    <td  ><% =rsLista("DSSILO") %></td>
                        <% end if    %>
					</tr>
					<% 
					rsLista.movenext
				wend %>
		<% end if %>		
	</tbody>
	<tfoot>
  		<td colspan="12"><div id="paginacion"></div></td>
  	</tfoot>	
</table>
<%
End Function
'-------------------------------------------------------------------------------------------------------------
Function recuperarPartesIntervinientes(fnDesde, fnHasta)
    strSQL= "Select A.DTCONTABLE, A.IDCAMION, A.NUCUITREM, VTITULAR.DSVENDEDOR DSTITULAR, A.INTERMEDIARIO, VINT.DSVENDEDOR DSINTERMEDIARIO, A.RTECOMERCIAL, VRC.DSVENDEDOR DSRTECOMERCIAL " &_
            "from (" &_
		    "        (" &_
			"            Select (YEAR(HC.DtContable)*10000 + Month(HC.DtContable)*100 + DAY(HC.DtContable)) DTCONTABLE, HC.IDCAMION, HC.NUCUITREM, INT.CDVENDEDOR INTERMEDIARIO, RC.CDVENDEDOR RTECOMERCIAL " &_
			"            from HCAMIONES HC " &_
			"            left join (Select * from HCUENTAYORDENESCAMIONES where SQORDEN=1) INT on HC.IDCAMION=INT.IDCAMION and HC.DTCONTABLE=INT.DTCONTABLE " &_
			"            left join (Select * from HCUENTAYORDENESCAMIONES where SQORDEN=2) RC on HC.IDCAMION=RC.IDCAMION and HC.DTCONTABLE=RC.DTCONTABLE" &_
		    "        ) UNION (" &_
			"            Select  " & myHoy & " DTCONTABLE, HC.IDCAMION, HC.NUCUITREM, INT.CDVENDEDOR INTERMEDIARIO, RC.CDVENDEDOR RTECOMERCIAL " &_
			"            from CAMIONES HC " &_
			"            left join (Select * from CUENTAYORDENESCAMIONES where SQORDEN=1) INT on HC.IDCAMION=INT.IDCAMION " &_
			"            left join (Select * from CUENTAYORDENESCAMIONES where SQORDEN=2) RC on HC.IDCAMION=RC.IDCAMION " &_
			"            ) " &_
		    "        ) A " &_
		    "        left join VENDEDORES VINT on VINT.CDVENDEDOR=A.INTERMEDIARIO " &_
		    "        left join VENDEDORES VRC on VRC.CDVENDEDOR=A.RTECOMERCIAL " &_
		    "        left join VENDEDORES VTITULAR on VTITULAR.NUDOCUMENTO=A.NUCUITREM " &_
            " where A.DTCONTABLE>=" & fnDesde & " and A.DTCONTABLE<=" & fnHasta            
    
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN",strSql)
    while (not rs.eof)
        session(g_strPuerto & "_" & rs("DTCONTABLE") & "_" & rs("IDCAMION")) = rs("NUCUITREM") & "|" & rs("DSTITULAR") & "|" & rs("INTERMEDIARIO") & "|" & rs("DSINTERMEDIARIO") & "|" & rs("RTECOMERCIAL") & "|" & rs("DSRTECOMERCIAL")
        rs.MoveNext()
    wend        
End Function
'-------------------------------------------------------------------------------------------------------------
Function leerPartesIntervinientes(dtContable, idCamion, ByRef cuitTitular, ByRef dsTitular, ByRef cdIntermediario, ByRef dsIntermediario, ByRef cdRteComercial, ByRef dsRteComercial)
    Dim rs, strSQL
    Dim datos
        
    cuitTitular = ""
    dsTitular = ""
    cdIntermediario = ""
    dsIntermediario = ""
    cdRteComercial = ""
    dsRteComercial = ""    
    if (session(g_strPuerto & "_" & dtContable & "_" & idCamion) <> "") then        
        'Hay datos grabados!
        datos = split(session(g_strPuerto & "_" & dtContable & "_" & idCamion), "|")
        cuitTitular = datos(0)
        dsTitular = datos(1)
        if (dsTitular <> "") then dsTitular = " - " & dsTitular
        'Para las cartas de porte viejas, si se cargo el CO1 (Intermediario) y no el CO2 (Rte Comercial) implica que el CO1 es en realidad el Rte Comercial.
        if (datos(4) = "") then
            cdRteComercial = datos(2)
            dsRteComercial = datos(3)
        else
            cdIntermediario = datos(2)
            dsIntermediario = datos(3)
            cdRteComercial = datos(4)
            dsRteComercial = datos(5)
        end if
    end if
End Function           
'------------------------------------------------------
'	COMIENZO PAGINA
'------------------------------------------------------
	dim strSQL, rsLista
	Dim pto, params, sortBy, accion
	pto = GF_PARAMETROS7("pto", "", 6)
	g_strPuerto = pto
	call addParam("pto", pto, params)
	dim idCamion, fecContableH, fecContableHD, fecContableHM, fecContableHA, fecContable, fecContableD, fecContableM, fecContableA, codCupo, nuCartaPorte, nuCartaPorte1, nuCartaPorte2, patChasis, patChasis1, patChasis2, patAcoplado, patAcoplado1, patAcoplado2
	dim cdCliente, dsCliente, cdCorredor, dsCorredor, cdVendedor, dsVendedor, cdEstado, cdProducto, cdChofer, cdTransportista, myHoy, cdCircuito, chkMuestrasAud
	'Armar filtros para la pagina
	
	accion = GF_PARAMETROS7("accion", "", 6)	
	
	'Id Camion
	idCamion = GF_PARAMETROS7("idCamion", "", 6)	
	if (idCamion <> "") then 
		idCamion = GF_nDigits(idCamion, 10)
		Call mkWhere(myWhere, "IDCAMION", idCamion, "=", 0)
		call addParam("idCamion", idCamion, params)
	end if
	call addParam("idCamion", idCamion, params)
	'Fecha Contable
	fecContableD = GF_PARAMETROS7("fecContableD", "", 6)
	if (fecContableD = "") then fecContableD=Day(Now())
	call addParam("fecContableD", fecContableD, params)

	fecContableM = GF_PARAMETROS7("fecContableM", "", 6)
	if (fecContableM = "") then fecContableM=Month(Now())
	Call addParam("fecContableM", fecContableM, params)

	fecContableA = GF_PARAMETROS7("fecContableA", "", 6)
	if (fecContableA = "") then fecContableA=Year(Now())
	Call addParam("fecContableA", fecContableA, params)


	fecContableHD = GF_PARAMETROS7("fecContableHD", "", 6)
	if (fecContableHD = "") then fecContableHD=Day(Now())
	call addParam("fecContableHD", fecContableHD, params)

	fecContableHM = GF_PARAMETROS7("fecContableHM", "", 6)
	if (fecContableHM = "") then fecContableHM=Month(Now())
	call addParam("fecContableHM", fecContableHM, params)

	fecContableHA = GF_PARAMETROS7("fecContableHA", "", 6)
	if (fecContableHA = "") then fecContableHA=Year(Now())
	call addParam("fecContableHA", fecContableHA, params)
	
	'Codigo de cupo
	codCupo = GF_PARAMETROS7("codCupo", "", 6)
	call addParam("codCupo", codCupo, params)
	if (codCupo <> "") then 
		if (len(codCupo) < 9) then 
			Call mkWhere(myWhere, "NUCUPO", "%" & ucase(codCupo) & "%", "LIKE", 0)
		else
			Call mkWhere(myWhere, "NUCUPO", ucase(codCupo), "=", 0)
		end if	
	end if	
	
	'Carta de Porte
	nuCartaPorte1 = GF_PARAMETROS7("nuCartaPorte1", "", 6)
	call addParam("nuCartaPorte1", nuCartaPorte1, params)
	nuCartaPorte2 = GF_PARAMETROS7("nuCartaPorte2", "", 6)
	call addParam("nuCartaPorte2", nuCartaPorte2, params)
	
	if ((nuCartaPorte1 = "" and nuCartaPorte2 <> "") or (nuCartaPorte1 <> "" and nuCartaPorte2 = "")) then
		Call setError(NRO_CTA_PTE_INCOMPLETO)
	else
		if (nuCartaPorte1 <> "" and nuCartaPorte2 <> "") then
			nuCartaPorte = nuCartaPorte1 & nuCartaPorte2 
		end if
	end if	
	if (nuCartaPorte <> "") then Call mkWhere(myWhere, "NUCARTAPORTE", nuCartaPorte, "=", 0)
	
	'Patente chasis
	patChasis1 = GF_PARAMETROS7("patChasis1", "", 6)
	call addParam("patChasis1", patChasis1, params)
	patChasis2 = GF_PARAMETROS7("patChasis2", "", 6)
	call addParam("patChasis2", patChasis2, params)
	if (patChasis1 = "" xor patChasis2 = "") then
		Call setError(NRO_CTA_PTE_INCOMPLETO)
	else
		if (patChasis1 <> "" and patChasis2 <> "") then patChasis = patChasis1 & patChasis2 
	end if	
	if (patChasis <> "") then Call mkWhere(myWhere, "CDCHAPACAMION", ucase(patChasis), "=", 0)
	
	'Patente aomplado
	patAcoplado1 = GF_PARAMETROS7("patAcoplado1", "", 6)
	call addParam("patAcoplado1", patAcoplado1, params)
	patAcoplado2 = GF_PARAMETROS7("patAcoplado2", "", 6)
	call addParam("patAcoplado2", patAcoplado2, params)
	if (patAcoplado1 = "" xor patAcoplado2 = "") then
		Call setError(NRO_CTA_PTE_INCOMPLETO)
	else
		if (patAcoplado1 <> "" and patAcoplado2 <> "") then patAcoplado = patAcoplado1 & patAcoplado2 
	end if	
	if (patAcoplado <> "") then Call mkWhere(myWhere, "CDCHAPAACOPLADO", ucase(patAcoplado), "=", 0)

	'Estado
	cdEstado = GF_PARAMETROS7("cdEstado", 0, 6)
	call addParam("cdEstado", cdEstado, params)
	if cdEstado <> 0 then Call mkWhere(myWhere, "TABLA.CDESTADO", cdEstado, "=", 1)
	chkMuestrasAud = GF_PARAMETROS7("chkMuestrasAud", 0, 6)
	call addParam("chkMuestrasAud", chkMuestrasAud, params)	
	'Producto
	cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)
	call addParam("cdProducto", cdProducto, params)
	if cdProducto <> 0 then Call mkWhere(myWhere, "TABLA.CDPRODUCTO", cdProducto, "=", 1)
	
	'Chofer
	cdChofer = GF_PARAMETROS7("cdChofer", "", 6)
	call addParam("cdChofer", cdChofer, params)
	dsChofer = GF_PARAMETROS7("dsChofer", "", 6)
	call addParam("dsChofer", dsChofer, params)
	'NO SE ESTA UTILIZANDO!!
	'if cdChofer <> "" then Call mkWhere(myWhere, "NUDOCUMENTO", cdChofer, "=", 1)
	'Transportista
	cdTransportista = GF_PARAMETROS7("cdTransportista", "", 6)
	call addParam("cdTransportista", cdTransportista, params)
	dsTransportista = GF_PARAMETROS7("dsTransportista", "", 6)
	call addParam("dsTransportista", dsTransportista, params)
	if cdTransportista <> "" then Call mkWhere(myWhere, "CDTRANSPORTISTA", cdTransportista, "=", 1)
	
	if ((fecContableD <> "") or (fecContableHD <> "") or (fecContableM <> "") or (fecContableHM <> "") or (fecContableA <> "") or ( fecContableHA <> "")) then
		ret = GF_CONTROL_PERIODO(fecContableD, fecContableHD, fecContableM, fecContableHM, fecContableA, fecContableHA)
		Select case (ret)
			case 0
				'Si el control resulto exitoso
				fecContable  = fecContableA & fecContableM & fecContableD
				fecContableH = fecContableHA & fecContableHM & fecContableHD
				Call mkWhere(myWhere, "DTCONTABLE", fecContable, ">=", 1)
				Call mkWhere(myWhere, "DTCONTABLE", fecContableH, "<=", 1)
			case 1
				Call setError(FECHA_INICIO_INCORRECTA)
			case 2
				Call setError(FECHA_FIN_INCORRECTA)
			case 3
				Call setError(PERIODO_ERRONEO)
		end select
	end if
	
	'Corredor o Vendedor o Cliente: si la empresa no es Toepfer, filtro por los tres campos. Siempre va a ver un campo (Dtcontable) que filtra antes 
	if (not IsToepfer(session("KCOrganizacion"))) then		 
		myWhere = myWhere & "	  AND ( TABLA.cdcliente in (Select CDCLIENTE from clientes where NUCUIT = '" & session("CuitOrganizacion") & "') "
		myWhere = myWhere & "     OR TABLA.cdvendedor in (Select CDVENDEDOR from VENDEDORES where NUDOCUMENTO = '" & session("CuitOrganizacion") & "') "
		myWhere = myWhere & "     OR TABLA.cdcorredor in (Select CDCORREDOR from CORREDORES where NUCUIT = '" & session("CuitOrganizacion") & "') )"	
	else    	
        'Cliente
		cdCliente = GF_PARAMETROS7("cdCliente", "", 6)
		call addParam("cdCliente", cdCliente, params)
		dsCliente = GF_PARAMETROS7("dsCliente", "", 6)
		call addParam("dsCliente", dsCliente, params)
        'Corredor
		cdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
		call addParam("cdCorredor", cdCorredor, params)
		dsCorredor = GF_PARAMETROS7("dsCorredor", "", 6)
		call addParam("dsCorredor", dsCorredor, params)
        'Vendedor
		cdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
		call addParam("cdVendedor", cdVendedor, params)
		dsVendedor = GF_PARAMETROS7("dsVendedor", "", 6)
		call addParam("dsVendedor", dsVendedor, params)
        if cdCliente <> "" then 
            Call mkWhere(myWhere, "TABLA.CDCLIENTE", cdCliente, "=", 1)
            dsCliente = getDsCliente(cdCliente)
        end if
		if cdCorredor <> "" then Call mkWhere(myWhere, "TABLA.CDCORREDOR", cdCorredor, "=", 1)
		if cdVendedor <> "" then Call mkWhere(myWhere, "TABLA.CDVENDEDOR", cdVendedor, "=", 1)
    end if
    
	cdCircuito = GF_PARAMETROS7("cdCircuito", 0, 6)
	if (cdCircuito = 0) then cdCircuito = CIRCUITO_CAMION_DESCARGA
	Call addParam("cdCircuito", cdCircuito, params)
	    
    
	if not hayError() then 
	    
	    myHoyD = Day(Now())
	    myHoyM = Month(Now())
	    myHoyY = Year(Now())
	    Call GF_STANDARIZAR_FECHA(myHoyD, myHoyM, myHoyY)
	    myHoy = myHoyY & myHoyM & myHoyD
				  
	    'Preparo las SQL para los camiones con muestras especiales para auditoria.
	    strSQLMuestras = ""
	    strSQLMuestrasH = ""
		if (chkMuestrasAud = MUESTRAS_AUDITORIA_ONLY) then 
            'Se debe editar la fecha ya que en todo el sistema el formato es AAAA-MM-DD y en esta tabla se creo mal y agregaron hh:mm:ss!
            strSQLMuestras = " INNER JOIN dbo.MUESTRASAUDCALADA MC on C.IDCAMION=MC.IDCAMION and  (YEAR(dtauditoria)*10000 + Month(dtauditoria)*100 + DAY(dtauditoria)) = '" & myHoy & "'"
            strSQLMuestrasH = " INNER JOIN dbo.MUESTRASAUDCALADA MC on C.IDCAMION=MC.IDCAMION and (YEAR(dtauditoria)*10000 + Month(dtauditoria)*100 + DAY(dtauditoria)) = (YEAR(C.dtcontable)*10000 + Month(C.dtcontable)*100 + DAY(C.dtcontable)) "
        end if  
		    	    		
		strSQL = "SELECT TABLA.*, DSPRODUCTO, DSCLIENTE, DSCORREDOR, DSVENDEDOR, DSESTADO, DSSILO FROM ("
		if ((cdCircuito = CIRCUITO_CAMION_TODOS) or (cdCircuito = CIRCUITO_CAMION_DESCARGA)) then
		    strSQL = strSQL & "(SELECT C.IDCAMION, " & myHoy & " AS DTCONTABLE, C.NUAUTSALIDA, CD.NUCARTAPORTE, NUCTAPTEDIG, CDPRODUCTO, CDCHAPACAMION, C.CDESTADO, NUCUPO, CD.CDCLIENTE, CD.CDCORREDOR, CD.CDVENDEDOR, CC.NUBARRAS, CC.SQCALADA ,"&CIRCUITO_CAMION_DESCARGA&" AS CIRCUITO, 0 Peso, C.CDSILO, C.CDTRANSPORTISTA "
		    strSQL = strSQL & " FROM dbo.CAMIONES C INNER JOIN dbo.CAMIONESDESCARGA CD ON C.IDCAMION=CD.IDCAMION "
		    strSQL = strSQL & strSQLMuestras		    
		    strSQL = strSQL & " LEFT JOIN (Select * from caladadecamiones X where SQCALADA = (Select MAX(SQCALADA) from CALADADECAMIONES Y where X.IDCAMION=Y.IDCAMION) ) CC on CC.IDCAMION=C.IDCAMION"		    
		    strSQL = strSQL & " where C.CDESTADO not in (Select CDESTADO from ESTADOSTERMINALES where CDTIPOCAMION=" & TIPO_TRANSPORTE_CAMION & "))"
		    strSQL = strSQL & " UNION"
		    strSQL = strSQL & " (SELECT C.IDCAMION, (YEAR(c.DtContable)*10000 + Month(c.DtContable)*100 + DAY(c.DtContable)) DTCONTABLE, C.NUAUTSALIDA, CD.NUCARTAPORTE, NUCTAPTEDIG, CDPRODUCTO, CDCHAPACAMION, C.CDESTADO, NUCUPO, CD.CDCLIENTE, CD.CDCORREDOR, CD.CDVENDEDOR, CC.NUBARRAS, CC.SQCALADA,"&CIRCUITO_CAMION_DESCARGA&" AS CIRCUITO, "
		    strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from dbo.HPesadasCamion pc  "
			strSQL = strSQL & "          where pc.dtContable = c.dtContable and pc.Idcamion = c.Idcamion and pc.cdPesada = 1 "
			strSQL = strSQL & "                and pc.sqpesada =  (select max(sqPesada) from dbo.HPesadasCamion "
			strSQL = strSQL & "                                     where dtcontable = pc.DtContable and pc.Idcamion = Idcamion and cdPesada = 1)) - "
			strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from dbo.HPesadasCamion pc "
			strSQL = strSQL & "          where pc.dtContable = c.dtContable  and pc.Idcamion = c.Idcamion and pc.cdPesada = 2 "
			strSQL = strSQL & "               and pc.sqpesada =  (select max(sqPesada) from dbo.HPesadasCamion where dtcontable = pc.DtContable and Idcamion = pc.Idcamion and cdPesada = 2)) -   "
			strSQL = strSQL & "         (select case when mc.vlMermaKilos is null then 0 else mc.vlMermaKilos end as vlMermaKilos from dbo.HMermasCamiones mc "
			strSQL = strSQL & "          where mc.dtContable = c.dtContable  and mc.Idcamion = c.Idcamion  "
			strSQL = strSQL & "                   and mc.sqpesada =  (select max(sqPesada) from dbo.HPesadasCamion where dtcontable = mc.DtContable and Idcamion = mc.Idcamion and cdPesada = 2)) as Peso, C.CDSILO, C.CDTRANSPORTISTA   "		
		    strSQL = strSQL & " FROM dbo.HCAMIONES C INNER JOIN dbo.HCAMIONESDESCARGA CD ON C.IDCAMION=CD.IDCAMION AND C.DTCONTABLE = CD.DTCONTABLE"
		    strSQL = strSQL & strSQLMuestrasH
		    strSQL = strSQL & " LEFT JOIN (Select * from hcaladadecamiones X where SQCALADA=(Select MAX(SQCALADA) from hcaladadecamiones Y where Y.DTCONTABLE=X.DTCONTABLE and X.IDCAMION=Y.IDCAMION) ) CC on CC.DTCONTABLE=C.DTCONTABLE and CC.IDCAMION=C.IDCAMION)"		    
		end if		
		if (cdCircuito = CIRCUITO_CAMION_TODOS) then strSQL = strSQL & " UNION "
		if ((cdCircuito = CIRCUITO_CAMION_TODOS) or (cdCircuito = CIRCUITO_CAMION_CARGA)) then
		    strSQL = strSQL & " (SELECT C.IDCAMION, " & myHoy & " AS DTCONTABLE, C.NUAUTSALIDA, CD.NUCARTAPORTE, NUCTAPTEDIG, CDPRODUCTO, CDCHAPACAMION, C.CDESTADO, NUCUPO, CD.CDCLIENTE, CD.CDCORREDOR, CD.CDVENDEDOR, CC.NUBARRAS, CC.SQCALADA,"&CIRCUITO_CAMION_CARGA&" AS CIRCUITO, "
            strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada "
            strSQL = strSQL & "          from dbo.PesadasCamion pc "
            strSQL = strSQL & "          where pc.Idcamion = C.Idcamion and pc.cdPesada = 1 "
            strSQL = strSQL & "                and pc.sqpesada =  (select max(sqPesada) from dbo.PesadasCamion  "
            strSQL = strSQL & "                                    where pc.Idcamion = Idcamion and cdPesada = 1)) -  "
            strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from dbo.PesadasCamion pc  "
            strSQL = strSQL & "          where pc.Idcamion = C.Idcamion and pc.cdPesada = 2  "
            strSQL = strSQL & "                and pc.sqpesada =  (select max(sqPesada) from dbo.PesadasCamion  "
			strSQL = strSQL & " 						           where Idcamion = pc.Idcamion and cdPesada = 2)) Peso, C.CDSILO, C.CDTRANSPORTISTA "
		    strSQL = strSQL & " FROM dbo.CAMIONES C INNER JOIN dbo.CAMIONESCARGA CD ON C.IDCAMION=CD.IDCAMION "
		    strSQL = strSQL & strSQLMuestras		    		    
		    strSQL = strSQL & " LEFT JOIN (Select * from caladadecamiones X where SQCALADA = (Select MAX(SQCALADA) from CALADADECAMIONES Y where X.IDCAMION=Y.IDCAMION)) CC on CC.IDCAMION=C.IDCAMION"
		    strSQL = strSQL & " where C.CDESTADO not in (Select CDESTADO from ESTADOSTERMINALES where CDTIPOCAMION=" & TIPO_TRANSPORTE_CAMION & "))"
		    strSQL = strSQL & " UNION"
		    strSQL = strSQL & " (SELECT C.IDCAMION, (YEAR(c.DtContable)*10000 + Month(c.DtContable)*100 + DAY(c.DtContable)) DTCONTABLE, C.NUAUTSALIDA, CD.NUCARTAPORTE, NUCTAPTEDIG, CDPRODUCTO, CDCHAPACAMION, C.CDESTADO, NUCUPO, CD.CDCLIENTE, CD.CDCORREDOR, CD.CDVENDEDOR, CC.NUBARRAS, CC.SQCALADA,"&CIRCUITO_CAMION_CARGA&" AS CIRCUITO, "
            strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada "
            strSQL = strSQL & "          from dbo.HPesadasCamion pc "
            strSQL = strSQL & "          where pc.dtContable = C.dtContable and pc.Idcamion = C.Idcamion and pc.cdPesada = 1 "
            strSQL = strSQL & "                and pc.sqpesada =  (select max(sqPesada) from dbo.HPesadasCamion  "
            strSQL = strSQL & "                                    where dtcontable = pc.DtContable and pc.Idcamion = Idcamion and cdPesada = 1)) -  "
            strSQL = strSQL & "         (select case when pc.vlPesada is null then 0 else pc.vlPesada end as vlPesada from dbo.HPesadasCamion pc  "
            strSQL = strSQL & "          where pc.dtContable = C.dtContable  and pc.Idcamion = C.Idcamion and pc.cdPesada = 2  "
            strSQL = strSQL & "                and pc.sqpesada =  (select max(sqPesada) from dbo.HPesadasCamion  "
			strSQL = strSQL & " 						           where dtcontable = pc.DtContable and Idcamion = pc.Idcamion and cdPesada = 2)) Peso, C.CDSILO, C.CDTRANSPORTISTA "
		    strSQL = strSQL & " FROM dbo.HCAMIONES C INNER JOIN dbo.HCAMIONESCARGA CD ON C.IDCAMION=CD.IDCAMION AND C.DTCONTABLE = CD.DTCONTABLE"
		    strSQL = strSQL & strSQLMuestrasH		    
		    strSQL = strSQL & " LEFT JOIN (Select * from hcaladadecamiones X where SQCALADA=(Select MAX(SQCALADA) from hcaladadecamiones Y where Y.DTCONTABLE=X.DTCONTABLE and X.IDCAMION=Y.IDCAMION)) CC on CC.DTCONTABLE=C.DTCONTABLE and CC.IDCAMION=C.IDCAMION)"
		end if
		strSQL = strSQL & ") AS TABLA "		
		strSQL = strSQL & " INNER JOIN dbo.ESTADOS E ON TABLA.CDESTADO=E.CDESTADO"
		strSQL = strSQL & " INNER JOIN dbo.PRODUCTOS P ON TABLA.CDPRODUCTO=P.CDPRODUCTO "
	    strSQL = strSQL & " INNER JOIN dbo.CLIENTES CL ON TABLA.CDCLIENTE=CL.CDCLIENTE "
	    strSQL = strSQL & " INNER JOIN dbo.CORREDORES CO ON TABLA.CDCORREDOR=CO.CDCORREDOR "
	    strSQL = strSQL & " INNER JOIN dbo.VENDEDORES VE ON TABLA.CDVENDEDOR=VE.CDVENDEDOR "
	    strSQL = strSQL & " LEFT JOIN dbo.SILOS SI ON TABLA.CDSILO=SI.CDSILO "
		strSQL = strSQL & myWhere	

		'Ordenamiento
		sortBy = GF_PARAMETROS7("sortBy", "", 6)
		if len(sortBy) > 0 then
			call addParam("sortBy", sortBy, params)
			strSQL = strSQL & " " & sortBy
		else
			strSQL = strSQL & " ORDER BY DTCONTABLE DESC, IDCAMION ASC"
		end if		
    	
    	'Response.Write strSQL
		Call GF_BD_Puertos(pto, rsLista, "OPEN",strSql)
		
		if (accion <> ACCION_PROCESAR) then
		    hayBusqueda = false
		    paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
		    if (paginaActual = 0) then paginaActual=1
		    mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
		    if (mostrar = 0) then mostrar = 50
		    Call setupPaginacion(rsLista, paginaActual, mostrar)
		    lineasTotales = rsLista.recordcount		
		else
		    mostrar = rsLista.recordcount
		    Call recuperarPartesIntervinientes(fecContable, fecContableH)
		end if		 
	end if	

%>