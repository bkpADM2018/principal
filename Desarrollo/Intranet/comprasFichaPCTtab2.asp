<%
	Dim rsCotizaciones, mensaje, textoFechaApertura, textoUsuarioApertura, myStyle, mySelected
	dim declino, presento,textoFechaLectura, textoUsuarioLectura,auxCRCEncryipt,auxCRC

%>
<html>
<head>
<script type="text/javascript" src="scripts/channel.js"></script>

<script type="text/javascript">		
var ch = new channel();	
	
	function asignarPresupuestoAbierto(fileno){
	ch.bind("comprasRegistrarLecturaAjax.asp?accion=<%=ACCION_ARCHIVO_LECTURA%>&fileno=" + fileno , "deleteCredito_callback()");
	ch.send();
	}
	function deleteCredito_callback(){}
	
	function verCotizacion(fileno) {
		<%	if (pct_idEstado >= ESTADO_PCT_ABIERTO) then %>
		asignarPresupuestoAbierto(fileno)
		window.open("comprasOpenArchivo.asp?idPedido=<% =pct_idPedido %>&fileno=" + fileno , "_blank", "",false);		
		<%	else	%>
		alert("<% =GF_TRADUCIR("A�n no se ha cerrado el plazo de presentaci�n de cotizaciones.") %>");
		<%	end if %>
	}
	function consultarMail(idProveedor) {		
		window.open("comprasEnvioPCTMail.asp?idPedido=<% =pct_idPedido %>&accion=<% =ACCION_VISUALIZAR %>&idProveedor=" + idProveedor, "_blank", "location=no,menubar=no,statusbar=no,height=240,width=500",false);		
	}		
	
	function subirCotizacion(idProveedor, idPedido) {	
		window.open("comprasCargarCotizacion.asp?idPedido=" + idPedido + "&idProveedor=" + idProveedor, "_blank", "location=no,menubar=no,statusbar=no,height=240,width=500",false);
	}
			
</script>
<div style="height:450px;">
</head>
<body>
<br>
<%
'strSQL="Select * from TOEPFERDB.TBLPCTCOTIZACIONES TC INNER JOIN TOEPFERDB.VWEMPRESAS VE on TC.IDPROVEEDOR = VE.IDEMPRESA WHERE TC.IDCOTIZACION=(Select idCotizacion from toepferdb.tblPCTCAbecera where idPedido=" & pct_idPedido & ")"
'Response.Write strsql
myStyle = "reg_Header"
imagen="CTZ-48x48.png"	
'Call GF_BD_COMPRAS(rs, oConn, "OPEN", strSQL)
Set rs = getCotizaciones(pct_idPedido, pct_idProveedorElegido)		
if not rs.eof then		%>
	<table class="<% =myStyle %>" align="center" width="100%">
			<tr>
				<td colspan="3" class="TDNOHAY"><b><% =GF_TRADUCIR("PRESUPUESTO SELECCIONADO") %> </b></td>			
			</tr>		
			<tr>
				<td colspan="3"><b><% =pct_idProveedorElegido & "-" & pct_dsProveedorElegido %> </b></td>			
			</tr>		
<%	while not rs.eof 
		mensaje = "verCotizacion('" & rs("IDCOTIZACION") & "')"
		if (rs("FECHAAPERTURA") <> "") then
			textoFechaApertura = GF_TRADUCIR("Fecha Apertura ") & ": " & GF_FN2DTE(rs("FECHAAPERTURA"))
			textoUsuarioApertura = GF_TRADUCIR("Responsables ") & ": " & getUsuariosApertura(pct_idPedido)
			textoFechaLectura = GF_TRADUCIR("Fecha Lectura") & ": " & GF_FN2DTE(rs("FECHALECTURA"))
			textoUsuarioLectura = GF_TRADUCIR("Usuario") & ": " & getUserDescription(rs("CDUSRLECTURA"))
		end if
		%>
		
			<tr>
				<td rowSpan="6" width="5%">&nbsp;</td>
				<td rowSpan="6" width="10%">
					<span style="cursor:pointer" onclick="<% =mensaje %>">
						<img src="images/compras/<% =imagen %>">
					</span>
				</td>
			</tr>							
			<tr class="<%=myStyle%>"><td><% =GF_TRADUCIR("Fecha Presentacion") %>: <% =GF_FN2DTE(rs("FECHAPRESENTACION")) %></td></tr>
			<tr class="<%=myStyle%>"><td><% =textoFechaLectura %></td></tr>
			<tr class="<%=myStyle%>"><td><% =textoUsuarioLectura %></td></tr>
			<tr class="<%=myStyle%>"><td><% =textoFechaApertura %></td></tr>		
			<tr class="<%=myStyle%>"><td><% =textoUsuarioApertura %></td></tr>								
		<%
		rs.movenext
	wend	%>
	</table>
	<hr>
<%
end if

%>
<table align="center" width="100%">
	<tr>
		<td colspan="3" class="TDNOHAY"><b><% =GF_TRADUCIR("PRESUPUESTOS PRESENTADOS") %> </b></td>			
	</tr>		
	<% 	
    Call initProveedores()
	while (readNextProveedor())
		declino = false
		presento = false        
		%>
		<tr>
			<td colspan="3"><b><% =pct_idProveedor & "-" & pct_dsProveedor %></b>&nbsp
<%
                myRol = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)
                if (myRol = FIRMA_ROL_GTE_COMPRAS) then
                    'auxCRC = "SV:"& SISTEMA_COMPRAS &"|PR:"& pct_idProveedor &"|P:"& pct_idPedido &"|F:V"
                    'auxCRCEncryipt = Trim(MD5(generarCRCByPCT(pct_idProveedor,pct_idPedido,"V")))                    
%>			                
				<a name="emailLink_<% =pct_idProveedor%>" id="emailLink_<% =pct_idProveedor%>" href="javascript:subirCotizacion(<% =pct_idProveedor %>, <% =pct_idPedido %>)">
                    <img src="images/compras/supplier_key-16x16.png">
                </a>&nbsp 
<%              end if %>                               
                <a href="javascript:consultarMail(<% =pct_idProveedor %>)">
                    <img src="images/compras/PCT_publish-16x16.png" title="Pedido listo para enviar a proveedores">
                </a>
            </td>
		</tr>
		<%		
		Set rsCotizaciones = getCotizaciones(pct_idPedido, pct_idProveedor)		
		if (not rsCotizaciones.eof) then	
			while (not rsCotizaciones.eof)	
				textoFechaApertura = ""
				textoUsuarioApertura = ""
				textoFechaLectura = ""
				textoUsuarioLectura = ""
				if (isNull(rsCotizaciones("CDUSRAPERTURA")) and pct_tipoCompra = TIPO_PCT_CONCURSO) then				
					textoFechaApertura = "&nbsp;"
					textoUsuarioApertura = "&nbsp;"									
					textoFechaLectura = "&nbsp;"
					textoUsuarioLectura = "&nbsp;"
					if (rsCotizaciones("PATHCOTIZACION") <> ACCION_PCT_RETIRARSE) then
						'El sobre no fue abierto
						imagen="Bid_purchase-48x48.png"
						if (pct_idEstado >= ESTADO_PCT_ABIERTO) then
							'Ya se hizo la apertura de sobres
							mensaje="alert('" & GF_TRADUCIR("La apertura de sobres ya fue realizada pero este sobre no fue considerado.") & "')"
						else
							mensaje= "alert('" & GF_TRADUCIR("Aun no se ha hecho la apertura de los sobres!") & "')"						
						end if
						presento = true
					else
						imagen="CTZR-48x48.png"
						textoFechaApertura = "<B>" & GF_TRADUCIR("El proveedor decidi� no participar de la cotizac�n.") & "</B>"
						mensaje = "alert('El proveedor decidi� no participar de la cotizac�n.');"
						declino = true
					end if
				else		
					if (rsCotizaciones("PATHCOTIZACION") <> ACCION_PCT_RETIRARSE) then		
						imagen="CTZ-48x48.png"						
						mensaje = "verCotizacion('" & rsCotizaciones("IDCOTIZACION") & "')"
						if (rsCotizaciones("FECHAAPERTURA") <> "") then																	
							textoFechaApertura = GF_TRADUCIR("Fecha Aperutura ") & ": " & GF_FN2DTE(rsCotizaciones("FECHAAPERTURA"))
							textoUsuarioApertura = GF_TRADUCIR("Responsables ") & ": " & getUsuariosApertura(pct_idPedido)							
						end if
						if not IsNull(rsCotizaciones("FECHALECTURA"))then
							textoFechaLectura =  GF_TRADUCIR("Fecha Lectura ") & ": " & GF_FN2DTE(rsCotizaciones("FECHALECTURA")) 
							textoUsuarioLectura = GF_TRADUCIR("Usuario ") & ": " & getUserDescription(rsCotizaciones("CDUSRLECTURA"))								
						end if
						presento = true						
					else
						imagen="CTZR-48x48.png"
						textoFechaApertura = "<B>" & GF_TRADUCIR("El proveedor decidi� no participar de la cotizac�n.") & "</B>"
						mensaje = "alert('El proveedor decidi� no participar de la cotizac�n.');"
						declino = true						
					end if
				end if
				%>
				<tr>
					<td rowSpan="6" width="5%"></td>
					<td rowSpan="6" width="10%">
						<span style="cursor:pointer" onclick="<% =mensaje %>">
							<img src="images/compras/<% =imagen %>">
						</span>
					</td>
				</tr>		
				<tr><td>
						<label <% if ( (GF_DTEDIFF(rsCotizaciones("FECHAPRESENTACION"), GF_DTE2FN(pct_FechaCierre), "D") < 0) and presento) then %>
							class="reg_header_error round_border_all" title="Cotizacion cargada fuera de termino" style="cursor:pointer; padding:2px;"
						<% end if %> >
							<% =GF_TRADUCIR("Fecha Presentacion") %>: <% =GF_FN2DTE(rsCotizaciones("FECHAPRESENTACION")) %>&nbsp;
						</label>
					</td>
				</tr>
				<tr><td><% =textoFechaApertura %></td></tr>
				<tr><td><% =textoUsuarioApertura %></td></tr>				
				<tr><td><% =textoFechaLectura %></td></tr>
				<tr><td><% =textoUsuarioLectura %></td></tr>
				<%
				rsCotizaciones.MoveNext()
			wend
		else	
			%>
			<tr><td colspan="3" class="TDNOHAY"><% =GF_TRADUCIR("No hay cotizaciones") %></td></tr>
			<%
		end if 							
		%>
		<tr><td colspan="3"><hr></td></tr>
	<% 	
	wend 
	%>			
</table>
</body>	
</div>
