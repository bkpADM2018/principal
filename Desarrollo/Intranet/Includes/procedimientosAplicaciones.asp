<%
'******************************************************************************************************************
'
'	PROCEDIMIENTO APLICACIONES : CONTIENE LAS CLASES ENCARGADAS DE CREAR LA PAGINA DE INICIO DE LA INTRANET,
'								 CADA UNA DE ESTAS CLASES REPRESENTAN APLICACIONES DEL SISTEMA.
'		
'******************************************************************************************************************
'Constantes que representan a cada cuadrante
Const QUADRANT_HOME_AUTORIZACIONES = 1
Const QUADRANT_HOME_COMPLIANCE	   = 2
Const QUADRANT_HOME_POSEIDON	   = 3
Const QUADRANT_HOME_PERMISOS	   = 4

Const QUADRANT_HOME_CONTRATOS = 5
Const QUADRANT_HOME_PAGOS	  = 6
'Const QUADRANT_HOME_CAMIONES  = 7 - JAS: No se usa mas.
Const QUADRANT_HOME_AFIP	  = 8
Const QUADRANT_HOME_ADUANA	  = 9

Set oDicQuadrant = Server.CreateObject("Scripting.Dictionary")
'-------------------------------------------------------------------------------------------------------------------
'Funcion existQuadrant: verifica si ya fue instanciado el Quadrante pasado por parametro
'			True  : ya fue cargado 
'			False : no fue cargado 
Function existQuadrant(pCdQuadrant)
	existQuadrant = false
	if oDicQuadrant.Exists(pCdQuadrant) then existQuadrant = true	
End Function
'----------------------------------------------------------------------------------------------------------------------
'clsAutorizaciones :
'			Esta clase carga los datos de las Autorizaciones que tiene cada usuario.
'				- En caso de que no encuentre el cuadrante solicitado, se lo dibujará con los valores por defecto.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsAutorizaciones
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo
		'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*- FUNCIONES PRIVADAS DEL CUADRANTE *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-
		
		'Dibuja la structura HTML del cuadrante
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function		
		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "authorize-200.png"
			v_name			= "Autorizaciones"
			v_title			= "Pendientes"
			v_link			= "comprasAutorizaciones.asp"
			v_linkResumen_1 = "comprasAutorizaciones.asp"
			v_linkResumen_2 = "comprasAutorizaciones.asp"
			v_linkResumen_3 = "comprasAutorizaciones.asp"
			v_strResumen_1  = "-"
			v_strResumen_2  = "-"
			v_strResumen_3  = "-"
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_cdQuadrant    = QUADRANT_HOME_AUTORIZACIONES
		End Function
		
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_AUTORIZACIONES) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_AUTORIZACIONES, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function
	End Class 
'----------------------------------------------------------------------------------------------------------------------
'clsPoseidon : 
'			Esta clase carga los datos para acceder al Poseidon.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsPoseidon
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo
		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "poseidon-200.png"
			v_name			= "Poseidon"
			v_title			= "Poseidon"
			if (SITE_INTRANET) then
				v_linkResumen_1 = URL_INTRANET_ARROYO
				v_linkResumen_2 = URL_INTRANET_TRANSITO
				v_linkResumen_3 = URL_INTRANET_BAHIA
			else
				v_linkResumen_1 = "Poseidon/panelPuertos.asp?pto=" & DBSITE_ARROYO
				v_linkResumen_2 = "Poseidon/panelPuertos.asp?pto=" & DBSITE_TRANSITO
				v_linkResumen_3 = "Poseidon/panelPuertos.asp?pto=" & DBSITE_BAHIA
			end if
			v_strResumen_1  = "Arroyo"
			v_strResumen_2  = "Pto. San Mart&iacuten"
			v_strResumen_3  = "Bah&iacutea Blanca"
			v_target_1		= "_blank"
			v_target_2		= "_blank"
			v_target_3		= "_blank"
			v_link			= ""
			v_cdQuadrant	= QUADRANT_HOME_POSEIDON
		End Function
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function		
		
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_POSEIDON) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_POSEIDON, "")
				Call loadPanel()
				Call drawStructPanel()			
			end if	
		End Function
	End Class 
'----------------------------------------------------------------------------------------------------------------------
'clsCompliance : 
'			Esta clase carga los datos para acceder al Poseidon.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsCompliance
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo

		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "compliance-200.png"
			v_name			= "Compliance"
			v_title			= "Compliance"						
			v_linkResumen_1 = "compliance/normas.asp?p_baseFolder=Alert%20Line&p_titulo=Alert%20Line"
			v_strResumen_1  = "Alert Line"
			v_linkResumen_2 = "compliance/normas.asp?p_baseFolder=Documentos&p_titulo=Centro%20de%20Pol%EDticas"
'			v_linkResumen_2 = "http://inside.adm.com/es-ES/Policies/Paginas/default.aspx"
			v_strResumen_2	= "Centro de Politicas"
			v_linkResumen_3 = "compliance/contacts.asp"
			v_strResumen_3  = "Contactos"
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_link			= "compliance/index.asp"
			v_cdQuadrant	= QUADRANT_HOME_COMPLIANCE
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
	Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function			
	
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_COMPLIANCE) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_COMPLIANCE, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function		
	End Class 
	
'----------------------------------------------------------------------------------------------------------------------
'clsPermisos : 
'			Esta clase carga los datos para acceder a los Permisos.
'				- En caso de que el usuario sea un Jefe de Área podrá ver los permisos de su sector
'----------------------------------------------------------------------------------------------------------------------		
	Class clsPermisos
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo
		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "perfil-200.png"
			v_name			= "Permisos"
			v_title			= "Permisos"
			v_linkResumen_1 = "#"
			v_strResumen_1  = "-"
			v_linkResumen_2 = "#"
			v_strResumen_2  = "-"
			v_linkResumen_3 = "#"
			v_strResumen_3  = "-"
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_link			= "AUPSectores.asp"
			v_cdQuadrant	= QUADRANT_HOME_PERMISOS
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function		
	
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_PERMISOS) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_PERMISOS, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function
	End Class 

'----------------------------------------------------------------------------------------------------------------------
'clsContratos : 
'			Esta clase carga el cuadrante de contratos.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsContratos
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo

		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "../images/contratos-200.png"
			v_name			= "Contratos"
			v_title			= "Contratos"						
			v_linkResumen_1 = "cor-MenuCto.asp"
			v_strResumen_1  = "Contratos y Descargas"
			v_linkResumen_2 = "contratosCabeceraConfirma.asp"
			v_strResumen_2	= "Confirmación de Contratos"
			v_linkResumen_3 = "mercaderias/AdministrarF1116A.asp"
			v_strResumen_3  = "Formularios F1116A"
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_link			= "cor-MenuCto.asp"
			v_cdQuadrant	= QUADRANT_HOME_CONTRATOS
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function			
	
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_CONTRATOS) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_CONTRATOS, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function		
	End Class 
'----------------------------------------------------------------------------------------------------------------------
'clsContratos : 
'			Esta clase carga el cuadrante de contratos.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsAFIP
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo

		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "../images/AFIP-200.png"
			v_name			= "AFIP"
			v_title			= "Información Impositiva"						
			v_linkResumen_1 = "cor-ImpRet.asp"
			v_strResumen_1  = "Retenciones"
			v_linkResumen_2 = "Documentos/LegajoImpositivo.htm"
			v_strResumen_2	= "Información Impositiva"
			v_linkResumen_3 = ""
			v_strResumen_3  = "&nbsp;"
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_link			= "cor-SitImpMenu.asp"
			v_cdQuadrant	= QUADRANT_HOME_AFIP
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function		
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_AFIP) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_AFIP, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function		
	End Class 
	
	'----------------------------------------------------------------------------------------------------------------------
'clsContratos : 
'			Esta clase carga el cuadrante de contratos.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsPagos
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_strResumen_1		'Primera descripción del resumen
		private v_linkResumen_1		'Primer link del resumen
		private v_strResumen_2		'Segunda descripción del resumen
		private v_linkResumen_2		'Segundo link del resumen
		private v_strResumen_3		'Tercera descripción del resumen
		private v_linkResumen_3		'Tercer link del resumen
		private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo

		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		= "../images/Pagos-200.png"
			v_name			= "Pagos/Cobros"
			v_title			= "Pagos y Cobros"						
			v_linkResumen_1 = "cor-ordenesPago.asp"
			v_strResumen_1  = "Pagos"
			v_linkResumen_2 = "interfacturas/interfacturasConsulta.asp"
			v_strResumen_2	= "Facturas"
			v_linkResumen_3 = ""
			v_strResumen_3  = "<b>IMPORTANTE:</b> Los pagos correspondientes al dia de la fecha estarán disponibles a partir de las 12 Hs."
			v_target_1		= "_self"
			v_target_2		= "_self"
			v_target_3		= "_self"
			v_link			= "cor-ordenesPago.asp"
			v_cdQuadrant	= QUADRANT_HOME_PAGOS
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>
				<div class="boxer_btn">
				</div>
			</div>
		<%
		End Function		
	
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_PAGOS) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_PAGOS, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function		
	End Class 		
'----------------------------------------------------------------------------------------------------------------------
'clsAduana : 
'			Esta clase carga el cuadrante de Aduana.
'----------------------------------------------------------------------------------------------------------------------		
	Class clsAduana
		private v_imgName			'Ruta de la imagen del cuadrante
		private v_name				'Titulo de la imagen del cuadrante
		private v_title				'Titulo del cuadrante
		private v_linkResumen_1	'Primer link del resumen
        private v_linkResumen_2	'Primer link del resumen
        private v_linkResumen_3	'Primer link del resumen
        private v_strResumen_1		'Primera descripción del resumen
        private v_strResumen_2		'Segunda descripción del resumen
		private v_strResumen_3		'Tercera descripción del resumen
        private	v_target_1			'Target del primer link del resumen
		private	v_target_2			'Target del segundo link del resumen
		private	v_target_3			'Target del tercer link del resumen
		private v_link				'Link maestro del cuadrante
		private v_cdQuadrant		'Es el codigo del cuadrante para identificarlo

		
		'Funcion loadPanel : se encarga de cargar los atributos de Autorizaciones(Imagen, Alt, Titulo, Link, etc)
		Private Function loadPanel()
			v_imgName		    = "poseidon-200.png"
			v_name			    = "Embarques"
			v_title			    = "Embarques"						
			v_linkResumen_1  = "javascript:abrirRegistrosBalanza('" & TERMINAL_ARROYO & "')"			
            v_linkResumen_2  = "javascript:abrirRegistrosBalanza('" & TERMINAL_TRANSITO & "')"
			v_linkResumen_3  = "javascript:abrirRegistrosBalanza('" & TERMINAL_PIEDRABUENA & "')"
            v_strResumen_1      = "Arroyo"
			v_strResumen_2      = "Pto. San Mart&iacute;n "
			v_strResumen_3      = "Bah&iacute;a Blanca"
			v_target_1		    = ""
			v_target_2		    = ""
			v_target_3		    = ""
			v_link			    = "#"
			v_cdQuadrant	= QUADRANT_HOME_ADUANA
		End Function		
		
		'Dibuja la structura HTML del cuadrante de Compliance
		Private Function drawStructPanel()
		%>
			<div class="boxer">
				<div class="boxer_img">
					<% if (v_link <> "") then %>
    				<a href="<%=v_link%>">
    				<% end if %>
    					<img src="images/<%=v_imgName%>" alt="<%=v_name%>" />
					<% if (v_link <> "") then %>    					
    				</a>
    				<% end if %>
					<span><%=v_name%></span>
				</div>
			    <div class="boxer_table">
		    		<h2 id="titleQuadrant_<%=v_cdQuadrant%>"><%=v_title%></h2>
					<li>
						<a href="<%=v_linkResumen_1%>" target="<%=v_target_1%>"><%=v_strResumen_1%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_2%>" target="<%=v_target_2%>"><%=v_strResumen_2%></a>
					</li>
					<li>
						<a href="<%=v_linkResumen_3%>" target="<%=v_target_3%>"><%=v_strResumen_3%></a>
					</li>
				</div>				
			</div>
		<%
		End Function			
	
		'Funcion drawPanel : dibuja un nuevo panel en la pagina
		Public Function drawPanel()
			if not oDicQuadrant.Exists(QUADRANT_HOME_ADUANA) then 
				Call oDicQuadrant.Add(QUADRANT_HOME_ADUANA, "")
				Call loadPanel()
				Call drawStructPanel()
			end if	
		End Function		
	End Class 

%>