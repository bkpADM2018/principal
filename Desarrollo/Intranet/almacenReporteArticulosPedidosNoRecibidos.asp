<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosPDF.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->

<%
Dim ultimaLinea,nroPagina,oPDF

Const COMIENZO_COL_ARTICULO  = 15
Const COMIENZO_COL_PROVEEDOR = 240
Const COMIENZO_COL_CANTP = 480
Const COMIENZO_COL_CANTR = 530

Const FILTRO_TODOS = 0
Const FILTRO_SIN_OBRA = 1
Const FILTRO_CON_OBRA = 2
Const FILTRO_MANTENIMIENTO = 3
Const FILTRO_INVERSIONES = 4
'-----------------------------------------------------------------------
Function dibujarEncabezado(p_titulo,p_subtitulo)
	'Devuelve la posicion donde se puede seguir escribiendo
	Dim parametros(),rs1,rs2,dscategoria,dsalmacen,busqueda
	redim parametros(4)
	
	Call GF_setFont(oPDF,"ARIAL",16,8)
	Call GF_writeTextAlign(oPDF,5,25,p_titulo, 580 , PDF_ALIGN_CENTER)
	Call GF_setFont(oPDF,"ARIAL",8,0)
	Call GF_writeTextAlign(oPDF,5,42,p_subtitulo, 580 , PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,2,65,590)
	Call GF_setFont(oPDF,"COURIER",8,0)
	GP_CONFIGURARMOMENTOS
	Call GF_writeTextAlign(oPDF,5,5,GF_FN2DTE(session("MmtoSistema")), 580 , PDF_ALIGN_RIGHT)
	Call GF_writeTextAlign(oPDF,5,5+pdf_currentFontSize,session("Usuario"), 580 , PDF_ALIGN_RIGHT)
	'obtengo las descripciones de los parametros de busqueda
	if (nroPagina = 1) then
		
		
		parametros(0)= "Division"  & "," & getDivisionDS(division)
		if (CAB_dsProveedor <> "" ) then
			parametros(1) = "Proveedor" & "," & CAB_dsProveedor
		else
			parametros(1) = "Proveedor" & "," & "TODOS"
		end if
		if (ARTDS <> "") then
			parametros(2) = "Articulo"  & "," & ARTDS
		else
			parametros(2) = "Articulo"  & "," & "TODOS"
		end if
		if (pedido <> "") then
			parametros(3) = "Pedido"    & "," & pedido
		else
			parametros(3) = "Pedido"    & "," & "TODOS"
		end if
		select case (aux_filtro)
			case FILTRO_TODOS :
				parametros(4) = "Asociado a"    & "," & "TODOS"
			case FILTRO_SIN_OBRA :
				parametros(4) = "Asociado a"    & "," & "SIN OBRA"
			case FILTRO_CON_OBRA :
				parametros(4) = "Asociado a"    & "," & "CON OBRA"
			case FILTRO_MANTENIMIENTO :
				parametros(4) = "Asociado a"    & "," & "OBRA / MANTENIMIENTO"
			case FILTRO_INVERSIONES :
				parametros(4) = "Asociado a"    & "," & "OBRA / INVERSIONES"
		end select

		dibujarEncabezado = dibujarFiltros(parametros)
	else
		dibujarEncabezado = 75
	end if
End Function
'-----------------------------------------------------------------------
Function dibujarPagina()
	Dim y_comienzo
	

		
	Call GF_squareBox(oPDF, 2, 2, 590, 833, 0, "", "#0B3B0B", 2, PDF_SQUARE_ROUND)
	
	Call GF_writeImage(oPDF, Server.MapPath("images\kogge64.gif"), 20, 10, 48, 48, 0)
	
	y_comienzo =  dibujarEncabezado("REPORTE ARTICULOS PEDIDOS NO RECIBIDOS","")
	
	Call GF_squareBox(oPDF, COMIENZO_COL_ARTICULO, y_comienzo, 225, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_PROVEEDOR, y_comienzo, 240, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_CANTP, y_comienzo, 50, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	Call GF_squareBox(oPDF, COMIENZO_COL_CANTR, y_comienzo, 50, 13, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
	
	Call GF_setFontColor("#FFFFFF")
	Call GF_setFont(oPDF,"ARIAL",8,FONT_STYLE_BOLD)
	
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_ARTICULO , y_comienzo+2, "ARTICULO" , 225 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_PROVEEDOR , y_comienzo+2, "PROVEEDOR" , 240 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTP , y_comienzo+2, "PEDIDO" , 50 , PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTR , y_comienzo+2, "FALTANTE" , 50, PDF_ALIGN_CENTER)
	
	Call GF_setFontColor("#000000")
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF, 10 , 840, "Pagina  " & nroPagina		 , 580 , PDF_ALIGN_RIGHT)
	
	dibujarPagina = y_comienzo +16
End Function
'-----------------------------------------------------------------------
Function cargarDatos(ultimaLinea)
	Dim i,color,articulo,proveedor,cantr,cantp,unidad
	
	i = 0
	Call GF_setFontColor("#000000")
	Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
	if (not rsArticulos.eof) then
		while not rsArticulos.EOF
			if i mod 2 then 
				color = "#dcf7dc"			
			else
				color = "#FFFFFF"
			end if
			y = ultimaLinea+(i*(pdf_currentFontSize+separacion))
			Call GF_squareBox(oPDF, COMIENZO_COL_ARTICULO , y-1 ,565, pdf_currentFontSize, 0, color, color, 1, PDF_SQUARE_NORMAL)
			
			articulo  = rsArticulos("idarticulo") & "-" & rsArticulos("dsarticulo")
			if (len(trim(articulo))>42) then articulo = left(articulo,42) & "..."
			proveedor = rsArticulos("idproveedor") & "-" & rsArticulos("dsproveedor")
			if (len(trim(proveedor))>42) then proveedor  = left(proveedor ,42) & "..."
			if (isnull(rsArticulos("cantidadp"))) then cantp = 0 else cantp = rsArticulos("cantidadp") end if
			if (isnull(rsArticulos("cantidadr"))) then cantr = rsArticulos("cantidadp") else cantr = cdbl(rsArticulos("cantidadp"))-cdbl(rsArticulos("cantidadr")) end if
			unidad = rsArticulos("unidad")

			Call GF_writeTextAlign(oPDF, COMIENZO_COL_ARTICULO+5 , y-2, articulo , 215 , PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF, COMIENZO_COL_PROVEEDOR+5 , y-2, proveedor , 230 , PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTP+5 , y-2, cantp  , 20 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTP+30 , y-2, unidad , 20 , PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTR+5 , y-2, cantr  , 20 , PDF_ALIGN_RIGHT)
			Call GF_writeTextAlign(oPDF, COMIENZO_COL_CANTR+30 , y-2, unidad , 20 , PDF_ALIGN_LEFT)
			
			if (ultimaLinea+(i*(pdf_currentFontSize+separacion)) >= 810) then
					ultimaLinea = nuevaPagina()
					i = -1
					Call GF_setFontColor("#000000")
					Call GF_setFont(oPDF,"COURIER",8,FONT_STYLE_NORMAL)
			end if
				
			i = i +1
			rsArticulos.MoveNext
		wend	
	else
		Call GF_squareBox(oPDF, COMIENZO_COL_ARTICULO, ultimaLinea-16, 570, 15, 0, "#517b4a", "#000000", 1, PDF_SQUARE_NORMAL)
		Call GF_setFont(oPDF,"ARIAL",8,0)
		Call GF_setFontColor("#FFFFFF")
		
		Call GF_writeTextAlign(oPDF, COMIENZO_COL_ARTICULO , ultimaLinea-13, "NO SE ENCONTRARON DATOS EN LA BUSQUEDA" , 570 , PDF_ALIGN_CENTER)
	end if
End Function
'-----------------------------------------------------------------------------------------------
Function nuevaPagina()
	nroPagina = nroPagina +1
	Call GF_newPage(oPDF)
	nuevaPagina = dibujarPagina()
End Function
'-----------------------------------------------------------------------
Function crearPDF()
	Set oPDF = GF_createPDF("")
	Call GF_setPDFMODE(PDF_STREAM_MODE)
	nroPagina = 1
	ultimaLinea = dibujarPagina()
	Call cargarDatos(ultimaLinea)
	Call GF_closePDF(oPDF)
End Function
'-----------------------------------------------------------------------
Function dibujarFiltros(p_parametros)
'Funcion que dibuja los parametros de busqueda 
'recibe como parametro un vector con la siguiente estructura en cada posicion:
'	"nombreBusqueda,valorBusqueda"
'Devuelve la posicion donde se puede seguir escribiendo
	Dim aux,x_inicial,y_inicial,font_size
	Dim max_len
	
	my_font_size = 8
	x_inicial = 20
	y_inicial = 75
	
	max_len = 0
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
		if (len(aux(0)) > max_len) then
			max_len = len(aux(0))
		end if
	next
	
	for i = 0 to ubound(p_parametros)
		aux = split(p_parametros(i),",")
		Call GF_setFont(oPDF,"COURIER",my_font_size,0)
		Call GF_writeTextAlign(oPDF,x_inicial,y_inicial+(i*my_font_size), completarEspacios(aux(0),max_len) & ": " & aux(1) ,580, PDF_ALIGN_LEFT)
	next 
	
	dibujarFiltros = y_inicial+(i*my_font_size)+my_font_size
	
End Function
'-----------------------------------------------------------------------------------------------
Function completarEspacios(p_palabra,p_len)
	Dim rtrn
	rtrn = p_palabra
	for i = len(p_palabra) to p_len
		rtrn = rtrn & "."
	next
	completarEspacios = rtrn
End Function
'-----------------------------------------------------------------------
Function crearSelect(byref rs,pNombre,pValor,pTexto)
	Dim rtrn
	rtrn = "<select class='selects' id='" & pNombre & "' name='" & pNombre & "'>"
	
	while not rs.eof
		rtrn = rtrn & "<option " 
		if ( cstr(rs(pValor))=GF_PARAMETROS7(pNombre,"",6) ) then
			rtrn = rtrn & "selected='selected'"
		end if
		rtrn = rtrn & "value=" & rs(pValor) & ">" & rs(pTexto) & "</option>"
		rs.MoveNext
	wend
	
	rtrn = rtrn & "</select>"
	crearSelect = rtrn
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
	Dim origen,accion,CAB_idProveedor,CAB_dsProveedor,IT_artID,division,pedido
	Dim rsArticulos, auxfiltro, filtro
	
	origen = GF_PARAMETROS7("origen"   , "", 6)
	accion = GF_PARAMETROS7("accion"   , "", 6)
	CAB_idProveedor = GF_PARAMETROS7("CAB_idProveedor",0,6)	
	CAB_dsProveedor = GF_PARAMETROS7("CAB_dsProveedor","",6)
	ARTID = GF_PARAMETROS7("ARTID",0,6)	
	ARTDS = GF_PARAMETROS7("ARTDS","",6)
	division = GF_PARAMETROS7("division","",6)
	pedido = GF_PARAMETROS7("pedido","",6)
	aux_filtro = GF_PARAMETROS7("filtro",0,6)
	if (aux_filtro <> FILTRO_TODOS) then
		if (aux_filtro = FILTRO_SIN_OBRA) then
			filtro = " and (OBR.IDOBRA is null or OBR.IDOBRA = 0) "
		else
			filtro = " and OBR.IDOBRA > 0 "
		end if
		if (aux_filtro = FILTRO_MANTENIMIENTO) then
			filtro = filtro & " and OBR.ESINVERSION <> '" & OBRA_INVERSION & "' "
		elseif (aux_filtro = FILTRO_INVERSIONES) then
			filtro = filtro & " and OBR.ESINVERSION <> '" & OBRA_MANTENIMIENTO & "' "
		end if
	end if

	if (accion = ACCION_PROCESAR) then		
		Set rsArticulos = obtenerArticulosPedidosNoRecibidos(division, CAB_idProveedor, ARTID, pedido, filtro)
		Call crearPDF()	
	end if

	
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento sin t&iacute;tulo</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="CSS/MagicSearch.css" type="text/css">
<script type="text/javascript" src="scripts/channel.js"></script>
<script language="javascript" src="scripts/magicSearchObj.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript">	
	function bodyOnLoad() {			
		tb = new Toolbar('toolbar', 6,'images/almacenes/');		
		tb.addButton("printer-16x16.png", "Generar", "submitInfo()");
		tb.addButton("Previous-16x16.png", "Volver", "cerrar()");
		tb.draw();	
		startMagicSearch();		
	}
	function cerrar(){
		<% if (origen = "COMPRAS") then%>
			location.href='comprasReportes.asp'
		<% else %>
			location.href='almacenReportes.asp'
		<%end if%>
	}
	function submitInfo() {		
		document.getElementById("frm").submit();
	}	
	function startMagicSearch(){
		
		var msSolicitante;
		var urlArticulos;

		msSolicitante = new MagicSearch("", "companyName0", 30, 2, "");
		msSolicitante.setNewURL("comprasStreamElementos.asp?tipo=empresas");		
		msSolicitante.setMinChar(3);
		msSolicitante.setToken(";");
		msSolicitante.onBlur = SeleccionarProveedor;
		msSolicitante.setValue(document.getElementById("CAB_dsProveedor").value);
		
		msArticulos = new MagicSearch('', 'articuloItem0' , 30, 2, "");
		urlArticulos = "comprasStreamElementos.asp?tipo=articulos&all=1";
		msArticulos.setNewURL(urlArticulos);		
		msArticulos.setMinChar(1);
		msArticulos.setToken(";");
		msArticulos.onBlur = SeleccionarArticulo;
		msArticulos.setValue(document.getElementById("ARTDS").value);	
			
	}
	function SeleccionarProveedor(ms){
		var desc = ms.getSelectedItem();
		if (desc.indexOf('-') != -1) {
			var arr = desc.split('-');
			document.getElementById("CAB_idProveedor").value = arr[0];
			document.getElementById("CAB_dsProveedor").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("CAB_idProveedor").value = 0;
				document.getElementById("CAB_dsProveedor").value = "";
				ms.setValue("");
			}	
		}				
	}
	function SeleccionarArticulo(ms){
		var desc = ms.getSelectedItem();
		if (desc.indexOf('|') != -1) {
			var arr = desc.split('|');
			document.getElementById("ARTID").value = arr[0];
			document.getElementById("ARTDS").value = arr[1];
			ms.setValue(arr[1]);
		} else {
			if (desc == ""){
				document.getElementById("ARTID").value = "";
				document.getElementById("ARTDS").value = "";
			}	
		}				
	}
</script>
<style type="text/css">
	.selects { width: 192px;}
	#inputs { width: 192px;}
</style>
</head>

<body onLoad="bodyOnLoad()">

    
<% call GF_TITULO2("kogge64.gif","Reporte articulos pedidos no recibidos") %>
<div id="toolbar"></div>
<br>
<form name="frm" id="frm" action="almacenReporteArticulosPedidosNoRecibidos.asp" method="get">
		<input type="hidden" name="accion" id="accion" value="<%=ACCION_PROCESAR%>"  />
        <table width="80%" border="0" align="center" class="reg_header">
          <tr>
            <td colspan="4" class="reg_header_nav round_border_top_left round_border_top_right">
                <font class="big"><%=GF_Traducir("Reporte Articulos Pedidos No Recibidos")%></font>    </td>
          </tr>
          <tr>
            <td width="10%" class="reg_header_navdos"><%=GF_Traducir("Division")%></td>
            <td width="41%"><%
                            strSQL = "select iddivision id,dsdivision ds from tbldivisiones where iddivision <> 1"
                            Call executeQueryDB(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
                            response.write crearSelect(rs,"division","id","ds") 
                            %>
            </td>
            <td width="11%" class="reg_header_navdos"><%=GF_Traducir("Proveedor")%></td>
            <td width="39%">
                <input type="hidden" id="CAB_idProveedor" name="CAB_idProveedor" value="<%=CAB_idProveedor%>">
                <input type="hidden" id="CAB_dsProveedor" name="CAB_dsProveedor" value="<%=CAB_dsProveedor%>">
                <div id="proveedoresList1"></div>
                <div id="companyName0"></div>
            </td>
          </tr>
          <tr>
            <td class="reg_header_navdos"><%=GF_Traducir("Articulo")%></td>
            <td>
                <div id="articuloItem0"></div>
                <input type="hidden" id="ARTID" name="ARTID" value="<%=ARTID%>">
                <input type="hidden" id="ARTDS" name="ARTDS" value="<%=ARTDS%>">
            </td>
            <td class="reg_header_navdos"><%=GF_Traducir("Pedido")%></td>
            <td><input name="pedido" id ="inputs" type="text" value="<%=pedido%>"/></td>
          </tr>
          <tr>
            <td class="reg_header_navdos"><%=GF_Traducir("Asociado a")%></td>
            <td colspan="3">
				<select id="filtro" name="filtro">
					<option value="<%=FILTRO_TODOS        %>" selected><% =GF_Traducir("Todos") %></option>
					<option value="<%=FILTRO_SIN_OBRA     %>"><% =GF_Traducir("Sin Obra") %></option>
					<option value="<%=FILTRO_CON_OBRA     %>"><% =GF_Traducir("Con Obra") %></option>
					<option value="<%=FILTRO_MANTENIMIENTO%>"><% =GF_Traducir("Obra/Mantenimientos") %></option>
					<option value="<%=FILTRO_INVERSIONES  %>"><% =GF_Traducir("Obra/Inversiones") %></option>
				</select>
            </td>
          </tr>
        </table>
</form>
</body>
</html>
