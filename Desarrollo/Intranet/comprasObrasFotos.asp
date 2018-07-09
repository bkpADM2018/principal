<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
CONST CARPETA_OBRAS = "Obras Puertos"
Dim objFS, BaseFolder,idobra,ext,extHabilitadas,nomObra,existeCarpeta,origen
Dim linkVolver,subcarpeta,subcarpeta1,ruta,tituloAddImg,tituloAddFolder,accion

'------------------------------------------------------------------
Function obtenerSubdirectorios(pPath)
	Dim fso,carpetaRaiz,raiz,coleccionCarpetas,rtrn
	
	set fso = server.createObject("Scripting.FileSystemObject") 
	set carpetaRaiz = fso.getFolder(pPath)
	set coleccionCarpetas = carpetaRaiz.subFolders 
	
	if (subCarpeta1 <> "") then
		rtrn = "<a href=comprasObrasFotos.asp?idObra=" & idObra & "&origen="&origen & ">" & GF_TRADUCIR("Principal") & "</a> | "
	else
		rtrn = GF_TRADUCIR("Principal") & " | "
	end if
	for each carpeta in coleccionCarpetas
		if (ucase(subCarpeta1) <> ucase(carpeta.name)) then
			rtrn = rtrn & "<a href=comprasObrasFotos.asp?idObra=" & idObra & "&subcarpeta=" &  carpeta.name & "&origen="&origen & ">" & replace(carpeta.name,"_"," ") & "</a> | "
		else
			rtrn = rtrn & replace(carpeta.name,"_"," ") & " | "
		end if
		
	next 
	
	rtrn = left(rtrn,len(rtrn)-2)
	
	obtenerSubdirectorios = rtrn
End Function
'------------------------------------------------------------------
Function extensionHabilitada(pExtension)
	Dim rtrn
	rtrn = false

	if ( instr( ucase(extHabilitadas),ucase(pExtension))  ) then rtrn  = true
	
	extensionHabilitada = rtrn
End Function

'*****************************************************************
'					INICIO PAGINA
'*****************************************************************

extHabilitadas = "jpg,gif,png"
existeCarpeta = true

idobra = GF_Parametros7("idObra","",6)
origen = GF_Parametros7("origen","",6)
accion = GF_Parametros7("accion","",6)
borrarPic = GF_Parametros7("urlborrar","",6)
subCarpeta1 = GF_Parametros7("subcarpeta","",6)

if (subCarpeta1 <> "") then subCarpeta = "/" & subCarpeta1

ruta = idObra & subCarpeta

if (origen = "" or origen = "comprasTableroObra") then 
	origen = "comprasTableroObra"
	linkVolver = "comprasTableroObra.asp?idObra=" & idObra
else
	linkVolver = origen & ".asp"
end if


set objFS = Server.createObject("Scripting.FileSystemObject")

if (accion = "borrar") then 
	dim fs
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
	
	auxChar = 0
	i = 0
	while auxChar = 0
		if (mid(borrarPic,len(borrarPic)-i,1) = "/") then auxChar = i
		i = i + 1
	wend
	
	pathFileBorrar = Server.MapPath(CARPETA_OBRAS & "\" &  ruta & right(borrarPic,auxChar+1))	

    if fs.FileExists(pathFileBorrar) then fs.DeleteFile(pathFileBorrar)
    Set fs = Nothing
end if

if (not objFS.FolderExists(Server.MapPath(CARPETA_OBRAS & "\" &  idobra))) then 
	objFS.CreateFolder(Server.MapPath(CARPETA_OBRAS & "\" &  idobra))
end if

set baseFolder = objFS.getFolder(Server.MapPath(CARPETA_OBRAS & "\" & ruta))


tituloAddImg    = "<table><tr><td><img src='images/compras/Add-Picture-icon-16x16.png' alt='agregar'></td><td class='titulo'>&nbsp;&nbsp;"&GF_TRADUCIR("Agregar Foto")&"</td></tr></table>"
tituloAddFolder = "<table><tr><td><img src='images/compras/new_folder-16x16.png' alt='agregar'></td><td class='titulo'>&nbsp;&nbsp;"&GF_TRADUCIR("Nuevo Trabajo")&"</td></tr></table>"


%>

<html>
<head>
	<title>:: Fotos de las Obras ::</title>
    <link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/galleriffic-2.css" type="text/css" />
    <link type="text/css" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" rel="stylesheet" />	
    
    <style type="text/css">
		body { font-size: 62.5%; }

		.titulo{
			color:#FFFFFF;
			font-weight:bold;
		}
	</style>
    
	<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery.galleriffic.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery.opacityrollover.js"></script>
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>

	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.7.3.custom.min.js"></script>
    
    <script type="text/javascript">
		function addNewWork(){
			var puw = new winPopUp('PopUpAddNewWork', 'comprasObrasFotosNuevaCarpeta.asp?idobra=<%=idObra%>', '500', '120', '<% =GF_TRADUCIR("Agregar Carpeta") %>', 'location.reload();');
		}

		function addNewPicture(){
			var puw = new winPopUp('PopUpAddNewPicture', 'comprasObrasFotosUpload.asp?idObra=<%=idobra%>&subcarpeta=<%=subCarpeta1%>', '600', '420', '<% =GF_TRADUCIR("Agregar Foto") %>', 'location.reload();');
		}

		function cerrar() {
			$('#agregar').dialog('close'); 
		}
	
		function borrarImg(){
			if (document.getElementById("actual")) {
				imgUrl = document.getElementById("actual").src;
				location.href = "comprasObrasFotos.asp?idobra=<%=idobra%>&origen=<%=origen%>&accion=borrar&urlborrar="+imgUrl+"&subcarpeta=<%=subCarpeta1%>"
			}
		}
	
		document.write('<style>.noscript { display: none; }</style>');
    </script>
  
</head>

<body>
<p class="nav" align="center">
<table width="770px" border="0" align="center">
  <tr>
    <td align="center"><a href="<%=linkVolver%>"><img src="images/compras/Previous-16x16.png" border="0" alt="<%=GF_TRADUCIR("Atras")%>"></a></td>
    <td align="center">&nbsp;<a href="<%=linkVolver%>">Volver</a></td>
    <td width="700" align="center"><h1><%=getDescripcionObra(IdObra)%></h1></td>
    <td>
     
    </td>

  </tr>
</table>
<table width="770px" border="0" align="center" class="reg_header">
  		<tr>
  		  <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
            <tr class="reg_header">
              <td width="93%" align="center" class="reg_header_nav round_border_top_left">Trabajos</td>
              <td width="7%" align="center" class="reg_header_nav round_border_top_right" style="cursor:pointer;">
				<img border="0" title="<%=GF_TRADUCIR("Agregar Trabajo")%>" src="images/compras/new_folder-16x16.png" onClick="addNewWork();">
              </td>
            </tr>
            <tr class="reg_header">
              <td colspan="2" align="center" bgcolor="#FFFAF0" class="round_border_bottom"><%=obtenerSubdirectorios(Server.MapPath(CARPETA_OBRAS & "\" &  idobra))%> </td>
            </tr>
			
			
            
          </table></td>
  </tr>
  		<tr>
  		  <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="reg_header">
            <tr bgcolor="#FFFAF0">
            	<td >
                    <table align="right">
                        <tr>
							<% if (baseFolder.Files.Count > 0) then %> 
							<td><%=GF_TRADUCIR("Borrar Imagen")%></td>
							<td align="left">
								<img src="images/compras/remove-16x16.png" title="Borrar" alt="Borrar Foto" onclick="borrarImg();" style="cursor:pointer">
							</td>
							<% end if %>
                            <td><%=GF_TRADUCIR("Agregar")%></td>
                            <td style="cursor:pointer;">
								<img src="images/compras/Add-Picture-icon-16x16.png" title="<%=GF_TRADUCIR("Agregar Foto")%>" alt="agregar" onClick="addNewPicture();">
                            </td>
                            <td width="12px"></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
<% if (baseFolder.Files.Count > 0) then %>   
              <td bgcolor="#FFFAF0" class="round_border_all">
            <div id="page">
                <div id="container">
                  <!-- Start Advanced Gallery Html Containers -->
                  <div id="gallery" class="content">
                    <div id="controls" class="controls"></div>
                    <div class="slideshow-container">
                      <div id="loading" class="loader"></div>
                      <div id="slideshow" class="slideshow"></div>
                    </div>
                  </div>
                  <div id="thumbs" class="navigation">
                    <ul class="thumbs noscript">
                      <%
						dim i
						if (baseFolder.Files.Count > 0) then%>
                      <%for each Foto in baseFolder.Files%>
                      <% 
                                ext = ucase(objFS.GetExtensionName(CARPETA_OBRAS & "/"&baseFolder.name&"/"&Foto.name))
                                if (extensionHabilitada(ext)) then %>
                      <li> <a class="thumb" name="leaf" href="<%=CARPETA_OBRAS%>/<%=ruta&"/"&Foto.name%>" > <img src="<%=CARPETA_OBRAS%>/<%=ruta&"/"&Foto.name%>" alt="" width="75" height="75"/> </a> </li>
                      <%end if%>
                      <%next%>
                      <%end if%>
                    </ul>
                  </div>
                  <div style="clear: both;"></div>
                </div>
        </div>
        </td>
            <% else %>
            	<td bgcolor="#FFFAF0" class="reg_header_navdos round_border_bottom" align="center">
					La obra no tiene fotos actualmente.                </td>
			<% end if %>
            </td>
            </tr>
          </table></td>
  </tr>
</table>
    
</p>
<br />

</div>
<% if (baseFolder.Files.Count > 0) then %>   
	<script type="text/javascript">
			jQuery(document).ready(function($) {
				// We only want these styles applied when javascript is enabled
				$('div.navigation').css({'width' : '200px', 'float' : 'left'});
				$('div.content').css('display', 'block');

				// Initially set opacity on thumbs and add
				// additional styling for hover effect on thumbs
				var onMouseOutOpacity = 0.67;
				$('#thumbs ul.thumbs li').opacityrollover({
					mouseOutOpacity:   onMouseOutOpacity,
					mouseOverOpacity:  1.0,
					fadeSpeed:         'fast',
					exemptionSelector: '.selected'
				});
				
				// Initialize Advanced Galleriffic Gallery
				var gallery = $('#thumbs').galleriffic({
					delay:                     2500,
					numThumbs:                 10,
					preloadAhead:              10,
					enableTopPager:            true,
					enableBottomPager:         true,
					maxPagesToShow:            7,
					imageContainerSel:         '#slideshow',
					controlsContainerSel:      '#controls',
					captionContainerSel:       '#caption',
					loadingContainerSel:       '#loading',
					renderSSControls:          true,
					renderNavControls:         true,
					playLinkText:              '<%=GF_TRADUCIR("Iniciar Diapositivas")%>',
					pauseLinkText:             '<%=GF_TRADUCIR("Parar Diapositivas")%>',
					prevLinkText:              '&lsaquo; <%=GF_TRADUCIR("Anterior")%>',
					nextLinkText:              '<%=GF_TRADUCIR("Siguiente")%> &rsaquo;',
					nextPageLinkText:          '<%=GF_TRADUCIR("Siguiente")%> &rsaquo;',
					prevPageLinkText:          '&lsaquo; <%=GF_TRADUCIR("Anterior")%>',
					enableHistory:             false,
					autoStart:                 false,
					syncTransitions:           true,
					defaultTransitionDuration: 900,
					onSlideChange:             function(prevIndex, nextIndex) {
						// 'this' refers to the gallery, which is an extension of $('#thumbs')
						this.find('ul.thumbs').children()
							.eq(prevIndex).fadeTo('fast', onMouseOutOpacity).end()
							.eq(nextIndex).fadeTo('fast', 1.0);
					},
					onPageTransitionOut:       function(callback) {
						this.fadeTo('fast', 0.0, callback);
					},
					onPageTransitionIn:        function() {
						this.fadeTo('fast', 1.0);
					}
				});
			});
	</script>
<% end if %>   
</body>
</html>
