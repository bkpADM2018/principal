<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
Dim idObra,subcarpeta

idObra = GF_PARAMETROS7("idobra","",6)
subCarpeta = GF_PARAMETROS7("subcarpeta","",6)

CONST CARPETA_OBRAS = "Obras Puertos"

'---------------------------------------------------------------------------
Function obtenerSubdirectorios(pPath,pPathLimpio)
	Dim fso,carpetaRaiz,raiz,coleccionCarpetas,rtrn,i
	
	set fso = server.createObject("Scripting.FileSystemObject") 
	set carpetaRaiz = fso.getFolder(pPath)
	set coleccionCarpetas = carpetaRaiz.subFolders 
	
	if (subCarpeta = "") then
		rtrn = "<option value='"&pPathLimpio&"' selected >" & GF_TRADUCIR("Principal") & "</option>"
	else
		rtrn = "<option value='"&pPathLimpio&"' >" & GF_TRADUCIR("Principal") & "</option>"
	end if
	
	i = 2
	for each carpeta in coleccionCarpetas
		if (ucase(subCarpeta) <> ucase(carpeta.name)) then
			rtrn = rtrn & "<option value='"&  pPathLimpio&"\"&carpeta.name &"' >"&replace(carpeta.name,"_"," ")&"</option>"
		else
			rtrn = rtrn & "<option value='"&  pPathLimpio&"\"&carpeta.name &"' selected >"&replace(carpeta.name,"_"," ")&"</option>"
		end if
		i = i + 1
	next 

	obtenerSubdirectorios = rtrn
End Function
'---------------------------------------------------------------------------
%>

<html>
	<head>

    <link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
    <link rel="stylesheet" href="css/uploadManager.css" type="text/css">
    <link rel="stylesheet" href="css/jquery.fileupload-ui.css" type="text/css">
	<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css">
    
    <style type="text/css">
    	.combo1{
			width:195px;
		}
    </style>
	
    <script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
	<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery.fileupload.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery.fileupload-ui.js"></script>
	<script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jQueryUpload.js"></script>
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
	
		
       
    <script type="text/javascript">
		var myUpload = new JQueryUpload(
								{
									id:"subirIMG",
									folder:"<% =CARPETA_OBRAS & "\/" &  idObra & "\/" & subCarpeta%>",
									filter:"images",
									showFolderFiles:false
								});
		
		function bodyOnLoad(){
			myUpload.show();
			document.getElementById('comboCarpeta').focus() 
		}
		
		function seleccionarCarpeta(){
			var carpeta = document.getElementById('comboCarpeta').value;
			myUpload.setFolder(carpeta);
		}
		
		
		
    </script>
    
    
	</head>
<body onLoad="bodyOnLoad()">

   	<table border="0" align="center" cellpadding="0" cellspacing="1" class="reg_header">
       	<tr>
       	  <td align="right" class="reg_header_nav round_border_top_left"><%=GF_TRADUCIR("Seleccione el trabajo")%>&nbsp;</td>
   	      <td>
          	<select class="combo1" id="comboCarpeta" name="comboCarpeta" onChange="seleccionarCarpeta()">
            	<%=obtenerSubdirectorios(Server.MapPath(CARPETA_OBRAS & "\" &  idobra),CARPETA_OBRAS & "\" &  idobra)%>
            </select>          </td>
		</tr>
		<tr><td>&nbsp;</td></tr>
      	<tr>
            <td colspan="2" align="center"><div id="subirIMG"></div></td>
      </tr>
	</table>
	<br>
	
	
	<table class="reg_header" align="center">
       	<tr><td><% =GF_TRADUCIR("Una vez cargada la imagen cierre esta ventana para continuar") %></td></tr>
	</table>
</body>
</html>