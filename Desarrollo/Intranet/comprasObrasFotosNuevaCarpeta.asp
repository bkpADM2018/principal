<!--#include file="Includes/procedimientosMG.asp"-->	
<!--#include file="Includes/procedimientostraducir.asp"-->	

<%
	Dim idObra, nombreCarpeta,fso,carpeta,carpetaCreada,msgError,submitio
	CONST CARPETA_OBRAS = "Obras Puertos"
	
	
	nombreCarpeta = GF_Parametros7("folderName","",6)
    idObra = GF_Parametros7("idObra","",6)
	submitio = GF_Parametros7("aceptar","",6)
	msgError = ""
	
	
	carpetaCreada = false
	if (submitio <> "") then
		if (nombreCarpeta <> "") then
			nombreCarpeta = replace(nombreCarpeta," ","_")
			
			carpeta = Server.MapPath(CARPETA_OBRAS & "\" & idobra & "\" & nombreCarpeta)
			
			Set fso = CreateObject("Scripting.FileSystemObject")
		
			if (Not fso.FolderExists(carpeta)) then
				Set fol = fso.CreateFolder(carpeta)
				carpetaCreada = True
			else
				msgError = GF_TRADUCIR("El trabajo ya existe")
			end if
		else
			msgError = GF_TRADUCIR("Debe darle un nombre al trabajo")
		end if
	end if
%>
<html>
<head>
	<title></title>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">

    <style type="text/css">
    	.nombre{
			width:200px;
		}
    </style>
    
	<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
    <script type="text/javascript">
		var popUpNewWork;
    	function bodyOnLoad(){
    		popUpNewWork = getObjPopUp('PopUpAddNewWork');
    		<% if (carpetaCreada) then%>
    			popUpNewWork.hide();
    		<% end if %>
			document.getElementById('folderName').focus()
		}
    </script>
	
</head>

<body onLoad="bodyOnLoad()">
    <form method="POST">
    
    <table width="365" border="0" align="center" cellpadding="1" cellspacing="1" class="reg_header">
      <tr>
    
        <td width="150" align="right" class="reg_header_nav round_border_top_left"><%=GF_TRADUCIR("Nombre Trabajo")%>&nbsp;</td>
        <td width="200"><input name="folderName" type="text" class="nombre round_border_top_right" id="folderName" value="" maxlength="20"></td>
      </tr>
        <% if (msgError <> "") then%>
          <tr class="reg_header_error">
            <td colspan="2"><%=msgError%></td>
          </tr>
        <% end if %>
      <tr>
        <td height="25" colspan="2" align="right" class="reg_header_nav round_border_bottom">
            <input type="submit" id="aceptar" name="aceptar" value="<%=GF_TRADUCIR("Aceptar")%>" class="round_border_bottom_right">&nbsp;
        </td>
      </tr>
    </table>
    <input type="hidden" id="idobra" name="idobra" value="<%=idObra%>">
    </form>
</body>

</html>
