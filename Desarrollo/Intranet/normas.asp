<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->

<%
On Error Resume Next

dim oFS, oBaseFolder, contTablas

CONST URL_WEB = "//BAI-VM-INTRA-1/ActiSAintra"

p_baseFolder = GF_Parametros7("p_baseFolder", "", 6)
p_titulo = GF_Parametros7("p_titulo", "", 6)%>
<html>
<head>
	<title>TOEPFER INTERNATIONAL</title>
	<link rel="stylesheet" href="CSS/ActiSAintra-1.css">
	<script language="javascript">
		var isIE = (window.navigator.userAgent.indexOf('MSIE')> 0); 
		var display = new String();
		function mostrarOcultar(p_idTabla) {
			var oTabla = document.getElementById(p_idTabla);
			var oHdn = document.getElementById('hdn' + p_idTabla);

			if (oHdn.value == 'none') {
                //Tengo q mostrar su contenido
				document.getElementById('imgArbol' + p_idTabla).src = 'images/Tminus.gif';
				document.getElementById('imgCarpeta' + p_idTabla).src = 'images/CPTOPN.gif';
				if (isIE){
					display = "block";
				}
				else
				{
					display = "table-row";
				}
				oHdn.value = display;
				var strDisplayStyle=display;
				var strPositionStyle='relative';
			} else {
                //Tengo q ocultar su contenido
				document.getElementById('imgArbol' + p_idTabla).src = 'images/Tplusik.gif';
				document.getElementById('imgCarpeta' + p_idTabla).src = 'images/CPTCLSE.gif';
				oHdn.value = 'none';
    			var strDisplayStyle='none';
				var strPositionStyle='absolute';
			}
			oTabla.style.display = strDisplayStyle;
			oTabla.style.position = strPositionStyle;
		}
	</script>
</head>
<body>
	<%response.write GF_Titulo_4(p_titulo)
	set oFS = Server.CreateObject("Scripting.FileSystemObject")
	set oBaseFolder = oFS.GetFolder(Server.MapPath(p_baseFolder))
	contTablas = 0
	call mostrarContenidoCarpeta(oBaseFolder, URL_WEB & "/" & replace(p_baseFolder, "\", "/"), false)
	set oBaseFolder = nothing%>
</body>
</html>
<%'************************************************************************************************
sub mostrarContenidoCarpeta(byref p_oBaseFolder, byval p_urlPredecesor, byval p_boolVisible)
	dim oChildFolder, oArch

	contTablas = contTablas + 1%>
    <table width="90%" cellpadding=0 cellspacing=2>
	<%for each oChildFolder in p_oBaseFolder.SubFolders
		if p_boolVisible then
			strStyle = "display:block;position:relative;"
			strHdnValue = "table-row"
		    strImagenCarpeta = "CPTOPN"
		    strImagenArbol = "Tminus"
		else
		    strStyle = "display:none;position:absolute;"
		    strHdnValue = "none"
		    strImagenCarpeta = "CPTCLSE"
			strImagenArbol = "Tplusik"
		end if%>
    	<tr>
        	<td colspan=2 align=left>
        	    <a href="javascript:mostrarOcultar(<%=contTablas%>);">
					<img id="imgArbol<%=contTablas%>" src="images/<%=strImagenArbol%>.gif" align=absmiddle><img id="imgCarpeta<%=contTablas%>" src="images/<%=strImagenCarpeta%>.gif" align=absmiddle>
					&nbsp;<font size=4><%=GF_Traducir(oChildFolder.Name)%></font>
				</a>
			</td>
    	</tr>
		<tr id="<%=contTablas%>" style="<%=strStyle%>">
		<input type="hidden" id="hdn<%=contTablas%>" value="<%=strHdnValue%>"/>
	    <td colspan=2 align=left style="padding-left:24px;"><%call mostrarContenidoCarpeta(oChildFolder, p_urlPredecesor & "/" & oChildFolder.Name, false)%></td>
		</tr>
	<%next%>
	<%for each oArch in p_oBaseFolder.Files
		fileDesc = oArch.Name		
		'Se saca el código de nombrado de archivos.		
		if (mid(filedesc, 11,1) = "_") then fileDesc = right(fileDesc,Len(fileDesc)-11)
	%>
		<tr>
  			<td align=left style="padding-left:17px;">
			  	<a href="<%=p_urlPredecesor & "/" & oArch.Name %>" target="_blank">
			  		<img src="images/docs_20.gif" align=absmiddle>&nbsp;<font size=4><%=GF_Traducir(fileDesc)%></font>
			  	</a>
  			</td>
		</tr>
	<%next
	set oArch = nothing
	set oBaseFolder = nothing%>
	</table>
<%end sub%>
