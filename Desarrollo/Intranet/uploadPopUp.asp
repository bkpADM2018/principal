<!--#include file="Includes/procedimientosMG.asp"-->
<%
Dim divId,folder,subFolder,size,files,myfilter,showFolderFiles,index
	
	divId 		= GF_PARAMETROS7("divId","",6)		
	folder 		= GF_PARAMETROS7("folder","",6)		
	subFolder 	= GF_PARAMETROS7("subFolder","",6)		
	size 		= GF_PARAMETROS7("size","",6)		
	showFolderFiles	= GF_PARAMETROS7("showFolderFiles","",6)		
	myfilter 	= GF_PARAMETROS7("filter","",6)	
	maxFiles 	= GF_PARAMETROS7("maxFiles",0,6)	
	index 	= GF_PARAMETROS7("index","",6)	
	
	set objFS = Server.createObject("Scripting.FileSystemObject")
	
	if (objFS.FolderExists(Server.MapPath(folder & subFolder))) then 
		'Obtengo los nombres de los archivos de la carpeta
		set baseFolder = objFS.getFolder(Server.MapPath(folder & subFolder))
		
		for each myFile in baseFolder.Files
			'response.write objFS.GetExtensionName(Server.MapPath("/")&myFile.name)
			files = files & myFile.name & ","
		next
		if (files <> "") then files = left(files,len(files)-1)
		
	end if
%>
<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
		<link rel="stylesheet" href="css/jquery.fileupload-ui.css"	 type="text/css">
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
		
		<script type="text/javascript" src="scripts/channel.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
		<script type="text/javascript" src="scripts/JQueryUpload.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery.fileupload.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery.fileupload-ui.js"></script>
		<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>	
		
		<script>
		
		var myUpload = new JQueryUpload(
								{
									id:"<%=divId%>",
									folder:"<%=folder & subFolder%>",
									subFolder:"<%=subFolder%>",
									size:"big",
									filter:"<%=myfilter%>",
									inPopUp:true,
									showFolderFiles:<%=showFolderFiles%>,
									parentIndex:<%=index%>,
									maxFiles:<%=maxFiles%>
								});
		
		function bodyOnLoad(){
			myUpload.show();
		}
		</script>
	</head>
	<body onload="bodyOnLoad()">
		<div id="<%=divId%>"></div>
	</body>
</html>