<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosJSON.asp"-->
<%
Const MENU_ITEM_NODE = "UC"
Const MENU_ITEM_LEAF = "SP"

Const MENU_ITEM_TYPE = "type"
Const MENU_ITEM_PARENT = "parent"
Const MENU_ITEM_NAME = "name"
MENU_ITEM_TEXT = "desc"
MENU_ITEM_LINK = "link"
MENU_ITEM_TARGET= "target"

Function LoadNodeChilds(parentName) 

	Dim krRel, krNode, dsRel, dsNode, oConn, ret
	Dim krParent, dsParent, link, target, nodetype
	Dim nodeName
	'Tomo el KR de la relacion.
	Call GF_MGKS("SR","EXEC", krRel, dsRel)
	'Tomo el KR del nodo.
	Call GF_MGKS(MENU_ITEM_NODE,parentName, krParent, dsParent)
	
	'Tomo los nodos y hojas dependientes del nodo en análisis
	strSQL = "Select * From RelacionesConsulta WHERE  SRO1KR = " & krRel & " AND SRO2KR = " & krParent & " AND SRO3KM IN('UC','SP') and SRVALOR not in ('*', 'HIDDEN') ORDER BY SRO3KM desc, SRO3DS asc, SRO3KC asc"	
	Call GF_BD_CONTROL (rs ,oConn ,"OPEN",strSQL)
	Set jsa = jsArray()
	while (not rs.eof)
		nodetype = ""	'Esta variable es de control para que no se agreguen nodos u hojas sin texto que mostrar o que no se deban mostrar.
		Set jsa(Null) = jsObject()
		if (rs("SRO3KM") = MENU_ITEM_NODE) then
			'Cuelgo el nodo solo si debe ser mostrado.
			if (UCase(GF_DT1("READ","*MOSTRAR","","",rs("SRO3KM"),rs("SRO3KC"))) <> "NO") then
				nodetype = MENU_ITEM_NODE
				'Antes de colgar el nodo se busca la descripcion a mostrar.
				dsNode = GF_DT1("READ","DSCVER","","",rs("SRO3KM"),rs("SRO3KC"))
				if (dsNode = "?") then dsNode = rs("SRO3DS")
				nodeName = rs("SRO3KC")
				target=""
				link=""
			end if
		else
			'Cuelgo la hoja
			nodetype = MENU_ITEM_LEAF
			nodeName = ""
		    dsNode = GF_DT1("READ","DSCVER","","",rs("SRO3KM"),rs("SRO3KC"))
			if (dsNode = "?") then dsNode = rs("SRO3DS")
		    'Obtengo el target del procedimiento
		    target = GF_DT1("READ","TARGET","","",rs("SRO3KM"),rs("SRO3KC"))	
	        if (target = "?") then target = "MainFrame"
			link = "MG523.asp?P_KR=" & rs("SRO3KR")
		    'JAS --> P_File.WriteLine(P_VAR & ".AddLeaf('" & P_strPadre & "','" & GF_TRADUCIR(strDSCargo) & "','MG523.asp?P_KR=" & rs("SRO3KR") & "','" & strDSTarget & "','','" & GF_IMAGENPATH(rs("SRO3KR"),1) & "');")
		end if  	
		if (nodetype <> "") then
			jsa(Null) (MENU_ITEM_TYPE) = trim(nodetype)
			jsa(Null) (MENU_ITEM_PARENT) = trim(parentName)
			jsa(Null) (MENU_ITEM_NAME) = trim(nodeName)
			jsa(Null) (MENU_ITEM_TEXT) = trim(dsNode)
			jsa(Null) (MENU_ITEM_LINK) = trim(link)
			jsa(Null) (MENU_ITEM_TARGET) = trim(target)
		end if
		rs.MoveNext()
	wend	
	Set LoadNodeChilds = jsa
   
End Function
'*****************************************************************
'*****		COMIENZO DE LA PAGINA			 *****
'*****************************************************************
Dim nodeName, action

nodeName = GF_PARAMETROS7("node", "", 6)
action = GF_PARAMETROS7("action", "", 6)

if (action <> "") then 	
	LoadNodeChilds(nodeName).Flush	
	response.end
end if

%>
<HTML>
<HEAD>
<link rel="stylesheet" href="css/main.css" type="text/css">
<link rel="stylesheet" href="css/Menu.css" type="text/css">

<script src="scripts/jquery.min.js"></script>
<script language="javascript" src="scripts/Menu.js"></script>
<script language="javascript">
	$(document).ready(function(){
		var vMenu= new Menu('<% =Request.QueryString("P_MNU") %>');		
		vMenu.Run();
	});
</script>

</HEAD>
<BODY leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>
</BODY>
</HTML>
