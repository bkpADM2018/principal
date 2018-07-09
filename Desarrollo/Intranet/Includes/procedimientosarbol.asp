<%
'Autor: Javier A. Scalisi
'Fecha: 15/01/2003
function GF_IMAGEN(P_intKR,P_intNumero)
' Esta funcion devuelve la imagen correspondiente para el dato en formato HTML.
Dim strImagen
strImagen = GF_IMAGENPATH(P_intKR,P_intNumero)
if (strImagen = "") then
      GF_IMAGEN = ""
   ELSE
   	  GF_IMAGEN = "<IMG src='" & strImagen & "' border='0'>"
end if
end function
'--------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 15/01/2003
function GF_IMAGENPATH(P_intKR,P_intNumero)
' Esta funcion devuelve la imagen asociada a un registro.
Dim strImagen,strDato
select case CInt(P_intNumero)
   case 1: strDato="__IMAGEN"
   case 2: strDato="__IMAGEN2"
end select   
strImagen= GF_DT1KR("READ",GF_SESSIONKR("SD",strDato),"","",P_intKR)
if (strImagen = "?") then
   GF_IMAGENPATH = ""
else
   GF_IMAGENPATH = strImagen
end if		
end function
'---------------------------------------------------------------------------------------------
Sub LP_OBTENER_HIJOS(P_Kr, ByRef rs, intSRKR)

Dim strSQL, con
strSQL = "Select * From RelacionesConsulta WHERE  SRO1KR = " & intSRKR & " AND SRO2KR = " & P_Kr & " AND SRO3KM IN('UC','SP') and SRVALOR<>'*' "
strSQL = strSQL & "and SRVALOR <> 'HIDDEN' and SRVALOR <> 'Hidden' and SRVALOR <> 'hidden' ORDER BY SRO3KM desc, SRO3KC asc,SRO3DS asc"
GF_BD_CONTROL rs ,con ,"OPEN",strSQL 
end sub
'---------------------------------------------------------------------------------------------
sub LP_COLGAR(rs,P_strPadre, ByRef P_intIX, ByRef P_tblCargos,ByRef P_File)

Dim strNodo, strDSCargo, strDSTarget

while not rs.eof 
   strNodo= "N" & rs("SRO3KC")
   if rs("SRO3KM") = "UC" then
	  if (P_strPadre <> strNodo) then 	  
	     'Cuelgo el cargo solo si debe ser mostrado.
	     if (UCase(GF_DT1("READ","*MOSTRAR","","",rs("SRO3KM"),rs("SRO3KC"))) <> "NO") then
		    'Antes de colgar el nodo se busca la descripcion a mostrar.
		    strDSCargo = GF_DT1("READ","DSCVER","","",rs("SRO3KM"),rs("SRO3KC"))
		    if (strDSCargo = "?") then strDSCargo = rs("SRO3DS")
            P_File.WriteLine("mD.neu(new VE('" & strNodo & "','" & P_strPadre & "','" & GF_TRADUCIR(strDSCargo) & "','','" & GF_IMAGENPATH(rs("SRO3KR"),1) & "','" & GF_IMAGENPATH(rs("SRO3KR"),2) & "',''));")
		 else
		    'Si no se muestra el cargo se setea como padre de sus hijos al padre 
			'del nodo cargo
		    strNodo = P_strPadre
		 end if
		 P_tblCargos(P_intIX,1)= strNodo        'Padre.
		 P_tblCargos(P_intIX,2)= rs("SRO3KR")   'kr del Padre.
	     P_intIX=P_intIX+1
		 if (P_intIX = 100) then P_intIX=0
	   end if
	else
	   'Antes de colgar el nodo se busca la descripcion a mostrar.
	    strDSCargo = GF_DT1("READ","DSCVER","","",rs("SRO3KM"),rs("SRO3KC"))
		if (strDSCargo = "?") then strDSCargo = rs("SRO3DS")
	    'Obtengo el target del procedimiento
	    strDSTarget = GF_DT1("READ","TARGET","","",rs("SRO3KM"),rs("SRO3KC"))
		if (strDSTarget = "?") then strDSTarget = "RightFrame"
	    P_File.WriteLine("mD.neu(new LE('" & P_strPadre & "','" & GF_TRADUCIR(strDSCargo) & "','MG523.asp?P_KR=" & rs("SRO3KR") & "','" & strDSTarget & "','" & GF_IMAGENPATH(rs("SRO3KR"),1) & "',''));")
   end if     
   rs.movenext
wend

end sub   
'---------------------------------------------------------------------------------------------
Sub LP_ARMAR_ARBOL(P_intKR,ByRef P_File,p_rel)
        
Dim intINDX, intINDX2
Dim rs,intSRKR
Dim tblCargos(99,2)

		intINDX=0
		intINDX2=0
		intSRKR = GF_sessionKr("SR",p_rel)
		'Genero el primer arbol.
        LP_OBTENER_HIJOS P_intKR,rs,intSRKR
        LP_COLGAR rs,"root",intINDX,tblCargos,P_File
        'Cuelgo todos hijos en los cargos que corresponde.
		while (intINDX > intINDX2) 
           LP_OBTENER_HIJOS tblCargos(intINDX2,2),rs,intSRKR
           LP_COLGAR rs,tblCargos(intINDX2,1),intINDX,tblCargos,P_File
           intINDX2=intINDX2+1
        wend

end sub		
'-------------------------------------------------------------------------
sub CrearArbol(strKMCargo,strKCCargo,intKRCargo,P_chkExternos,P_REL) 

Dim strSQL,intIdiomaAUX,oConn,rsIdiomas,intKR,fso,Archivo,strDSCargo

if (GF_MGKS(strKMCargo,strKCCargo,intKRCargo,strDSCargo)) then ' {2}
   'Se tiene un usuario valido.
   GP_CONFIGURARMOMENTOS
   'Se setea el idioma del arbol.
   strSQL= "Select Id_Idioma from TablaIdiomas"
   GF_BD_CONTROL rsIdiomas,oConn,"OPEN",strSQL
   intIdiomaAUX =GF_GET_IDIOMA()
   while not(rsIdiomas.eof)
   GF_SET_IDIOMA(rsIdiomas("Id_Idioma"))
   Set fso = CreateObject("Scripting.FileSystemObject")
   'Creo el Archivo con el arbol del Usuario.
   Set Archivo = fso.CreateTextFile(server.mappath(".") & "\UC-TREES\A-" & strKCCargo & "-" & rsIdiomas("Id_Idioma") & ".js",true)
   'Se inicializa la funcion.
   Archivo.WriteLine("function CrearArbol()")
   Archivo.WriteLine("// Esta funcion crea el arbol del cargo " & strKCCargo)
   Archivo.WriteLine("// Descripcion del Arbol: " & strDSCargo)
   Archivo.WriteLine("// Momento inicio Generacion: " & now())
   Archivo.WriteLine("{")
   Archivo.WriteLine("mD=new Satz();")
   Archivo.WriteLine("mD.neu(new HVE('root','Home Buenos Aires','globus-0.gif','Home Buenos Aires'));")   
   Archivo.WriteLine("// Nodos del Cargo Publico")   
   'Se cuelgan los nodos publicos.
   if (P_chkExternos = "CHECKED") then
      intKR = GF_sessionKr("UC","PUBLIC") 'Para usuarios externos.
   else
      intKR = GF_sessionKr("UC","PUBLICB") 'Para usuarios internos.
   end if 	  
   Call LP_ARMAR_ARBOL(intKR,Archivo,P_REL)
   Archivo.WriteLine("// Nodos del Usuario")   
   'Se cuelgan todos los nodos que corresponden a el usuario.
   Call LP_ARMAR_ARBOL(intKRCargo,Archivo,P_REL)
   'Se cierra la Funcion.
   Archivo.WriteLine("// Momento fin Generacion: " & now())
   Archivo.WriteLine("}")
   'Se cierra el archivo.
   Archivo.Close
   rsIdiomas.MoveNext
   wend
   GF_SET_IDIOMA(intIdiomaAUX)
end if ' {2}    
end sub
%>
