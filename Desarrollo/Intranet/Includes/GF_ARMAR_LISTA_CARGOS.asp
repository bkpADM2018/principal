<%
'Autor: Javier A. Scalisi
'Fecha: 27/02/2003

Sub LP_OBTENER_CARGOS(P_Kr, ByRef rs,intSRKR)

Dim strSQL
Dim con

strSQL = "Select * From RelacionesConsulta WHERE  SRO1KR = " & intSRKR & " AND SRO2KR = " & P_Kr & " AND SRO3KM='UC' AND SRVALOR<>'*' ORDER BY SRO3KR"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
'GF_BD_CONTROL rs,con,"OPEN",strSQL 
end sub
'---------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 27/02/2003
sub LP_ADD_A_LISTA(rs,ByRef P_intIX, ByRef P_tblCargos)

Dim strNodo, strDS

while not rs.eof 
   strNodo=rs("SRO3KR")
   if rs("SRO3KM") = "UC" then
	  'Verifico que el cargo no este en la lista.
	  if (InStr(session("LISTADECARGOS"),strNodo) = 0) then 	  
	     'Cuelgo el cargo.
		 if (session("LISTADECARGOS") = "(") then
		    session("LISTADECARGOS")= session("LISTADECARGOS") & strNodo
		 else 
		    session("LISTADECARGOS")= session("LISTADECARGOS") & "," & strNodo
	     end if		
		 P_tblCargos(P_intIX,1)= strNodo        
		 P_intIX=P_intIX+1
		 if (P_intIX = 100) then P_intIX=0
	   end if
   end if  
   rs.movenext
wend
end sub   
'---------------------------------------------------------------------------------------------\
'Autor: Javier A. Scalisi
'Fecha: 27/02/2003
function GF_ARMAR_LISTA_CARGOS(P_intKRObjeto)
	    
Dim intINDX, intINDX2
Dim rs,intKRAcceso
Dim tblCargos(99,2)
   
		intINDX=0
		intINDX2=0
		intKRAcceso=GF_SESSIONKR("SR","EXEC")
        session("LISTADECARGOS")="("
		'Tomo los cargos de asociacion directa.
		LP_OBTENER_CARGOS P_intKRObjeto,rs,intKRAcceso
        LP_ADD_A_LISTA rs,intINDX,tblCargos
        'Cuelgo todos los demas cargos.
		while (intINDX > intINDX2) 
           LP_OBTENER_CARGOS tblCargos(intINDX2,1),rs,intKRAcceso
           LP_ADD_A_LISTA rs,intINDX,tblCargos
           intINDX2=intINDX2+1
        wend
		session("LISTADECARGOS")=session("LISTADECARGOS") & ")"				

end function		
%>
