<%
'                        ******************************
'                        **  PROCEDIMIENTOS EXECUTE  **
'                        ******************************
'						 Autor: Javier A. scalisi
'						 Fecha: 13/08/2003
'                        Desc : Procedimientos que estandarizan 
'								el uso de Server.Execute
'--------------------------------------------------------------------------
sub  GP_SERVEREXECUTE(p_asp,ByRef p1,ByRef p2,ByRef p3,ByRef p4, ByRef p5,ByRef p6,ByRef p7,ByRef p8,ByRef p9,ByRef p10)
' Procedimiento que graba parametros en la session para que un server.Execute
' los lea y ejecute sus procedimientos de acuerdo a los parametros.
' El server.execute debe actualizar los parametros en la session a los efectos
' que este procedimiento los retorne.
Dim fso
'Se controla la existencia de la pagina.
Set fso = CreateObject("Scripting.FileSystemObject")   
if (fso.FileExists(SERVER.MapPath(p_asp))) then
	GP_SERVERGUARDAR p_asp,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10
	SERVER.EXECUTE(P_ASP)
	GP_SERVERRECUPERAR p_asp,p1,p2,p3,p4,p5,p6,p7,p8,p9,p10
else
	Response.Redirect "MGMSG.ASP?P_MSG=La pagina " & P_asp & " No existe."	
end if
end sub  
'---------------------------------------------------------------------------------------------
sub  GP_SERVERGUARDAR(p_asp,BYVAL p1,BYVAL p2,BYVAL p3,BYVAL p4, BYVAL p5,BYVAL p6,BYVAL p7,BYVAL p8,BYVAL p9,BYVAL p10)
DIM S,My_Key
S = ")("
My_Key = P_ASP & "/PARAMETROS"
'Response.Write "<br>KEY(" & My_Key & ")"
session(My_key) = S & P1 & S & P2 & S & P3 & S & P4 & S & P5 & S & P6 & S & P7 & S & P8 & S & P9 & S
'Response.Write "KEY(" & session(My_Key) & ")"
end sub
'---------------------------------------------------------------------------------------------
sub  GP_SERVERRECUPERAR(p_asp,ByRef p1,ByRef p2,ByRef p3,ByRef p4, ByRef p5,ByRef p6,ByRef p7,ByRef p8,ByRef p9,ByRef p10)
	DIM S,p,My_Key
	S = ")("
	My_Key = P_ASP & "/PARAMETROS"
	'Response.Write "KEY(" & session(My_Key) & ")"
	p = Split(SESSION(my_key) & s & "" & s & "" & s & "" & s & "" & s & "" & s & "" & s & "" & s & "" & s & "" & s & "",s)
	'response.write session(my_key)
	p1 = p(1)
	p2 = p(2)
	p3 = p(3)
	p4 = p(4)
	p5 = p(5)
	p6 = p(6)
	p7 = p(7)
	p8 = p(8)
	p9 = p(9)
	p10 = p(10)
END sub  
'---------------------------------------------------------------------------------------------
%>
