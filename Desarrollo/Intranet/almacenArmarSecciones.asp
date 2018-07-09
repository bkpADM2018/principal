<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosUser.asp"-->
<%
'**********************************
'****	COMIENZO DE LA PAGINA
'**********************************
Dim seccion, pagina, regXPag, conn, rs, strSQL, verTodos, i, rs1, rrn

seccion = GF_PARAMETROS7("seccion", 0 ,6)
verTodos = GF_PARAMETROS7("todos", "" ,6)
pagina = GF_PARAMETROS7("numeroPagina",0,6)
if (pagina = 0) then pagina = 1
regXPag = GF_PARAMETROS7("registrosPorPagina",0,6)
if (regXPag = 0) then regXPag = 10
rrn = ((pagina-1)*regXPag)+1

Select Case(cint(seccion))
	Case 0:		%>
		<!--#include file="almacenTabAlmacenes.asp"-->
<%	Case 1:		%>
		<!--#include file="comprasTabCategorias.asp"-->
<%	Case 2:		%>
		<!--#include file="comprasTabUnidades.asp"-->
<%	Case 3:		%>
		<!--#include file="comprasTabArticulos.asp"-->
<%	Case 4:		%>
		<!--#include file="comprasTabResponsables.asp"-->
<% 
End Select
%>