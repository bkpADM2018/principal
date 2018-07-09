<!--#include file="Includes/procedimientosMap.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 03/12/2008                         -->
'<!--               Pagina     : MapSaveComment.ASP                 -->
'<!--               Descripcion: Ajax - Guarda Comentarios			-->
'<!------------------------------------------------------------------->
dim FrmDic, i
dim rs, conn, sql, myWhere
dim myNow
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary")
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
call GP_ConfigurarMomentos

if FrmDic("pAction") = 0 then
	sql = "Insert into Comments values(" & FrmDic("pIdShape") & ",'" & FrmDic("pUser") & "', '" & mid(FrmDic("pComment"),1,3999) & "'," & session("momentosistema") & ")"
elseif FrmDic("pAction") = 1 then
	sql = "Update Comments set [user] = '" & FrmDic("pUser") & "', Comment= '" & mid(FrmDic("pComment"),1,3999) & "', moment=" & session("momentosistema") & " where idShape=" & FrmDic("pIdShape")
elseif FrmDic("pAction") = 2 then
	sql = "Delete from Comments where idShape=" & FrmDic("pIdShape")
elseif FrmDic("pAction") = 3 then
	sql = "Delete from Comments where idShape=" & FrmDic("pIdShape")
	call GF_BD_Control_Map(rs, conn, "EXEC", sql)
	sql = "Delete from Drawings where idShape=" & FrmDic("pIdShape")
end if	
call GF_BD_Control_Map(rs, conn, "EXEC", sql)
Response.write sql 
%>