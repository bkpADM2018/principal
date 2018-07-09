<!--#include file="Includes/procedimientosMap.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 03/12/2008                         -->
'<!--               Pagina     : MapLoadAjax.ASP                    -->
'<!--               Descripcion: Ajax - Carga de figuras			-->
'<!------------------------------------------------------------------->
dim FrmDic, i
dim rs, conn, sql, myWhere
dim myShapes
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary")
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
if len(FrmDic("byGroup")) > 0 then 
	if (FrmDic("byGroup")="ALL") then
		myWhere = myWhere & " and D.drawingGroup <> 0 "
	else
		myWhere = myWhere & " and D.drawingGroup in(" & FrmDic("byGroup") & ")"	
	end if	
end if	
if len(FrmDic("byName")) > 2 then myWhere = myWhere & " and D.dsShape in("  & FrmDic("byName") & ")"
if len(FrmDic("byUser")) > 2  then myWhere = myWhere & " and D.owner in(" & FrmDic("byUser") & ")"
if len(FrmDic("byCommentsByUser")) > 2  then myWhere = myWhere & " and [user] in(" & FrmDic("byCommentsByUser") & ")"
if len(FrmDic("byComment")) > 0  then myWhere = myWhere & " and C.comment like '%" & FrmDic("byComment") & "%'"
if len(myWhere) > 0 then
	sql = "Select D.drawingType, D.idShape, D.dsShape, D.drawingCoords, D.drawingGroup, D.drawingColor, D.drawingIcon, D.drawingComment, C.comment from Drawings D left join Comments C on D.idShape = C.idShape where 1=1 " & myWhere
	'Response.Write sql
	'Response.end
	call GF_BD_Control_Map(rs, conn, "OPEN", sql)
		while not rs.eof
				myShapes = myShapes & "[" & rs("drawingType") & "||" & rs("drawingCoords") & "||" & rs("idShape") & "||" & rs("dsShape") & "||" &  rs("drawingGroup") & "||" & rs("drawingColor") & "||" & rs("drawingIcon")  & "||" & rs("drawingComment") & "||" & rs("comment") & "]"
			rs.movenext
		wend
	call GF_BD_Control_Map(rs, conn, "CLOSE", sql)
	Response.write myShapes 
end if
'------------------------------------------------------------------------------------------------
function getComments(pId)
dim myComments, rs2
	sql = "Select * from Comments where idShape=" & pId & " order by moment desc"
	'Response.Write sql
	'Response.End 
	call GF_BD_Control_Map(rs2, conn, "OPEN", sql)
		while not rs2.eof
				myComments = myComments & "//" & rs2("idComment") & "::" & rs2("user") & "::" & rs2("comment") & "::" & rs2("moment")
			rs2.movenext
		wend
getComments = myComments 		
end function
%>