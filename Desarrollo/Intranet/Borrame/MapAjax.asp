<!--#include file="Includes/procedimientosMap.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<%
'<!------------------------------------------------------------------->
'<!--                        INTRANET ACTISA                        -->
'<!--                                                               -->
'<!--               Programador: Ezequiel A. Bacarini               -->
'<!--               Fecha      : 03/12/2008                         -->
'<!--               Pagina     : MapAjax.ASP	                    -->
'<!--               Descripcion: Ajax - Almacenamiento de figuras	-->
'<!------------------------------------------------------------------->
dim FrmDic, i, j, myDescomposicion, myDescomposicion2
dim rs, conn, sql
dim myIdElement, myDsElement, myDrawingType, myDrawingCoords, myDrawingGroup, myDrawingComment, myDrawingColor, myDrawingIcon, myOwner, myDtCreation
'Se crea el diccionario de parametros.
set FrmDic= CreateObject ("Scripting.Dictionary")
For Each i in Request.QueryString
   FrmDic.Add  i,Request.QueryString(i).item
Next
'Response.write "<hr>" + FrmDic("param") + "<hr>"
myDrawingGroup = 0
if FrmDic("param") <> "" then
	myDescomposicion = split(FrmDic("param"),"[")
	for i=1 to ubound(myDescomposicion)
		myDescomposicion2 = split(myDescomposicion(i),"||")
		myDrawingType = left(myDescomposicion2(0),1)
		'if myDrawingType = "P" then
			'myIdElement = getLastElement()
			myDsElement = myDescomposicion2(1)
			myDrawingCoords = right(myDescomposicion2(0),len(myDescomposicion2(0))-1)
			myDrawingCoords = replace(myDrawingCoords,"((","(")
			myDrawingCoords = replace(myDrawingCoords,"))",")") 
			myDrawingCoords = replace(myDrawingCoords,"$","") 
			myDrawingGroup = myDescomposicion2(2)
			myDrawingColor = myDescomposicion2(3)
			myDrawingIcon = myDescomposicion2(4)
			myOwner = session("Usuario")
			myDtCreation = session("momentoSistema")
			myDrawingComment = myDescomposicion2(5)
			
			'for j=0 to ubound(myDescomposicion2)-1
			'	Response.Write "<hr>I(" & j & ")" & myDescomposicion2(j)
			'next
			call GF_BD_Control_Map(rs, conn, "OPEN", "Select max(idShape) as maxID from Drawings")
			if not rs.eof then
				if isnull(rs("maxID")) then
					myIdElement = 1
				else
					myIdElement = rs("maxID") + 1
				end if	
			else
				myIdElement = 1
			end if	
			call GF_BD_Control_Map(rs, conn, "CLOSE", "")
			sql = "Insert Into Drawings values(" & myIdElement & ",'" & myDsElement  & "','" & myDrawingType & "','" & _
					myDrawingCoords & "'," & myDrawingGroup & ",'" & myDrawingColor & "','" & myDrawingIcon & "','" & _
					myDrawingComment & "','" &  myOwner & "'," & myDtCreation & ")"
			'Response.Write "<hr>" & sql					
			call GF_BD_Control_Map(rs, conn, "EXEC", sql)		
		'end if	
	next
else
	Response.Write "No hay datos para guardar"	
end if
%>
<!--<table width="90%" border=0 align="center" class="reg_header" cellpadding="2" cellspacing="1">
		<td width="50%" align=left>
			<font>
				Nada
			</font>
		</td>
</table>-->
