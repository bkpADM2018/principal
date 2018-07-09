<!--#include file="procedimientosConexion.asp"-->
<%
Const CONEXION_MAP = "MAP"
'---------------------------------------------------------------------------------------------
function GF_BD_Control_Map(byref pRs,byref pCon,pOperacion,byref pSql)
'Está función genera la conexión con la base de datos Black Rock
on error resume next
GF_BD_Control_Map = false
session("strSQL")=pSql
	if(IsEmpty(session("conn" & CONEXION_MAP &  "Alias")))then Call loadConfigFile(CONEXION_MAP)	
	if pOperacion = "CLOSE" THEN
		pRs.close
		pCon.close
		set pCon = nothing
		GF_BD_Control_Map = true
	end if
	if pOperacion = "OPEN" or pOperacion = "UPDATE" then
		set pCon = server.CreateObject("ADODB.connection")
		set pRs = server.CreateObject("ADODB.Recordset")
		pCon.open session("conn" & CONEXION_MAP &  "Alias"), session("conn" & CONEXION_MAP &  "User"), session("conn" & CONEXION_MAP &  "Key")		
		pRs.Open pSql,pCon,1,1
    	GF_BD_Control_Map = true
	end if
	'Se ejecuta una sentencia strSQL sobre la base. 
	if ((pOperacion = "EXECUTE") or (pOperacion = "EXEC")) and (pSql <> "") then 
		set pCon = server.CreateObject("ADODB.connection")
		pCon.open session("conn" & CONEXION_MAP &  "Alias"), session("conn" & CONEXION_MAP &  "User"), session("conn" & CONEXION_MAP &  "Key")		
		'Response.Write "<br> SQL: " & psql
		pCon.execute pSql
		pCon.close
		GF_BD_Control_Map = true
	end if  
end function
'----------------------------------------------------------------
function GF_BD_Control_Paginacion_Map(byref pRs, byref pCon, pOperacion, byref pSql, pPagina, pRegPorPagina)
'Está función genera la conexión con la base de datos y devuelve solo la pagina deseada
dim rsAux,i, element, col, auxArray, intRows, intRecord
call GF_BD_Control_Map(pRs,pCon,pOperacion,pSql)
if ucase(pOperacion) <> "CLOSE" then
	if not pRs.EOF then
		pRs.PageSize     = pRegPorPagina
		pRs.AbsolutePage = pPagina
	end if	
end if	
end function
%>
