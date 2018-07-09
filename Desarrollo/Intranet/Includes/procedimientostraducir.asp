<%

'               **************************************
'               *    PROCEDIMIENTOS DE TRADUCION     *
'               **************************************
'________________________________________________________________________________________________________
LANG_SPANISH = 1
LANG_ENGLISH = 2

'Autor: Javier A. Scalisi
'Fecha: 28/04/2003
function GF_SET_IDIOMA(P_ID)
'Esta funcion setea el idioma de la pagina.
   if ( P_ID = "") then
       session("UsuarioIdiomaCodigo")= LANG_SPANISH
   else
	   session("UsuarioIdiomaCodigo")= P_ID
   end if	
end function
'-------------------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 28/04/2003
function GF_GET_IDIOMA()
   if (session("UsuarioIdiomaCodigo") = "") then
      GF_SET_IDIOMA(LANG_SPANISH)
      GF_GET_IDIOMA=session("UsuarioIdiomaCodigo")
   else
      GF_GET_IDIOMA=session("UsuarioIdiomaCodigo")
   end if
end function
'-----------------------------------------------------------------------
function GF_Traducir(P_texto)
dim my_texto, con, sql, rs, my_idioma, my_existe, my_id, ret
ret = p_texto
my_idioma = CInt(session("usuarioidiomacodigo"))

if my_idioma > LANG_SPANISH and my_idioma < 4 and len(p_texto) > 0 then
   my_existe = 0
   while my_existe = 0
   sql = "SELECT * FROM TABLATEXTOS1 WHERE TEXTO = '" & P_texto & "'"
   Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", SQL)
   IF  rs.eof then 
   	   Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", "SELECT max(id_texto) as max_id FROM TABLATEXTOS1 ")
       my_id = rs("max_id") + 1
       Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXECUTE", "Insert Into tablatextos1 (id_texto, Texto)  Values ( " & my_id & ", '" & p_texto & "')") 
       Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", SQL) 
	else	  
      my_existe = 1
	  my_id = rs("ID_TEXTO")
	  Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", SQL) 
   end if
   wend
   my_existe = 0
   while my_existe = 0
   SQL = "SELECT * FROM TABLATEXTOS" & my_idioma & " WHERE ID_TEXTO = " & MY_ID 
   Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", SQL) 
   if rs.eof then   
      sql = "Insert Into tablatextos" & my_idioma & " (id_texto, Texto)  Values ( " & my_id & ", '?" & P_TEXTO & "')"
      Call executeQueryDb(DBSITE_SQL_INTRA, rsIns, "EXECUTE", sql) 	  
	else
	  my_existe = 1
      ret = RS("TEXTO")
      Call executeQueryDb(DBSITE_SQL_INTRA, rs, "CLOSE", SQL)
   end if
   wend
END IF
GF_Traducir = ret
end function	
%>
