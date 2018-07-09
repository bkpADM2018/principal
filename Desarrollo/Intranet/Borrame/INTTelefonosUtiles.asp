<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"--><!--#include file="../ActiSAIntra/Includes/procedimientostraducir.asp"--> 
<!--#include file="Includes/procedimientosfechas.asp"-->
<script language="javascript">
   function SetearOperacion(P_OPR)
   {
      frmAdd.action=frmAdd.action + '?P_OPR=' + P_OPR
   }
</script>
<% 
ProcedimientoControl "TELUTIL"

'-----------------------------------------------
function GF_Memo(p_operacion, p_kr, byref p_MemoCargado)
dim cn, rs, strSQL, strSQLins, strSQLupd
  strSQL = "Select *From MGOrmemo where mg_kr=" & P_KR
  GF_BD_CONTROL rs, cn, "OPEN", strSQL
  if p_operacion = "READ" then
	 if not rs.eof then
	    P_MemoCargado = rs("Or_memo")
	    GF_Memo = true
	 else
	    P_MemoCargado = ""
	    GF_Memo = false
	 end if
  elseif p_operacion = "WRITE" then
  p_MemoCargado = replace(p_MemoCargado, "'", "*")
     if rs.eof then
        strSQLins = "Insert Into MGormemo (mg_kr,OR_memo) Values(" & P_KR & ",'" & P_MemoCargado & "')"
	    cn.execute strSQLins
     else   
        strSQLupd = "Update MGormemo SET or_memo='" & P_MemoCargado & "' where mg_kr=" & P_KR
	    cn.execute strSQLupd
     end if
     GF_Memo = true
  end if
  GF_BD_CONTROL rs, cn, "CLOSE", strSQL
end function
'-----------------------------------------------
Dim P_strAccion,strSQL,oConn,rsTelefonos,intKR,i
Dim P_KC,rsDatos,strMEMO,strTexto,strDS,intAccesoST

dim My_Kc, My_Kr, My_Ds

'Se obtienen los parametros.
P_strAccion= GF_PARAMETROS("P_OPR","")
'Tomo el nivel de acceso al maestro 'ST'.
intAccesoST = GF_controlAccesoKS("ACCESO","SM","ST","","")
'Se procesa la accion pedida.
if (P_strAccion = "ADD") then
   'Se Agrega un nuevo telefono a la tabla.
   if (GF_PARAMETROS("","strKC") <> "") and (GF_PARAMETROS("","strDS") <> "") then
      GF_MGADD "ST", GF_PARAMETROS("","strKC"), GF_PARAMETROS("","strDS"), intKR
      if (GF_PARAMETROS("","strMEMO") <> "") then GF_Memo "WRITE", intKR, GF_PARAMETROS("","strMEMO")
      Response.Redirect("INTTelefonosUtiles.asp?P_OPR=VER&P_KC=" & GF_PARAMETROS("","strKC"))
   end if   
end if
if (P_strAccion = "DLT") then
   MY_KC = GF_Parametros7("P_KC", "", 6) 
   call GF_MGC ("ST", My_KC, My_Kr, My_Ds) 
   strSQL="Delete MGORMEMO where MG_KR=" & My_Kr 
	GF_BD_CONTROL "", oConn, "EXEC", strSQL
   strSQL="Delete MG where MG_KR=" & My_Kr 
	GF_BD_CONTROL "", oConn, "EXEC", strSQL
end if

'Se obtiene todos los datos a mostrar.
strSQL="Select * from MG where MG_KM='ST' order by MG_KC"
GF_BD_CONTROL rsTelefonos,oConn,"OPEN",strSQL

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="CSS/ActisaIntra-1.css" type="text/css">
</HEAD>
<BODY>
<% =GF_TITULO("Telefono.jpg","Telefonos para Requerir Servicios Tecnicos")%>
<% if (intAccesoST > 1) then %>
<table border="0" width="100%"><tr><td align="right">
<a href="INTTelefonosUtiles.asp?P_OPR=NEW">
<img border="0" src="images/Agregar.gif" alt="<% =GF_TRADUCIR("Nuevo") %>">
</a>
</td></tr></table>
<% end if %>
<%
   if not(rsTelefonos.EOF) then
        'Armo el indice.
%>
        <DIV><% =GF_TRADUCIR("INDICE") %>:</DIV><br>
        <table border="0" width="70%">
<%
         while not(rsTelefonos.EOF)
            i=0
%>
            <tr>
<%         
            while not(rsTelefonos.EOF) and (i<5) 
%>
               <td align="center" width="25%"><a href="INTTelefonosUtiles.asp?P_OPR=VER&P_KC=<% =rsTelefonos("MG_KC") %>">
                         <% =rsTelefonos("MG_KC") %></a></td> 
<%   
               i=i+1
               rsTelefonos.MoveNext
            wend
%>
            </tr>
<%
         wend 
%>
         </table>      
         <hr>
<%
   end if
%>
   <form name="frmAdd" method="POST" action="INTTelefonosUtiles.asp">
<%
   if (P_strAccion = "MOD") or (P_strAccion = "NEW") then
      if (P_strAccion = "MOD") then 
         'SI SE PIDE MODIFICACION.     
         P_KC=GF_PARAMETROS("P_KC","")
         'Se obtiene todos los datos a mostrar.
         strSQL="Select * from MG where MG_KM='ST' and MG_KC='" & P_KC &"'"
         GF_BD_CONTROL rsDatos,oConn,"OPEN",strSQL
         strDS=rsDatos("MG_DS")
         GF_Memo "READ", rsDatos("MG_KR"), strMEMO
         strTexto="Aceptar Cambios"
      else
         'SI SE PIDE AGREGAR UN NUEVO NUMERO.     
         P_KC=""
         strDS=""
         strMEMO=""
         strTexto="Agregar"
      end if
%>    
    <table border="0" width="100%">
    <tr><td width="10%">Codigo:</td><td><input type="text" name="strKC" maxlength="10" value="<% =P_KC %>"></td>
        <td align="right" width="30%"><input class="NOBORDER" type="image" src="images/Aceptar.gif" name="btnAccion" onClick="SetearOperacion('ADD')" alt="<% =GF_TRADUCIR(strTexto) %>"></td></tr>
    <tr><td width="10%">Descripcion:</td><td><input type="text" name="strDS" size="52" value="<% =strDS %>"><td></tr>
    <tr><td colspan="3"><textarea name="strMemo" rows="7" cols="50"><% =strMEMO %></textarea></td></tr>
    </table>
    <hr>
<%
   end if
%>
   </form>
<%
   if (P_strAccion = "VER") then
      P_KC=GF_PARAMETROS("P_KC","")
      'Se obtiene todos los datos a mostrar.
      strSQL="Select * from MG where MG_KM='ST' and MG_KC='" & P_KC &"'"
      GF_BD_CONTROL rsDatos,oConn,"OPEN",strSQL
      GF_Memo "READ", rsDatos("MG_KR"), strMEMO
      strMEMO=replace(strMEMO,chr(13),"<br>")
%>
   <table width="100%">
      <tr><td>
         <font FACE="Book Antiqua" SIZE="5" COLOR="#008000"><u><% =P_KC %></u></font>
      </td>
      <td align="right">
<%       if (intAccesoST > 1) then %>         
         <a href="INTTelefonosUtiles.asp?P_OPR=DLT&P_KC=<% =P_KC %>">
         <img src="images/Eliminar.gif" border="0" alt="<% =GF_TRADUCIR("Eliminar") %>">&nbsp;
         </a>
         <a href="INTTelefonosUtiles.asp?P_OPR=MOD&P_KC=<% =P_KC %>">
         <img src="images/Modificar.gif" border="0" alt="<% =GF_TRADUCIR("Modificar") %>">
         </a>
<%       end if                    %>      
      </td>     
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td>
         <strong><% =rsDatos("MG_DS") %></strong>
      </td></tr>   
      <tr><td>&nbsp;</td></tr>
      <tr><td><% =strMEMO %></td></tr>   
   </table>
   <hr>
<%
   end if
%>
</BODY>
</HTML>
