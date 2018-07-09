<!--#include file="procedimientosParametros.asp"-->
<!--#include file="procedimientosConexion.asp"-->
<!--#include file="procedimientosTitulos.asp"-->
<%
'Const DBSITE_SQL_MG = "MGACTISA"

sub ProcedimientoControl(Procedimiento)
dim con, rs, strSQL, my_kr, my_ds

If Procedimiento <> "" then
   IF Procedimiento <> session("procedimiento") then
      IF GF_ControlAccesoKS("EXEC","SP",Procedimiento,my_kr,my_ds) <= "0" THEN 			
    	 Response.Redirect "MGMSG.ASP?P_MSG=" & GF_TRADUCIR("ACCESO DENEGADO")
		 else
		 session("procedimiento") = procedimiento
	  END IF
   END IF
END IF	  
end sub
'-------------------------------------------------------------
FUNCTION GF_ControlAccesoKs(P_Acceso, byref P_Km, byref P_kc, byref P_kr, byref P_ds)
dim  v
v = "0"
IF GF_MGKS(P_KM, p_KC,p_kR, p_DS) THEN  v = GF_ControlAcceso( P_Acceso, P_Km,P_kc, P_Kr)   
gf_controlaccesoks = v
END Function
'-------------------------------------------------------------
FUNCTION GF_ControlAccesokr(P_Acceso, byref P_KR, byref P_KM, byref P_KC, byref P_DS)
dim  v
v = "0"
IF GF_MGKR (p_kR, P_KM, p_KC, p_DS) THEN v = GF_ControlAcceso( P_Acceso,P_km,P_kc, P_Kr) 
gf_controlaccesokr = v
END Function
'-------------------------------------------------------------
FUNCTION GF_ControlAcceso(P_Acceso,byref P_Km,byref P_kc,P_kr)
DIM RS , CON , strSQL , My_Acceso, s, my_AccesoNuevo
dim my_krAcceso, my_dsAcceso,intSMKR
Dim i,strsss,My_Acceso_Maestro
Dim strListaCargos
IF session("usuario") = "" THEN Response.Redirect (session("home"))


IF session("usuario") = "ADMIN" THEN  
   GF_ControlAcceso = "9"
else
   if P_KM <> "SM" AND session("Acc_KM_" & P_km ) = "9" then
      My_Acceso = "9"
   else	  
      S = "Acc_" & P_kr & P_Acceso
      if session(S) <> ""   then
         My_Acceso = session(S) 'Retorna el acceso que ya buscó
       ELSE  
	     ' Hay que buscar el maximo acceso al registro solicitado 
         My_Acceso = "0"
		 strListaCargos=session("LISTADECARGOS")
		 'Response.Redirect "MGMSG.ASP?P_MSG=" & strListaCargos
		 if (strListaCargos = "") then Response.Redirect (session("home"))
		 if (strListaCargos <> "()") then 
		    ' Obtener clave reducida del acceso solicitado y del acceso all 
   		    My_KrAcceso = GF_SessionKR("SR", P_acceso)
  		    'Busco el acceso al registro.
			My_Acceso = GF_SessionAccesoBuscar(My_KrAcceso,P_KR)
			if (My_Acceso = "0") and (P_KM <> "SM") then
			   'Busco el acceso al maestro.
			   GF_MGKS "SM",P_KM,intSMKR,""
			   My_Acceso=GF_SessionAccesoBuscar(My_KrAcceso,intSMKR)
			   session("Acc_KM_" & P_km ) =My_Acceso
			end if
			session(S)=My_Acceso
		 end if
	  end if
   end if
   gf_ControlAcceso = My_Acceso
end if   
END FUNCTION
'--------------------------------------------------------------------------------------------
function GF_SessionAccesoBuscar(p_AccesoKr,P_objetoKR)
	dim My_Sql, My_rs, My_cn
	My_sql = "SELECT TOP 1 * FROM relacionesconsulta WHERE "
	my_SQL = My_SQL & "     SRO1KR = " & P_AccesoKR 
	My_SQL = My_SQL & " AND SRO3KR = " & P_objetoKR 
	My_SQL = My_SQL & " AND SRO2KR IN " &  session("LISTADECARGOS")
	MY_sql = My_sql & " ORDER BY SRVALOR DESC"
	'Response.write MY_sql
	call executeQueryDb(DBSITE_SQL_INTRA, My_rs, "OPEN", My_sql)
	IF My_rs.EOF THEN 
	      GF_SessionAccesoBuscar = "0"
	else
		  GF_SessionAccesoBuscar = MY_RS("SRVALOR")
	END IF	  
end function
'--------------------------------------------------------------------------------------------
FUNCTION GF_SessionKr(byref P_KM,byref P_KC)
' Desde la session retornar un KR Para una clave simbolica
DIM V,DS,KR
p_km = ucase(p_km)
p_kc = ucase(p_kc)
v = "KR_" & P_KM & P_KC  
KR = SESSION(V) + 0
IF KR < 1 THEN 
   IF GF_MGKS( P_KM, P_KC, KR, DS ) THEN   
      SESSION(V) = KR
   END IF
END IF
GF_SessionKR = kr
end function   
'---------------------------------------------------------------------------------------------
Function GF_MGSR(P_o1kr, p_o2kr, P_o3kr, byref P_3okr)
DIM CON, RS, strSQL
  strSQL = "SELECT * FROM RelacionesConsulta where sro1kr = " & p_o1kr & "  and sro2kr = " & p_o2kr & " and sro3kr = " & p_o3kr & " order by SRMMDT desc"
  call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
  'Response.Write strSQL
  IF RS.EOF then 
     P_3OKR = ""
     GF_MGSR = FALSE
  ELSE
	 p_3okr = rs("sr3okr")
	 GF_MGSR = TRUE
  end if	   
end function
'-------------------------------------------------------------------------------------------------------------------
Function GF_MGSR_EXISTE(P_o1kr, p_o2kr, byref p_o3kr)
DIM CON, RS, strSQL
  strSQL = "SELECT * FROM MGSR where sro1kr = " & p_o1kr & " and sro2kr = " & p_o2kr & " and SRValor <> '*'  order by SRMMDT desc" 
  call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
  IF RS.EOF then 
     P_o3KR = ""
     GF_MGSR_EXISTE = FALSE
  ELSE
	 p_o3kr = rs("sro3kr")
	 GF_MGSR_EXISTE = TRUE
  end if	   
end function
'--------------------------------------------------------------------------------------------
Function GF_DT1(P_Accion,P_dato,P_MMDato, P_MMSistema, P_km1, P_kc1)
dim my_mgkr1, my_dtkr, my_mgds1
' Obtener las claves del objeto y del dato
if gf_mgkS (P_km1,P_kc1, my_mgkr1, my_mgds1 ) THEN
   gf_mgADD "SD", p_DATO,"?", MY_DTKR 
   GF_DT1 = gf_dt1kr(P_ACCION, MY_DTKR, P_MMdATO, P_MMsISTEMA, MY_MGKR1)
ELSE 
  GF_DT1 = "?"
END IF
END function
'--------------------------------------------------------------------------------------------
Function GF_DT1KR(P_ACCION,P_DTKR ,BYREF P_MMD, BYREF P_MMS, P_MGKR)
DIM CON, RS, strSQL  
   IF P_mmD = "" then P_MMD = session("MomentoDato")
   IF P_MMS = "" then P_MMS = session("MomentoSistema")
   strSQL = "SELECT TOP 1 DT_VALOR FROM MGDT WHERE DT_KO = " & P_MGKR & " AND DT_Objetos = 1 AND DT_KR = " & P_DTKR & " and dt_mmmv <= " & P_MMD & " and dt_MmSy <= " & P_MMS & " ORDER BY dt_mmmv DESC, dt_mmSY DESC"
   call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
   if rs.eof then 
      gf_dt1kr = "?"
	  else
	  GF_DT1KR =  rs("DT_Valor") 
   end if
END FUNCTION
'--------------------------------------------------------------------------------------------
Function GF_DT1W(P_dato, P_km1, P_kc1,P_VALOR, byref p_mmmv)
DIM CON, RS, strSQL
dim my_mgkr1, my_dtkr, my_Valor, my_mgds1
if p_MmMv = "" then p_MmMv = session("MomentoDato")
' Obtener las claves del objeto y del dato
if gf_mgkS(P_km1,P_kc1, my_mgkr1, my_mgds1 ) THEN
   gf_mgADD "SD", p_DATO,"?", MY_DTKR  
   GF_DT1W = GF_DT1W_KR( MY_DTKR, MY_MGKR1, P_VALOR, p_mmmv)
 ELSE 
   GF_DT1W = FALSE
END IF
ENd function
'--------------------------------------------------------------------------------------------
Function GF_DT1W_kr(P_dtkr, P_mgkr, P_VALOR, p_mmmv)
DIM CON, RS, strSQL,R
dim aux,  My_MMMS 
if p_MmMv = "" then p_MmMv = session("MomentoDato")
' Obtener las claves del objeto y del dato
if len(ltrim(rtrim(P_valor))) < 1 then
     R = FALSE
  else
     My_MmMs = session("MomentoSistema")
	 strSQL = "SELECT TOP 1 DT_VALOR FROM MGDT WHERE DT_KO = " & p_mgkr & " AND DT_Objetos = 1 AND DT_KR = " & P_DTKR & " and dt_mmmv <= " & p_MmMv & " and dt_MmSy <= " & My_mmmS & " ORDER BY dt_mmmv DESC"
     call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
     if NOT (rs.eof = TRUE AND RS.BOF = TRUE) then AUX = rs("dt_valor")
     if P_VALOR <> AUX  then 
	     strSQL = "Insert Into MGDT (DT_KO, DT_Objetos, DT_KR, DT_VALOR, dt_MmMv, dt_MmSy ) "
    	 strSQL = strSQL & " Values ( " & p_mgKR & ",1, " & p_DTKR  & ", '" & P_valor & "', " &  p_MmMv & " , " & My_mmmS & ")"
	     call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
		 'gf_bd_control rs,con,"CLOSE",""
		 R = TRUE
   	    ELSE
	     'gf_bd_control rs,con,"CLOSE",""
     END IF
end if	 
GF_DT1W_kr = R
end function
'--------------------------------------------------------------------------------------------
Function GF_MGADD (p_KM, p_KC, p_DS, byref P_KR)
DIM CON, RS, strSQL, MY_ID, CONT, where, strSQL2
p_kr = 0
while p_kr < 1
  strSQL = "SELECT * FROM Mg where MG_KM = '" & P_KM & "' AND MG_KC = '" & P_KC & "' AND MG_MMMV <= " & session("MomentoDato")
  'response.write strSQL
  call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
  IF  RS.EOF then
	 'Si va agregar un articulo del kiosco, va a partir de 90000 sino a partir de 20000
 	if p_KM = "LK" then 
	  where = " MG_KR > 90000 "
	else 
	  where = " MG_KR < 90000 " 
	end if  
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", "SELECT max(mg_kr) as max_id FROM MG WHERE " & where)
	strSQL2 = "SELECT max(mg_kr) as max_id FROM MG WHERE " & where
	p_kr = rs("max_id") + 1
    strSQL = "Insert Into MG (mg_MmMv,mg_MmSy,mg_km,mg_kc,mg_ds,mg_kr )  Values ( " & session("MomentoDato") & "," & session("MomentoSistema") & ",'"  & p_km & "','" & P_kc & "','" &  P_ds & "'," & p_kr & ")"
    'response.write strSQL
    call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
   else 
    if rs("MG_DS") <> P_DS and P_ds <> "?" and p_ds <> "" THEN
       p_kr = rs("mg_kr") 
	   strSQL = "Update MG set mg_ds = '" & P_ds & "'" & " WHERE MG_kr = " & p_kr & " and mg_mmmv =" & rs("mg_mmmv")
	   'response.write  strSQL
	   call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", strSQL)
	END IF
    p_kr = rs("MG_KR")
   end if	   
WEND
GF_MGADD = TRUE
end function
'------------------------------------
FUNCTION GF_ControlarInputKc(byref p_kc)
' Funcion que intenta evitar el ingreso de codigo malicioso
'response.write "inputkc entra kc(" & p_kc & ")"
dim return, i, strTexto
strTexto = array(" OR ", " GO ", "'", chr(34), "SELECT ", "INSERT ", "UPDATE ", "@", "FROM ", " FROM", " SELECT", " INSERT", " INTO ", " UPDATE", "/", "\", " UNION ", " AND " )
  return = ucase(p_kc)
  for i = 1 to ubound(strTexto)
    return = replace (return, strTexto(i), " ")
  next
  GF_ControlarInputKc = return
 'response.write "inputkc sale kc(" & p_kc & ")"
END FUNCTION
'-----------------------------------
FUNCTION GF_MG_Acceso(byref P_KM,byref P_KC,BYREF P_KR,BYREF P_DS, BYREF p_kmkr, BYREF p_kmds, BYREF P_SMACCESO, P_MSGERROR)
DIM my_acceso,MY_DS,My_existe,v
my_existe = GF_MGKS( P_KM,p_kc,P_KR,MY_DS )
p_SMAcceso = "0"
p_kmds = ""
p_kmkr = 0
P_MSGERROR = ""
IF P_KM = "" THEN 
   P_MSGERROR = "Falta indicar maestro"
else 
   P_SMACCESO = GF_CONTROLACCESOKS("ACCESO","SM",(P_KM),p_kmkr,p_kmds) 
   IF P_SMACCESO < "1" then
      P_MSGERROR = "No tiene acceso al Maestro: " & P_KM 
   else
      if my_existe then
    	  my_acceso = GF_CONTROLACCESOKS("ACCESO",(p_km),(P_Kc),P_KR,P_DS) 
		  IF (P_SMACCESO > my_acceso) THEN my_acceso=P_SMACCESO
         IF my_ACCESO < "1" then p_msgerror = "No tiene acceso al registro: '" & P_KM & "' '" & P_KC & "'" 
	  else
	     p_msgerror = "No existe el registro: '" & P_KM & "' '" & P_KC & "'" 
	  end if	 
   end if     
END IF   
if p_msgerror = "" then 
   GF_MG_Acceso = true
   P_DS = MY_DS
   ELSE
   GF_MG_Acceso =  false
   P_DS = ""
END IF   
END FUNCTION
'-------------------------------------------------------------------------------------
Function GF_MGKR (p_KR, BYREF P_KM, BYREF P_KC, byref p_ds)
' Obtener la clave simbolica desde una clave reducida.del maestro general
DIM CON, RS, strSQL 
P_km = ""
P_kc = ""
p_DS = ""
GF_MGKR = FALSE 
if not isnumeric(p_kr) then p_kr=0
if (p_kr+0) > 0 then
   strSQL = "SELECT * FROM Mg where MG_Kr = " & P_Kr 
   'Response.Write strsql
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)   
	if not rs.eof then   
		P_km = rs("mg_km") 
		P_kc = rs("mg_kc") 
		P_DS = RS("MG_DS")
		GF_MGKR = true 
    end if	  
end if
END FUNCTION
'-------------------------------------------------------------------------------------
Function GF_MGKS (P_KM, P_KC, byref P_KR, byref P_DS)
' Obtener la clave reducida de un registro del maestro general
DIM CON, RS, strSQL, elemento, dic 
IF LEFT(P_KC,1) = "?" THEN  
  set dic = server.createobject("Scripting.Dictionary")
  dic.removeall
  if request.form.count = 0 then
    for each elemento in request.querystring()
      if (left(request.querystring(elemento),1) = "?") and (len(request.querystring(elemento))>1) then 
        dic(elemento) = "?"
      else 
        dic(elemento) = replace(request.querystring(elemento)," ","%20")
      end if
	next
  else
    for each elemento in request.form()
      if (left(request.form(elemento),1) = "?") and (len(request.form(elemento))>1) then 
        dic(elemento) = "?"
      else 
        dic(elemento) = replace(request.form(elemento)," ","%20")
      end if
	next
  end if
  dic("$pagina") = request.servervariables("URL")
  session("nombres") =  join(dic.Keys, "$$")
  session("valores") =  join(dic.Items, "$$")
  RESPONSE.REDIRECT("MGBKS.ASP?P_KC=" & P_KM & "&" & "P_OPERACION=SELECT&P_DSSEL=" & mid(P_KC,2,len(P_KC)))
end if
strSQL = "SELECT * FROM Mg where MG_KM = '" & P_KM & "' AND MG_KC = '" & P_KC & "'"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
'gf_bd_control rs,con,"OPEN",strSQL
if rs.eof then  
      P_kr = 0
	  p_DS = ""
      GF_MGKS = FALSE 
	else
	  P_kr = rs("mg_kr") 
	  P_DS = rs("mg_ds")
	  GF_MGKS = true 
end if
'gf_bd_control rs,con,"CLOSE",strSQL
'GP_LOG "GF_MGKS", "SALE=KM(" & p_km & ") KC(" &  P_KC & ") DS(" & P_DS & ") kr(" & p_kr & ")"
END FUNCTION
'---------------------------------------------------------------
Function GF_MGC(byref P_KM, byref P_KC, byref P_KR, byref P_DS)
' Leer un registro por cualquiera de sus claves.
IF NOT ISNUMERIC(P_KR) THEN P_KR = 0
if (p_kr) < 1 then 
    P_KR = GF_SESSIONKR( P_KM, gf_controlarinputkc(P_KC) )	
	else 
	p_kc = gf_controlarinputkc(P_KC)
END IF
GF_MGKR p_kr,p_km,p_kc,P_ds
IF P_KR > 0 THEN 
   GF_MGC = true
   else
   GF_MGC = FALSE
END IF    
END FUNCTION
'-------------------------------------------------------------------------------------------------
function Editar_Importe(p_importe)
    dim my_valor
    my_valor = replace(p_importe,",",".")
    if not isnumeric(my_valor) then my_valor = 0
    my_valor = formatnumber(my_valor,2,true,false,true)    
    my_valor = replace(my_valor,",","#")
    my_valor = replace(my_valor,".",",")
    my_valor = replace(my_valor,"#",".")
    Editar_Importe = my_valor
end function
%>

