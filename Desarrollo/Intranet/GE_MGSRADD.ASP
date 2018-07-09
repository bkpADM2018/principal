<!--#include file="Includes/procedimientosMG.asp"-->
<%
dim  P_o1kr, P_o2kr, P_o3kr, P_val, P_3okr
DIM V,S,My_Rtrn,My_key
S = ")("
my_key = "GE_MGSRADD.ASP/PARAMETROS"
V = SPLIT( SESSION(my_key) , s ) ' Tomar parametros para este execute
P_O1KR = V(1)
P_O2KR = V(2)
P_O3KR = V(3)
P_VAL  = V(4)
P_3OKR = V(5)
DIM CON, RS, SQL,R
dim My_O1Km,My_O1Kc,My_O2Km,My_O2Kc,My_O3Km,My_O3Kc,MY_O1DS,MY_O2DS,MY_O3DS
dim my_id, my_val
my_val = p_val
P_3okr = 0
IF GF_MGKR(P_o1KR, MY_o1KM, MY_o1KC, MY_o1DS) AND GF_MGKR(P_o2KR, MY_o2KM, MY_o2KC, MY_o2DS) AND GF_MGKR(P_o3KR, MY_o3KM, MY_o3KC, MY_o3DS) THEN
   if len(my_val) = 0 or my_val = "" or isnull(my_val) then my_val = "0"
   sql = "SELECT * FROM relacionesconsulta WHERE "
   SQL = SQL & "     SRO1KR = " & P_O1KR 
   SQL = SQL & " AND SRO2KR = " & P_O2KR 
   SQL = SQL & " AND SRO3KR = " & P_O3KR 
   GF_BD_CONTROL rs,con,"OPEN",sql 
   IF RS.EOF THEN 
      GF_BD_CONTROL rs,con,"CLOSE","" 
      GF_BD_CONTROL rs,con,"OPEN", "SELECT max(sr3okr) as MY_ID FROM relacionesconsulta " 
	  P_3OKR = RS("MY_ID") + 1
	  Sql = "Insert Into MGSR(sro1kr,sro2kr,sro3kr,sr3okr,srvalor)"
	  sql = sql & " Values ( " & P_o1kr & "," & P_o2kr & "," & P_o3kr & "," & p_3okr 
	  sql = sql & ",'" & my_val & "')"
      MY_RTRN = true
   ELSE 
	  p_3okr = rs("sr3okr")
	  Sql = "UPDATE MGSR SET srvalor = '" & my_val & "' WHERE sr3okr= " & p_3Okr 
      MY_RTRN = false
   END IF
   R = TRUE
   'Response.Write "<br>SQL(" & sql & ")" 
   con.execute sql 
else
   R = FALSE
END IF
' Retornar los parametros de este execute
SESSION(my_key) = my_RTRN & S & P_O1KR & S & P_O2KR & S & P_O3KR & S & P_VAL & S & P_3OKR & s & " " & s & " " & s & " " & s & " " & s & " " & s & " "

'-------------------------------------------------------------------------------------------------------------------
%>