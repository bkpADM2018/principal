<!--#include file="procedimientosfechas.asp"-->
<%
Function GF_MGSRADD( P_o1kr, P_o2kr, P_o3kr, P_val, byref P_3okr)
DIM CON, RS, strSQL,R
dim My_O1Km,My_O1Kc,My_O2Km,My_O2Kc,My_O3Km,My_O3Kc,MY_O1DS,MY_O2DS,MY_O3DS
dim my_id, my_val
my_val = p_val
P_3okr = 0
call GP_CONFIGURARMOMENTOS()
IF GF_MGKR(P_o1KR, MY_o1KM, MY_o1KC, MY_o1DS) AND GF_MGKR(P_o2KR, MY_o2KM, MY_o2KC, MY_o2DS) AND GF_MGKR(P_o3KR, MY_o3KM, MY_o3KC, MY_o3DS) THEN
	if len(my_val) = 0 or my_val = "" or isnull(my_val) then my_val = "0"
    GF_BD_Control rs,con,"OPEN", "SELECT max(sr3okr) as MY_ID FROM MGSR"
	P_3OKR = RS("MY_ID") + 1
	strSQL = "Insert Into MGSR(sro1kr,sro2kr,sro3kr,sr3okr,srvalor,srmmdt,srmmsy,srUser)"
	strSQL = strSQL & " Values ( " & P_o1kr & "," & P_o2kr & "," & P_o3kr & "," & p_3okr
	strSQL = strSQL & ",'" & my_val & "'," & session("MomentoDato") & "," & session("MomentoSistema") & ",'" & session("Usuario") & "')"
    GF_MGSRADD = true
	GF_BD_Control rs,con,"EXEC", strSQL	
END IF
end function
'-------------------------------------------------------------------------------------------------------------------
%>
