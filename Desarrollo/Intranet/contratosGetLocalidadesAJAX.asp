<!--#include file="Includes/procedimientosAS400.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/json.asp"-->
<%
dim rs, oConn, strSQL, ret, str
Dim member
str = UCASE(GF_PARAMETROS7("term","",6))
if len(trim(str)) > 0 then
	Set member = jsObject()
	if isNumeric(str) then
		'strSQL = "Select PROC.CODIPC, PROC.AUXIPC, PROC.DESCPC, PROV.CODIPO, PROV.DESCPO from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE PROC.CODIPC LIKE '%" & str & "%' ORDER BY DESCPC"
		strSQL = "Select PROC.CODIPC as id, PROC.DESCPC as label, PROC.AUXIPC as value, concat(concat('(',PROV.DESCPO), ')') as desc from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE PROC.CODIPC LIKE '%" & str & "%' ORDER BY DESCPC"
	else
		'strSQL = "Select PROC.CODIPC, PROC.AUXIPC, PROC.DESCPC, PROV.CODIPO, PROV.DESCPO from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE DESCPC LIKE '%" & str & "%' ORDER BY DESCPC"
		'strSQL = "Select PROC.CODIPC as id, concat(concat(PROC.DESCPC,'|'), PROV.CODIPO) as value from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE DESCPC LIKE '%" & str & "%' ORDER BY DESCPC"
		strSQL = "Select PROC.CODIPC as id, PROC.DESCPC as label, PROC.AUXIPC as value, concat(concat('(',PROV.DESCPO), ')') as desc from MERFL.MER142F1 PROC INNER JOIN MERFL.MER1K2F1 PROV ON PROC.PROVPC=PROV.CODIPO WHERE DESCPC LIKE '%" & str & "%' ORDER BY DESCPC"
	end if
	call QueryToJSON(strSQL).Flush
end if
'--------------------------------------------------------------------------------------------
Function QueryToJSON( sql)
        Dim rs, jsa, col
        Call GF_BD_AS400_2(rs,oConn,"OPEN",strSQL)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null) (lcase(col.Name)) = trim(col.Value)
                Next
			rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function
%>
