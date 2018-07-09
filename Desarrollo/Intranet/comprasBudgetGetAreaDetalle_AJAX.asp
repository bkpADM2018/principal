<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->


<%
dim myObra, myBudgetArea, myBudgetDetalle, myObraFechaInicio, myFechaFin
dim rs, conn, strSQL, myRtrn, dicVales

myObra = GF_Parametros7("idObra", 0, 6)
myBudgetArea = GF_Parametros7("idBudgetArea", 0, 6)
myObraFechaInicio = GF_Parametros7("obraFechaInicio", "", 6)
myFechaFin = GF_Parametros7("fechaFin", "", 6)


strSQL="SELECT * FROM TBLBUDGETOBRAS WHERE IDOBRA=" & myObra & " AND IDAREA=" & myBudgetArea & " AND IDDETALLE<>0 ORDER BY IDDETALLE"
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)

while not rs.eof
    Call executeProcedureDb(DBSITE_SQL_INTRA, rs1, "TBLBUDGETOBRAS_GET_SALDO_BY_IDOBRA", myObra & "||" & myBudgetArea & "||" & rs("IDDETALLE"))		
	if (not rs1.eof) then
        importePIC = Cdbl(rs1("IMPORTEPICDOLARES"))
        myVlVales = Cdbl(rs1("IMPORTEVALESDOLARES"))	        
    else
        importePIC = 0
        myVlVales = 0
	end if		
	gastosFacturados = calcularGastosFacturados(myObra,myBudgetArea,rs("IDDETALLE"),"","",MONEDA_DOLAR)	
	myRtrn = myRtrn & "//" & rs("IDDETALLE") & ";" & trim(rs("DSBUDGET")) & ";" & myVlVales/100 & ";" & importePIC/100 & ";"  & gastosFacturados/100 & ";"  & cdbl(rs("DLBUDGET"))/100 & ";"  & trim(rs("CDCUENTA")) & ";" & trim(rs("CCOSTOS"))
	rs.movenext
wend
Response.Write myRtrn
%>

