<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<%
Dim strSQL, conn,rs,i,codigo,auxTotalTrim
Dim totalTrim(5),tipoFormulario


Const KEY_CODIGO = "cd_"
Const KEY_DESCRIPCION = "ds_"
Const KEY_CUENTA = "cuenta_"
Const KEY_CCOSTO = "cc_"
Const KEY_IMPORTE = "bg_"
Const KEY_TRIMESTRES = "trim_"
Const KEY_DETALLE = "ta_"
Const KEY_DIV = "div_"

Const TRIMESTRE_1 = 0
Const TRIMESTRE_2 = 1
Const TRIMESTRE_3 = 2
Const TRIMESTRE_4 = 3

pIdObra = GF_PARAMETROS7("idobra", 0, 6)
pIdArea = GF_PARAMETROS7("idarea", 0, 6)
ptheArea = GF_PARAMETROS7("theArea", 0, 6)
indiceLineas = GF_PARAMETROS7("indiceLineas", 0, 6)
tipoFormulario = GF_PARAMETROS7("tipoFormulario", 0, 6)

strSQL = "select distinct(deta.idarea),deta.iddetalle,obra.* from tblbudgetobras obra "
strSQL = strSQL & "left join tblbudgetobrasdetalle deta on obra.idobra = deta.idobra and obra.idarea = deta.idarea and obra.iddetalle = deta.iddetalle "
strSQL = strSQL & " where obra.idobra = "&pIdObra&" and obra.idarea = " & pIdArea & " and obra.iddetalle <> 0"
Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
i = 0

while not rs.EoF 
	%>
	nuevoDetalle(<%=ptheArea%>);
	document.getElementById("<% =KEY_IMPORTE  & ptheArea & "_" & indiceLineas %>").value = editarImporte("<% =CDbl(rs("DLBUDGET"))/100 %>");
	
	totalTrim5 += <%=CDbl(rs("DLBUDGET"))/100%>;
	<%
    if (tipoFormulario = OBRA_FORM_TRIM) then
		for i = TRIMESTRE_1 to TRIMESTRE_4
			auxTotalTrim  = cdbl(obtenerImporteTrimestres(pIdObra,rs("IDAREA"),rs("IDDETALLE"),i,MONEDA_DOLAR))/100
		%>
			totalTrim<%=i%> = totalTrim<%=i%> + <%=auxTotalTrim%>
			document.getElementById("<% =KEY_TRIMESTRES & "_" & i & "_" & ptheArea & "_" & indiceLineas %>").value = editarImporte("<%=auxTotalTrim %>");
	<%	next 
	end if
	%>
	document.getElementById("<% =KEY_CODIGO & ptheArea & "_" & indiceLineas %>").value = "<% =rs("IDDETALLE") %>";
	document.getElementById("<% =KEY_DETALLE & ptheArea & "_" & indiceLineas %>").value = "<% =Trim(rs("DSDETALLE")) %>".replace(/<% =ENTER_SYMBOL%>/g, '\n');
	document.getElementById("<% =KEY_CUENTA & ptheArea & "_" & indiceLineas %>").value = "<% =Trim(rs("CDCUENTA")) %>";
	<% if (tipoFormulario <> OBRA_FORM_ANUAL) then %>
		document.getElementById("<% =KEY_CCOSTO & ptheArea & "_" & indiceLineas %>").value = "<% =rs("CCOSTOS") %>";
	<% end if %>
	document.getElementById("<% =KEY_DIV & ptheArea & "_" & indiceLineas %>").innerHTML = "<% =rs("IDDETALLE") %>";
	agregarImagenDel(<%=ptheArea%>,<% =indiceLineas %>);
	
	msArray["<% =KEY_DESCRIPCION & ptheArea & "_" & indiceLineas %>"].setValue("<% =rs("DSBUDGET") %>");
	shadowItems["<%=ptheArea & "_" & indiceLineas %>"] = "<% =rs("DSBUDGET") %>";
<%
	indiceLineas = indiceLineas +1
	
	rs.MoveNext
wend



%>