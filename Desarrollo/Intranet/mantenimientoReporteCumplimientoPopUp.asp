<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<%        
    Dim listaDivisionesDefault, myAnio, myMes, myFecha
    
    Call initTaskAccessInfo(TASK_SMW_REPO_CUMPLIMIENTO, "")
    
    Call GP_CONFIGURARMOMENTOS()
    
    listaDivisionesDefault = GetTaskDivisionAccessList(TASK_SMW_REPO_CUMPLIMIENTO, session("Usuario"))
    
    myFecha = GF_DTEADD(session("MmtoDato"), -1, "M")
    
    myAnio = Left(myFecha, 4)
    myMes = Right(Left(myFecha, 6), 2)
%>
<html>
<head>
    <title>Sistema de Mantenimiento - Reporte de Cumplimiento</title>    
    <link rel="stylesheet" href="css/main.css" type="text/css" /> 
    <script type="text/javascript">

        function executeReport() {
            var aaaa = document.getElementById("anio").value;
            var mm = document.getElementById("mes").value;
            var si = document.getElementById("txtIdDivision").selectedIndex;
            var divi = document.getElementById("txtIdDivision").options[si].value;
            window.open("mantenimientoReporteCumplimientoPrint.asp?div=" + divi + "&anio=" + aaaa + "&mes=" + mm);
        }
    </script>
</head>
<body>
    <table border="0">
        <tr><td>&nbsp;</td></tr>
        <tr>
            <td>Divisi&oacute;n:</td>
            <td>
                <select name="txtIdDivision" id="txtIdDivision">					
					<%
					call executeProcedureDb(DBSITE_SQL_INTRA, rsList, "TBLDIVISIONES_GET_BY_LIST", listaDivisionesDefault)
					while not rsList.eof
						%>	
							<option value="<%=rsList("IDDIVISION")%>"><%=rsList("DSDIVISION")%></option>
						<%	
						rsList.movenext
					wend	
					%>
				</select>
            </td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr>
            <td>A�o:</td>
            <td><input type="text" size="5" name="anio" id="anio" maxlength="4" value="<% =myAnio %>"/> (aaaa)</td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr>
            <td>Mes:</td>
            <td><input type="text" size="5" name="mes" id="mes" maxlength="2" value="<% =myMes %>"/> (mm)</td>
        </tr>
        <tr><td>&nbsp;</td></tr>
        <tr>
            <td><input type="button" value="Generar" onclick="javascript:executeReport()" /></td>
        </tr>
    </table>
</body>
</html>