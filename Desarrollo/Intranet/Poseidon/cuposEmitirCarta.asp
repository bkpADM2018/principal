<!--#include file="../Includes/procedimientosMail.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="cartaCuposPrint.asp"-->
<%
Const PERIODO_CUPOS = 9 
'********************************************************************************************************************************
'********************************************************* INICIO DE PAGINA *****************************************************
'********************************************************************************************************************************
Dim  rs, diasCupos,contratoPro, contratoSuc, contratoOpe, contratoNum, contratoCos
Dim strSQL,  cuitCupeador, myWhere, puedeAgregar, maxCuposDisponibles, cantCupo, fechaCupo

Call GP_CONFIGURARMOMENTOS

accion = GF_PARAMETROS7("accion", "", 6)
myPlanta = GF_PARAMETROS7("pto", 0, 6)
fechaDesde = GF_PARAMETROS7("fd", "", 6)
if (fechaDesde = "") then fechaDesde = GF_DTEADD(Left(Session("MmtoDato"),8), 1, "D")
fechaHasta = GF_DTEADD(fechaDesde, PERIODO_CUPOS, "D")
diasCupos = GF_DTEDIFF(fechaDesde ,fechaHasta ,"D")

if (accion <> "") then    
    contratoPro = GF_PARAMETROS7("contratoPro", 0, 6)
    contratoSuc = GF_PARAMETROS7("contratoSuc", "", 6) 'Va texto dado que 0 es un valor valido.
    contratoOpe = GF_PARAMETROS7("contratoOpe", "", 6) 'Va texto dado que 0 es un valor valido.
    contratoNum = GF_PARAMETROS7("contratoNum", 0, 6)
    contratoCos = GF_PARAMETROS7("contratoCos", 0, 6)

    if (contratoPro > 0) and (contratoSuc <> "") and (contratoOpe <> "") and (contratoNum > 0) and (contratoCos > 0) then
        Redim arrCupos(diasCupos)
        Redim arrFecha(diasCupos)
        For i = 0 to diasCupos-1
            arrCupos(i) = GF_PARAMETROS7("cupo_" & i, 0, 6)
            arrFecha(i) = GF_PARAMETROS7("colFecha_" & i, 0, 6)        
        Next
        
        myFile = armarPDF(arrFecha, arrCupos, contratoPro, contratoSuc, contratoOpe, contratoNum, contratoCos, myPlanta, PDF_FILE_MODE)
        myMail = getTaskMailList(TASK_POS_ADMIN_CUPOS, MAIL_TASK_SENDER)
        Call GP_ENVIAR_MAIL_ATTACHMENT("Carta Cupos - Contrato " & GF_EDIT_CONTRATO(contratoPro, contratoSuc, contratoOpe, contratoNum, contratoCos), "Se adjunta la carta de cupos generada para el contrato " & GF_EDIT_CONTRATO(contratoPro, contratoSuc, contratoOpe, contratoNum, contratoCos), myMail, myMail, myFile)
    end if        
end if

%>
<html>
<head>
<title>Sistema de Cupos - Emisi�n de carta cupos</title>
<link rel="stylesheet" href="../css/tabs.css" TYPE="text/css" MEDIA="screen">
<link rel="stylesheet" href="../css/tabs-print.css" TYPE="text/css" MEDIA="print">
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
    .selectorColumn {
        background: #FF6666;
        color: #FFFFFF;
    }
    .inputImgNominacion {
        background:none;
        border:none;
    }
</style>

<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>

<script type="text/javascript">
    var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");
    
    function MostrarCalendario(p_objID, funcSel) {
		var dte= new Date();		    	    
		var elem= document.getElementById(p_objID);
		if (calendar != null) calendar.hide();		
		var cal = new Calendar(false, dte, SeleccionarCal, CerrarCal);
	    cal.weekNumbers = false;
		cal.setRange(1993, 2045);
		cal.create();
		calendar = cal;		
	    calendar.setDateFormat("dd/mm/y");
	    calendar.showAtElement(elem);
	}
	function SeleccionarCal(cal, date) {
		var str= new String(date);
		document.getElementById("fdVisible").value = str;
        document.getElementById("fd").value = str.substring(6, 10).concat(str.substring(3, 5).concat(str.substring(0, 2)));
		if (cal) cal.hide();
		submit();
	}
	function CerrarCal(cal) {
		cal.hide();
    }
    function submit() {
        document.getElementById("frmSel").submit();
    }
</script>
</head>
<body>
    <form id="frmSel" name="frmSel" action="cuposEmitirCarta.asp" method="POST">                         
    <div class="tableaside size100">
	    <h3> EMISI&Oacute;N DE CARTA CUPOS </h3>
        <br>        
	        <div id="searchfilter" class="tableasidecontent">	
	            <div class="col26 reg_header_navdos"> Desde </div>
                <div class="col26"> 
                    <input type="text" id="fdVisible" onclick="javascript:MostrarCalendario('fdVisible')" value="<% =GF_FN2DTE(fechaDesde) %>" />
                    <input type="hidden" id="fd" name="fd" value="<% =fechaDesde %>" />                    
                </div>
            </div>       
            <div id="searchfilter" class="tableasidecontent">
                <div class="col26 reg_header_navdos"> Planta: </div>
                <div class="col26"> 
                    <%  Call executeQuery(rsPlanta, "OPEN", "Select * from MERFL.MER192F1 where CODIDE in (10, 91, 36) order by DESCDE") %>
	                <select name="pto" id="pto" >
		                <option value=""> - Seleccione - </option>
                    <%  while (not rsPlanta.eof)                                            
                            mySelected = ""					                                                    				                
                            if (CInt(rsPlanta("CODIDE")) = myPlanta) then mySelected = "SELECTED"%>
                            <option value="<%=rsPlanta("CODIDE")%>" <%=mySelected%>> 
                                <% =rsPlanta("DESCDE")  %>
                             </option>
                    <%      rsPlanta.MoveNext()
                        wend       %>
	                </select>
                </div>
            </div>	                                    
    </div>    
    <table class="datagrid" align="center">
        <thead>
            <tr>
               <th>Contrato</th>
             <% i = 0
                auxDesde = fechaDesde 
                while (auxDesde < fechaHasta)
                  auxDesde = GF_DTEADD(fechaDesde, i, "D")                           %>
                  <th align="center" width="50px" id="th_<%=i %>">
                      <%=getDayName(auxDesde) & "<br>" & LEFT(GF_FN2DTE(auxDesde), 5) %>
                      <input type="hidden" id="colFecha_<%=i %>" name="colFecha_<%=i %>"  value="<%=auxDesde %>">
                  </th>
            <%     i = i + 1
                wend %>
            </tr>
        </thead>
        <tbody>
                <!-- ******************************* NUEVA CARGA ******************************* -->
                <tr id="trCupos">
                    <td>
                        <input type="text" id="contratoPro" name="contratoPro" value="" maxlength="2" size="1" title="Producto" onKeyPress="return controlDatos(this, event, 'N');">-
                        <input type="text" id="contratoSuc" name="contratoSuc" value="" maxlength="1" size="1" title="Sucursal" onKeyPress="return controlDatos(this, event, 'N');">-
                        <input type="text" id="contratoOpe" name="contratoOpe" value="" maxlength="2" size="1" title="Operacion" onKeyPress="return controlDatos(this, event, 'N');">-
                        <input type="text" id="contratoNum" name="contratoNum" value="" maxlength="5" size="4" title="Contrato" onKeyPress="return controlDatos(this, event, 'N');">/
                        <input type="text" id="contratoCos" name="contratoCos" value="" maxlength="2" size="1" title="Cosecha" onKeyPress="return controlDatos(this, event, 'N');">
                    </td>
                    <% i = 0
                        while (i <= diasCupos) %>
                            <td align="center">
                                <input type="text" id="cupo_<%=i %>" name="cupo_<%=i %>"  size="3"/>
                            </td>
                    <%     i = i + 1
                        wend %>
                </tr>                       
        </tbody>
    </table>                         
    <div class="col26"></div>
    <span class="btnaction">
        <input type="submit" value="Emitir Carta Cupos" ></input>
    </span>         
    <input type="hidden" name="accion" value="<% =ACCION_PROCESAR %>" />          
    </form>            
</body>
</html>