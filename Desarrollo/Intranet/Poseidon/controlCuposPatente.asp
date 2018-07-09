<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosformato.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<!--#include file="../Includes/procedimientosCupos.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<%
Call initTaskAccessInfo(TASK_POS_CONSULTA_CUPO_PATENTE, session("DIVISION_PUERTO"))
'******************************************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'********************************************************************
'					INICIO PAGINA
'********************************************************************
dim pto,rsGeneral,accion,strSQLPro,flagReport, fdesde, myWhere, patente, codigocupo, saveFechaCupo, saveCodigoCupo, savePatente, releaseCupo, releaseFcupo, reg, regSinPatente, saveMax
Dim myLog

'Por ahora solo controlo que ingrese un usuario de TOEPFER.
if (not IsToepfer(session("KCOrganizacion"))) then response.redirect "/actisaintra/comprasAccesoDenegado.asp"

totalVagones = 0
totalKilosNetos = 0
Call GP_CONFIGURARMOMENTOS()

g_strPuerto = GF_PARAMETROS7("pto", "", 6)
accion = GF_PARAMETROS7("accion", "", 6)
patente = GF_PARAMETROS7("patente", "", 6)
codigocupo = GF_PARAMETROS7("codigocupo", "", 6)
fdesde = GF_PARAMETROS7("fdesde", "", 6)
if ((fdesde = "") and (codigocupo = "") and (patente = ""))  then fdesde = Left(session("MmtoSistema"), 8)

Set myLog = new classLog
myLog.fileName = "CUPOS-MOVILE-WEB-" & Left(Session("MmtoDato"),8)
if (accion = ACCION_GRABAR) then    
    saveMax = GF_PARAMETROS7("saveMax", 0, 6)
    if (saveMax > 0) then
        For i = 1 to saveMax
            savePatente = GF_PARAMETROS7("saveP" & i, "", 6)
            saveCodigoCupo = GF_PARAMETROS7("saveC" & i, "", 6)
            saveFechaCupo = GF_PARAMETROS7("saveF" & i, "", 6)
            if (Trim(savePatente) <> "") then
                'Se agrego una nueva asociación entre cupo y patente.
                strSQL="Update CODIGOSCUPO set PATENTE='" & UCase(Trim(savePatente)) & "', MMTO=" & session("MmtoDato") & ", MOVIL='Manualmente usuario " & session("Usuario") & "' where FECHACUPO= " & saveFechaCupo & " and CODIGOCUPO='" & UCase(saveCodigoCupo) & "'"
                Call GF_BD_Puertos(g_strPuerto, rsX, "EXEC", strSQL)
                myLog.info("Cupo Asignado: " & UCase(saveCodigoCupo) & ", Patente: " & UCase(Trim(savePatente)) & ", Usuario: " & session("Usuario"))
            end if                
        Next            
    end if        
else
    if (accion = ACCION_BORRAR) then        
        releaseCupo = GF_PARAMETROS7("releaseCupo", "", 6)
        releaseFcupo = GF_PARAMETROS7("releaseFcupo", "", 6)
        strSQL="Update CODIGOSCUPO set PATENTE='', MMTO=" & session("MmtoDato") & ", MOVIL='' where FECHACUPO= " & releaseFcupo & " and CODIGOCUPO='" & UCase(releaseCupo) & "'"                
        Call GF_BD_Puertos(g_strPuerto, rsX, "EXEC", strSQL)
        myLog.info("Cupo Liberado: " & UCase(releaseCupo) & ", Usuario: " & session("Usuario"))
    end if
end if

'Se toman las horas máximas de atraso permitido para el cupo.
strSQL="Select * from parametros where CDPARAMETRO='QTHORASCAMRETCONCUPO'" 
Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)
if ((not rs.eof) and (fdesde <> "")) then fechaCupoLimite = GF_DTEADD(fdesde, CLng(Trim(rs("VLPARAMETRO")))*-1, "H")

strSQL= "Select FECHACUPO, CODIGOCUPO, DSPRODUCTO, CL.CDCLIENTE, DSCLIENTE, CC.CDCORREDOR, DSCORREDOR, CC.CDVENDEDOR, DSVENDEDOR, Case when PATENTE is Null then '' else PATENTE END PATENTE, MOVIL, CM.CDESTADO ESTADO, " &_
        "   (Select TOP 1 CDESTADO from HCAMIONES where DTCONTABLE >= DATEADD(D, -1 ,CONVERT(datetime, CONVERT(CHAR(8), FECHACUPO), 112)) and DTCONTABLE <= DATEADD(D, 1, CONVERT(datetime, CONVERT(CHAR(8), FECHACUPO), 112)) and NUCUPO = CC.CODIGOCUPO order by DTEGRESO desc) HESTADO " &_        
        "   from CODIGOSCUPO CC " &_
	    "       inner join PRODUCTOS P on P.CDPRODUCTO=CC.CDPRODUCTO " &_
	    "       inner join (Select * from CLIENTES X where CDCLIENTE = (Select MIN(CDCLIENTE) from CLIENTES Y where Y.NUCUIT=X.NUCUIT)) CL on CL.NUCUIT=CC.CUITCLIENTE " &_
	    "       left join CORREDORES C on C.CDCORREDOR=CC.CDCORREDOR " &_
	    "       left join VENDEDORES V on V.CDVENDEDOR=CC.CDVENDEDOR "	&_    
	    "       left join CAMIONES CM on CM.NUCUPO=CC.CODIGOCUPO"	    
myWhere = " where ESTADO > " & CUPO_PROVISORIO
if (fdesde <> "") then      Call mkWhere(myWhere, "FECHACUPO", fdesde, "=", 1)
if (codigocupo <> "") then  Call mkWhere(myWhere, "CODIGOCUPO", UCASE(CODIGOCUPO), "=", 3)
if (patente <> "") then     Call mkWhere(myWhere, "PATENTE", UCASE(patente), "=", 3)
strSQL= strSQL & myWhere
strSQL= strSQL & " ORDER BY FECHACUPO DESC, CODIGOCUPO"
Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL) 
%>

<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=9">

<title><%=GF_TRADUCIR("Puertos - Patentes Asignadas a Cupos")%></title>
<link rel="stylesheet" href="../css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="../css/iwin.css" type="text/css">
<link rel="stylesheet" href="../css/calendar-win2k-2.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="../css/main.css" type="text/css"> 

<script type="text/javascript" src="../scripts/formato.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/calendar-1.js"></script>
<script type="text/javascript" src="../scripts/paginar.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script language="javascript">
    var errores = new Array();
    
    function CerrarCal(cal) {
        cal.hide();
    }
    function MostrarCalendario(p_objID, funcSel) {
        var dte = new Date();
        var elem = document.getElementById(p_objID);
        if (calendar != null) calendar.hide();
        var cal = new Calendar(false, dte, funcSel, CerrarCal);
        cal.weekNumbers = false;
        cal.setRange(1993, 2045);
        cal.create();
        calendar = cal;
        calendar.setDateFormat("y-mm-dd");
        calendar.showAtElement(elem);
    }

    function SeleccionarCalDesde(cal, date) {
        var str = new String(date);
        document.getElementById("fdesdeP").value = str;
        document.getElementById("fdesde").value = str.replace(/-/g, '');
        if (cal) cal.hide();
        swapInputEraser(fdesde);
    }
    
    function submitInfo() {
        document.getElementById("frmSel").submit();
    }

    function swapInputEraser(rootId) {
        if (document.getElementById(rootId).value == "") {
            document.getElementById(rootId + "Img").style.display = "none";
        } else {
            document.getElementById(rootId + "Img").style.display = "inline";
        }
    }

    function eraseInput(rootId) {
        document.getElementById(rootId).value = "";
        document.getElementById(rootId + "P").value = "";
        swapInputEraser(rootId);
    }

    function bodyOnLoad() {
        swapInputEraser("fdesde");
    }

    function savePatente() {
        var totalErrores = 0;
        for (var i in errores) { totalErrores += errores[i]; }
        if (totalErrores == 0) {
            var s = document.getElementsByName("btnSavePatente");
            var l = document.getElementsByName("loadingPatente");            
            for (var i = 0; i < s.length; i++) {
                s[i].style.display = "none";
                l[i].style.display = "inline";    
            }
            document.getElementById('accion').value = '<% =ACCION_GRABAR %>';
            submitInfo();
        } else {
            alert("Hay " + totalErrores + " patente/s con error.");
        }
    }

    function releasePatente(cupo, fcupo) {
        if (confirm("Desea realmente liberar el cupo " + cupo + " de la fecha " + fcupo + "?")) {
            document.getElementById('releaseCupo').value = cupo;
            document.getElementById('releaseFcupo').value = fcupo;
            document.getElementById('accion').value = '<% =ACCION_BORRAR %>';
            submitInfo();
        }          
    }
    
    function controlPatente(obj, id) {
        var pat = obj.value.trim().toUpperCase();
        obj.style.background = "";
        obj.style.color = "";
        errores[id] = 0;
        if (pat != "") {                        
            var le = pat.length;
            if (le == 6) {
                //Foramto viejo.
                var let = pat.substring(0, 3);
                var num = pat.substring(3, 6);
                if (!isNaN(let) || isNaN(num)) {
                    obj.style.background = "#FF0000";
                    obj.style.color = "#FFFFFF";
                    errores[id] = 1;              
                }
            } else {
                if (le == 7) {
                    //Foramto viejo.
                    var let = pat.substring(0, 2);
                    var num = pat.substring(2, 5);
                    var let2 = pat.substring(5, 7);
                    obj.style.background = "#FFFFFF";
                    if (!isNaN(let) || isNaN(num) || !isNaN(let2)) {
                        obj.style.background = "#FF0000";
                        obj.style.color = "#FFFFFF";
                        errores[id] = 1;
                    }
                } else {
                    obj.style.background = "#FF0000";
                    obj.style.color = "#FFFFFF";
                    errores[id] = 1;
                }
            }
        }      
    }
  </script>
</head>

<body onload="bodyOnLoad()">

<div id="toolbar"></div>

<form method="POST" name="frmSel" id="frmSel" action="controlCuposPatente.asp">	
<div class="tableaside size100"> <!-- BUSCAR -->
    <h3> filtro - <%=GF_Traducir("Patentes Asignadas a Cupos")%> </h3>
    
    <div id="searchfilter" class="tableasidecontent">
        <div class="col16 reg_header_navdos"> <% = GF_TRADUCIR("Fecha Cupo") %> </div>    
        <div class="col16"> 
			<input type="text" name="fdesdeP" id="fdesdeP" size="15" readonly onClick="javascript:MostrarCalendario('fdesdeP', SeleccionarCalDesde)" value="<% =GF_FN2DTE(fdesde) %>">
			<input type="hidden" name="fdesde" id="fdesde" value="<% =fdesde %>">
			<img id="fdesdeImg" src="../images/icon_del.gif" style="display: none; cursor: pointer" onclick="eraseInput('fdesde')"/>
        </div>                        
        <div class="col16 reg_header_navdos"> C&oacute;digo Cupo </div>    
        <div class="col16"> 
		    <input type="text" name="codigocupo" id="codigocupo" size="15" maxLength="12" value="<% =codigocupo %>">			    
        </div>  
        <div class="col16 reg_header_navdos"> Patente </div>    
        <div class="col16"> 
		    <input type="text" name="patente" id="patente" size="15" maxLength="10" value="<% =patente %>">
        </div>  
        <span class="btnaction"><input type="submit" value="Buscar" id=submit1 name=submit1></span>
    </div>
</div><!-- END BUSCAR -->

<div class="col66"></div>
        <TABLE class="datagrid" id="TABLE1" align="center" width="300px">
            <thead >
		        <th>Cupos Totales</th>
		        <th>Cupos Con Patente</th>
		    </thead>
		    <tbody>
		        <tr>
		            <td align="center" width="150px"><div id="cuposTotales"></div></td>
		            <td align="center"><div id="cuposPatentes"></div></td>
		        </tr>		        
		    </tbody>
		<TABLE class="datagrid" id="TAB1" align="center" width="90%">
		<thead >
		    <th>Fecha</th>
		    <th>C&oacute;digo de Cupo</th>
		    <th>Producto</th>		    
		    <th>Destinatario</th>
		    <th>Corredor</th>
		    <th>Vendedor</th>
		    <th>Patente</th>
		    <th>ID Movil</th>
		    <th>Estado</th>
		    <th>.</th>
		</thead>
        <tbody>
			<%reg = 0
			regSinPatente = 0
			regProcesable = 0
			while not rs.eof
				reg = reg + 1 
%>
                <tr>
                    <td align="center"><%  =GF_FN2DTE(rs("FECHACUPO")) %></td>
                    <td align="center"><%  =rs("CODIGOCUPO") %></td>
                    <td align="center"><%  =rs("DSPRODUCTO") %></td>
                    <td><%  =rs("DSCLIENTE") %></td>
                    <td><%  if (CLng(rs("CDCORREDOR")) <> 0) then
                                response.Write rs("DSCORREDOR") 
                            end if%>
                    </td>
                    <td><%  if (CLng(rs("CDVENDEDOR")) <> 0) then
                                response.Write rs("DSVENDEDOR") 
                            end if%>
                    </td>
                    <td align="center">
                    <%  if (Trim(rs("PATENTE")) = "") then                    
                            regSinPatente = regSinPatente + 1                            
                            if ((IsNull(rs("HESTADO"))) and (IsNull(rs("ESTADO")))) then
                                if (CheckAccess(TASK_POS_ASOCIAR_CUPO_PATENTE, session("DIVISION_PUERTO"))) then                                                                                       
                                regProcesable = regProcesable + 1        
                    %>                        
                        <input type="text" name="saveP<% =regProcesable %>" id="saveP<% =regProcesable %>" size="10" maxlength="10" onblur="controlPatente(this, <% =regSinPatente %>)" />                         
	                    <input type="hidden" id="saveF<% =regProcesable %>" name="saveF<% =regProcesable %>" value="<% =rs("FECHACUPO") %>">
	                    <input type="hidden" id="saveC<% =regProcesable %>" name="saveC<% =regProcesable %>" value="<% =rs("CODIGOCUPO") %>">                    
                    <%          end if    
                            end if
                        else                            
                            response.Write rs("PATENTE")
                        end if                            
                    %>
                    </td>
                    <td align="center"><%  =rs("MOVIL") %></td>
                    <td align="center">
                    <%  if (not IsNull(rs("HESTADO"))) then 
                            response.Write "FINALIZADO"
                        else
                            if (not IsNull(rs("ESTADO"))) then 
                                response.Write "INGRESADO"
                            else
                                response.write "NO ARRIBADO"
                            end if
                        end if
                    %>
                    </td>
                    <td>
                    <%  if ((IsNull(rs("HESTADO"))) and (IsNull(rs("ESTADO")))) then
                            if (Trim(rs("PATENTE")) = "") then    
                                if (CheckAccess(TASK_POS_ASOCIAR_CUPO_PATENTE, session("DIVISION_PUERTO"))) then
                        %>
                                <img onclick="javascript:savePatente();" name="btnSavePatente" src="images/save-16.png" style="cursor: pointer;"/>                                                        
                        <%      end if
                            else                           
                                if (CheckAccess(TASK_POS_LIBERAR_CUPO_PATENTE, session("DIVISION_PUERTO"))) then
                        %> 
                                <img onclick="javascript:releasePatente('<%  =rs("CODIGOCUPO") %>', '<%  =rs("FECHACUPO") %>');" name="btnReleasePatente" src="images/cancel-16x16.png" style="cursor: pointer;"/>                        
                        <%      end if
                            end if  
                        end if
                        %>                            
                        <img name="loadingPatente" src="../images/loading_small_green.gif" title="Guardando" alt="Guardando" style="display:none;"/>                        
                    </td>
				</tr>
			<%
				rs.MoveNext()
			wend %>	  	                 	
  	<%
	if (reg = 0) then		
		%>
		<tr>
			<td colspan="10" align="center"><% =GF_TRADUCIR("No se encontraron resultados.") %></td>
		</tr>
		<%
	end if 
	%>	
	    </tbody>
	</TABLE>	
	<input type="hidden" id="accion" name="accion" value="<% =ACCION_SUBMITIR %>">
	<input type="hidden" id="pto" name="pto" value="<% =g_strPuerto %>">			
	<input type="hidden" id="saveMax" name="saveMax" value="<% =regProcesable %>">	
	<input type="hidden" id="releaseCupo" name="releaseCupo" value="">
	<input type="hidden" id="releaseFcupo" name="releaseFcupo" value="">				
</form>
<script type="text/javascript">
    document.getElementById("cuposTotales").innerHTML = <% =reg %>;
    document.getElementById("cuposPatentes").innerHTML = <% =reg - regSinPatente %> + ' (' + <% if (reg > 0) then response.write round((reg - regSinPatente)*100/reg, 0) else response.write 0 end if %> + ' %)';
</script>
</body>
</html>
