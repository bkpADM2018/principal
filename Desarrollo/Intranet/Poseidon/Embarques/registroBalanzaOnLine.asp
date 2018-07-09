<!--#include file="../../Includes/procedimientosSQL.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientosFechas.asp"-->
<!--#include file="../../Includes/procedimientosPuertos.asp"-->
<!--#include file="../../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../../Includes/procedimientosFormato.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<%
'------------------------------------------------------------------------------------------------
Function getTotalEmbarqueActual(pPto, pBza)
    Dim rs, strSQL, salir, totalGeneral, myBuque, myProducto, salirProd, totalProducto, myWhere, bloqueProducto
        
    'Obtengo el ultimo buque.        
    if (pBza <> "") then myWhere = "where Balanza = '" & pBza & "'"
    strSQL="Select BUQUE, CONCAT(COMMODITY, CONCAT('/', Exportador)) COMMODITY, Left(MMTO, 8) Fecha, TURNO, BODEGA, PESO from EMBARQUESREGISTROPESO " & myWhere & " order by MMTO desc"
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)            
    if (not rs.Eof) then
        myBuque = Trim(rs("BUQUE"))
        totalGeneral = 0
        salir = false    
        while ((not rs.eof) and (not salir))
            myProducto = Trim(rs("COMMODITY"))
            salirProd = false
            bloqueProducto = ""
            totalProducto = 0
            while ((not rs.eof) and (not salir) and (not salirProd))
                if (myBuque = Trim(rs("BUQUE"))) then
                    if (myProducto = Trim(rs("COMMODITY"))) then                
                        bloqueProducto = space(5) & "Fecha...: " & GF_FN2DTE(rs("FECHA")) & " Turno...: " & GF_nChars(rs("TURNO"), 10, " ", CHR_AFT) & " Bodega..: " & GF_nChars(rs("BODEGA"), 10, " ", CHR_AFT) & " Cargado: " & GF_nChars(GF_EDIT_DECIMALS(rs("PESO"), 0), 10, " ", CHR_FWD) & " Kg<br>" & bloqueProducto           
                        totalProducto = totalProducto + CLng(rs("PESO"))        
                        rs.MoveNext()
                    else
                        salirProd = true
                    end if            
                else
                    salir = true                
                end if                
            wend            
            bloqueProducto = space(3) & "Producto/Expo.: " & GF_nChars(myProducto, 16, " ", CHR_AFT) & "<br>" & bloqueProducto           
            bloqueProducto = bloqueProducto & GF_nChars("-----------------------", 91, " ", CHR_FWD) & "<br>"            
            bloqueProducto = bloqueProducto & GF_nChars("SUBTOTAL: " & GF_nChars(GF_EDIT_DECIMALS(totalProducto, 0), 10, " ", CHR_FWD), 87, " ", CHR_FWD) & " Kg<br>"
            getTotalEmbarqueActual = bloqueProducto & getTotalEmbarqueActual
            totalGeneral = totalGeneral + totalProducto        
        wend
        getTotalEmbarqueActual = "Buque...: " & GF_nChars(myBuque, 20, " ", CHR_AFT) & "<br>" & getTotalEmbarqueActual    
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<br>" & GF_nChars("TOTAL BUQUE: " & GF_nChars(GF_EDIT_DECIMALS(totalGeneral, 0), 10, " ", CHR_FWD), 87, " ", CHR_FWD) & " Kg"
    end if
End function
'-----------------------------------------------------------------------------------------------
Dim cdBalanza, g_strPuerto, idRegistro, accion, rsBza
Dim lastRecord, data, strTotales

Call initTaskAccessInfo(TASK_POS_BZA_EMB, session("DIVISION_PUERTO"))

Call GP_CONFIGURARMOMENTOS

g_strPuerto = GF_PARAMETROS7("pto","",6)
cdBalanza= GF_PARAMETROS7("bza", "", 6)
idRegistro = GF_PARAMETROS7("reg", 0, 6) 
accion = GF_PARAMETROS7("accion", "", 6) 
if (accion = "") then
    'Tomo el ultimo registro emitido como punto de inicio para la informacion.
    strSQL="Select MAX(IDREGISTRO) IDREGISTRO from EMBARQUESREGISTROBALANZA where BALANZA='" & cdBalanza & "'"
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)        
    'lastRecord=0
    if (not rs.eof) then lastRecord=CLng(rs("IDREGISTRO"))
    strTotales = getTotalEmbarqueActual(g_strPuerto, cdBalanza)
else if (accion=ACCION_PROCESAR) then    
        strTotales = getTotalEmbarqueActual(g_strPuerto, cdBalanza)
        'Se solicitan los registros que se hayan generado desde la ultima consulta.
        strSQL="Select IDREGISTRO, REGISTRO from EMBARQUESREGISTROBALANZA where BALANZA='" & cdBalanza & "' and IDREGISTRO > " & idRegistro        
        Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)                    
        data = ""         
        lastRecord = idRegistro   
        if (not rs.eof) then                         
            while (not rs.eof) 
                data = data & rs("REGISTRO") & "#"                        
                lastRecord = rs("IDREGISTRO")
                rs.MoveNext()
            wend            
        end if  
        data = lastRecord & "|" & data & "|" & strTotales & "|" & GF_FN2DTE(session("MmtoSistema"))            
        response.write data
        response.end
    end if
end if    
%>

<html>
<head>
<style>
div {
    width: 100%;
    font-family: Courier;
}
#infoPesada {
    text-align:left;           
    height: 250px;
    overflow: auto;                
    border-width : 1px;
    border-style: solid;
    border-color: #CCCCCC;
    font-size : medium;
}
            
#infoTotales {                     
    height: 350px;
    overflow: auto;           
    text-align: left;
    border-width : 1px;
    border-style: solid;
    border-color: #CCCCCC;
    text-weight: bold;
    color: #000000;
    font-size : medium;    
}

#infoUpdate {                                
    text-align: right;
    color: #777777;
    font-size : small;
    padding-top : 3px;
}

body{
    font-family:'Courier New';
    font-size:20px;
    text-align:center;
    -webkit-user-select: none;
    -khtml-user-select: none;
    -moz-user-select: none;
    -o-user-select: none;
    -ms-user-select: none;
    user-select: none;
}
.titulo {
    color:white;
    float: left;
    width: calc(100%/3 - 2*1em - 2*1px);
    width: 100%;
    margin: auto;
    background-color: #2e6b4d;
    border: 1px solid;
    border-color:lightgray;
    border-radius: 7px;
}
</style>
    <script type="text/javascript" src="../../scripts/channel.js"></script>
    <script language="Javascript" type="text/javascript">
        var lastRecord = <% =lastRecord %>;
        var ch = new channel();
            
        function bodyOnLoad() {
            setInterval(loadPesada, 20000);
            var info = document.getElementById("infoUpdate");                                
            info.innerHTML = "Ult. Actualizaci&oacute;n <% =GF_FN2DTE(session("MmtoSistema")) %>";
        }

        function loadPesada() {
            ch.bind('registroBalanzaOnLine.asp?pto=<% =g_strPuerto %>&bza=<% =cdBalanza %>&reg=' + lastRecord + '&accion=<% =ACCION_PROCESAR %>','loadPesada_callBack()');
			ch.send();   
        }
            
        function loadPesada_callBack() {
            var resp = ch.response();                
            var data = resp.split('|');                
            lastRecord = data[0];
            if (data[1] != "") {                
                resp = data[1].replace(/#/g,"<br>");                
                resp = resp.replace(/ /g,"&nbsp;");    
                var dest = document.getElementById("infoPesada");
                dest.innerHTML = dest.innerHTML + resp;                    
                dest.scrollTop = dest.scrollHeight;
            }             
            var tot = document.getElementById("infoTotales");                
            tot.innerHTML = data[2].replace(/ /g,"&nbsp;");
            tot.scrollTop = tot.scrollHeight;
            var info = document.getElementById("infoUpdate");                
            info.innerHTML = "Ult. Actualizaci&oacute;n " + data[3];
        }
        document.oncontextmenu = function () { return false }
    </script>
</head>
<body onload="bodyOnLoad()" style="cursor:default">
    <h3><% =getNombrePuerto(g_strPuerto) %></h3>
    <h3>CABEZAL: <% =cdBalanza %></h3>
    <div id="infoTotales">Cargando Valores....</div>
    <div id="infoPesada"></div>
    <div id="infoUpdate"></div>
</body>
</html>
