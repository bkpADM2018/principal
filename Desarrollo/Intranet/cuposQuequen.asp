<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/includeGeneracionArchivos.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<%

Const PUERTO_TERMINAL_QUEQUEN = "20, 63"
Const PRODUCT_UNKNOWN = "ERROR"

Function generateFile(pCdProducto, pDteDesde, pDteHasta)
    
    Dim strPath, fs, arch, strSQL, rs, myWhere
    
    'Establezco la ruta y el nombre del archivo a crear
    strPath = Server.mapPath("temp\cuposQuequen.txt")
    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath) Then  call fs.deleteFile(strPath, true)
    set arch = fs.CreateTextFile(strPath)    
    'Se escriben los titulos.    
    Call writeFile(arch, "Fecha", false)
    Call writeFile(arch, "Cereal", false)
    Call writeFile(arch, "Cargador", false)
    Call writeFile(arch, "Corredor", false)
    Call writeFile(arch, "Cantidad", true)
    'Se aplican los filtros para los datos
    myWhere = "where CUCDES in (" & PUERTO_TERMINAL_QUEQUEN & ")"
    if (pCdProducto > 0) then   Call mkWhere(myWhere, "CUCPRO",  pCdProducto, "=", 1)
    if (pDteDesde > 0) then     Call mkWhere(myWhere, "CUFCCP",  pDteDesde, ">=", 1)
    if (pDteHasta > 0) then     Call mkWhere(myWhere, "CUFCCP",  pDteHasta, "<=", 1)
    'Se obtienen los datos de cupos.
    strSQL= "Select CUFCCP, CUCPRO, CUCOOR, NRODOC, sum(CUCCCP) CUPOS from " & _
            "(Select CUFCCP, CUCPRO, CUCOOR, CUCCCP, Case CUCCOR when 15 then CUCVEN else CUCCOR end CORREDOR from MERFL.MER517F1 " & myWhere & ") A " & _            
            " inner join MERFL.TCB6A1F1 B on A.CORREDOR=B.NROPRO " & _            
            " group by CUFCCP, CUCPRO, CUCOOR, NRODOC" & _
            " order by CUFCCP, CUCPRO, CUCOOR, NRODOC"
    Call executeQuery(rs, "OPEN", strSQL)
    while (not rs.eof)
        Call writeFile(arch, rs("CUFCCP"), false)
        Call writeFile(arch, translateProduct(CInt(rs("CUCPRO"))), false)
        Call writeFile(arch, rs("CUCOOR"), false)
        Call writeFile(arch, rs("NRODOC"), false)
        Call writeFile(arch, rs("CUPOS"), true)
        rs.MoveNext()
    wend    
    arch.close
    set arch=nothing
    set fs=nothing
    
    generateFile = strPath
End Function
'---------------------------------------------------------------------------------------------
'Funcion responsable de escribir el archivo respetando el formato de salida.
Sub writeFile(byref p_arch, p_strValor, isLast)    
    Dim suffix
    
    suffix = ","
    if (isLast) then suffix = chr(13) & chr(10)
    p_arch.write chr(34) & p_strValor & chr(34) & suffix
    
End Sub
'---------------------------------------------------------------------------------------------
'Funcion responsable de traducir los c{odigos de producto de TOEPFER a Terminal Quequen
Function translateProduct(pCdProducto)
   
    Select Case pCdProducto
        Case 15
            translateProduct = "TRIGO"
        Case 19
            translateProduct = "MAIZ"
        Case 23
            translateProduct = "SOJA"
        Case 17
            translateProduct = "CEBADACE"
        Case 24
            translateProduct = "CEBADAFO"
        Case Else
            translateProduct = PRODUCT_UNKNOWN
    End Select
    
End Function
'/*******************************************\
' *         COMIENZO DE LA PAGINA           *
'\*******************************************/ 
Dim dteDesde, dteHasta, action, cdProducto, fileName

Call GP_ConfigurarMomentos()

action = GF_PARAMETROS7("action", "", 6)
cdProducto = GF_PARAMETROS7("cdProducto", 0, 6)

dteDesde = GF_PARAMETROS7("txtFechadesde","",6)
if (dteDesde = "") then dteDesde = GF_FN2DTE(left(session("MmtoDato"), 8))

dteHasta = GF_PARAMETROS7("txtFechaHasta","",6)
if (dteHasta = "") then dteHasta = dteDesde

if (action = ACCION_PROCESAR) then        
    fileName = generateFile(cdProducto, GF_DTE2FN(dteDesde), GF_DTE2FN(dteHasta))    
    Call Descargar(fileName)
    response.end
end if

%>
<html>
<head>
    <title>Sistema de Mercaderias - Cupos Terminal Quequen</title>
    
    <link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
    <!-- Estilo de Calendario -->
    <link rel="stylesheet" type="text/css" media="all" href="CSS/calendar-win2k-2.css" title="win2k-2" />
    
    <script type="text/javascript" src="scripts/channel.js"></script>
    <!-- Script del calendario -->
    <script type="text/javascript" src="scripts/calendar.js"></script>
    <!-- Modulo de Lenguaje -->
    <script type="text/javascript" src="scripts/calendar-<% =GF_GET_IDIOMA() %>.js"></script>
    
    <script type="text/javascript">
        /* Variable globales de calendarios */
        var nombre = null;
        var cals = new Array();
        cals[0] = null;
        cals[1] = null;
        /* Funcion de Seleccion del Calendario */
        function SeleccionarCal(cal, date) {
            var obj = document.getElementById(nombre);
            var str = new String(date);
            obj.value = str;
            cal.hide();
        }
        /* Funcion de Cierre del Calendario */
        function CerrarCal(cal) {
            nombre = null;
            cal.hide();
        }
        /* Funcion para mostrar calendario */
        function MostrarCalendario(p_cal, p_name) {
            var dte = new Date();
            var obj = document.getElementById(p_name);
            var strFecha = obj.value;
            var arrFecha = strFecha.split("/");
            dte.setFullYear(arrFecha[2], arrFecha[1] - 1, arrFecha[0]);
            if (cals[p_cal] != null) {
                // Ya hay un calendario creado, se oculta.
                cals[p_cal].hide();
            }
            else {
                cal = new Calendar(false, dte, SeleccionarCal, CerrarCal);
                cal.weekNumbers = false;
                cal.setRange(2010, 2050);
                cal.create();
            }
            nombre = p_name;
            cals[p_cal] = cal;
            cals[p_cal].setDateFormat("dd/mm/y");
            cals[p_cal].showAtElement(obj);            
        }      
    </script>    
</head>
<body>
<form name="frmSel" method="POST">
    <% Call GF_TITULO2("kogge64.gif", "Exportacion de Cupos Terminal Quequen") %>	
    <table class="reg_header" width="300px" align="center" border="0">
        <tr class="reg_header_nav">
            <td colspan="3">Filtros</td>
        </tr>        
        <tr class="reg_header_navdos" cellspacing="1" cellpadding="2">
            <td><% =GF_TRADUCIR("Fecha Desde")%></td>
            <td align="center"><input type="text" name="txtFechaDesde" id="txtFechaDesde" size="10" value="<% =dteDesde %>"></td>
            <td align="center" width="16px"><img align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario(0,'txtFechaDesde')"></td>
        </tr>
        <tr class="reg_header_navdos" cellspacing="1" cellpadding="2">
            <td><% =GF_TRADUCIR("Fecha Hasta")%></td>
            <td align="center"><input type="text" name="txtFechaHasta" id="txtFechaHasta" size="10" value="<% =dteHasta %>"></td>
            <td align="center" width="16px"><img align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:pointer" onclick="MostrarCalendario(1,'txtFechaHasta')"></td>
        </tr>
        <tr class="reg_header_navdos" cellspacing="1" cellpadding="2">
            <td><% =GF_TRADUCIR("Producto")%></td>
            <td align="center" colspan="2">
                <select name="cdProducto" id="cdProducto">
                    <option value="0">- Todos -</option>
                    <option value="17">Cebada Cervecera</option>
                    <option value="24">Cebada Forrajera</option>
                    <option value="19">Maiz</option>
                    <option value="15">Trigo</option>
                    <option value="23">Soja</option>
                </select>            
            </td>            
        </tr>
        <tr>
            <td colspan="3" align="center"><input type="submit" value="Exportar" /></td>
        </tr>
    </table>
    <input type="hidden" name="action" id="action" value="<% =ACCION_PROCESAR %>" />
</form>
</body>
</html>
