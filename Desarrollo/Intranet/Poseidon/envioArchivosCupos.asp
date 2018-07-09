<!-- #include file="../Includes/procedimientosUnificador.asp"-->
<!-- #include file="../Includes/procedimientosParametros.asp"-->
<!-- #include file="../Includes/procedimientosfechas.asp"-->
<!-- #include file="../Includes/procedimientosMail.asp" -->
<!-- #include file="../Includes/procedimientosSeguridad.asp" -->
<!-- #include file="../Includes/procedimientosFormato.asp" -->
<!-- #include file="../Includes/procedimientos.asp" -->
<!-- #include file="../Includes/procedimientosCupos.asp" -->
<!-- #include file="../Includes/procedimientosPuertos.asp" -->
<!-- #include file="../Includes/includeGeneracionArchivos.asp" -->
<!-- #include file="../Includes/procedimientostraducir.asp"-->
<%
Server.ScriptTimeout = 1200
Const TIPO_CORREDOR = "C"
Const TIPO_VENDEDOR = "V"
'*****************************************************************************************
function getEnterpriseData(pTipo, cdEmpresa, ByRef dsEmpresa, ByRef cuitEmpresa)
    dim strSQl, rs

    if (session("PLAYA_CUIT_" & cdEmpresa) = "") then
        cuitEmpresa = "00-00000000-0"
        dsEmpresa = ""
        if isnumeric(cdEmpresa) then
            if (CLng(cdEmpresa) <> 0) then
                if (pTipo = TIPO_VENDEDOR) then
                    strSQL = "select NUDOCUMENTO, DSVENDEDOR DS from VENDEDORES where CDVENDEDOR=" & cdEmpresa
                else
                    strSQL = "select NUCUIT NUDOCUMENTO, DSCORREDOR DS from CORREDORES where CDCORREDOR=" & cdEmpresa
                end if
                Call executeQueryDb(DBSITE_BAHIA, rs, "OPEN", strSQL)
                if not rs.eof then
                    cuitEmpresa = GF_STR2CUIT(Trim(rs("NUDOCUMENTO")))
                    dsEmpresa = Trim(rs("DS"))
                end if
            end if            
        end if
        session("PLAYA_CUIT_" & cdEmpresa) = cuitEmpresa
        session("PLAYA_DESC_" & cdEmpresa) = dsEmpresa
    else
        cuitEmpresa = session("PLAYA_CUIT_" & cdEmpresa)
        dsEmpresa = session("PLAYA_DESC_" & cdEmpresa)
    end if        
end function
'*****************************************************************************************
Function generarArchivo(p_dteDate)
    dim arch, fs, strPath, CU_KCVEN, CU_KCCOR, CU_KCCOR_CUIT, CU_KCVEN_CUIT, CU_DS_KCCOR, CU_DS_KCVEN
	dim CU_NroCupo, CU_NroCupoDesde, CU_NroCupoHasta, CU_NroCupoTxt, CU_CantCamiones, CU_producto
	dim CU_auxFecha, CU_CuitDestinatario, CU_dsDestinatario, CU_lugarDescarga	
	
    'Establezco la ruta y el nombre del archivo a crear
    strPath = Server.mapPath("..\temp\cuposadmagro.txt")	
    'Si existe la borro
    set fs = Server.CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(strPath) Then  call fs.deleteFile(strPath, true)
    set arch = fs.CreateTextFile(strPath)
    set fs = nothing
    
    g_strPuerto = DBSITE_BAHIA
        
    Call escribirCabeceraArchivo(arch)
    
    strSQL="Select C.*, P.DSPRODUCTO from CODIGOSCUPO C inner join PRODUCTOS P on C.CDPRODUCTO=P.CDPRODUCTO where FECHACUPO >= " & p_dteDate & " and ESTADO > " & CUPO_CANCELADO    
    Call executeQueryDB(DBSITE_BAHIA, rs, "OPEN", strSQL)
    
    Call escribirDetalleArchivo(arch, rs, DBSITE_BAHIA)

    arch.close()
    set arch = Nothing
    generarArchivo = strPath
end function
'*****************************************************************************************    
Function escribirCabeceraArchivo(ByRef arch)
    
    Call escribirArchivo(arch, "99999999",8)
    Call escribirArchivo(arch, getDsClienteByCUIT(CUIT_TOEPFER),20) 'Destinatario = Toepfer
    Call escribirArchivo(arch, "",131) 'El resto en blanco
    Call escribirArchivo(arch, "!" & chr(13) & chr(10),3)    
    g_cant = 0            
    
End Function    
'*****************************************************************************************
Function escribirDetalleArchivo(ByRef arch, rs, pPto)
            
    'dim dicCupos
	Dim cuposMaxInforme, myTotalEnviados, myPatente
	'Set dicCupos = Server.CreateObject("Scripting.Dictionary")	
	CU_CantCamiones = 1
	CU_lugarDescarga = getNumeroPuerto(pPto)				
    while (not rs.eof)               
        CU_CuitDestinatario = rs("CUITCLIENTE")
        CU_dsDestinatario = getDsClienteByCUIT(CU_CuitDestinatario)        
		CU_auxFecha = right(rs("FECHACUPO"),2) & mid(rs("FECHACUPO"),5,2) & left(rs("FECHACUPO"), 4)		
	    CU_KCVEN = CLng(rs("CDVENDEDOR"))
	    CU_KCCOR = CLng(rs("CDCORREDOR"))		    
	    Call getEnterpriseData(TIPO_VENDEDOR, CU_KCVEN, CU_DS_KCVEN, CU_KCVEN_CUIT)		
	    Call getEnterpriseData(TIPO_CORREDOR, CU_KCCOR, CU_DS_KCCOR, CU_KCCOR_CUIT)		    
		
		call escribirArchivo(arch, CU_auxFecha,8)
		call escribirArchivo(arch, CU_dsDestinatario,20)
		call escribirArchivo(arch, rs("DSPRODUCTO"), 20)
		call escribirArchivo(arch, CU_lugarDescarga,2)
		call escribirArchivo(arch, CU_Cantcamiones,10)
		call escribirArchivo(arch, rs("CODIGOCUPO"),10)
		call escribirArchivo(arch, CU_DS_KCVEN,20)
		call escribirArchivo(arch, CU_DS_KCCOR,20)
		call escribirArchivo(arch, GF_STR2CUIT(CU_CuitDestinatario),13)
		call escribirArchivo(arch, CU_KCVEN_CUIT,13)
		call escribirArchivo(arch, CU_KCCOR_CUIT,13)
		call escribirArchivo(arch, rs("PATENTE"),10)
		call escribirArchivo(arch, "!" & chr(13) & chr(10),3)			
		g_cant = g_cant + 1					
        rs.movenext
    wend    
End Function        
'*****************************************************************************************
sub escribirArchivo(byref p_arch, p_strValor, p_longitud)
    dim k
    if len(p_strValor) <= p_longitud then
        p_arch.write p_strValor
        for k = len(p_strValor) + 1 to p_longitud
            p_arch.write " "
        next
    else
        p_arch.write left(p_strValor, p_longitud)
    end if
end sub
'*****************************************************************************************
sub enviarMailCupos()
    Dim strAsunto, strPathAttachment, strBody


    'completo los datos del mail
    strAsunto = "Cupos ADM AGRO"
    strPathAttachment = Server.mapPath("..\temp\cuposadmagro.txt")    
    strBody = "Se adjuntan los cupos asigandos. Total:" & g_cant	
	
	Call SendMail(TASK_POS_CUPOS_TYL, MAIL_TASK_INFO_LIST, strAsunto, strBody, strPathAttachment)    

end sub
'*****************************************************************************************
sub writeLog(p_type, p_message)
    dim archLog, fs, strPathLog

    strPathLog = Server.mapPath("LOGS\logEnvioAutomaticoCupos.txt")
    set fs = Server.CreateObject("Scripting.FileSystemObject")
    if not fs.FileExists(strPathLog) then
        set archLog = fs.CreateTextFile(strPathLog)
    else
        set archLog = fs.OpenTextFile(strPathLog, 8)
    end if
    archLog.writeline replace(session("MomentoSistema"),"'","") & "  " & p_type & "  " & p_message
    archLog.close

    set fs = Nothing
    set archLog = Nothing
end sub
'*****************************************************************************************
Function Bach(p_dteDate)

dim conn, strSQL, rs, strPath

call writeLog("INF","Inicio proceso")
Call writeLog("INF","Generando Archivo")
call generarArchivo(p_dteDate)
Call writeLog("INF","Enviando Mails...")
call enviarMailCupos()
call writeLog("INF","Fin proceso")
call writeLog("INF", "--------------------------------------------------")

End Function
'*****************************************************************************************
'******                           COMIENZO DE LA PAGINA                             ******
'*****************************************************************************************
'Este dictionary lo uso como cache para las ds de los productos, para no tener
'q ir a buscar a la BD
dim g_diccProductos, dteFecha, intFecha, g_strPuerto
dim p_oper, strPath, g_cant, g_diccPatentes, diccNominaciones

p_oper = GF_PARAMETROS7("P_ACTION","",6)
if (p_oper = "") then p_oper="BACH"
dteFecha = GF_PARAMETROS7("txtFechadesde","",6)

set g_diccProductos = Server.CreateObject("Scripting.Dictionary")
set g_diccPatentes = Server.CreateObject("Scripting.Dictionary")

'Abro una sesion para tener los momentos y un usuario para las consultas a la base
session("Usuario")="GUEST"
GP_CONFIGURARMOMENTOS
Call GF_SET_IDIOMA(1)

intFecha = mid(session("MomentoSistema"),2, 8)
if (dteFecha = "") then dteFecha = GF_FN2DTE(intFecha)
if (p_oper = "BACH") then
    Call Bach(intFecha)
end if
if (p_oper <> "BACH") then
%>
<html>
<head>
  <title></title>
  <link rel="stylesheet" href="CSS/ActiSAIntra-1.css" type="text/css">
  <!-- Estilo de Calendario -->
  <link rel="stylesheet" type="text/css" media="all" href="CSS/calendar-win2k-2.css" title="win2k-2" />
  <!-- Script del calendario -->
  <script type="text/javascript" src="scripts/calendar.js"></script>
  <!-- Modulo de Lenguaje -->
  <script type="text/javascript" src="scripts/calendar-<% =GF_GET_IDIOMA() %>.js"></script>
  <script language="javascript">
        /* Variable globales de calendarios */
        var nombre = null;
        var cals = new Array();
        cals[0]= null;
        cals[1]= null;
        /* Funcion de Seleccion del Calendario */
        function SeleccionarCal(cal,date) {
            var obj = document.getElementById(nombre);
            var str= new String(date);
            obj.value = str;
            cal.hide();
        }
        /* Funcion de Cierre del Calendario */
        function CerrarCal(cal)
        {
          nombre = null;
		  cal.hide();
	    }
        /* Funcion para mostrar calendario */
        function MostrarCalendario(p_cal, p_name)
        {
	           var dte= new Date();
               var obj = document.getElementById(p_name);
               var strFecha = obj.value;
               var arrFecha = strFecha.split("/");
               dte.setFullYear(arrFecha[2],arrFecha[1]-1,arrFecha[0]);
               if (cals[p_cal] != null)
	           {
                    // Ya hay un calendario creado, se oculta.
                    cals[p_cal].hide();
	           }
	           else
	           {
                    cal = new Calendar(false, dte, SeleccionarCal,CerrarCal);
                    cal.weekNumbers = false;
                    cal.setRange(1993, 2033);
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
<form id="frmContenedor" name="frmContenedor" method="GET">
<table width="90%" cellspacing="0" cellpadding="0" align="center" border="0">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
           <td width="73%"><td>
           <td></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r2_c1.gif"></td>
           <td align=center valign="center"><font class="big" color="#517b4a"><% =GF_Traducir("Cupos")%></font></td>
           <td width="8"><img src="images/marco_r2_c3.gif"></td>
           <td></td>
           <td></td>
       </tr>
       <tr>
           <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
           <td></td>
           <td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="5"></td>
           <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r1_c3.gif"></td>
       </tr>
       <% if (errMsg <> "") then %>
       <tr>
            <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
            <td class="TDERROR" width="100%" align=center colspan="3"><% =GF_TRADUCIR(errMsg) %></td>
            <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
       </tr>
       <% end if%>
       <tr>
            <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
            <td width="100%" align=center colspan="3">&nbsp;</td>
            <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
       </tr>
       <tr>
           <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
           <td colspan="3">
            <table width="100%" height="100%" cellpadding="2" cellspacing="0" border="0">
                <tr>
                    <td width="10%">&nbsp;</td>
                    <td COLSPAN="2">
                        <font><% =GF_TRADUCIR("Fecha Desde")%></font><br>
                    </td>
                    <td width="10%">&nbsp;</td>
                </tr>
                <tr>
                    <td align="right"><img align="absMiddle" src="images/DATE.gif" alt="Seleccionar Fecha" style="cursor:hand" onclick="MostrarCalendario(0,'txtFechaDesde')"></td>
                    <td >
                        <input type="text" name="txtFechaDesde" value="<% =dteFecha %>">
                    </td>
                    <td ><input type=submit value="<%=GF_Traducir("Generar Archivo")%>"></td>
                    <td width="10%">&nbsp;</td>
                </tr>
            </table>
           </td>
           <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
       </tr>
       <tr>
            <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
            <td width="100%" align=center colspan="3">&nbsp;</td>
            <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
       </tr>
       <tr>
            <td height="100%"><img src="images/marco_r2_c1.gif" width="8" height="100%"></td>
            <td width="100%" align=center colspan="3">&nbsp;</td>
            <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
       </tr>
       <tr>
           <td width="8"><img src="images/marco_r3_c1.gif"></td>
           <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
           <td width="8"><img src="images/marco_r3_c3.gif"></td>
       </tr>
</table>
<input type="HIDDEN" name="P_ACTION" value="GENERATE">
</form>
</body>

</html>
<%
end if
if (p_oper = "GENERATE") then
    strPath = GenerarArchivo(GF_DTE2FN(dteFecha))
    Call Descargar(strPath)
end if
%>