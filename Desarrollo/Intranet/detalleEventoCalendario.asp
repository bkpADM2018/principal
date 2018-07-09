<!--#include file="includes/procedimientosMG.asp"-->
<!--#include file="includes/procedimientosfechas.asp"-->
<!--#include file="includes/procedimientossql.asp"-->
<!--#include file="includes/procedimientostraducir.asp"-->
<html>
<head>
	<title></title>
  	<script language="javascript">
        function modificarEvento(p_idEvento) {
			parent.location.href='pv_carga.asp?action=MODIFICAR&idEvento=' + p_idEvento;
		};
		
		function eliminarEvento(p_idEvento, p_anio) {
            parent.location.href='pv_carga.asp?action=BORRAR&idEvento=' + p_idEvento + '&txta=' + p_anio;
		}
  	</script>
</head>

<body style="background-color:white;" topmargin="0px" leftmargin="0px" rightmargin="0px">
<%dim o1kr, o2kr, o3kr, my_3okr

call GF_MGC("SR", "exec", o1kr, strDS)
call GF_MGC("UP", session("usuario"), o2kr, strDS)
call GF_MGC("SP", "PV_CARGAR", o3kr, strDS)
tienePermisosModificacion = GF_MGSR(o1kr, o2kr, o3kr, my_3okr)

p_dia = GF_Parametros7("dia", "", 6)
p_mes = GF_Parametros7("mes", "", 6)
p_anio = GF_Parametros7("anio", "", 6)

if p_anio & p_mes & p_dia <> "" then
    call GF_STANDARIZAR_FECHA(p_dia,p_mes,p_anio)
	fecha = p_anio & p_mes & p_dia%>
	<table cellpadding="0" cellspacing="0" width="100%">
	    <tr bgcolor="red" height="20px" valign="middle">
	        <td class="tituloDetalle" style="padding-left:10px;color:FFFFFF;font-weight:bold;font-size:14px;font-family:Verdana;"><%=GF_Traducir("Pagos y Vencimientos para el")%>&nbsp;<%=GF_FN2DTE(fecha)%></td>
	    </tr>
	    <tr>
	        <td style="font-size:12px;font-family:Verdana;padding-left:15px;padding-top:10px;">
		    <%strSQL = "select id, Descripcion from eventosPagosVencimientos where Fecha='" & fecha & "'"
			call GF_BD_Control(rs, conn, "OPEN", strSQL)
			while not rs.eof%>
			    <li>
				<%response.write rs("Descripcion")
				if tienePermisosModificacion then%>
					&nbsp;<span style="color:red;font-size:10px;cursor:pointer;" onClick="modificarEvento(<%=rs("id")%>);">[modificar]</span>&nbsp;<span style="color:red;font-size:10px;cursor:pointer;" onClick="eliminarEvento(<%=rs("id")%>, <%=p_anio%>);">[eliminar]</span>
				<%end if%>
                </li>
				<%rs.movenext
			wend%>
			</td>
		</tr>
	</table>
	<div style="position:absolute;left:370px;top:0;font-family:sans-serif;color:white;cursor:pointer;" title="Cerrar" onClick="window.parent.ocultarEvento();"><b>X</b></div>
<%end if%>
</body>
</html>
