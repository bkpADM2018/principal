<!--#include file="includes/procedimientosMG.asp"-->
<!--#include file="includes/procedimientosfechas.asp"-->
<!--#include file="includes/procedimientossql.asp"-->
<!--#include file="includes/procedimientostraducir.asp"-->
<!--#include file="includes/procedimientoscalendario.asp"-->
<%dim p_anio, vecMeses(11, 1)

p_anio = GF_Parametros7("p_anio", 0, 6)

vecMeses(0, 0) = GF_Traducir("Enero")
vecMeses(1, 0) = GF_Traducir("Febrero")
vecMeses(2, 0) = GF_Traducir("Marzo")
vecMeses(3, 0) = GF_Traducir("Abril")
vecMeses(4, 0) = GF_Traducir("Mayo")
vecMeses(5, 0) = GF_Traducir("Junio")
vecMeses(6, 0) = GF_Traducir("Julio")
vecMeses(7, 0) = GF_Traducir("Agosto")
vecMeses(8, 0) = GF_Traducir("Setiembre")
vecMeses(9, 0) = GF_Traducir("Octubre")
vecMeses(10, 0) = GF_Traducir("Noviembre")
vecMeses(11, 0) = GF_Traducir("Diciembre")

vecMeses(0, 1) = 31
if (p_anio mod 4 = 0) or (p_anio mod 400 = 0) then
	vecMeses(1, 1) = 28
else
    vecMeses(1, 1) = 28
end if
vecMeses(2, 1) = 31
vecMeses(3, 1) = 30
vecMeses(4, 1) = 31
vecMeses(5, 1) = 30
vecMeses(6, 1) = 31
vecMeses(7, 1) = 31
vecMeses(8, 1) = 30
vecMeses(9, 1) = 31
vecMeses(10, 1) = 30
vecMeses(11, 1) = 31

%>
<html>
<head>
  <title><%=GF_Traducir("Calendario de Pagos y Vencimientos")%></title>
  <link href="../../ActiSAIntra/CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
  <link href="../../ActiSAIntra/CSS/calendarioPagos.css" rel="stylesheet" type="text/css">
  <script language="javascript">

		function mostrarEvento(p_obj, p_anio, p_mes, p_dia) {
			 var x,y;
			 var vecCoords;
			 
			 document.getElementById('ifrmDetalleEvento').src = 'detalleEventoCalendario.asp?anio=' + p_anio + '&mes=' + p_mes + '&dia=' + p_dia;
             var divEvento = document.getElementById('divEvento');

			 vecCoords = findPos(p_obj);
			 divEvento.style.left = vecCoords[0] + 15 + 'px'; //15 es el width del span, pero no lo toma por HTML DOM
			 if ((vecCoords[1] - document.body.scrollTop) < 450) {
             	divEvento.style.top = vecCoords[1] + 'px';
			 } else {
                divEvento.style.top = 450 + document.body.scrollTop + 'px';
			 }
			 divEvento.className = 'evento visible';
   		}

        function ocultarEvento() {
			document.getElementById('divEvento').className = 'evento oculto';
		}

        function findPos(obj) {
			var curleft = curtop = 0;
			if (obj.offsetParent) {
				curleft = obj.offsetLeft
				curtop = obj.offsetTop
				while (obj = obj.offsetParent) {
					curleft += obj.offsetLeft
					curtop += obj.offsetTop
				}
			}
			return [curleft,curtop];
		}
		
		function agregarEvento(p_dia, p_mes, p_anio) {
			window.location.href='pv_carga.asp?txtd=' + p_dia + '&txtm=' + p_mes + '&txta=' + p_anio;
		}
  </script>
</head>

<body onClick="ocultarEvento();">
      <%=GF_Titulo_4(GF_Traducir("Calendario de Pagos y Vencimientos para el ") & p_anio)%>
	<%dim o1kr, o2kr, o3kr, my_3okr
	call GF_MGC("SR", "exec", o1kr, strDS)
	call GF_MGC("UP", session("usuario"), o2kr, strDS)
	call GF_MGC("SP", "PV_CARGAR", o3kr, strDS)
	tienePermisosModificacion = GF_MGSR(o1kr, o2kr, o3kr, my_3okr)
	if tienePermisosModificacion then%>
		<div align="right"><a href="pv_carga.asp"><%=GF_Traducir("Agregar eventos")%></a></div>
	<%end if%>
	<table cellspacing=3 cellpadding=3>
	    <tr valign="top">
	        <td><%call dibujarCuadroMes(p_anio, 0, CalcDayOfWeek(p_anio, 1, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 1, CalcDayOfWeek(p_Anio, 2, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 2, CalcDayOfWeek(p_anio, 3, 1))%></td>

	    </tr>
	    <tr valign="top">
	        <td><%call dibujarCuadroMes(p_anio, 3, CalcDayOfWeek(p_anio, 4, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 4, CalcDayOfWeek(p_anio, 5, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 5, CalcDayOfWeek(p_anio, 6, 1))%></td>
	    </tr>
	    <tr valign="top">
	        <td><%call dibujarCuadroMes(p_anio, 6, CalcDayOfWeek(p_anio, 7, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 7, CalcDayOfWeek(p_anio, 8, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 8, CalcDayOfWeek(p_anio, 9, 1))%></td>
	    </tr>
	    <tr valign="top">
	        <td><%call dibujarCuadroMes(p_anio, 9, CalcDayOfWeek(p_anio, 10, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 10, CalcDayOfWeek(p_anio, 11, 1))%></td>
	        <td><%call dibujarCuadroMes(p_anio, 11, CalcDayOfWeek(p_anio, 12, 1))%></td>
	    </tr>
	</table>
	<div id="divEvento" class="evento oculto">
    	<iframe id="ifrmDetalleEvento" src="detalleEventoCalendario.asp?" frameborder=0 scrolling="auto" style="border-width:1px;border-style:solid;width:400px;height:150px;"></iframe>
	</div>
</body>
</html>
<%'*************************************************************************************
sub dibujarCuadroMes(byval p_anio, byval p_mes, byval p_comienzoMes)
	dim auxMes
	
	auxMes = p_mes + 1
	call GF_STANDARIZAR_FECHA("02",auxMes,p_anio)
	strSQL = "select distinct right(Fecha, 2) as Dia from eventosPagosVencimientos where Fecha like '" & p_anio & auxMes & "__' ORDER BY Dia ASC"
	call GF_BD_Control(rs, conn, "OPEN", strSQL)

'	response.write strSQL%>
	<table class="tablaMes" cellspacing=0 cellpadding=3>
	    <tr class="encabezadoMes">
	        <td colspan="7" align="center"><%=GF_Traducir(vecMeses(p_mes, 0))%></td>
	    </tr>
	    <tr class="encabezadoDias">
	        <td><%=left(GF_Traducir("Domingo"), 2)%></td>
	        <td><%=left(GF_Traducir("Lunes"), 2)%></td>
	        <td><%=left(GF_Traducir("Martes"), 2)%></td>
	        <td><%=left(GF_Traducir("Miercoles"), 2)%></td>
	        <td><%=left(GF_Traducir("Jueves"), 2)%></td>
	        <td><%=left(GF_Traducir("Viernes"), 2)%></td>
	        <td><%=left(GF_Traducir("Sabado"), 2)%></td>
	    </tr>
		<%contDia = 0
		diaSemana = 0
		acum = 0

  		while (acum < 42)

			if (diaSemana = 0) then
				response.write "<tr><td class='diaDomingo'"
			else
			    response.write "<td"
			end if

			if ((acum >= p_comienzoMes) and (contDia < vecMeses(p_mes, 1))) then
                contDia = contDia + 1
                if not rs.eof then
                    if (cint(rs("Dia")) < cint(contDia)) then
						if tienePermisosModificacion then response.write " ondblClick='agregarEvento(" & contDia & ", " & auxMes & ", " & p_anio & ");'"
						response.write ">" & contDia
                        rs.movenext
					elseif (cint(rs("Dia")) = cint(contDia)) then%>
						 ><span style="cursor:default;" onMouseOver="mostrarEvento(this, <%=p_anio%>, <%=auxMes%>, <%=contDia%>);"><%=contDia%></span>
						<%rs.movenext
					else 'Si el dia es mayor al proximo evento, avanzo el dia
                        if tienePermisosModificacion then response.write " ondblClick='agregarEvento(" & contDia & ", " & auxMes & ", " & p_anio & ");'"
						response.write ">" & contDia
					end if
				else
                    if tienePermisosModificacion then response.write " ondblClick='agregarEvento(" & contDia & ", " & auxMes & ", " & p_anio & ");'"
					response.write ">" & contDia
				end if
			else
			    response.write ">&nbsp;"
			end if
			acum = acum + 1
			diaSemana = diaSemana + 1
			response.write "</td>"

			if (diaSemana = 7) then
				response.write "</tr>"
				diaSemana = 0
			end if
  		wend
		%>
	</table>
<%
end sub
'*************************************************************************************
function prepararTexto(p_texto)
	dim ret
	
	ret = replace(p_texto, "'", "\'")
	ret = replace(ret, "\", "\\")
	prepararTexto = ret
end function
%>
