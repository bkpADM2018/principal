<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<html>
<head>
<title>Intranet ActiSA</title>
<link rel="stylesheet" href="css/main.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%
'ProcedimientoControl "INTCUMPLE"
Dim intDia,intPaso,vector(100),indx
Dim strDS,intKR, intMes
Dim strSQL,rsDatos,bolSalir,strMes
Dim i,k,strKC, oConn, krToepfer, dsToepfer
Dim strKMSector, strKCSector,strDSSector,intKRSector

'Se recupera el parametro
intMes=CInt(Request.QueryString("P_MES"))
if (intMes="") then intMes=1
'Se setea el ultimo dia del mes.
select case intMes
   case 1: intDia=31
           strMes=GF_TRADUCIR("ENERO") 
   case 2: intDia=29
           strMes=GF_TRADUCIR("FEBRERO") 
   case 3: intDia=31
           strMes=GF_TRADUCIR("MARZO") 
   case 4: intDia=30
           strMes=GF_TRADUCIR("ABRIL") 
   case 5: intDia=31
           strMes=GF_TRADUCIR("MAYO") 
   case 6: intDia=30
           strMes=GF_TRADUCIR("JUNIO") 
   case 7: intDia=31
           strMes=GF_TRADUCIR("JULIO") 
   case 8: intDia=31
           strMes=GF_TRADUCIR("AGOSTO") 
   case 9: intDia=30
           strMes=GF_TRADUCIR("SEPTIEMBRE") 
   case 10: intDia=31
           strMes=GF_TRADUCIR("OCTUBRE") 
   case 11: intDia=30
           strMes=GF_TRADUCIR("NOVIEMBRE") 
   case 12: intDia=31
           strMes=GF_TRADUCIR("DICIEMBRE")                        
end select
'Estandarizo el mes a dos digitos
if (Len(intMes) < 2) then intMes= "0" & intMes
Call GF_MGC("OR", "07431", krToepfer, dsToepfer)
strSQL = "Select idProfesional, (Apellido + ', ' + Nombre) as Nombre, substring(FechaNacimiento,7,2) as dia, mg_kc as sectorKc, mg_ds as sectorDs from ((Personas inner join Profesionales on idPersona=idProfesional) inner join mg on sector=mg_kr) where egresoValido='F' and FechaNacimiento like '____" & intmes & "%' and Empresa=" & krToepfer & " order by dia asc, Nombre asc, Apellido asc"
Call executeQueryDb(DBSITE_SQL_INTRA, rsEmpleadosArgentina,"OPEN",strSQL)
%>
<table border=0 width="80%" align="center">
<tr>
	<td align="center" colspan="4"><img src="images/feliz_cumple.jpg"></td>
</tr>
<% intPaso=false
indx=0		
%>
<tr><td><div align="center" class="Month" colspan="4"><% =strMes %></div></td></tr>
<tr><td valign=top width=70%>
<table border="0" width="100%">
<%if not rsEmpleadosArgentina.eof then
	diaAnterior = ""
	while not rsEmpleadosArgentina.eof%>
	   <tr>
		   <td width="10%" align=right class="Birthday">
		   		<%if (diaAnterior = rsEmpleadosArgentina("dia")) then
				   response.write "&nbsp;"
		   		else
	   				response.write rsEmpleadosArgentina("dia")
				    diaAnterior = rsEmpleadosArgentina("dia")
				end if%>
		   </td>
           <td width=5% align=center>-</td>
		   <td class="Birthday"><%=rsEmpleadosArgentina("Nombre")%><td>
		   <td class="Birthday">
		   		<%select case ucase(rsEmpleadosArgentina("sectorKc"))
		   		    case "01" 'Arroyo seco
		   		        strDSSector = "(Arroyo Seco)"
                    case "12" 'Rosario
					    strDSSector = "(Rosario)"
					case "16" 'Transito
					    strDSSector = "(Transito)"
					case "19" 'Piedra Buena
					    strDSSector = "(Piedrabuena)"
					case else
					    strDSSector = "&nbsp;"
		   		end select
       			response.write strDSSector%>
		   </td>
		   <%rsEmpleadosArgentina.MoveNext
	wend
else%>
      <tr>
	    <td align=center class="BirthDay" colspan="4">
		   <% =GF_TRADUCIR("NO") %><br><br>
		   <% =GF_TRADUCIR("HAY CUMPLEAÑOS") %><br><br>
		   <% =GF_TRADUCIR("ESTE MES") %>
		</td>
	  </tr>	  
<% end if %>

</table>
</td>
</tr>
</table>
</body>
</html>
