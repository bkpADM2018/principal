<%
'                             **************************************
'                             *    PROCEDIMIENTOS DE PAGINACION    *
'                             **************************************
'                             **     Autor: Javier A. Scalisi     **
'                             **     Fecha: 19/03/2003            **
'                             ************************************** 
'________________________________________________________________________________________________________
function GF_LINK_LETRAS(P_strPaginaActual,ByRef P_strLinkPagina)
'Esta Funcion genera los links para cada pagina mostrando letras.

Dim i, strLinkPagina, intAux, strLink

'Averiguo si el Link ya posee parametros.
intAux= inStr(P_strLinkPagina,"?")
if (intAux > 0) then
   P_strLinkPagina=P_strLinkPagina & "&P_PAGINA_ACTUAL=" 
else
   P_strLinkPagina=P_strLinkPagina & "?P_PAGINA_ACTUAL=" 
end if	 
response.write "<table>"
response.write "<tr>"
response.write "<td>" & GF_TRADUCIR("Paginas") & ": </td>"
response.write "<td width='100%'>"
     For i= asc("A") to asc("Z")
          if (chr(i) <> P_strPaginaActual) then
  	         strLinkPagina= P_strLinkPagina & chr(i) 
 		     response.write "<a href=" & strLinkPagina & " style=text-decoration:none;> " & chr(i) & " </a>"
          else
     	     'Si el numero coincide con la pagina mostrada no posee ningun Link.
             strLink = P_strLinkPagina & chr(i) 
			 response.write "<B><font color='RED'>[" & chr(i) & "]</font></B>"
			end if
		next
response.write "</td>"  
response.write "</tr>"
response.write "</table>"
P_strLinkPagina=strLink
end function
'________________________________________________________________________________________________________
function GF_LINK_NUMEROS(P_strPaginaActual,P_strLinkPagina, P_TotalPaginas,P_intPaginaINI,P_intPaginaFIN,P_intSectorINI)
'Esta Funcion genera los links para cada pagina mostrando Numeros.

Dim i, strLinkPagina, intAux, strLink
Dim intIntervalo,intPaginaFinIntervalo
Dim intCount

'Averiguo si el Link ya posee parametros.
intAux= inStr(P_strLinkPagina,"?")
if (intAux > 0) then
   P_strLinkPagina=P_strLinkPagina & "&P_PAGINA_ACTUAL=" 
else
   P_strLinkPagina=P_strLinkPagina & "?P_PAGINA_ACTUAL=" 
end if
intIntervalo= 24
response.write "<table id=tblPaginacion>"
response.write "<tr>"
response.write "<td>" & GF_TRADUCIR("Paginas") & ":</td>"
response.write "<td width='100%'>"
         'Se crea el link para cambiar a un sector anterior solo si no se esta en el primero.
	 if (CInt(P_intSectorINI) <> 1) then
	    strLinkPagina= P_strLinkPagina & "S" & P_intSectorINI-125
	    response.write "<a href=" & strLinkPagina & " style=text-decoration:none;><font color=Navy>&lt;&lt; prev - </font> </a>"	    
	 end if
	 'Se crean los intervalos de agrupacion de paginas del sector visualizado
	 ' anteriores al intervalo expandido         
	 if (CInt(P_intPaginaINI) <> CInt(P_intSectorINI)) then
	    i=CInt(P_intSectorINI)
	    intCount=0
	    while (CInt(P_intPaginaINI) > i)
               strLinkPagina= P_strLinkPagina & i & "-" & i + intIntervalo
	       response.write "<a href=" & strLinkPagina & " style=text-decoration:none;><font color=Navy>[" & i & " - " & i + intIntervalo & "]</font> </a>"	    
               i=i+intIntervalo+1
	       intCount= intCount+1
            wend		  
         end if
	 'Se dibuja el intevalo exdpandido.
	 if (CInt(P_intPaginaFIN) < CInt(P_TotalPaginas)) then
	    intPaginaFinIntervalo=P_intPaginaFIN
         else
	    intPaginaFinIntervalo= P_TotalPaginas
         end if			
	 For i= P_intPaginaINI to intPaginaFinIntervalo
          if (i <> CInt(P_strPaginaActual)) then
  	      strLinkPagina= P_strLinkPagina & i 
 	      response.write "<a href=" & strLinkPagina & " style=text-decoration:none;> " & i & " </a>-"
          else
     	     'Si el numero coincide con la pagina mostrada no posee ningun Link.
	     strLink = P_strLinkPagina & i
             response.write "&nbsp;<B style='background-color:#FF0000;color=#FFFFFF;'>[" & i & "]</B> -"
  	  end if
	 next 
	 'Sumo uno para posicionarme sobre el inicio del nuevo intervalo.
	 i=intPaginaFinIntervalo+1
	 while (P_TotalPaginas >= i) and (intCount < 4)
           strLinkPagina= P_strLinkPagina & i & "-" & i + intIntervalo
	   if (P_TotalPaginas < (i + intIntervalo)) then
	      intPaginaFinIntervalo= P_TotalPaginas
	   else   
	      intPaginaFinIntervalo= i + intIntervalo
	   end if
	   response.write "<a href=" & strLinkPagina & " style=text-decoration:none;><font color=Navy>[" & i & " - " & intPaginaFinIntervalo & "]</font></a>"	    
           i=i+intIntervalo+1
	   intCount = intCount + 1
        wend		  	
	if (P_TotalPaginas >= i) then
	    strLinkPagina= P_strLinkPagina & "S" & i
	    response.write "<a href=" & strLinkPagina & " style=text-decoration:none;><font color=Navy> - next &gt;&gt;</font> </a>"	    
	end if	
response.write "</td>"  
response.write "</tr>"
response.write "</table>"
end function
'________________________________________________________________________________________________________
function GF_LINK_PAGINAS(P_strTipo,P_PaginaActual,P_intTotalPaginas,P_strLinkPagina,P_intPaginaINI,P_intPaginaFIN,P_intSectorINI)

Select case UCASE(P_strTipo)
       case "L","LETRAS","LETRA": Call GF_LINK_LETRAS(P_PaginaActual,P_strLinkPagina)
       case "N","NUMEROS","NUMERO": Call GF_LINK_NUMEROS(P_PaginaActual,P_strLinkPagina,P_intTotalPaginas,P_intPaginaINI,P_intPaginaFIN,P_intSectorINI)
end select

end function
'________________________________________________________________________________________________________
function GF_CONTROL_PAGINACION(P_strTipo,ByRef rs,ByRef P_intPaginaActual,P_intTotalPaginas,P_intCantidadPorPagina,ByRef P_intPaginaINI,ByRef P_intPaginaFIN, ByRef P_intSectorINI)

Dim intAUX

'Obtengo el parametro
P_intPaginaActual= Request.QueryString("P_PAGINA_ACTUAL")
if (P_intPaginaActual = "") then P_intPaginaActual= Request.Form("P_PAGINA_ACTUAL")

if (UCASE(P_strTipo)="N") or (UCASE(P_strTipo)="NUMEROS")  or (UCASE(P_strTipo)="NUMERO") then 

'Determinacion de los registros por pagina.
rs.PageSize= P_intCantidadPorPagina
rs.CacheSize = P_intCantidadPorPagina
'Determinacion de la cantidad de paginas del listado.
P_intTotalPaginas= rs.PageCount

if (P_intPaginaActual = "") then 
   P_intSectorINI=1
   P_intPaginaActual=1
   P_intPaginaINI=1
   P_intPaginaFIN=25
else
   'Averiguo si es un sector 
   intAUX= inStr(P_intPaginaActual,"S")
   if (CInt(intAUX) > 0) then		
      'Es un sector
      P_intSectorINI= right(P_intPaginaActual,len(P_intPaginaActual)-1) 	
      P_intPaginaINI= P_intSectorINI
      P_intPaginaFIN= P_intSectorINI + 24
      P_intPaginaActual= P_intPaginaINI 	  	
   else
      P_intSectorINI=session("P_intSectorINI")
      'Averiguo si es un intervalo
      intAUX= inStr(P_intPaginaActual,"-")
      if (CInt(intAUX) > 0) then
         'Es un Intervalo
	  P_intPaginaINI= left(P_intPaginaActual,intAUX-1)   
	  P_intPaginaFIN= right(P_intPaginaActual,len(P_intPaginaActual)-intAUX)
	  P_intPaginaActual= P_intPaginaINI 	  
      else
         'Se tiene un numero de pagina valida se asignan los limites del ultimo intervalo abierto
	  P_intPaginaINI= session("P_intPaginaINI")
	  P_intPaginaFIN= session("P_intPaginaFIN")
      end if
   end if
end if

'Posicionamiento en la pagina solicitada
if (P_intTotalPaginas > 0) then
   'Posicionamiento en la pagina solicitada
   if (CInt(P_intPaginaActual) > P_intTotalPaginas) then
      'Si cambio la pagina actual debo cambiar los limites del intervalo a mostrar.
	  P_intPaginaActual=P_intTotalPaginas
      while (CInt(P_intPaginaINI) > CInt( P_intPaginaActual))
	      P_intSectorINI= P_intSectorINI - 25
	      P_intPaginaINI= P_intPaginaINI - 25
   	      P_intPaginaFIN= P_intPaginaFIN - 25
      wend		
	  rs.AbsolutePage = P_intTotalPaginas
   else
      rs.AbsolutePage = P_intPaginaActual
   end if
else 
   'Se setea como minimo una pagina para mostrar mensaje.
   P_intTotalPaginas = 1
end if
else
   'Se eligio LETRAS
   if (P_intPaginaActual = "") then P_intPaginaActual="A"
end if
'Salvo los limites del rango actual.
session("P_intPaginaINI")=P_intPaginaINI  
session("P_intPaginaFIN")=P_intPaginaFIN
'Salvo el sector actual
session("P_intSectorINI")=P_intSectorINI

response.write "<input type=hidden name=P_PAGINA_ACTUAL value='" & P_intPaginaActual & "'>"
end function
'________________________________________________________________________________________________________
function GF_REGISTROS_MOSTRADOS(ByRef P_intCantidad,ByRef P_strLinkPagina,ByRef P_intMaximo)
' Esta funcion genera los links para mostrar registros en pantalla.

Dim i,intPaso, intAux, strLink

'Obtengo el parametro
P_intCantidad= Request.QueryString("P_REGISTROS_A_MOSTRAR")
if (P_intCantidad = "") then P_intCantidad= Request.Form("P_REGISTROS_A_MOSTRAR")
if (P_intCantidad = "") then P_intCantidad=10

if (P_intMaximo = "") then 
   intPaso=10
   P_intMaximo=50
else
   if (P_intMaximo <= 100) then intPaso=10
   if (P_intMaximo > 100) and (P_intMaximo <= 500) then intPaso=50
   if (P_intMaximo > 500) then intPaso=100
end if
response.write "<table border=0 width='90%'><tr align=Right><td>" & GF_TRADUCIR("Paginar de a")  &  "&nbsp;"
'Averiguo si el Link ya posee parametros.
intAux= inStr(P_strLinkPagina,"?")
if (intAux > 0) then
   P_strLinkPagina=P_strLinkPagina & "&P_REGISTROS_A_MOSTRAR=" 
else
   P_strLinkPagina=P_strLinkPagina & "?P_REGISTROS_A_MOSTRAR=" 
end if	 
i=intPaso
while (i <= Cint(P_intMaximo))
   if (i = CInt(P_intCantidad)) then
      strLink= P_strLinkPagina & i
	  response.write "<B><font color='RED'>[" & i & "]</font> - </B>"
   else   
      response.write "<a href=" & P_strLinkPagina & i & " style=text-decoration:none;>" & i & "</a> - "   
   end if
   i=i+intPaso	  
wend
response.write "</td></tr></table>"
response.write "<input type=hidden name=P_REGISTROS_A_MOSTRAR value='" & P_intCantidad & "'>"
P_strLinkPagina=strLink
end function
'________________________________________________________________________________________________________
function GF_PAGINAR(P_strTipo,P_strLinkPagina,ByRef P_intMostrar,P_intMaximo,ByRef P_RS)
' OJO!!!!
' Esta funcion debe llamarse antes de recorrer el recordset para mostrar los datos dado que
' para posicionarse en una pagina se utiliza una propiedad del recodset.
'P_strTipo = Tipo de Links numericos o de Letras.
'P_RS = recordset del cual se obtienen los datos a paginar.
'P_intMostrar= Indica cuantos registros se muestran en una pagina, se devuelve para poder 
'              frenar cuando se recorre el recordset en la pagina.  
'P_intMaximo = Maximo numero de registros por pagina que se pueden ver.
'P_strLinkPagina = El Link.
'P_intCantidadPaginas= Indica el numero de links de paginas maximo a mostrar de una vez.

Dim intTotalPaginas, intPaginaActual, intPaginaINI, intPaginaFIN, intSectorINI

   P_strTipo=UCASE(P_strTipo)
   if (P_strTipo = "N") or (UCASE(P_strTipo)="NUMEROS")  or (UCASE(P_strTipo)="NUMERO") then 
      Call GF_REGISTROS_MOSTRADOS(P_intMostrar,P_strLinkPagina,P_intMaximo)
   end if	  	  
   Call GF_CONTROL_PAGINACION(P_strTipo,P_RS,intPaginaActual,intTotalPaginas,P_intMostrar,intPaginaINI,intPaginaFIN,intSectorINI)
   Call GF_LINK_PAGINAS(P_strTipo,intPaginaActual,intTotalPaginas,P_strLinkPagina,intPaginaINI,intPaginaFIN,intSectorINI)
   if (P_strTipo = "L") or (UCASE(P_strTipo)="LETRAS")  or (UCASE(P_strTipo)="LETRA") then P_intMostrar=intPaginaActual
end function
%>
