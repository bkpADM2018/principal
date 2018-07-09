<%
'/**
' * Funcion    : GF_EDIT_FECHA
' * Descripcion: Esta funcion da formato a una fecha
' * Parametros : pFecha
' * Valor Devuelto:
' * dd/mm/aaaa
' * Autor: Ezequiel A. Bacarini
' * Fecha: 20/01/2011
' */
Function GF_EDIT_FECHA(pFecha)
	Dim aux
	if isDate(pFecha) then
		aux = day(pFecha) & "/" & month(pFecha) & "/" & year(pFecha)
	end if	
	GF_EDIT_FECHA = aux
End Function
'/**
' * Funcion: GF_GET_DELIMITADOR
' * Descripcion: Busca el Delimitador de una fecha y lo devuelve
' * Parametros: p_dteFecha  [in] Fecha en formato DD/MM/AAAA HH:MM:SS
' * Autor: Javier A. Scalisi
' * Fecha: 15/01/2003
' */
function GF_GET_DELIMITADOR(p_dteFecha)

Dim intDelimitador, ret

         ret = ""
         intDelimitador= InStr(1,p_dteFecha,"/",1)
	 if (intDelimitador = 0) then
	   'Si no encontro la barra el que se utiliza es el guion.
	   ret= "-"
	 else
	   ret= "/"
	 end if
	 GF_GET_DELIMITADOR  = ret
	 
end function

'Autor: Javier A. Scalisi
'Fecha: 15/01/2003
function GF_CONTROL_FECHA(byref P_intDia, byref P_intMes, byref P_intAnio)
' Esta funcion se encarga de controlar las fechas y de dejarlas en el formato DD MM AAAA 
' TRUE = La fecha es correcta
' FALSE= La fecha es incorrecta 
   dim intDelimitador,intDiasFebrero,intDiasMes
   GF_CONTROL_FECHA=false
   if (isNumeric(P_intDia) and isNumeric(P_intMes) and isNumeric(P_intAnio)) then	
        '1º - Se le da formato al año (YYYY), mes (MM) y dia (DD).
        GF_STANDARIZAR_FECHA P_intDia,P_intMes,P_intAnio
		'2º - Determinacion de los dias del mes de febrero
		if ((P_intAnio mod 4) = 0) then
		   'El año es biciesto.
		   intDiasFebrero= 29
		else
		   'El año no es biciesto.
		   intDiasFebrero= 28
		end if
		'3º - Determinacion de los dias del mes.
		intDiasMes = 0
		select case P_intMes
		   case "01","03","05","07","08","10","12": intDiasMes = 31
		   case "02": intDiasMes = intDiasFebrero
		   case "04","06","09","11": intDiasMes = 30
		end select
		'4º - Se controla la fecha
		GF_CONTROL_FECHA=true
		if (intDiasMes = 0) then GF_CONTROL_FECHA=false		
		if (CInt(P_intDia) < 1) or (CInt(P_intDia) > intDiasMes) then GF_CONTROL_FECHA=false 
   end if

end function
'------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 15/01/2003
function GF_CONTROL_FECHA_2(byref P_dteFecha)
' Esta funcion se encarga de controlar las fechas y de dejarlas en el formato DD/MM/AAAA 
' TRUE = La fecha es correcta
' FALSE= La fecha es incorrecta 

    dim arrFecha,arrFechaHora,arrHora
	dim intDelimitador
	dim strDelimitador
	dim bolTieneHora 	
    GF_CONTROL_FECHA_2=false
	'Se borran los posibles espacios en blanco tanto a derecha como a izquierda.
	P_dteFecha= Trim(P_dteFecha) 
	'Si es posible que sea una fecha se controla.
    if (isDate(P_dteFecha)) then
		'1º - Separo la hora de la fecha.
		'Posicion 0= Fecha. Posicion 1= Hora.
		bolTieneHora= false
		if (len(P_dteFecha) > 10) then bolTieneHora= true
		arrFechaHora= Split(P_dteFecha," ",3,1)   		
		'2º - Se busca el delimitador utilizado en la fecha.
	        strDelimitador= GF_GET_DELIMITADOR(arrFechaHora(0))
		'3º - Se separa la fecha en dia, mes y año.
		arrFecha= Split(arrFechaHora(0),strDelimitador,3,1)
		'4º - Se controla la fecha
		GF_CONTROL_FECHA_2 = GF_CONTROL_FECHA(arrFecha(0),arrFecha(1),arrFecha(2))   
		'5º - Rearmo la fecha		
        P_dteFecha= arrFecha(0) & strDelimitador & arrFecha(1) & strDelimitador & arrFecha(2) 
        if (bolTieneHora) then	
			arrHora= Split(arrFechaHora(1),":",3,1)
			CALL GF_STANDARIZAR_MM(arrHora(0),arrHora(1),arrHora(2))
			P_dteFecha= P_dteFecha & " " & arrHora(0) & ":" & arrHora(1) & ":" & arrHora(2)
			GF_CONTROL_FECHA_2 = GF_ControlHora(arrHora(0),arrHora(1),arrHora(2))
    	end if
	end if		
end function
'------------------------------------------------------------------------------------------
'Autor: Ezequiel Bacarini
'Fecha: 22/02/2011
function GF_STANDARIZAR_FECHA_RTRN(byVal pDate)
'Esta funcion le da formato a la fecha DD MM AAAA.
dim auxDate, intYear, intMonth, intDay
GF_STANDARIZAR_FECHA_RTRN = "ERROR"
auxDate = CDate(pDate)

intYear = Year(auxDate)
intMonth = Month(auxDate)
intDay = Day(auxDate)

if (len(intYear) < 4) then
   if (intYear < 30) then 
      intYear= intYear + 2000
   else 
      intYear= intYear + 1900
   end if
end if
if (len(intMonth) = 1) then intMonth= "0" & intMonth
if (len(intDay) = 1) then intDay= "0" & intDay
GF_STANDARIZAR_FECHA_RTRN = intDay & "/" & intMonth & "/" & intYear 
end function
'------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 20/02/2003
function GF_STANDARIZAR_FECHA(ByRef P_intDia,ByRef P_intMes,ByRef P_intAnio)
'Esta funcion le da formato a la fecha DD MM AAAA.
 
if (len(P_intAnio) < 4) then
   if (P_intAnio < 30) then 
      P_intAnio= P_intAnio + 2000
   else 
      P_intAnio= P_intAnio + 1900
   end if
end if
if (len(P_intMes) = 1) then P_intMes= "0" & P_intMes
if (len(P_intDia) = 1) then P_intDia= "0" & P_intDia

end function
'------------------------------------------------------------------------------------------		   
'Autor: Javier A. Scalisi
'Fecha: 06/08/2003
function GF_STANDARIZAR_MM(ByRef P_intHora,ByRef P_intMin,ByRef P_intSeg)
'Esta funcion le da formato a la hora HH MM SS.
 
if (len(P_intHora) = 1) then P_intHora= "0" & P_intHora
if (len(P_intMin) = 1) then P_intMin= "0" & P_intMin
if (len(P_intSeg) = 1) then P_intSeg= "0" & P_intSeg

end function
'/**
' * Funcion: GF_FN2DTCONTABLE
' * Descripcion: Convierte una fecha del formato string numerico
' *              a un string de fecha para la base de datos de los puertos.
' * Parametros:  P_intFecha [in] String Numerico a transformar.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 06/08/2015
' */
Function GF_FN2DTCONTABLE(P_intFecha)
    GF_FN2DTCONTABLE = "0000-00-00"
    if (Len(Cstr(P_intFecha)) >= 8) then
        GF_FN2DTCONTABLE = left(P_intFecha,4) & "-" & Mid(P_intFecha, 5, 2) & "-" & Mid(P_intFecha, 7, 2)
    end if
            
End Function
'/**
' * Funcion: GF_DT2DTCONTABLE
' * Descripcion: Convierte una fecha de cualquier formato al formato YYYY-MM-DD
' * Parametros:  p_DtFecha [in] Fecha a transformar.
' */
Function GF_DT2DTCONTABLE(p_DtFecha)
dim pDia, pMes, pAnio
    GF_DT2DTCONTABLE = "0000-00-00"
    if isDate(P_dtFecha) then
		pDia = day(p_DtFecha)
		pMes = month(p_DtFecha)
		pAnio = year(p_DtFecha)
		call GF_STANDARIZAR_FECHA(pDia, pMes, pAnio)
        GF_DT2DTCONTABLE = pAnio & "-" & pMes & "-" & pDia
    end if
End Function

'/**
' * Funcion: GF_FN2DTE
' * Descripcion: Convierte una fecha del formato string numerico
' *              a un string de fecha.
' * Parametros:  P_intFecha [in] String Numerico a transformar.
' * Valor Devuelto
' * Devuelve un string con formato de fecha (XX/XX/XXXX HH:MM:SS).
' *
' * Autor: Javier A. Scalisi
' * Fecha: 12/04/2004
' */
function GF_FN2DTE(P_intFecha)
'El formato de la fecha supuesto es AAAAMMDD o AAAAMMDDHHMMSS
dim intAnio,intMes,intDia
dim intHora, intMin, intSeg
dim intLongitud,rtn
rtn=P_intFecha
intLongitud= Len(P_intFecha)
if (intLongitud >= 8) then
	intAnio= left(P_intFecha,4)
	intMes = mid(P_intFecha,5,2)
	intDia = mid(P_intFecha,7,2)
	if (GF_CONTROL_FECHA(intDia,intMes,intAnio)) then
	   rtn= intDia & "/" & intMes & "/" & intAnio
	   if (intLongitud > 8) then
	      intHora = mid(P_intFecha,9,2)
	      intMin = mid(P_intFecha,11,2)
	      intSeg = Right(P_intFecha,2)
	      if (GF_ControlHora(intHora,intMin,intSeg)) then
	         rtn= rtn & " " & intHora & ":" & intMin & ":" & intSeg
	      else
	         rtn= "-"
	      end if
	   end if
	else
	    rtn="-"
	end if
end if
GF_FN2DTE= rtn
end function
'/**
' * Funcion: GF_DTE2FN
' * Descripcion: Convierte un string de formato fecha
' *              a un string numerico.
' * Parametros:  P_intFecha [in] String Numerico a transformar.
' * Valor Devuelto
' * Devuelve un string con formato numerico (AAAAMMDDHHMSS).
' *
' * Autor: Javier A. Scalisi
' * Fecha: 12/04/2004
' */
function GF_DTE2FN(p_dteFecha)

Dim arrFechaHora,arrFecha,arrHora
dim intLongitud,rtn

    rtn=p_dteFecha
    intLongitud= Len(p_dteFecha)
    if (intLongitud > 0) then
             if (GF_CONTROL_FECHA_2(p_dteFecha)) then
               arrFechaHora= Split(P_dteFecha," ",3,1)
     	       'Se busca el delimitador utilizado en la fecha.
               strDelimitador= GF_GET_DELIMITADOR(arrFechaHora(0))
               'Se separa la fecha en dia, mes y año.
               arrFecha= Split(arrFechaHora(0),strDelimitador,3,1)
               call GF_STANDARIZAR_FECHA(arrFecha(0),arrFecha(1),arrFecha(2))
               rtn = arrFecha(2) & arrFecha(1) & arrFecha(0)
               if (UBound(arrFechaHora) = 1) then
                  arrHora= Split(arrFechaHora(1),":",3,1)
                  call GF_STANDARIZAR_MM(arrHora(0),arrHora(1),arrHora(2))
                  rtn = rtn & arrHora(0) & arrHora(1) & arrHora(2)
               end if
             else
                 rtn="#ERROR FECHA#"
             end if
    end if

GF_DTE2FN = rtn
End Function
'------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 24/02/2003
function GF_CONTROL_PERIODO(ByRef P_intDiaDesde,ByRef P_intDiaHasta,ByRef P_intMesDesde,ByRef P_intMesHasta,ByRef P_intAnioDesde,ByRef P_intAnioHasta)
'Esta funcion se encaraga de comprobar que el periodo de fehcas pasado como parametro sea valido.
'0:OK  1:Fecha Desde Incorrecta   2:Fecha Hasta Incorrecta    3:Periodo No Valido
Dim R,blnFecha

R=0
'1.- Controlo las fechas.
blnFecha=GF_CONTROL_FECHA(P_intDiaDesde,P_intMesDesde,P_intAnioDesde)
if not(blnFecha) then R=1
blnFecha=GF_CONTROL_FECHA(P_intDiaHasta,P_intMesHasta,P_intAnioHasta)
if not(blnFecha) then R=2
'2.- Controlo el Periodo si no hubo errores.
if (R=0) then
   if (P_intAnioDesde > P_intAnioHasta) then R=3
   if (P_intAnioDesde = P_intAnioHasta) and (P_intMesDesde > P_intMesHasta) then R=3
   if (P_intAnioDesde = P_intAnioHasta) and (P_intMesDesde = P_intMesHasta) and (Cint(P_intDiaDesde) > Cint(P_intDiaHasta)) then R=3
end if
GF_CONTROL_PERIODO=R   
end function
'------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 24/02/2003
function GF_CONTROL_PERIODO_2(ByRef P_dteFechaDesde,ByRef P_dteFechaHasta)
'Esta funcion se encaraga de comprobar que el periodo de fehcas pasado como parametro sea valido.
'Se supone que los delimitadores usados en ambas fechas son iguales.(/,-)
'0:OK  1:Fecha Desde Incorrecta   2:Fecha Hasta Incorrecta    3:Periodo No Valido
Dim strDelimitador,arrFechaDesde,arrfechaHasta,intDelimitador

        '1º - Se busca el delimitador utilizado en la fecha.
	    strDelimitador= GF_GET_DELIMITADOR(P_dteFechaDesde)
		'2º - Se separa la fecha en dia, mes y año.		
		arrFechaDesde= Split(P_dteFechaDesde,strDelimitador,3,1)
		arrFechaHasta= Split(P_dteFechaHasta,strDelimitador,3,1)		
		'3º - Se controla el periodo		   
		GF_CONTROL_PERIODO_2 = GF_CONTROL_PERIODO(arrFechaDesde(0),arrFechaHasta(0),arrFechaDesde(1),arrFechaHasta(1),arrFechaDesde(2),arrFechaHasta(2))   
		'4º - Rearmo las Fechas		
        P_dteFechaDesde= arrFechaDesde(0) & strDelimitador & arrFechaDesde(1) & strDelimitador & arrFechaDesde(2) 
        P_dteFechaHasta= arrFechaHasta(0) & strDelimitador & arrFechaHasta(1) & strDelimitador & arrFechaHasta(2) 
		
end function
'------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 28/02/2003
'Change Log: Javier A. Scalisi - 02/06/2003
function GF_MDA2DMA(P_dteFecha)
'Esta funcion recibe una fecha en formato MM/DD/AAAA HH:MM:SS y como resultado devuelve la misma fecha 
'pero con el formato DD/MM/AAAA HH:MM:SS.
Dim R
Dim strHora

if (Len(P_dteFecha) > 10) then
   if (instr(1,P_dteFecha,"M",1) = 0) and (instr(1,P_dteFecha,"m",1) = 0) then
      strHora= Right(P_dteFecha,8)
   else
      strHora= Right(P_dteFecha,11)
   end if   
   R= day(P_dteFecha) & "/" & month(P_dteFecha) & "/" & year(P_dteFecha) & " " & strHora
else
   R= day(P_dteFecha) & "/" & month(P_dteFecha) & "/" & year(P_dteFecha)
end if
GF_MDA2DMA=R
end function
'--------------------------------------------------------------------------------------------
function GF_MOMENTOSISTEMA()
    GF_MOMENTOSISTEMA = session("MomentoSistema")
end function
'--------------------------------------------------------------------------------------------
function GF_MOMENTODATO()
    GF_MOMENTODATO = session("MomentoDato")
end function
'--------------------------------------------------------------------------------------------
function GF_VERFECHASISTEMA()
     Dim v
     v=mid(session("MmtoSistema"),7,2) & "/" & mid(session("MmtoSistema"),5,2) & "/" & left(session("MmtoSistema"),4)
     v= v & " " & mid(session("MmtoSistema"),9,2) & ":" & mid(session("MmtoSistema"),11,2) & ":" & mid(session("MmtoSistema"),13,2)
     GF_VERFECHASISTEMA=v
end function
'--------------------------------------------------------------------------------------------
function GF_VERFECHADATO()
     Dim v          
     v=mid(session("MmtoDato"),7,2) & "/" & mid(session("MmtoDato"),5,2) & "/" & left(session("MmtoDato"),4)
     v= v & " " & mid(session("MmtoDato"),9,2) & ":" & mid(session("MmtoDato"),11,2) & ":" & mid(session("MmtoDato"),13,2)
     GF_VERFECHADATO=v
end function
'--------------------------------------------------------------------------------------------
Sub GP_ConfigurarMomentos

Call GF_setMomentoDato(day(date()),month(date()),year(date()),hour(time),minute(time()),second(time()))
Call GF_setMomentoSistema(day(date()),month(date()),year(date()),hour(time),minute(time()),second(time()))

end sub
'------------------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 11/08/2003
'Change Log: Se agrego la estandarizacion de hora.
'Autor: Javier A. Scalisi
'Fecha: 12/08/2004
function GF_ControlHora(ByRef P_h,ByRef P_m,ByRef P_s)
'Esta funcion controla la hora.
'Devuelve true si todo esta ok de otra manera devuelve false.
        Dim ret

	ret=true
 	if (CInt(P_h) > 23) and (CInt(P_h) < 0) then ret=false
	if (CInt(P_m) > 59) and (CInt(P_m) < 0) then ret=false
	if (CInt(P_s) > 59) and (CInt(P_s) < 0) then ret=false
	if (ret) then
           Call GF_STANDARIZAR_MM(P_h,P_m,P_s)
        end if
        GF_ControlHora= ret
        
end function
'-------------------------------------------------------------------------------------------------------------------
'/**
' * Funcion: GF_DTEDIFF
' * Descripcion: Obtiene la diferencia entre 2 fechas
' * Parametros:  p_intInit [in] Fecha de Inicio
' *              p_intEnd  [in] Fecha de Fin
' *              p_type    [in] String representando la unidad
' *                             en la que se expresa el resultado.
' * Valor Devuelto
' * Devuelve la diferencia entre las 2 fechas en la unidad
' * indicada segun la siguiente tabla:
' *         p_type   unidad
' *           D       Dias
' *           M       Meses
' *           A       Anios
' *           H       Horas
' *           MM      Min.
' *           S       Seg
' *
' * Autor: Javier A. Scalisi
' * Fecha: 01/06/2004
' */
Function GF_DTEDIFF(p_intInit,p_intEnd,p_type)		
         Dim dteInicio, dteFin, restype, ret
         
         locale=session.lcid
		 session.lcid=2057	'Formato dd/mm/aaaa
	
         dteInicio = GF_MDA2DMA(GF_FN2DTE(p_intInit))
         dteFin = GF_MDA2DMA(GF_FN2DTE(p_intEnd))
         'Obtengo el tipo de resultado esperado
         restype = LCase(p_type)
         if (restype = "mm") then restype = "n"
         if (restype = "a") then restype = "yyyy"
         ret = datediff(restype, dteInicio, dteFin)
         
         session.lcid=locale
         GF_DTEDIFF = ret

End Function
'/**
' * Funcion: GF_DTEADD
' * Descripcion: Suma una cantidad especifica a una fecha
' * Parametros:  p_intDate [in] Fecha de Inicio
' *              p_intAdd  [in] Cantidad a Agregar
' *              p_type    [in] String representando la unidad
' *                             de la cantidad a agregar.
' * Valor Devuelto
' * Agrega a una fecha la cantidad deseada en la unidad
' * indicada segun la siguiente tabla:
' *         p_type   unidad
' *           D       Dias
' *           M       Meses
' *           A       Anios
' *           H       Horas
' *           MM      Min.
' *           S       Seg
' *
' * Autor: Javier A. Scalisi
' * Fecha: 02/06/2004
' */
Function GF_DTEADD(p_intDate,p_intAdd,p_type)
	
	Dim dteInicio, restype, locale, ret
	
	locale=session.lcid
	session.lcid=2057	'Formato dd/mm/aaaa
	
	dteInicio = GF_FN2DTE(p_intDate)		
	'Obtengo el tipo de resultado esperado
	restype = LCase(p_type)
	if (restype = "mm") then restype = "n"
	if (restype = "a") then restype = "yyyy"		
	ret = GF_DTE2FN(dateadd(restype, p_intAdd, dteInicio))        
	
	session.lcid=locale
	GF_DTEADD = ret
	
End Function

'/**
' * Funcion: GF_INT2MES
' * Descripcion: Transforma un numero en su mes correspondiente.
' * Parametros:  p_intNumeroMEs   [in] Numero entre [1-12].
' *
' * Valor Devuelto
' * Devuelve el mes indicado en letras
' *
' * Autor: Javier A. Scalisi
' * Fecha: 08/06/2004
' */

Function GF_INT2MES(p_intNumeroMes)
         
         GF_INT2MES = getNameOfMonth(p_intNumeroMes)
         
End Function
'----------------------------------------------------------------------------------------
function getNameOfMonth(pMes)
dim auxName
	select case cint(pMes)
		case 1
			auxName = "Enero"
		case 2
			auxName = "Febrero"
		case 3
			auxName = "Marzo"
		case 4
			auxName = "Abril"
		case 5
			auxName = "Mayo"
		case 6
			auxName = "Junio"
		case 7
			auxName = "Julio"
		case 8
			auxName = "Agosto"
		case 9
			auxName = "Septiembre"
		case 10
			auxName = "Octubre"
		case 11
			auxName = "Noviembre"
		case 12
			auxName = "Diciembre"
		case else
		    auxName = "ERROR getNameOfMonth"
	end select
    getNameOfMonth = auxName
end function
'/**
' * Funcion: GF_DateGet
' * Descripcion: Obtiene parte de una fecha.
' * Parametros:  p_par   [in] Parte de la fecha a obtener.
' *              p_fecha [in] La fecha.
' *
' * Valor Devuelto
' * Obtiene parte de una fecha se gun se indique
' *          p_par   unidad
' *           D       Dias
' *           M       Meses
' *           A       Anios
' *           H       Horas
' *           MM      Min.
' *           S       Seg
' *
' * Autor: ?
' * Fecha: ?
' */
function GF_DateGet(p_par, p_fecha)

select case ucase(p_par)
       case "A","ANIO"
            GF_DateGet = left(p_fecha,4)
       case "M","MES"
            GF_DateGet = mid(p_fecha,5,2)
       case "D","DIA"
            GF_DateGet = mid(p_fecha,7,2)
       case "H","HORA"
            GF_DateGet = mid(p_fecha,9,2)
       case "MM","MIN","MINUTOS"
            GF_DateGet = mid(p_fecha,11,2)
       case "S","SEG","SEGUNDOS"
            GF_DateGet = mid(p_fecha,13,2)
end select
end function
'/**
' * Funcion: GF_MMTO2SQL
' * Descripcion: Devuelve un Momento en una exprecion LIKE SQL
' * Parametros:  p_field [in] Nombre del campo en la BD.
' *              p_anio  [in] El anio de la fecha.
' *              p_mes   [in] El mes de la fecha.
' *              p_dia   [in] El dia de la fecha.
' *
' * Valor Devuelto
' * Devuelve una expresion del tipo LIKE para una sentencia SQL
' * Autor: Pavlo Henzel
' * Fecha: 10/08/2004
' */
function GF_MMTO2SQL (p_field,anio, mes, dia)
Dim diasql, messql, aniosql, auxAnio
auxAnio = anio 'esto se hace para controlar la salida de anio de GF_Control_fecha que da 2000 cuando no se selecciona anio
if (dia = "") then
    diasql = "__"
else
    diasql=dia
end if
if (mes = "") then
   messql="__"

else
   messql=mes
end if
if (anio = "") then
   aniosql="____"
else
   aniosql=anio
end if
GF_MMTO2SQL = GF_LIKE(p_field,aniosql & messql & diasql)
end function
'--------------------------------------------------------------------------------------------
Function GF_setMomentoDato(P_Dia,P_Mes,P_Anio,P_Hora,P_Min,P_Seg)
dim v
'Se trabaja con la fecha
if (P_Dia <> "") and (P_Mes <> "") and (P_Anio <> "") then
        if (GF_CONTROL_FECHA(P_Dia,P_Mes,P_Anio)) then
           session("MomentoDato") =  P_Mes & "/" & P_Dia & "/" & P_Anio
           session("MmtoDato") = P_Anio & P_Mes & p_Dia
           'Se trabaja con la hora
           v= "000000"
	   if (P_Hora <> "") and (P_Min <> "") and (P_Seg <> "") then
                if (GF_ControlHora(P_Hora,P_Min,P_Seg)) then
                   v = P_Hora & P_Min & P_Seg
                end if
           end if
           session("MmtoDato") = session("MmtoDato") & v
           session("MomentoDato") = "'" & session("MmtoDato") & "'"
	end if
end if
End Function
'--------------------------------------------------------------------------------------------
Function GF_setMomentoSistema(P_Dia,P_Mes,P_Anio,P_Hora,P_Min,P_Seg)
dim v
'Se trabaja con la fecha
if (P_Dia <> "") and (P_Mes <> "") and (P_Anio <> "") then
        if (GF_CONTROL_FECHA(P_Dia,P_Mes,P_Anio)) then
           session("MmtoSistema") = P_Anio & P_Mes & p_Dia
           'Se trabaja con la hora
           v= "000000"
	   if (P_Hora <> "") and (P_Min <> "") and (P_Seg <> "") then
                if (GF_ControlHora(P_Hora,P_Min,P_Seg)) then
                   v = P_Hora & P_Min & P_Seg
                end if
           end if
           session("MmtoSistema") = session("MmtoSistema") & v
           session("MomentoSistema") = "'" & session("MmtoSistema") & "'"
	end if
end if
End Function
'**************************************
'Funcion encargada de determinar el unltimo día del mes indicado.
Function LastDayOfMonth(pYear, pMonth)
    LastDayOfMonth = getLastDayOfMonth(pYear,pMonth)	
End Function
'----------------------------------------------------------------------------------------
function getLastDayOfMonth(pAnio,pMes)
    Dim auxDay
    if (pAnio = "") or (pMes = "") then
        auxDay = ""
    else        
	    select case cint(pMes)
		    case 2
			    if ((CInt(pAnio) mod 4)=0) and ((CInt(pAnio) mod 100)<>0 or (CInt(pAnio) mod 400)=0) then
				    auxDay = 29
			    else
				    auxDay = 28
			    end if	
		    case 4,6,9,11	
			    auxDay = 30	
		    case else
			    auxDay = 31	
	    end select
    end if	    
    getLastDayOfMonth = auxDay
end function
'----------------------------------------------------------------------------------------
Function getDayName(pFn)
    Dim x
    
    x = WeekDay(GF_FN2DTE(pFn))
    
    Select case CInt(x)
        case 1
            getDayName = "Domingo"
        case 2
            getDayName = "Lunes"
        case 3
            getDayName = "Martes"
        case 4
            getDayName = "Miercoles"
        case 5
            getDayName = "Jueves"
        case 6
            getDayName = "Viernes"            
        case else
            getDayName = "Sabado"
            
    end select        
    
    
End Function
%>
