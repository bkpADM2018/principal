<%
'//--   INTERFAZ DE LA UNIDAD
'//--
'//-- function GF_CONTROL_EMAIL(P_strEMAIL)
'//--    true= OK  false=Invalido
'//--
'//-- function GF_CONTROL_PASSWORD(P_strPass)
'//--    true= OK  false=Invalido
'//--
'//-- function GF_CONTROL_CUIT(P_strCUIT)
'//--    true= OK  false=Invalido
'----------------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 26/02/2003
function GF_CONTROL_EMAIL(P_strEMAIL)
'Esta funcion sirve para controlar un email.

Dim aux,lng,R,aux2,aux3, backMail

R=false
if (P_strEMAIL <> "") then
   R=true
   'Borro todos los espacos en blanco a derecha e izquierda.
   P_strEMAIL = Trim(P_strEMAIL)
   lng=Len(P_strEMAIL)
   'Controlo que exista un @
   aux=InStr(P_strEMAIL,"@")
   if (aux = 0) then R=false
   'Si el @ esta en la primera letra tampoco es valido
   if (aux = 1) then R=false
   backMail = Right(P_strEMAIL,lng-aux)
   lng=Len(backMail)
   'Controlo que no exista mas que una @
   if (InStr(backMail,"@") <> 0) then R=false
   'Controlo que se haya ingresado un punto   
   aux2=InStr(backMail,".")
   'if (aux2 < 4) then R=false   
   if (aux2 = lng) then R=false
   'Verifico si hay otro punto y si lo hay controlo que entre ellos haya 3 caracters ".com.ar"
   'aux3=InStr(Right(backMail,lng-aux2),".")  
   'if (aux3 <> 0) then
      'Hay otro punto, no puede haber mas
  '	  aux3= aux3 + aux2 'Transformo la posicion a la longitud real desde el primer caracter
	'  if (InStr(Right(backMail,lng-aux3),".") <> 0) then R=false
	  'Controlo la descripcion de 2 caracteres
	'  if (mid(backMail,lng-2,1) <> ".") then R=false
	'  'Controlo la descripcion de 3 caracteres
'	  if (mid(backMail,lng-6,1) <> ".") then R=false
 '  else
  '    if ((lng-aux2) > 3) then R=false
  ' end if	     
end if 
GF_CONTROL_EMAIL=R
end function
'----------------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 26/02/2003
function GF_CONTROL_PASSWORD(P_strPass)
'Esta funcion se encarga de controlar passwords.

GF_CONTROL_PASSWORD=false
if (P_strPass <> "") then
   GF_CONTROL_PASSWORD=true
   '1.- Borro todos los espacos en blanco a derecha e izquierda.
   P_strPass = Trim(P_strPass)       
   '2.- Controlo que no haya espacios en blanco
   if (InStr(P_strPass," ") <> 0) then 
      GF_CONTROL_PASSWORD=false
   else
      '3.- Controlo que no haya determinados caracteres especiales
      if (InStr(P_strPass,"*") <> 0) then 
         GF_CONTROL_PASSWORD=false
      else
         if (InStr(P_strPass,"'") <> 0) then 
            GF_CONTROL_PASSWORD=false
         else
            if (InStr(P_strPass,"@") <> 0) then 
               GF_CONTROL_PASSWORD=false
            else
               if (InStr(UCase(P_strPass),"SELECT") <> 0) then 
                  GF_CONTROL_PASSWORD=false
               else
                  if (InStr(UCase(P_strPass),"INSERT") <> 0) then 
                     GF_CONTROL_PASSWORD=false
                  else
                     if (InStr(UCase(P_strPass),"UPDATE") <> 0) then GF_CONTROL_PASSWORD=false
                  end if
               end if
            end if
         end if
      end if
   end if   
end if
end function
'----------------------------------------------------------------------------------------------------
'Autor: Javier A. Scalisi
'Fecha: 26/02/2003
function GF_CONTROL_CUIT(P_strCUIT)
'Esta funcion controla un numero de CUIT o CUIL.
Dim strCUIT, arrNumeros(11), k, suma, resto, digito

GF_CONTROL_CUIT=false
if (P_strCUIT <> "") then
   strCUIT = trim(replace(P_strCUIT,"-",""))
   'Vrifico la longitud
   if (Len(P_strCUIT) = 11) then
      'Verifico que el string sea numerico
   	  if (isNumeric(Left(strCUIT,2))) and (isNumeric(Mid(strCUIT,3,8))) and (isNumeric(Right(strCUIT,1))) then
        'Separo los numeros
        for k= 1 to 11
            arrNumeros(k-1) = Mid(strCUIT,k,1)
        next
        'Calculo el digito verificador.
        suma=0
        suma = suma + (arrNumeros(0) * 5)
        suma = suma + (arrNumeros(1) * 4)
        suma = suma + (arrNumeros(2) * 3)
        suma = suma + (arrNumeros(3) * 2)
        suma = suma + (arrNumeros(4) * 7)
        suma = suma + (arrNumeros(5) * 6)
        suma = suma + (arrNumeros(6) * 5)
        suma = suma + (arrNumeros(7) * 4)
        suma = suma + (arrNumeros(8) * 3)
        suma = suma + (arrNumeros(9) * 2)
        resto = suma mod 11
        digito = 11 - resto
        if (digito > 9) then
            digito=0
        end if
        'Comparo con el digito verificador
        if (Cint(digito) = CInt(arrNumeros(10))) then
            GF_CONTROL_CUIT=true
        end if
      end if
   end if
end if
end function
%>