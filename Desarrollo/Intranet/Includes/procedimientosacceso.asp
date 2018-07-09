<%
'                         ********************************
'                         *     Funciones de Acceso      *
'                         ********************************
'                         *      Fecha: 25/03/2003       * 
'                         ********************************
'Niveles de Acceso:
'0 : No tiene Acceso.
'1 : Acceso de Solo Lectura.
'2 : Acceso de Update.
'3 : Acceso de ADD.
'9 : Acceso Total.
'-1: No hay informacion sobre el acceso.
function GF_MG_ACCESO_REGISTRO(P_strACCESO,ByRef P_strKM, ByRef P_strKC,ByRef P_intKR,ByRef P_strDS)
'Esta funcion devuelve el nivel de acceso a un registro.
'P_strACCESO= Tipo de Acceso requerido.
'P_strKM    = KM del registro al que se quiere tener acceso.
'P_strKC    = KC del registro al que se quiere tener acceso.
'P_intKR    = KR del registro al que se quiere tener acceso.
'P_strDS    = DS del registro al que se quiere tener acceso.

Dim intAcceso
Dim intAccesoTemp,strListaCargos
Dim strClaveaccesoRegistro

    
   'Controlo que haya un usuario valido.
   if (session("usuario") = "") then Response.Redirect (session("home"))

   intAcceso=-1
   'Verifico si el registro existe.
   if (GF_MGC(P_strKM,P_strKC,P_intKR,P_strDS)) then              '[1]
	   if (session("usuario") = "ADMIN") then                     '[2]
	      intAcceso=9
	   else	     
	      strClaveAccesoRegistro=session("usuario") & "/" & P_strACCESO & "/" & P_intKR
          'Verifico si ya no se pidio este acceso.
	      intAccesoTemp=session(strClaveAccesoRegistro)
		  if (intAccesoTemp = "") then                            '[3]
             'Busco el nivel de acceso.
  	         strListaCargos=session("LISTADECARGOS")
			 if (strListaCargos = "") then Response.Redirect(session("home"))
			 intAccesoTemp= GF_SESSIONACCESOBUSCAR(GF_SESSIONKR("SR",P_strACCESO),P_intKR)
			 if (intAccesoTemp = "") then 
			    intAcceso=-1
			 else 
			    intAcceso= intAccesoTemp
			 end if
			 'Se devuelve el valor del nivel de acceso.
             session(strClaveAccesoRegistro)=intAcceso
	      else
	         intAcceso= intAccesoTemp
	      end if                                              '[/3]
	   end if                                                 '[/2]
   end if                                                    '[/1]      	  	  
   GF_MG_ACCESO_REGISTRO= intAcceso

end function
'--------------------------------------------------------------------------------------------------------
function GF_ACCESO_DATO(P_strMGKM,P_strMGKC,P_intMGKR,P_strKCDATO)
' Esta funcion devuelve el nivel de acceso a un dato para un registro determinado.
'P_strMGKM  = Maestro del registro sobre el que se quiere acceder un dato.
'P_strMGKC  = Clave del registro sobre el que se quiere acceder a un dato.
'P_strMGKR  = Clave reducida del registro sobre el que se quiere acceder a un dato.
'P_strKCDATO= El dato al que se quiere tener acceso.
Dim strDS
Dim intValorAcceso,intValor
Dim strNombreRaiz

    strNombreRaiz = Trim(P_strKCDATO) & Trim(P_strMGKM) & Trim(P_strMGKC) & Trim(P_intMGKR)
    'Traigo el valor de acceso para este registro.
    intValorAcceso= session(strNombreRaiz & "ACC")
	if (intValorAcceso = "") then                                          '[1]     
   	  'Si no hay datos sobre el acceso al registro consulto que acceso tengo al maestro.
      intValorAcceso=GF_MG_ACCESO_REGISTRO("ACCESO","SM", P_strMGKM,"","")
      if (CInt(intValorAcceso) = -1) then                                  '[2]
		 'Si no hay datos de acceso ni al registro ni al maestro, se le niega acceso al registro.
		 intValorAcceso=0                                     
      end if                                                               '[/2]
	   '// CONSULTA DE ACCESO AL DATO //
	   if (CInt(intValorAcceso) > 0) then                                  '[3]
         'Si no hay datos sobre el acceso al dato consulto que acceso tengo al maestro.
          intValor=GF_MG_ACCESO_REGISTRO("ACCESO","SM", "SD","","")
   	      if (CInt(intValor) = -1) then                                    '[4] 
		     'Si no hay datos de acceso ni al dato ni al maestro, se le niega acceso al dato.
		     intValor=0
		  end if                                                           '[/4]
	      'Tomo el menor de los accesos  
          if (CInt(intValor) < CInt(intValorAcceso)) then intValorAcceso=intValor          
		end if                                                             '[/3]     
    end if                                                                 '[/1]
    session(strNombreRaiz & "ACC") = intValorAcceso
    GF_ACCESO_DATO= intValorAcceso

end function
'--------------------------------------------------------------------------------------------------------
function GF_MANAGER_DATO(P_strKCDATO,P_strMGKM,P_strMGKC,P_intMGKR)
'Esta funcion administra los datos del sistema para cualquier registro.
Dim intValorAcceso,intKR
Dim ValorNEW,ValorOLD
Dim strNombreRaiz

strNombreRaiz = Trim(P_strKCDATO) & Trim(P_strMGKM) & Trim(P_strMGKC) & Trim(P_intMGKR)
'Tomo el valor de acceso al dato.
intValorAcceso= request.Form(strNombreRaiz & "ACC")
if (intValorAcceso = "") then 
  intValorAcceso = CInt(GF_ACCESO_DATO(P_strMGKM,P_strMGKC,P_intMGKR,P_strKCDATO))
  response.write("<input type='hidden' name='" & strNombreRaiz & "ACC' value='" & intValorAcceso & "'>")
  if (intValorAcceso > 0) then 
     ValorNEW= GF_DT1("READ",P_strKCDATO,"","",P_strMGKM,P_strMGKC)  
     ValorOLD=ValorNEW
  end if
else
  ValorOLD= request.Form(strNombreRaiz & "OLD")
  if(intValorAcceso > 1) then
     ValorNEW= request.Form(strNombreRaiz & "NEW")
     if (ValorOLD <> ValorNew) then
        'Actualizo el dato.
        GF_DT1W P_strKCDATO,P_strMGKM,P_strMGKC,ValorNEW,P_dteMmtoDato
		ValorOLD=ValorNEW
     end if	  
  end if
end if  
'METODO XXX
if (intValorAcceso > 1) then
   'Muestro el dato.
   response.write("<input type='text' name='" & strNombreRaiz & "NEW' value='" & ValorNEW & "'>")
else
   if (intValorAcceso = 1) then response.write(ValorOLD)
end if	            
response.write("<input type='hidden' name='" & strNombreRaiz & "OLD' value='" & ValorOLD & "'>")

end function
%>