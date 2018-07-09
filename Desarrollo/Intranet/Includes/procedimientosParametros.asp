<%

FUNCTION GF_PARAMETROS7(Parametro,Decimales,Opciones)
' Retorna el valor de un parametro. donde Parametro=nombredelparametro
' Decimales=Cantidad de decimales si es numerico o blanco si es alfabetico
' Opciones,= 4 Parmaetro, 2 formulario, 1 Entorno
DIM R, LEIDO ' Valor a retornar

'response.write "#" & Parametro & "#"
R = ""
LEIDO=false
IF OPCIONES >= 4 and not(LEIDO)THEN 'Intenta tomar parametro
   R = request.querystring(Parametro)
   OPCIONES = OPCIONES - 4
END IF   
IF OPCIONES >= 2 then
   if R = "" then R = request.form(Parametro) ' Tomar de formulario o pagina
   OPCIONES = OPCIONES - 2
END IF
IF OPCIONES = 1 THEN 
   ' Leer el dato del entorno guardado en la session
   DIM N ' Nombre para buscar en session
   N =  session("procedimiento") & "/Prmtr/" & Parametro
   DIM A
   A = session(N) ' leer el valor del entorno   
   IF R = "" THEN 
      R = A ' No se leyo valor de parametro ni formulario asume de entorno
   ELSE
      ' hay un valor por parametro o formulario si es distinto al entorno
	  ' modificar entorno 
	  IF R <> A THEN  session(N) = R
   END IF		 
END IF
if Decimales <> "" then 
   R = replace(R,",",".")
   if IsNumeric(trim(R)) then
		R = FormatNumber(CDbl(R),Decimales,true,false,false) 			
   else
		R = 0
   end if
   if (Decimales = 0) then
		R = CDbl(R)
   else
		R = CDbl(R)
   end if
END IF
GF_PARAMETROS7 = R
END FUNCTION
'-------------------------------------
FUNCTION  GF_PARAMETROSNUMEROS(P,F,L)
DIM R
R = GF_PARAMETROS(P,F)
WHILE LEN(R) > 0 AND LEN(R) < L
      R = "0" & R
WEND
GF_PARAMETROSNUMEROS = R	  
END FUNCTION
'-------------------------------------
function GF_ParametrosNumericosEnteros(P)
gf_ParametrosNumericosEnteros = Int(GF_ParametrosNumericos(P))
end function
'-------------------------------------
function GF_ParametrosNumericos2D(P)
dim my_valor 
gf_parametrosnumericos2d =  gf_2d(gf_parametrosnumericos(p))
end function
'-------------------------------------
function GF_ParametrosNumericos(P)
dim My_ValorNumerico
My_ValorNumerico = GF_ParametrosForm(P)
My_ValorNumerico = replace(My_ValorNumerico,".",",")
if not isnumeric(My_ValorNumerico) then  My_ValorNumerico = 0
GF_ParametrosNumericos = My_ValorNumerico + 0 
end function
'-------------------------------------
FUNCTION GF_ParametrosForm(p)
DIM R
R = GF_Parametros(p,p)
IF R = "" THEN r = GF_ParametrosUser(p)
gf_ParametrosForm = r
end function
FUNCTION  GF_PARAMETROS(P,F)
DIM R
R = ""
IF p <> "" THEN R = request.querystring(P)
IF F <> "" AND R = "" THEN R = request.form(F)
GF_PARAMETROS = R
END FUNCTION
function GF_PARAMETROSUSER(P)
' Retorna el ultimo valor utilizado por el usuario para el parametro en la pagina
GF_PARAMETROSUSER = ""
END FUNCTION

%>
