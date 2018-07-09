<!--#include file="../includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/md5.asp"-->
<%
Function PwdAleatorio ( Longitud, Repetir )
'----------------------------------------------------------
' por Carlos de la Orden Dijs, Abril 2001 - 100% gratis! ;-)
'----------------------------------------------------------
' Devuelve una cadena con una secuencia de caracteres
' aleatoria, de longitud especificada.
' Si Repetir = True la secuencia puede contener caracteres
' repetidos. Si Repetir = False, todos los caracteres son
' �nicos.
' Para a�adir m�s caracteres posibles, a�adirlos al vector
' vCaracteres simplemente separando como comas, como los
' que est�n ya escritos, y la funci�n los escoger�.
'----------------------------------------------------------
Dim vPass(), I, J ' nuestro vector y dos contadores
Dim vNumeros()	  ' vector para guardar lo que llevamos
Dim n, bRep
Dim vCaracteres	  ' vector donde est�n los posibles caract.

vCaracteres = Array("a","b", "c", "d", "e", "f", "g", "h", _
"i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", _
"u", "v", "w", "x", "y", "z", "A","B", "C", "D", "E", "F", _
"G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
"S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", _
"5", "6", "7", "8", "9", "0")

'Establezco la longitud del vector
Redim vPass(Longitud-1)
'Y del vector auxiliar que guarda los caracteres ya escogidos
Redim vNumeros(Longitud-1)
I = 0
'Inicializo los n�s aleatorios
Randomize
'Hasta que encuentre todos los caracteres
do until I = Longitud
	'Hallo un n�mero aleatorio entre 0 y el m�ximo indice
	' del vector de caracteres.
	n = int(rnd*Ubound(vCaracteres))
	'Si no puedo repetir...
	if not Repetir then
		bRep = False
		'Busco el numero entre los ya elegidos
		for J = 0 to UBound(vNumeros)
			if n = vNumeros(J) then
			'Si esta, indico que ya estaba
				bRep = True
			end if
		next
		'Si ya estaba, tengo que repetir este caracter
		'as� que resto 1 a I y volvemos sobre la misma
		'posici�n.
		if bRep then
			I = I - 1
		else
			vNumeros(I) = n
			vPass(I) = vCaracteres(n)
		end if
	else
	'Me da igual que est� o no repetido
		vNumeros(I) = n
		vPass(I) = vCaracteres(n)
	end if
'Siguiente car�cter!
I = I + 1
loop

'Devuelvo la cadena. Join une los elementos de un vector
'utilizando como separador el segundo par�metro: en este
'caso, nada -> "".
PwdAleatorio = Join(vPass, "")

End Function 'PwdAleatorio
'==========================================================
Function generarLlave()
	Dim myIP,token,rs,strSQL		
	myIP  = Request.ServerVariables("REMOTE_ADDR") 'OBTENGO LA IP
	token = UCASE(PwdAleatorio(10,false))	
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", "DELETE FROM TBLLOGINTOKENS WHERE  TOKENOWNER= '" & myIP & "'")			
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "EXEC", "INSERT INTO TBLLOGINTOKENS(TOKENOWNER, TOKEN, ISSUEDATE, VALIDTO) VALUES('" & myIp & "','" & token & "',GETDATE(), DATEADD(minute,1,GETDATE()))")	
	generarLlave = token	
End Function
'==========================================================
%>