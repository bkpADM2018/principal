<%
'********************************************************************
'                    procedimientosFormato
'          Funciones para conversion de cadenas y numeros
'********************************************************************
'/**
' * DEFINICION DE CONSTANTES
' */
Const CHR_AFT = 0
Const CHR_FWD = 1
'/**
' * Funcion    : nChars
' * Descripcion: Esta funcion estandariza un string a la
' *              cantidad de caracteres especificados.
' * Parametros : p_string [in] String a estandarizar
' *              p_size   [in] Longitud a alcanzar
' *              p_char   [in] Caracter/es a anteponer
' *              p_pos    [in] CHR_FWD:Delante CHR_AFT:Detras
' * Valor Devuelto:
' * Devuelve un string con la longitud especificada completando de
' * ser necesario con tantos p_char como sea necesario.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 10/05/2004
' */
Function GF_nChars(p_string,p_size,p_char,p_pos)
   Dim ret,prefix,suffix,cant

   'Se preparan las variables auxiliares de trabajo
   ret=Trim(p_string)
   cant=p_size
   if (ret <> "") then cant = cant - Len(ret)   
   if (cant < 0) then
        ret = left(ret, p_size)
        cant = 0
   end if
   prefix=""
   suffix=""   
   if (p_pos = CHR_FWD) then
      prefix=string(cant,p_char)
   else
      suffix=string(cant,p_char)
   end if
   'Se completa la longitud
   if (CInt(cant) > 0) then ret= prefix & ret & suffix
   GF_nChars= ret

End Function
'/**
' * Funcion    : GF_nDigits
' * Descripcion: Esta funcion estandariza un string numerico
' *              a la cantidad de digitos especificados.
' * Parametros : p_number [in] Numero a estandarizar
' *              p_size   [in] Longitud a alcanzar
' * Valor Devuelto:
' * Devuelve un string numerico con la longitud especificada
' * anteponiedo tantos ceros (0) como sea necesario.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 10/05/2004
' */
Function GF_nDigits(p_number,p_size)
   GF_nDigits= GF_nChars(p_number,p_size,"0",CHR_FWD)
End Function
'/**
' * Funcion: GF_STR2CUIT
' * Descripcion: Convierte un string a CUIT.
' * Parametros:  p_strText [in] String a formatear
' * Valor Devuelto
' *       Devuelve el string formateado como CUIT
' *       XX-XXXXXXXX-X.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 20/05/2004
' */
Function GF_STR2CUIT(p_strText)
         Dim ret
		 ret = p_strText
		 if (ret <> "") then
			 ret =  "ERR FORMATO"
			 if (len(p_strText) = 11) then
				ret = left(p_strText,2) & "-" & mid(p_strText,3,8) & "-" & right(p_strText,1)
			 end if
		 end if
         GF_STR2CUIT = ret
End Function
'/**
' * Funcion: GF_EDIT_CBTE
' * Descripcion: Edita un string a Comprobante (Factura)
' * Parametros: p_nbr  [in] String a formatear
' * Valor Devuelto
' *       Devuelve el string con formato de nro de factura
' *       XXXX-XXXXXXXX.
' */
Function GF_EDIT_CBTE(p_nbr)
    Dim nbr
    nbr = GF_nDigits(p_nbr, 12)
	GF_EDIT_CBTE= left(nbr,4) & "-" & right(nbr,len(nbr)-4)
End function

'/**
' * Funcion    : GF_EDIT_CONTRATO
' * Descripcion: Esta funcion Arma el numero de contrato dandole
' *              el formato alfanumerico correcto
' * Parametros : p_cto1 [in] 2 digitos (producto)
' *              p_cto2 [in] 1 digito  (sucursal)
' *              p_cto3 [in] 2 digitos (operacion)
' *              p_cto4 [in] 5 digitos (secuencia)
' *              p_cto5 [in] 2 digitos (cosecha)
' * Valor Devuelto:
' * Si todos los parametros son correctos devuelve el numero
' * de contrato con el formato XX-X-XX-XXXXX/XX.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 20/05/2004
' */
Function GF_EDIT_CONTRATO(p_cto1,p_cto2,p_cto3,p_cto4,p_cto5)

   Dim aux
   'Controlo que los parametros sean numericos
   aux = GF_nDigits(p_cto1,2) & "-" & GF_nDigits(p_cto2,1)
   aux = aux & "-" & GF_nDigits(p_cto3,2) & "-" & GF_nDigits(p_cto4,5) & "/" & GF_nDigits(p_cto5,2)
   GF_EDIT_CONTRATO = aux

End Function
'/**
' * Funcion    : GF_INT2CTO
' * Descripcion: Esta funcion transforma un string numerico en un
' *              contrato dandole el formato alfanumerico correcto
' * Parametros : p_cto [in] 12 digitos (formato : XXXXXXXXXXXX)
' * Valor Devuelto:
' * Si todos los parametros son correctos devuelve el numero
' * de contrato con el formato XX-X-XX-XXXXX/XX.
' *
' * Autor: Javier A. Scalisi
' * Fecha: 20/05/2004
' */
Function GF_INT2CTO(p_cto)

   Dim aux
   'Controlo que el string parado tenga la cant de digitos correcta.
   if (Len(p_cto) <> 12) then
      aux = GF_EDIT_CONTRATO("00","0","00","00000","00")
   else
      aux= GF_EDIT_CONTRATO(left(p_cto,2),mid(p_cto,3,1),mid(p_cto,4,2),mid(p_cto,6,5),right(p_cto,2))
   end if
   GF_INT2CTO = aux

End Function
'/**
' * Funcion    : GF_EDIT_DECIMALS
' * Descripcion: Edita un numero a la cantidad de decimales pedida
' * Parametros : p_nbr [in] Numero a editar
' *              p_dec [in] Decimales
' *
' * Autor: Javier A. Scalisi
' * Fecha: 10/06/2004
' *
' * Modifico: Di Santo, Eugenio
' * Fecha: 31/10/2006
' *
' * Modifico: Bacarini, Ezequiel
' * Fecha: 29/09/2010
' */
Function GF_EDIT_DECIMALS(p_nbr, p_dec)
dim rtrn		
		if IsNumeric(trim(p_nbr)) then
			rtrn = FormatNumber(cdbl(p_nbr)/(10^p_dec),p_dec) 
			rtrn = replace(rtrn,",","#" )
			rtrn = replace(rtrn,".","," )
			rtrn = replace(rtrn,"#","." )
		else
			rtrn = "ERR"
		end if
		GF_EDIT_DECIMALS = rtrn
End Function
'/**
' * Funcion    : GF_EDIT_CTAPTE
' * Descripcion: Esta funcion da formato a la Carta de Porte
' * Parametros : pCtaPte 16 digitos 
' * Valor Devuelto:
' * XXXX-XXXXXXXX-XXXX
' * Autor: Ezequiel A. Bacarini
' * Fecha: 20/01/2011
' */
Function GF_EDIT_CTAPTE(pCtaPte)
	Dim aux	
	aux = left(pCtaPte,4) & "-" & mid(pCtaPte,5,8) 
	if (Len(Trim(pCtaPte)) > 12) then aux = aux & "-" & right(pCtaPte,4)
	GF_EDIT_CTAPTE = aux
End Function
'/**
' * Funcion    : GF_EDIT_PATENTE
' * Descripcion: Esta funcion da formato a la Patente de un vehiculo
' * Parametros : pPatente 6 digitos 
' * Valor Devuelto:
' * AAA-000
' * Autor: Ezequiel A. Bacarini
' * Fecha: 20/01/2011
' */
Function GF_EDIT_PATENTE(pPatente)
	Dim aux
	pPatente = trim(pPatente)
	if len(pPatente) > 6 then
		aux = left(pPatente,2) & " " & mid(pPatente, 3,3) & " " & right(pPatente,2)
	else
		aux = left(pPatente,3) & "-" & right(pPatente,3)
	end if
	GF_EDIT_PATENTE = aux
End Function
'----------------------------------------------------------------------------------------------------------------------
Function splitRecordLines(ByRef p_oPDF, p_string, p_width)
        dim wtp_i,wth, strExpr, wordArray, coord_Y, strExprFinal, last_align
        Dim ret(), lineas, auxParrafo, i
		
		auxParrafo = split(p_string, chr(10))		
		'RESPONSE.WRITE UBound(auxParrafo) & "<br>"
		lineas=0
		for i = 0 to ubound(auxParrafo)
			'Determino el ancho del texto.(Sin componentes de formato)			
			MyText = Trim(auxParrafo(i))
			'RESPONSE.WRITE "(" & MyText & ")<br>"
			wth = Int(p_oPDF.Metrics.GetTextWidth(MyText, pdf_currentTextFont , pdf_currentFontSize ))					
			 'Si corresponde separo en renglones
			if (CLng(wth) > CLng(p_width)) then
				wordArray = Split(Trim(MyText)," ")
				strExpr = ""
				strExprFinal = ""            
				for wtp_i = LBound(wordArray) to UBound(wordArray)
					if (strExpr = "") and (wordArray(wtp_i)<>"") then
						strExpr = wordArray(wtp_i)
					else
						strExpr = strExpr & " " & wordArray(wtp_i)
					end if
					MyText = strExpr
					wth = Int(p_oPDF.Metrics.GetTextWidth(MyText, pdf_currentTextFont , pdf_currentFontSize ))
					if (wth >= p_width) then
						lineas=lineas+1				
						Redim preserve ret(lineas)
						'response.write lineas & "-" & strExprFinal & "<br>"
						ret(lineas-1) = Trim(strExprFinal)						
						strExprFinal = ""
						strExpr = wordArray(wtp_i)						
					end if                
					strExprFinal = strExpr                    
				next
				'Agerego la ultima linea				
				lineas=lineas+1				
				Redim preserve ret(lineas)
				ret(lineas-1) = Trim(strExprFinal)				
				'response.write lineas & "-" & strExprFinal & "<br>"
			else
				lineas=lineas+1
				Redim preserve ret(lineas)
				ret(lineas-1) = MyText
				'response.write lineas & "-" & MyText & "<br>"
			end if		
		Next
        splitRecordLines = ret
		'response.end
End Function
'----------------------------------------------------------------------------------------------------------------------
'MOD: 2017-10-04 - JAS
Function generateBodyText(ByRef p_oPDF, rsBody, Byref totalLineas, p_width)
	Dim auxParrafo, coord_Y, i, ret(1000, 2), lineas, aux, k
	
	lineas = 0	
	while (not rsBody.eof) 
		'Separo un regsitro de la BD en todas la lineas que corresponden a su texto.		
		if (Trim(rsBody("concepto")) <> "") then
			aux = splitRecordLines(p_oPDF, rsBody("concepto"), p_width)			
			for k = 0 to ubound(aux)-1
				if (trim(aux(k)) <> "") or ((Trim(aux(k)) = "") and (k <> ubound(aux)-1)) then
					ret(lineas, 0) = aux(k)
					ret(lineas, 1) = 0				
					lineas = lineas + 1
				end if
			next			
		else
			if (CDbl(rsBody("importe")) > 0) then lineas = lineas + 1			
		end if
		ret(lineas-1, 1) = rsBody("importe")
		rsBody.MoveNext()		
	wend
	totalLineas = lineas-1
	generateBodyText = ret
End Function
'******************************************************
'Funcion que formatea el numero a 8 o 12 caracteres numericos fijos (00000000 o 000000000000)
Function formatearNumeroFactura(pNumero)
	Dim rtrn,largo
	rtrn  = cstr(pNumero)
	largo = len(pNumero)
	if (largo<8) then
		for i = largo to 7
			rtrn = "0" & rtrn
		next
	end if
	
	formatearNumeroFactura = rtrn
End Function
'******************************************************
'Funcion que edita un texto para mantener los enters en la base de datos.
Function editText4DB(p_string)
	Dim ret
	ret = replace(p_string,"'","*")
	ret = replace(ret,chr(13),"")
	ret = replace(ret,chr(10),ENTER_SYMBOL)
	if (left(ret,4) = ENTER_SYMBOL) then
		ret = mid(ret,5,len(ret))
	end if
	editText4DB = ret
End Function
'******************************************************
'Funcion que edita un texto proveniente de la DB para mantener los enters en la pantalla.
Function editText4Input(p_string)
	Dim ret	
	ret = replace(p_string,ENTER_SYMBOL, chr(10))
	ret = replace(ret,"*","'")		
	editText4Input = ret
End Function
'******************************************************************
'//USADOS POR LA WEB
'******************************************************************
'/**
' * Funcion    : GF_EDIT_DECIMALS_POINT
' * Descripcion: Edita un numero a la cantidad de decimales pedida,separando con un punto
' * Parametros : p_nbr [in] Numero a editar
' *              p_dec [in] Decimales
' *
' * Autor: Javier A. Scalisi
' * Fecha: 10/06/2004
' *
' * Modifico: Henzel, Pavlo
' * Fecha: 10/12/2008

' */
Function GF_EDIT_DECIMALS_POINT(p_nbr, p_dec)
         Dim length, ret, decimals
		
         p_nbr = Clng(p_nbr)
         length = len(p_nbr)
         ret=""
         if (length <= p_dec) then
            ret ="0."
            decimals = p_dec - length
            while (decimals > 0)
               ret = ret & "0"
               decimals = decimals - 1
            wend
            ret = ret & p_nbr
         else
             ret = left(p_nbr,len(p_nbr) - p_dec) & "." & right(p_nbr,p_dec)
         end if
         GF_EDIT_DECIMALS_POINT =ret
End Function
'/**
' * Funcion    : GF_EDIT_INTEGER
' * Descripcion: Edita un numero entero poniendole el separador de unidades de mil
' * Parametros : p_nbr [in] Numero a editar
' *
' * Autor: Di Santo, Eugenio
' * Fecha: 31/10/2006
' */
Function GF_EDIT_INTEGER(p_nbr)        
        GF_EDIT_INTEGER = GF_EDIT_DECIMALS(CDbl(p_nbr), 0)
End Function

Function GF_SET_VALOR(p_valor, p_default)
    if (p_valor <> "") then
        GF_SET_VALOR = p_valor
    else
        GF_SET_VALOR = p_default
    end if
End Function

'/**
' * Funcion    : num2words
' * Descripcion: Convierte un numero a su exprecion literla en ingles.
' * Parametros : iNum [in] Numero a editar
' *
' * Autor: Javier A. Scalisi
' * Fecha: 06/05/2014
' */
Dim f_ss, f_ds, f_ts, f_qs
 
f_ss = "one,two,three,four,five,six,seven,eight,nine" 
f_ds = "ten,eleven,twelve,thirteen,fourteen,fifteen,sixteen," & _ 
     "seventeen,eighteen,nineteen" 
f_ts = "twenty,thirty,forty,fifty,sixty,seventy,eighty,ninety" 
f_qs = ",thousand,million,billion" 

Function num2words(iNum) 
	Dim i
	
    i = iNum 
    if i < 0 then b = true: i = i*-1 
    if i = 0 then 
        s="zero" 
    elseif i <= 2147483647 then 
        a = split(f_qs,",") 
        'Response.Write i & "<br>"
        for j = 0 to 3 
            iii = i mod 1000
            i = i \ 1000 
            if iii > 0 then s = nnn2words(iii) & " " & a(j) & " " & s 
        next 
    else 
        s = "out of range value" 
    end if 
    if b then s = "negative " & s    
    num2words = trim(s) 
End Function 

Function nnn2words(iNum) 
    a = split(f_ss,",") 
    i = iNum mod 10 
    if i > 0 then s = a(i-1) 
    ii = int(iNum mod 100)\10 
    if ii = 1 then  
        s = split(f_ds,",")(i) 
    elseif ((ii>1) and (ii<10)) then 
        s = split(f_ts,",")(ii-2) & " " & s 
    end if 
    i = (iNum \ 100) mod 10 
    if i > 0 then s = a(i-1) & " hundred " & s 
    nnn2words = s 
End Function 
 


'/**
' * Funcion    : numeroALetras
' * Descripcion: Convierte un numero a su exprecion literla en español.
' * Parametros : pnumero [in] Numero a editar
' *
' * Autor: Javier A. Scalisi
' * Fecha: 06/05/2014
' */

Dim xcen(9) 'centenas
Dim xdec(9) 'decenas
Dim xuni(9) 'unidades
Dim xexc(6) 'except
Dim ceros(9)

Function numeroALetras(pnumero)

Dim letras
Dim i
Dim c
Dim j
Dim xnumero
Dim xnum
Dim num
Dim digito
Dim numero_ent
Dim entero
Dim decimales
Dim temp
  
  xcen(2) = "dosc"
  xcen(3) = "tresc"
  xcen(4) = "cuatrosc"
  xcen(5) = "quin"
  xcen(6) = "seisc"
  xcen(7) = "setec"
  xcen(8) = "ochoc"
  xcen(9) = "novec"
  xdec(2) = "veinti"
  xdec(3) = "trei"
  xdec(4) = "cuare"
  xdec(5) = "cincue"
  xdec(6) = "sese"
  xdec(7) = "sete"
  xdec(8) = "oche"
  xdec(9) = "nove"
  xuni(1) = "uno"
  xuni(2) = "dos"
  xuni(3) = "tres"
  xuni(4) = "cuatro"
  xuni(5) = "cinco"
  xuni(6) = "seis"
  xuni(7) = "siete"
  xuni(8) = "ocho"
  xuni(9) = "nueve"
  xexc(1) = "diez"
  xexc(2) = "once"
  xexc(3) = "doce"
  xexc(4) = "trece"
  xexc(5) = "catorce"
  xexc(6) = "quince"
  ceros(1) = "0"
  ceros(2) = "00"
  ceros(3) = "000"
  ceros(4) = "0000"
  ceros(5) = "00000"
  ceros(6) = "000000"
  ceros(7) = "0000000"
  ceros(8) = "00000000"
  
  c = 1
  i = 1
  j = 0
  
  xnumero = cStr(pnumero)
If Cdbl(LTrim(RTrim(pnumero))) < 999999999.99 Then
    numero_ent = Int(Cdbl(pnumero))
    If Len(numero_ent) < 9 Then
        numero_ent = ceros(9 - Len(numero_ent)) & numero_ent
    End If
    entero = Cdbl(Int(numero_ent))
    decimales = (Cdbl(xnumero) - entero) * 100
    
	Do While i < 8
        temp = 0
        num = Cdbl(Mid(numero_ent, i, 3))
        xnum = Mid(numero_ent, i, 3)
        digito = Cdbl(Mid(xnum, 1, 1))
        
        '/* analizo el numero entero de a 3 */
        If xnum = "000" Then
            j = 0
        Else
        	j = 1
            If digito > 1 Then
                letras = letras & xcen(digito) & "ientos "
            End If
            If Mid(xnum, 1, 1) = "1" And Mid(xnum, 2, 2) <> "00" Then
                letras = letras & "ciento "
            ElseIf Mid(xnum, 1, 1) = "1" Then
                letras = letras & "cien "
            End If
  
  			'/* analisis de las decenas */
            digito = Cdbl(Mid(xnum, 2, 1))
            If digito > 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & xdec(digito) & "nta "
				temp = 1
            End If
            
			If digito > 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & xdec(digito) & "nta y "
				
            End If
            
			If digito = 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & "veinte "
				temp = 1
            ElseIf digito = 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & "veinti"
				
            End If
            
			If digito = 1 And Mid(xnum, 3, 1) >= "6" Then
                letras = letras & "dieci"
            ElseIf digito = 1 And Mid(xnum, 3, 1) < "6" Then
                letras = letras & xexc(Cdbl(Mid(xnum, 3, 1) + 1))
            	temp = 1
			End If
        End If
   
   		if temp = 0 then
   	'/* analisis del ultimo digito */
        digito = Cdbl(Mid(xnum, 3, 1))
            	If ((c = 1) Or (c = 2)) And xnum = "001" Then
                	letras = letras & "un"
            	Else
                	If ((c = 1) Or (c = 2)) And xnum >= "020" And Mid(xnum, 3, 1) = "1" Then
                    	letras = letras & "un"
                	Else
                    	If digito <> 0 Then
                        	letras = letras & xuni(digito)
                    	End If
                	End If
            	End If
		end if
  
  If j = 1 And i = 1 And xnum = "001" And c = 1 Then
    letras = letras & " millon "
  ElseIf j = 1 And i = 1 And xnum <> "001" And c = 1 Then
    letras = letras & " millones "
  ElseIf j = 1 And i = 4 And c = 2 Then
    letras = letras & " mil "
  End If
  i = i + 3
  c = c + 1
  Loop
  If letras = "" Then
  letras = "cero "
  End If
  If decimales <> 0 Then
    decimales = Round(decimales)
    
    letras = letras & " con " & CStr(decimales) & "/100"
  End If
  
End If

numeroALetras = letras
End Function
'-------------------------------------------------------------------------------------------------------------
'/**
' * Funcion    : GF_EDIT_FOLDER_CTO_GTIA
' * Descripcion: Convierte un string a un formato de Carpeta de Contrato de Garantia
' * Parametros : p_Folder
' *
' * Autor: Nahuel Ajaya
' * Fecha: 08/09/2015
' */
Function GF_EDIT_FOLDER_CTO_GTIA(p_Folder)
   GF_EDIT_FOLDER_CTO_GTIA = ""
   if (Len(Trim(p_Folder)) = 11) then GF_EDIT_FOLDER_CTO_GTIA = Left(Trim(p_Folder),8) &"-"& Mid(Trim(p_Folder),9,1) &"-"& Right(Trim(p_Folder),2)
End Function
'-------------------------------------------------------------------------------------------------------------
' Funcion responsable por formatear la cuenta de NNNNNNNNNNNN a NN-NN-NNNN
'Se asume el formato de cuenta contable MAGIC donde la longitud es fija de 12 y los último 4 son cero.
Function formatCuentaPantalla(cuentaN)
	
	Dim myCta

	myCta = cuentaN
	'Si no tiene 8 o más digitos, se completa con el mínimo necesario.
	if (Len(myCta) < 8) then myCta = GF_nChars(myCta, 8, "0", CHR_AFT)
	myCta = Left(myCta, 8)
	formatCuentaPantalla = Left(myCta, 2) & "-" & mid(myCta, 3, 2) & "-" & Right(myCta, 4)
End Function
%>
