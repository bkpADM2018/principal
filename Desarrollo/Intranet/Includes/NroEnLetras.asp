<%
function NroEnLetras(byval curNumero , blnO_Final)
'Devuelve un n�mero expresado en letras.
'El par�metro blnO_Final se utiliza en la recursi�n para saber si se debe colocar
'la "O" final cuando la palabra es UN(O)
    Dim dblCentavos 
    Dim lngContDec 
    Dim lngContCent 
    Dim lngContMil 
    Dim lngContMillon 
    Dim strNumLetras 
    Dim strNumero 
    Dim strDecenas 
    Dim strCentenas 
    Dim blnNegativo 
    Dim blnPlural 

    If Int(curNumero) = 0 Then 'Aca saque un # despuesdel cero
        strNumLetras = "CERO"
    End If

    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0 Then 'Aca saque otro #
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000 'Aca saque otro #
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000
    Loop

    Do While curNumero >= 1000 'Aca saque otro #
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000 'Aca saque otro #
    Loop

    Do While curNumero >= 100 'Aca saque otro #
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100 'Aca saque otro #
    Loop

    If Not (curNumero > 10 And curNumero <= 20) Then 'Aca saque dos #
        Do While curNumero >= 10 'Aca seque otro #
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10 'Aca saque otro #
        Loop
    End If

    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el n�mero es >1000000 usa recursividad
            strNumLetras = NroEnLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON"
        If blnPlural Then      'IIf(blnPlural, "ES ", " ")
           blnPlural = "ES "
        else
            blnPlural = " "
        End If
        strNumLetras = strNumLetras & blnPlural
    End If

    If lngContMil > 0 Then
        If lngContMil >= 1 Then   'si el n�mero es >100000 usa recursividad
            strNumLetras = strNumLetras & NroEnLetras(lngContMil, False)
            lngContMil = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
    End If

    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0 Then 'Aca saque otro # del ultimo cero
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If

    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If

        If lngContDec >= 3 And curNumero > 0 Then 'Aca saque otro # del ultimo cero
            strNumLetras = strNumLetras & " Y "
        End If
    Else
        If curNumero >= 0 And curNumero <= 20 Then 'Aca saque dos #
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1 And blnO_Final Then 'Aca saque otro #
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0 Then 'Aca saque otro #
                strNumLetras = Trim(strNumLetras) & " CON " & FormatNumber(CInt(dblCentavos * 100), "00") & "/100" 'Aca saque otro # del 100 que esta en rojo
            End If
            NroEnLetras = strNumLetras
            Exit Function
        End If
    End If

    If curNumero > 0 Then 'Aca saque otro #
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1 And blnO_Final Then 'Aca saque otro #
            strNumLetras = strNumLetras & "O"
        End If
    End If

    If dblCentavos > 0 Then 'Aca saque otro #
        strNumLetras = strNumLetras & " CON " + FormatNumber(CInt(dblCentavos * 100), "00") & "/100" 'Aca saque otro # del 100 que esta en rojo
    End If

    If blnNegativo Then
        NroEnLetras = "(" & strNumLetras & ")"
    else
        NroEnLetras = strNumLetras
    End If
  'NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
End Function
%>

