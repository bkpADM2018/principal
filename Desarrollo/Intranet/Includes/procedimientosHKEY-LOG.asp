<!--#include file="procedimientosLog.asp"-->
<%
Dim hk_sBox1, hk_sBox2, hk_sBox3, hk_sPassword, hk_sClassID, hk_data, hk_status, hk_serial, hk_type, hk_hkey_type

Const HKEY_4K_TYPE = "4K"
Const HKEY_8K_TYPE = "8K"

Const HK_NO_KEY = "00000000"

Const HK_MSG_VALID = "HK0"
Const HK_MSG_NOT_VALID = "HK1"
Const HK_ERR_NO_KEY = "HK2"
Const HK_ERR_UNKNOWN = "HK3"

Const HK_TYPE_USER = "002"
Const HK_TYPE_ADMIN = "001"

hk_hkey_type= HKEY_4K_TYPE

hk_sBox1 = array(&hf3,&hcf,&ha7,&h86,&he3,&h44,&hdb,&h33,&hcb,&hb4,&hff,&h78,&h2b,&h72,&h02,&h4b, _
&h53,&h52,&ha9,&h15,&hab,&h36,&h2d,&hd2,&h23,&haa,&h9d,&hc4,&h6b,&hde,&he1,&hc3, _
&h07,&h4a,&hb3,&hca,&h87,&h1c,&h65,&h0d,&h83,&hcc,&h63,&hd0,&h3b,&h3f,&h95,&hf8, _
&h1f,&h60,&hb5,&h50,&h43,&h5a,&hed,&h89,&hfb,&h6f,&h57,&h45,&h73,&h1e,&h10,&h88, _
&h22,&h4c,&haf,&h69,&h5b,&h61,&ha5,&h3c,&h28,&hbc,&h30,&hf4,&h06,&hf6,&h08,&hb2, _
&h3e,&h37,&h93,&h3a,&h21,&h3d,&h77,&hb6,&h41,&hc6,&hb9,&h70,&h39,&hf9,&h75,&h9f, _
&h40,&h42,&h58,&h67,&heb,&ha1,&h1a,&h76,&h00,&hef,&h99,&hc5,&h14,&h79,&h4e,&h2c, _
&h54,&h7e,&hc9,&h0a,&h62,&h5d,&h26,&h8f,&h04,&h51,&h48,&he7,&h5f,&h7c,&hf1,&hbd, _
&h68,&hb8,&hd9,&he5,&hd7,&h49,&h25,&hfe,&h74,&h7a,&h2a,&h92,&h0e,&h56,&h8d,&h05, _
&h8e,&hdf,&hd5,&hdc,&h90,&h7f,&hd3,&hb1,&h16,&h11,&h27,&h4f,&h84,&h38,&h13,&h98, _
&h96,&ha0,&hbb,&h0f,&h64,&he9,&h9e,&hc7,&ha2,&hfd,&h55,&h0c,&h80,&h7b,&h35,&h91, _
&h94,&hce,&h46,&hd8,&h34,&hd1,&h03,&h01,&hac,&h09,&hae,&hec,&had,&h8c,&h66,&h2f, _
&h81,&h29,&h2e,&h6d,&h5c,&hf0,&h8a,&hdd,&h0b,&hd6,&h32,&he0,&hc8,&h59,&hf7,&h6a, _
&hcd,&hc2,&hb0,&hb7,&ha4,&h18,&h5e,&hbf,&h7d,&h85,&h97,&h24,&hba,&h82,&h6c,&he4, _
&hc0,&hea,&hbe,&hfc,&he2,&h19,&hda,&h17,&hd4,&he8,&h9b,&h9c,&h1b,&he6,&ha8,&h71, _
&h6e,&h20,&ha6,&h12,&hee,&h47,&hf5,&hfa,&ha3,&hf2,&h8b,&h9a,&h4d,&h1d,&hc1,&h31)

hk_sBox2 = array(&h73,&h3d,&h9d,&h91,&hf3,&hc5,&h53,&h87,&h7b,&ha9,&h1b,&h82,&h97,&h58,&h29,&h4e, _
&h00,&heb,&h85,&hcf,&h67,&he6,&h51,&h70,&h2b,&hbc,&hf9,&h6e,&hfb,&h50,&h0e,&hef, _
&hb3,&h45,&h12,&h47,&h1c,&h5a,&h04,&hbe,&h26,&h24,&hcb,&h8b,&h5b,&h93,&he1,&h60, _
&h33,&h43,&had,&h48,&hd7,&hfe,&he3,&h17,&h10,&ha0,&h05,&h4f,&h31,&ha4,&ha5,&h8e, _
&h28,&hda,&hf1,&h84,&h3b,&hcc,&h2d,&h66,&h36,&hc4,&h34,&hea,&h2a,&h09,&hcd,&h75, _
&h40,&he7,&hd9,&hb6,&h16,&h49,&hdd,&h15,&h20,&h7f,&ha7,&h9b,&h0c,&h1d,&h6f,&h88, _
&h0d,&hac,&h08,&h61,&ha3,&h4a,&h65,&h37,&h11,&h07,&h89,&h52,&hb7,&h38,&h69,&h3e, _
&h2f,&hd3,&h4d,&h7e,&h23,&he4,&h6b,&h55,&h68,&h80,&h18,&hb5,&h25,&hf8,&hb9,&h9e, _
&h5d,&h41,&h1f,&h1a,&h71,&hd4,&hb1,&h9a,&h1e,&ha1,&h8f,&he8,&h46,&hab,&h62,&h96, _
&h8c,&hfa,&h32,&h9f,&h74,&h06,&h14,&h94,&h5f,&h54,&h98,&hfc,&hdb,&h6a,&hbd,&h4b, _
&h3c,&hbf,&hc9,&hd5,&h4c,&h0a,&h78,&ha2,&h8d,&hb8,&h92,&hae,&h30,&hb4,&hed,&haa, _
&h39,&hdf,&h64,&h3a,&h0b,&hf0,&h5e,&hb2,&hc3,&hc6,&h9c,&h42,&ha6,&h7c,&haf,&h83, _
&h79,&h86,&hbb,&hd8,&h5c,&h35,&h13,&he9,&h63,&h7a,&hc8,&h21,&hff,&h99,&hc7,&h77, _
&hc0,&h90,&h76,&hf7,&hb0,&h8a,&hfd,&hd0,&hca,&h7d,&he5,&h6d,&h27,&hee,&h19,&hc2, _
&hba,&h81,&hf5,&hd2,&hdc,&hec,&h22,&h3f,&h95,&hf6,&h44,&hc1,&h57,&h2c,&hd6,&h6c, _
&he0,&h2e,&hde,&hf4,&h59,&h01,&he2,&h02,&h0f,&ha8,&h56,&h72,&hce,&hd1,&hf2,&h03)

hk_sBox3 = array(&h69,&h42,&h62,&h7d,&h20,&hb3,&h41,&h30,&hc8,&h01,&h1b,&h2c,&h95,&ha1,&ha5,&hd4,&h47,&ha4,&h6a,&h61, _
&hd4,&h1e,&h6b,&haa,&hd6,&h29,&ha5,&h31,&h36,&h7d,&h71,&h87,&h4a,&h77,&haf,&ha0,&h3b,&h00,&h9a,&hf7, _
&hdb,&h54,&h02,&hee,&h58,&h18,&h6d,&h32,&h5b,&h70,&h80,&hd5,&h13,&h32,&hc6,&h13,&h59,&hf9,&hb9,&hee, _
&hee,&hd5,&hd9,&h2b,&h4d,&h57,&hd7,&hc3,&ha9,&h47,&hab,&hde,&h17,&h0d,&hec,&h5a,&h74,&h16,&h8f,&h93, _
&h92,&h21,&h8f,&h75,&h6e,&h84,&h7e,&h6f,&h91,&h5a,&h3b,&h4c,&h77,&h03,&hdc,&ha7,&h47,&h8f,&h0d,&h6a, _
&hfd,&h7b,&hdb,&hde,&hae,&h3c,&hcd,&h96,&h83,&h60,&h29,&h91,&hc0,&h16,&h46,&hdd,&hb5,&h50,&h6c,&h1a, _
&h78,&ha9,&hec,&h1f,&h44,&h5b,&h9d,&h38,&h23,&hd0,&hd8,&h9d,&h3f,&hc2,&h20,&h33,&h8b,&h48,&h0f,&h2a, _
&h20,&hdb,&h36,&h8b,&hea,&h4c,&h96,&hd7,&ha6,&h75,&h41,&h31,&h2b,&hd6,&h65,&hc3,&h07,&hd5,&hd0,&h5c, _
&h80,&hd1,&h96,&he9,&he5,&h40,&h82,&h5f,&h4b,&h7b,&h72,&h10,&h22,&h95,&hb8,&ha2,&hc9,&h39,&hc6,&h33, _
&h76,&h79,&h12,&h1b,&h22,&h94,&h11,&h1a,&hdf,&h91,&hbe,&hd7,&h34,&hb8,&h76,&h61,&he0,&hca,&ha3,&ha6)

hk_sPassword = "YpWiGtBZwbmjwaAZ"

hk_sClassID = "75F54931-E18B-43F7-A2D4-E396B57374F3"
  
hk_type=""

Call startLog(HND_FILE, MSG_DBG_LOG)
set myLog = new classLog
myLog.fileName = "HARDKEY-PRODUCCION.log"
myLog.path = "logs"	
myLog.debug("INICIANDO TEST!")
'------------------------------------------------------------------------------
'Funcion a llamar para obtener los datos necesarios para operar el componente que lee la HARD KEY
'Devuelve los parametros que hay que setear en la propiedad strBuffer del componenete de la llave.
Function HKEY()
	Dim sParameters, sHex
	
	sParameters = HK_buildParameterChain()		
	myLog.debug("Cadena Enviada Encriptada --> " & HK_encrypt(sParameters, hk_sPassword))
	HKEY = HK_encrypt(sParameters, hk_sPassword)		
End Function
'------------------------------------------------------------------------------
'Se arma la cadena de parametros a mandar a la llave en su inicialización.
Function HK_buildParameterChain()
	if (hk_hkey_type= HKEY_4K_TYPE) then
		HK_buildParameterChain = HK_buildParameterChain_4K()
	else
		HK_buildParameterChain = HK_buildParameterChain_8K()
	end if
End Function
'------------------------------------------------------------------------------
Function HK_binary2Hex(sBinary)
	Dim sReturn, c, i
	
	sReturn = ""		
    For i = 1 To Len(sBinary)
        c = Asc(Mid(sBinary, i, 1))
        If c > 15 Then
            sReturn = sReturn & Hex(c)
        Else
            sReturn = sReturn & "0" & Hex(c)
        End If
    Next 
    HK_binary2Hex = sReturn
End Function
'------------------------------------------------------------------------------
' Convierta cada valor en hexadecimal que viene como dos
' caracters ascii para obtener un vector de caracteres en binario.
Function HK_hex2Binary(sHex)
	dim sAux, i
	sAux = ""
	For i = 0 To 199
		sAux = sAux + Chr(HK_GetHex(Mid(sHex, (i * 2) + 1, 2)))
	Next 
	HK_hex2Binary = sAux
End Function
'------------------------------------------------------------------------------
' Esta rutina devuelve un valor binario a partir un número en
' hexadecimal expresado como dos caracteres ascii.
Function HK_GetHex(sTemp) 'As Byte
Dim i 'As Integer
Dim c 'As Byte
Dim nTemp 'As Integer
    nTemp = 0
	if( len(sTemp)>1 ) then
	    For i = 1 To 2
	        nTemp = nTemp * 16
	        c = Asc(Mid(sTemp, i, 1))
	        If (c >= Asc("0") And c <= Asc("9")) Then
	            nTemp = nTemp + c - Asc("0")
	        End If
	        If (c >= Asc("a") And c <= Asc("f")) Then
	            nTemp = nTemp + c - Asc("a") + 10
	        End If
	        If (c >= Asc("A") And c <= Asc("F")) Then
	            nTemp = nTemp + c - Asc("A") + 10
	        End If
	    Next 
	    nTemp = nTemp Mod 256
	end if
    HK_GetHex = nTemp
End Function
'------------------------------------------------------------------------------
' Esta rutina encripta la cadena original antes de pasarla al 
' componente ActiveX
Function HK_encrypt(buffer , password)
    
	if (hk_hkey_type= HKEY_4K_TYPE) then
		HK_encrypt = HK_encrypt_4K(buffer , password)
	else
		HK_encrypt = HK_encrypt_8K(buffer , password)
	end if
	
End Function
'------------------------------------------------------------------------------
' Esta rutina desencripta la cadena que devuelve el control ActiveX
Function HK_decrypt(buffer , password)
    
	if (hk_hkey_type= HKEY_4K_TYPE) then
	'Response.Write "IN(" & buffer & ")"
		HK_decrypt = HK_decrypt_4K(HK_hex2Binary(buffer) , password)
	else		
		HK_decrypt = HK_decrypt_8K(buffer , password)
	end if
	
End Function
'------------------------------------------------------------------------------
'Funcion que recibe los datos enviados por el cliente.
Function HK_receiveData()

	Dim sPlain, ctl, baseIndex
	
	baseIndex = 18
	if (hk_hkey_type = HKEY_4K_TYPE) then baseIndex = baseIndex + 2
	hk_status = false
	hk_data = Request.QueryString("HK")				
	if (Len(hk_data) = 0) then hk_data = Request.Form("HK")	
	if (Len(hk_data) > 0) then	
		'Hay datos, se procesan para que esten listos par usarse
		'Obtengo la respuesta a través de un control tipo TextBox		
		myLog.debug("Cadena Recibida Encriptada--> " & hk_data)
		sPlain = HK_decrypt(hk_data, hk_sPassword)	
		myLog.debug("Cadena Recibida --> " & sPlain)	
		'response.write "OUT(" & sPlain & ")<br>"		
		'hk_serial= HK_NO_KEY
		'ctl = mid(sPlain, baseIndex + 21, 3)		
		'if ((ctl = HK_TYPE_USER) or (ctl = HK_TYPE_ADMIN)) then
			'Llave OK						
			hk_serial = mid(sPlain, baseIndex, 8)
		'	hk_type = ctl 			
			hk_status=true
		'else			
		'	HK_automaticResponse(ctl)
		'end if
		if (session("Usuario") <> "XXX") then
		    '**************** PARCHE *******************
		    'Si enchufó una llave, se toma la que corresponde de la base de datos.
		    hk_status=false
		    strSQL="Select * from TOEPFERDB.TBLREGISTROFIRMAS where CDUSUARIO='" & session("Usuario") & "'"
		    Call executeQuery(rs, "OPEN", strSQL)
		    if (not rs.eof) then
		        hk_serial = rs("HKEY")
		        hk_status=true
		    end if    		
		    '************** FIN PARCHE *****************
		else
		    response.Write sPlain
		end if
	end if

End Function
'------------------------------------------------------------------------------
Function HK_isUser()
	HK_isUser (hk_type = HK_TYPE_USER)
End Function
'------------------------------------------------------------------------------
Function HK_isAdmin()
	HK_isUser (hk_type = HK_TYPE_ADMIN)
End Function
'------------------------------------------------------------------------------

Function HK_sendResponse(payload)
	Response.Write HK_MSG_VALID & "|" & payload
	Response.End
End Function
'------------------------------------------------------------------------------
Function HK_automaticResponse(ctl)
	Select case ctl
		case "000"
			Response.Write HK_ERR_NO_KEY & "|No hay llave conectada."
		case "212"
			Response.Write HK_MSG_NOT_VALID & "|Tarjeta no inicializada." & vbcrlf & "Contacte al administrador."
		case else
			Response.Write HK_ERR_UNKNOWN & "|Valor(" & ctl & ")."
	End Select
	Response.End
End Function
'------------------------------------------------------------------------------
'Funcion que permite saber si hay datos de una llave para trabajar del lado del servidor.
Function HK_isKeyReady()
	HK_isKeyReady = hk_status
End Function
'------------------------------------------------------------------------------
Function HK_readKey() 
	HK_readKey = hk_serial
End Function	
'------------------------------------------------------------------------------
Function HK_response(value)
	Response.Write value
End Function
'*************************************************************************************************************
'	4K UNIT FUNCTIONS
'*************************************************************************************************************
' Aquí se arma la cadena parámetro. Los primeros 10 bytes deben ser 
' valores random, luego un espacio y 8 caracteres en cero (ascci) 
' reservado para el nro de conexión, un espacio y 5 caracteres con
' la clave1 (en ascii), otro espacio y 5 caracteres con la clave2.
' El resto de los parámetros no son importantes para este caso y 
' deben dejarse en cero. Para una descripción más detallada de como
' se compone la cadena de parámetro, ver el capítulo "Interface de
' programación de aplicación (API)" del manual del usuario. 	
' Completar la cadena con espacio en blanco hasta llegar a los 200
' caracteres requeridos por la interfaz.
Function HK_buildParameterChain_4K()
	dim sOriginal, sAux, i
	
	randomize
	sOriginal = "" 
	For i = 1 to 10
		sOriginal = sOriginal + chr(int(rnd *255)+1)
	next
	sOriginal = sOriginal + " 00000 00000 00000000 0000 00 0"	
	sOriginal = "0000000000 00000 00000 0000 00 0"
	sOriginal = sOriginal + space(200)
	myLog.debug("Cadena Enviada --> " & sOriginal & "<--")	
	
	HK_buildParameterChain_4K = sOriginal
End Function
'------------------------------------------------------------------------------
' Esta rutina encripta la cadena original antes de pasarla al 
' componente ActiveX
Function HK_encrypt_4K(buffer , password)
    Dim i           'As Integer
    Dim ctemp       'As Integer
    Dim cAnterior   'As Integer
    Dim k           'As Integer
    Dim pw          'As Integer
    Dim bufEnc      'As String
	
	myLog.debug("CodePage:" & session.CodePage)	
	for i = 0 to 255
	    myLog.debug(i & "=" & chr(i))
	next
	tmpstr=""
	for i=LBound(hk_sBox1)+1 to UBound(hk_sBox1)
	    if (tmpstr <> "") then tmpstr = tmpstr & ","
		tmpstr = tmpstr & CStr(hk_sBox1(i))
	next	
	myLog.debug("sBOX1 --> " & tmpstr)	
	tmpstr=""
	for i=LBound(hk_sBox2)+1 to UBound(hk_sBox2)
	    if (tmpstr <> "") then tmpstr = tmpstr & ","
		tmpstr = tmpstr & CStr(hk_sBox2(i))
	next
	myLog.debug("sBOX2 --> " & tmpstr)	
	myLog.debug("Password --> " & password)	
	
    cAnterior = 0
    bufEnc = ""
    For i = 0 To 199
        ctemp = Asc(Mid(buffer, i + 1, 1))		
        jas1 = "1:" & Mid(buffer, i + 1, 1)                
        If (ctemp < 0) Then
            ctemp = ctemp + 256
        End If
        jas2 = "2:" & ctemp
        ctemp = ctemp Xor hk_sBox1(cAnterior)
        jas3 = "3:" & ctemp
        For k = 0 To 15
            pw = Asc(Mid(password, k + 1, 1))
            jas4 = "4:" & Mid(password, k + 1)
            jas5 = "5:" & pw            
            If ((k Mod 2) = 1) Then
                ctemp = ctemp Xor hk_sBox1(hk_sBox2(pw))
                jas6A = "6A:" & ctemp            
                ctemp = hk_sBox2(ctemp)
                jas7A = "7a:" & ctemp            
            Else
                ctemp = ctemp Xor hk_sBox2(hk_sBox1(pw))
                jas6B = "6B:" & ctemp            
                ctemp = hk_sBox1(ctemp)
                jas7B = "7B:" & ctemp         
            End If
        Next
        ctemp = ctemp Xor hk_sBox1(i)
        jas8 = "8:" & ctemp
        cAnterior = ctemp
        jas9 = "9:" & Chr(ctemp)
        bufEnc = bufEnc + Chr(ctemp)
        myLog.debug("Detalle -->" & jas1 & "|" & jas2 & "|" & jas3 & "|" & jas4 & "|" & jas5 & "|" & jas6A & "|" & jas6B & "|" & jas7A & "|" & jas7B & "|" & jas8 & "|" & jas9)
        jas1 = ""
        jas2 = ""
        jas3 = ""
        jas4 = ""
        jas5 = ""
        jas6A = ""
        jas7A = ""
        jas8 = ""
        jas9 = ""
        jas6B = ""
        jas7B = ""                
    Next        
    myLog.debug("Cadena Enviada Binaria --> " & bufEnc & "<--")
	HK_encrypt_4K = HK_binary2Hex(bufEnc)
End Function
'------------------------------------------------------------------------------
' Esta rutina desencripta la cadena que devuelve el control ActiveX
Function HK_decrypt_4K(buffer , password)
    Dim i           'As Integer
    Dim ctemp       'As Integer
    Dim cAnterior   'As Integer
    Dim k           'As Integer
    Dim pw          'As Integer
    Dim bufEnc      'As String
    cAnterior = 0
    bufEnc = ""
    For i = 0 To 199
        ctemp = Asc(Mid(buffer, i + 1, 1))
        If (ctemp < 0) Then
            ctemp = ctemp + 256
        End If
        ctemp = ctemp Xor hk_sBox1(cAnterior)
        For k = 0 To 15
            pw = Asc(Mid(password, k + 1, 1))
            If ((k Mod 2) = 1) Then
                ctemp = ctemp Xor hk_sBox1(hk_sBox2(pw))
                ctemp = hk_sBox2(ctemp)
            Else
                ctemp = ctemp Xor hk_sBox2(hk_sBox1(pw))
                ctemp = hk_sBox1(ctemp)
            End If
        Next
        ctemp = ctemp Xor hk_sBox1(i)
        cAnterior = Asc(Mid(buffer, i + 1, 1))
        bufEnc = bufEnc + Chr(ctemp)
    Next        
    HK_decrypt_4K = bufEnc	
End Function
'*************************************************************************************************************
'	8K UNIT FUNCTIONS
'*************************************************************************************************************
Function HK_buildParameterChain_8K()		
	Dim ret
	
	HK_buildParameterChain_8K = "0000000000 00000 00000 0000" & space(1024)
End Function
'------------------------------------------------------------------------------
' Esta rutina encripta la cadena original antes de pasarla al 
' componente ActiveX
Function HK_encrypt_8K(buffer , sPassword)
     Dim i               'As Integer
        Dim k               'As Integer
        Dim ctemp           'As Integer
        Dim cAnterior       'As Integer
        Dim pw              'As Integer
        Dim bufEnc          'As String
        Dim x
	    Dim bufHEX
        Dim s
        dim aux
                
        bufEnc = ""
        cAnterior = 0
        For i = 0 To 1023
            ctemp = Asc(Mid(buffer, (i \ 2) + 1, 1))
            If (ctemp < 0) Then
                ctemp = ctemp + 256
            End If
            ctemp = ctemp Xor hk_sBox1(cAnterior)
            For k = 0 To 15
                pw = Asc(Mid(sPassword, k + 1, 1))
                If ((k Mod 2) = 1) Then
                    ctemp = ctemp Xor hk_sBox1(hk_sBox2(pw))
                    ctemp = hk_sBox2(ctemp)
                Else
                    ctemp = ctemp Xor hk_sBox2(hk_sBox1(pw))
                    ctemp = hk_sBox1(ctemp)
                End If
            Next
	        ctemp = ctemp Xor hk_sBox1((i \ 2) Mod 256)
	        cAnterior = ((ctemp * (i Mod 2)) + (cAnterior * ((i + 1) Mod 2)))
	        x = ((i Mod 2) * hk_sBox3((i \ 2) Mod 200)) Xor ctemp                			
			s = HEX(x)
	        if len(s) = 1 then s = "0" & s	        
	        bufEnc = bufEnc & s
        Next
		HK_encrypt_8K = bufEnc
End Function      
'------------------------------------------------------------------------------
' Esta rutina desencripta la cadena que devuelve el control ActiveX
Function HK_decrypt_8K(buffer , sPassword)
    Dim i 			'As Integer
    Dim ctemp 		'As Integer
    Dim cAnterior 	'As Integer
    Dim k 			'As Integer
    Dim pw 			'As Integer
    Dim bufEnc 		'As String
    dim aux1
    cAnterior = 0
    bufEnc = ""
    For i = 0 To len(buffer) - 4
        ctemp = HK_GetHex(Mid(buffer, (i * 2) + 1, 2))		     
        aux1 = ctemp
        If (ctemp < 0) Then
            ctemp = ctemp + 256
        End If
        ctemp = ctemp Xor hk_sBox1(cAnterior)
        For k = 0 To 15
            pw = Asc(Mid(sPassword, k + 1, 1))
            If ((k Mod 2) = 1) Then
                ctemp = ctemp Xor hk_sBox1(hk_sBox2(pw))
                ctemp = hk_sBox2(ctemp)
            Else
                ctemp = ctemp Xor hk_sBox2(hk_sBox1(pw))
                ctemp = hk_sBox1(ctemp)
            End If
        Next
        ctemp = ctemp Xor hk_sBox1(i Mod 256)
        cAnterior = HK_GetHex(Mid(buffer, (i * 2) + 1, 2))
        bufEnc = bufEnc + Chr(ctemp)        
    Next
    HK_decrypt_8K = bufEnc
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Lineas que se ejecutan siempre para detectar si se submitieron datos a controlar de una llave
Call HK_receiveData()
%>
