
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosXML.asp"-->
<!--#include file="Includes/includeGeneracionArchivos.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosTitulos.asp"-->
<%

const PATH_DESTINO = "\ActisaIntra\Temp\"

'******************************************************************************************
sub crearObjetoToepfer(byref p_obj)

set p_obj = server.createObject("Scripting.Dictionary")

p_obj.add "CodProd", ""
p_obj.add "CodSuc", ""
p_obj.add "CodOpe", ""
p_obj.add "NroContrato", ""
p_obj.add "Cosecha", ""

end sub
'******************************************************************************************
sub crearObjetoDetalles(byref p_obj)

set p_obj = server.createObject("Scripting.Dictionary")

p_obj.add "Comision", ""
p_obj.add "Producto", ""
p_obj.add "FechaConcertacion", ""
p_obj.add "CantCamiones", ""
p_obj.add "Moneda", ""
p_obj.add "Transporte", ""
p_obj.add "FecEntregaDesde", ""
p_obj.add "FecEntregaHasta", ""
p_obj.add "Destino", ""
p_obj.add "FecPago", ""
p_obj.add "PorcPago", ""
p_obj.add "MercadProd", ""
p_obj.add "UnidadMedidaProducto", ""
p_obj.add "CantidadDesde", ""
p_obj.add "CantidadHasta", ""
p_obj.add "Precio", ""
p_obj.add "UnidadMedidaPrecio", ""
p_obj.add "NumeroSioGranos", ""
end sub
'*******************************************************************************************
sub crearObjetoContraParte(byref p_obj)

set p_obj = server.createObject("Scripting.Dictionary")

p_obj.add "V/C", ""
p_obj.add "CUIT", ""
p_obj.add "Contrato", ""

end sub
'*******************************************************************************************
function getToepferCUIT()
         getToepferCUIT = 30621973173
end function
'********************************************************************************************
sub obtenerDatosXML(p_oXML, byref p_datosToepfer, byref p_datosContraParte, byref p_contratoCorredor, byref p_datosDetalles)

'Obtengo datos de Toepfer del XML
call obtenerDatosToepferXML(p_oXML, p_datosToepfer)

'Obtengo Datos de la Contra Parte del XML
call obtenerDatosContraParteXML(p_oXML, p_datosContraParte)

'Obtengo Contrato del Corredor del XML
call obtenerDatosCorredorXML(p_oXML, p_contratoCorredor)

'Obtengo Datos del Detalle del XML
call obtenerDatosDetallesXML(p_oXML, p_datosDetalles)

end sub
'********************************************************************************************
sub obtenerDatosToepferXML(p_oXML, byref p_datosToepfer)
dim aux, aux2, index

index = 0
while isNode(GF_getNode(p_oXML, "Partes/" & index)) and cstr(GF_getPathValue(p_oXML,"Partes/" & index & "/CUIT"))<>cstr(getToepferCUIT())
	index = index + 1
wend
p_datosToepfer("NroContrato") = cstr(GF_getPathValue(p_oXML, "Partes/" & index & "/NroContratoInterno"))

end sub
'********************************************************************************************
sub obtenerDatosContraParteXML(p_oXML, byref p_datosContraParte)
dim aux, parte, index

parte = 0
while (isNode(GF_getNode(p_oXML, "Partes/" & parte)) and ( lcase(cstr(GF_getPathAttribute(p_oXML, "Partes/" & parte, "Caption")))="corredor" and (GF_getPathValue(p_oXML, "Partes" & parte & "/CUIT")=getToepferCUIT())))
	parte = parte + 1
wend

'Cargo los Datos de la contraparte con la parte que corresponda
p_datosContraParte("CUIT") = GF_getPathValue(p_oXML, "Partes/" & parte & "/CUIT")

aux = GF_getPathAttribute(p_oXML, "Partes/" & parte,"Caption")

if ucase(aux)="VENDEDOR" then
   p_datosContraParte("V/C") = "V"
else
    p_datosContraParte("V/C") = "C"
end if

p_datosContraParte("Contrato") = GF_getPathValue(p_oXML,"Partes/" & parte & "/NroContratoInterno")
'response.write p_datosContraParte("CUIT") & "  " & p_datosContraParte("V/C") & "  " & p_datosContraParte("Contrato") & "<br>"
end sub
'********************************************************************************************
sub obtenerDatosCorredorXML(p_oXML, byref p_contratoCorredor)
dim parte

parte = 0
while isNode(GF_getNode(p_oXML, "Partes/" & parte)) and lcase(cstr(GF_getPathAttribute(p_oXML, "Partes/" & parte, "Caption")))<>"corredor"
	parte = parte + 1
wend

p_contratoCorredor = GF_getPathValue(p_oXML, "Partes/" & parte & "/NroContratoInterno")
end sub
'********************************************************************************************
sub obtenerDatosDetallesXML(p_oXML, byref p_datosDetalles)
dim auxFecha

p_datosDetalles("Comision") = GF_getPathValue(p_oXML, "DetalleContrato/ComisionPorComprador")
p_datosDetalles("Producto") = GF_getPathAttribute(p_oXML, "DetalleContrato/Producto", "CodLista")
p_datosDetalles("FechaConcertacion") = GF_getPathValue(p_oXML, "DetalleContrato/FechaConcertacion")
p_datosDetalles("CantCamiones") = GF_getPathValue(p_oXML, "DetalleContrato/CantCamiones")
p_datosDetalles("Moneda") = GF_getPathAttribute(p_oXML, "DetalleContrato/Moneda", "CodLista")
p_datosDetalles("Transporte") = GF_getPathAttribute(p_oXML, "DetalleContrato/MedioTransporte","CodLista")

auxFecha = GF_getPathValue(p_oXML, "DetalleContrato/Entregas/EntregaDesde")
if isDate(auxFecha) then
   p_datosDetalles("FecEntregaDesde") = auxFecha
else
   p_datosDetalles("FecEntregaDesde") = "00/00/0000"
end if

auxFecha = GF_getPathValue(p_oXML, "DetalleContrato/Entregas/EntregaHasta")
if isDate(auxFecha) then
   p_datosDetalles("FecEntregaHasta") = auxFecha
else
    p_datosDetalles("FecEntregaHasta") = "00/00/0000"
end if

p_datosDetalles("Destino") = GF_getPathAttribute(p_oXML, "DetalleContrato/Destino", "CodLista")
p_datosDetalles("FecPago") = GF_getPathValue(p_oXML, "DetalleContrato/Pagos/FechaCondicionPago")
p_datosDetalles("PorcPago") = GF_getPathValue(p_oXML, "DetalleContrato/Pagos/PorcPago")
p_datosDetalles("MercadProd") = GF_getPathAttribute(p_oXML, "DetalleContrato/ProduccionVendedor", "CodLista")
p_datosDetalles("Precio") = GF_getPathValue(p_oXML, "DetalleContrato/Precio")
p_datosDetalles("UnidadMedidaPrecio") = GF_getPathAttribute(p_oXML, "DetalleContrato/UnidadMedidaPrecio", "CodLista")
p_datosDetalles("UnidadMedidaProducto") = GF_getPathAttribute(p_oXML, "DetalleContrato/UnidadMedida", "CodLista")
p_datosDetalles("CantidadDesde") = GF_getPathValue(p_oXML, "DetalleContrato/CantidadDesde")
p_datosDetalles("CantidadHasta") = GF_getPathValue(p_oXML, "DetalleContrato/CantidadHasta")
p_datosDetalles("NumeroSioGranos") = GF_getPathValue(p_oXML, "DetalleContrato/SioGranos/NumeroDeclaracion")
end sub
'********************************************************************************************
sub migrarDatos(p_oXML, p_datosToepfer, p_datosContraParte, p_contratoCorredor, p_datosDetalles)
dim cont

'Migro datos de Toepfer
call migrarDatosToepfer(p_datosToepfer)

'Migro datos de la Contra Parte
call migrarDatosContraParte(p_datosContraParte)

'Migro el contrato del Corredor
call imprimirString(p_contratoCorredor, 15)

'Migro datos del Detalle
call migrarDatosDetalles(p_datosDetalles)

'Migro Clausulas
cont = 0
while (isNode(GF_getNode(p_oXML, "Clausulas/" & cont)))
      arch.write lcase(validarCaracteres(GF_getPathValue(p_oXML, "Clausulas/" & cont)))
      cont = cont + 1
wend

arch.writeline ""
end sub
'********************************************************************************************
sub migrarDatosToepfer(p_datosToepfer)
 
call imprimirString(p_datosToepfer("NroContrato"), 20)

end sub
'********************************************************************************************
sub migrarDatosContraParte(p_datosContraParte)

call imprimirString(p_datosContraParte("V/C"), 1)
call imprimirNumero(p_datosContraParte("CUIT"),11,0)
call imprimirString(p_datosContraParte("Contrato"),15)

end sub
'********************************************************************************************
sub migrarDatosDetalles(p_datosDetalles)

call imprimirNumero(p_datosDetalles("Comision"), 3, 2)
call imprimirNumero(p_datosDetalles("Producto"),3, 0)
call imprimirString(p_datosDetalles("FechaConcertacion"), 10)
call imprimirNumero(p_datosDetalles("CantCamiones"), 2, 0)
call imprimirString(p_datosDetalles("Moneda"), 1)
call imprimirString(p_datosDetalles("Transporte"), 1)
call imprimirString(p_datosDetalles("FecEntregaDesde"), 10)
call imprimirString(p_datosDetalles("FecEntregaHasta"), 10)
call imprimirNumero(p_datosDetalles("Destino"), 2, 0)
call imprimirString(p_datosDetalles("FecPago"), 100)
call imprimirNumero(p_datosDetalles("PorcPago"), 5, 2)
call imprimirString(p_datosDetalles("MercadProd"), 1)
call imprimirNumero(p_datosDetalles("Precio"), 7, 2)
call imprimirString(p_datosDetalles("UnidadMedidaPrecio"), 1)
call imprimirString(p_datosDetalles("UnidadMedidaProducto"), 1)
call imprimirNumero(p_datosDetalles("CantidadDesde"),9,0)
call imprimirNumero(p_datosdetalles("CantidadHasta"),9,0)
call imprimirString(p_datosdetalles("NumeroSioGranos"),20)

end sub
'********************************************************************************************
sub imprimirNumero(p_valor, p_largo, p_dec)
dim aux
dim vec     'vec(0) -> parte entera
            'vec(1) -> parte decimal

if p_valor = "" or p_valor = "UNKNOWN" then p_valor = 0
vec = split(cstr(p_valor),".",-1)

if len(vec(0)) > cint(p_largo - p_dec) then
	'Response.End 
	msg = msg & "Error - numero que excede el tamaño -- " & p_valor & " largo=" & p_largo & " dec=" & p_dec
   response.write "Error - numero que excede el tamaño -- " & p_valor & " largo=" & p_largo & " dec=" & p_dec
else
    while len(vec(0)) < cint(p_largo - p_dec)
          vec(0) = "0" & vec(0)
    wend
end if

if ubound(vec) > 0 then
   aux = vec(1)
else
    aux = "0"
end if

if len(aux) < cint(p_dec) then
   while len(aux) < cint(p_dec)
         aux = aux & "0"
   wend
else
    aux = left(aux,p_dec)
end if

arch.write vec(0) & aux
end sub
'********************************************************************************************
sub imprimirString(p_str, p_largo)
dim aux

if p_str = "UNKNOWN" then
   p_str = ""
end if
 
if len(p_str) < cint(p_largo) then
   aux = len(p_str)
   arch.write p_str
   while cint(aux) < cint(p_largo)
         arch.write " "
         aux = aux + 1
   wend
else
   arch.write left(p_str,cint(p_largo))
end if

end sub
'********************************************************************************************
function validarCaracteres(p_str)
dim i

p_str = replace(p_str, "á", "a")
p_str = replace(p_str, "Á", "A")
p_str = replace(p_str,"é","e")
p_str = replace(p_str,"É","E")
p_str = replace(p_str,"í","i")
p_str = replace(p_str,"Í","I")
p_str = replace(p_str,"ó","o")
p_str = replace(p_str,"Ó","O")
p_str = replace(p_str,"ú","u")
p_str = replace(p_str,"Ú","U")
p_str = replace(p_str,"ñ","n")
p_str = replace(p_str,"Ñ","N")

for i = 1 to 31
    p_str = replace(p_str, chr(i),"")
next

validarCaracteres = p_str
end function

function subirArchivo() 
	ForWriting = 2
	adLongVarChar = 201
	lngNumberUploaded = 0

	'Get binary data from form
	noBytes = Request.TotalBytes
	binData = Request.BinaryRead (noBytes)
	'convert the binary data to a string
	Set RST = CreateObject("ADODB.Recordset")
	LenBinary = LenB(binData)
	if LenBinary > 0 Then
		RST.Fields.Append "myBinary", adLongVarChar, LenBinary
		RST.Open
		RST.AddNew
		RST("myBinary").AppendChunk BinData
		RST.Update
		strDataWhole = RST("myBinary")
	End if
	'Creates a raw data file for with all data sent.
	'Uncomment for debuging.
	'Set fso = CreateObject("Scripting.FileSystemObject")
	'Set f = fso.OpenTextFile(server.mappath(".") & "\raw.txt", ForWriting, True)
	'f.Write strDataWhole
	'set f = nothing
	'set fso = nothing
	'get the boundry indicator
	strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
	lngBoundryPos = instr(1,strBoundry,"boundary=") + 8
	strBoundry = "--" & right(strBoundry,len(strBoundry)-lngBoundryPos)
	'Get first file boundry positions.
	lngCurrentBegin = instr(1,strDataWhole,strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
	Do While lngCurrentEnd > 0
		'Get the data between current boundry an d remove it from the whole.
		strData = mid(strDataWhole,lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
		strDataWhole = replace(strDataWhole,strData,"")
		'Get the full path of the current file.
		lngBeginFileName = instr(1,strdata,"filename=") + 10
		lngEndFileName = instr(lngBeginFileName,strData,chr(34))
		'Make sure they selected at least one file.
		if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then
			msg = "Debe seleccionar un archivo"
			exit do
		End if
		'There could be one or more empty file boxes.
		if lngBeginFileName <> lngEndFileName Then
			strFilename = mid(strData,lngBeginFileName,lngEndFileName - lngBeginFileName)
			'Creates a raw data file with data betwe
			' en current boundrys. Uncomment for debug
			' ing.
			'Set fso = CreateObject("Scripting.FileSystemObject")
			'Set f = fso.OpenTextFile(server.mappath(".") & "\raw_" & lngNumberUploaded & ".txt", ForWriting, True)
			'f.Write strData
			'set f = nothing
			'set fso = nothing

			'Loose the path information and keep jus
			' t the file name.
			tmpLng = instr(1,strFilename,"\")
			Do While tmpLng > 0
				PrevPos = tmpLng
				tmpLng = instr(PrevPos + 1,strFilename,"\")
			Loop

			FileName = right(strFilename,len(strFileName) - PrevPos)
			
			'Get the begining position of the file d ata sent.
			'if the file type is registered with the
			' browser then there will be a Content-Typ
			' e
			lngCT = instr(1,strData,"Content-Type:")

			if lngCT > 0 Then
				lngBeginPos = instr(lngCT,strData,chr(13) & chr(10)) + 4
			Else
				lngBeginPos = lngEndFileName
			End if
			'Get the ending position of the file dat
			' a sent.
			lngEndPos = len(strData)

			'Calculate the file size.
			lngDataLenth = lngEndPos - lngBeginPos
			

		'Controlar de errores
			'Controlar que la extencion de archivo sea .doc, .pdf o .rtf
			vNombreSeparado = split(FileName, ".")
			strNombreTodosArchivos = strNombreTodosArchivos & " " &  FileName
			if (ubound(vNombreSeparado) <> 1) then
			   msg = "Nombre de archivo " & FileName & " es incorrecto<br>"
			else
				if (vNombreSeparado(1) <> "xml") then
					msg = "Tipo de archivo " & FileName & " es incorrecto<br>"
				end if
			end if
			'Controlar longitud de archivo
			if lngDataLenth = 0 then
				msg = msjError & "El archivo " & FileName & " no existe o su tamaño es nulo"
			end if
			'Si hubo errores, sale de ciclo
			if msg <> "" then
				exit do
			end if
			
			'Get the file data
			strFileData = mid(strData,lngBeginPos,lngDataLenth)
			'Create the file.
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set f = fso.OpenTextFile(server.mappath("..") & PATH_DESTINO & FileName, ForWriting, True)
			f.Write strFileData
			Set f = nothing
			Set fso = nothing

			lngNumberUploaded = lngNumberUploaded + 1

		End if

		'Get then next boundry postitions if any
		' .
		lngCurrentBegin = instr(1,strDataWhole,strBoundry)
		lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
	loop
	subirArchivo = server.mappath("..") & PATH_DESTINO & FileName
end function
'--------------------------------------------------------------------------------------
'Copia el archivo txt generado al sistema de archivos del AS400
'NOTA: esta funcion funciona correctamente, abre una conexion al AS400 (NET USE) y copia un archivo al directorio de Interfaz
'       Se dejo de usarse temporalmente debido a algunos problemas en el Interfaz del AS400 de produccion que no logra establecer la conexion
Function copyFileAS400(pPathFile)
    Dim FSO, oShell
	Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oShell = CreateObject("Wscript.Shell")
    oShell.exec("CMD.EXE NET USE "& PATH_AS400_CONFIRMA &" "&session("conn" & CONEXION_AS400 &  "Key")&" /USER:"&session("conn" & CONEXION_AS400 &  "User"))
    oShell.exec("CMD.EXE /C copy " & chr(34) & pPathFile & chr(34) & " " &  PATH_AS400_CONFIRMA & "\Contr.txt")
    Set oShell = nothing
    Set FSO = nothing
End Function 
'********************************************************************************************
'*      Aca empieza la Pagina
'********************************************************************************************
dim indice, datosToepfer, datosContraParte, datosDetalles, contratoCorredor
dim oXMLRaiz, oXML
dim fso, arch
dim txtDestino, txtOrigen, msg, paccion

'Levanto los parametros
paccion = GF_Parametros7("accion", "", 6)

if (trim(paccion)<>"") then   
	txtOrigen = subirArchivo()
	if (msg = "") then
	      'Abro el XML	  	  
		  set oXMLRaiz = GF_openXML(txtOrigen)	  
	      if not isNode(oXMLRaiz) then
	         msg = "Imposible abrir archivo de Origen.<br>Verifique la ruta y que sea un archivo XML."
	      else
	          on Error Resume Next          	  	
		  
	          'Abro el Archivo
	          Set fso = server.createobject("Scripting.FileSystemObject")			
		  
	      '#JAS#	
		  'Obtengo el path del origen	  
		  arrTemp = split(txtOrigen,"\",-1)
		  strPath = left(txtOrigen,len(txtOrigen)-len(arrTemp(UBound(arrTemp))))
		  txtDestino = strPath & "Confirma.txt"
		  '#END JAS#
	          '#EDD#
			  
	          set arch = fso.CreateTextFile(txtDestino)
	          if Err.number = 0 then
	             'creo los Objetos (Dictionary) para los Datos
	             call crearObjetoToepfer(datosToepfer)
	             call crearObjetoContraParte(datosContraParte)
	             call crearObjetoDetalles(datosDetalles)

	             indice=0
	             while (isNode(GF_getNode(oXMLRaiz, indice)))

	                 set oXML = GF_getNode(oXMLRaiz, indice)
	                 
	                 'Obtener Datos del XML
	                 call obtenerDatosXML(oXML, datosToepfer, datosContraParte, contratoCorredor, datosDetalles)

	                 'Pasar Datos al archivo
	                 Call migrarDatos(oXML, datosToepfer, datosContraParte, contratoCorredor, datosDetalles)

	                 indice = indice + 1
	             wend
			
	            'cierro el XML
	            call GF_closeXML(oXML)			
				call arch.close
				Set arch = nothing
                'Call copyFileAS400(txtDestino)
                call Descargar(txtDestino)
                Response.End
	            msg = "Archivo migrado con Exito"
	        else
		  
	             msg = Err.number & "--" & Err.description 
	        end if  
	      end if
	  end if
end if
%>
<HTML>
<HEAD>
      <link rel="stylesheet" href="css/main.css" type="text/css">
</HEAD>
<BODY>
<% call GF_TITULO("TablaMG.gif", GF_Traducir("Migrador de Contratos")) %>
<form ENCTYPE="multipart/form-data" name="form1" method="post" action="INTMigradorCont.asp?accion=run">
      <div class="tableaside size100">
        <h3> Importacion archivo confirma </h3>
         <div id="searchfilter" class="tableasidecontent">
             <div class="col36 reg_header_navdos"> <%=GF_Traducir("Archivo de Origen(XML):")%> </div>
             <div class="col36"><input type="file" name="file1" style="height:100%;padding:0;"></div>
             <input type="submit" value="<%=GF_Traducir("Migrar")%>">
             <div class="col36">
                 <font color=red><%=GF_Traducir(msg)%></font>
             </div>
         </div>
      </div>
</form>
</BODY>
</HTML>

