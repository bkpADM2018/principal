<!--#include file="../../Includes/procedimientosUnificador.asp"-->
<!--#include file="../../Includes/procedimientostraducir.asp"-->
<!--#include file="../../Includes/procedimientosfechas.asp"-->
<!--#include file="../../Includes/procedimientosformato.asp"-->
<!--#include file="../../Includes/procedimientosParametros.asp"-->
<!--#include file="../../Includes/procedimientos.asp"-->
<!--#include file="../../Includes/procedimientosExcel.asp"-->
<!--#include file="includes/procedimientosOperativos.asp"-->
<%
Function createFileSegment(pName)
	Dim fs, fadm
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	strPath = Server.MapPath(pName)
	If fs.FileExists(strPath) Then  Call fs.deleteFile(strPath, true)	
	Set fadm = fs.CreateTextFile(strPath)	
	Set fs = nothing
	Set fadm = nothing	
End Function
'*****************************************************************************************
'*	COMIENZO DE PAGINA
'*	ETAPA 1 - GENERACION DEL ARCHIVO DE TEXTO TEMPORAL
'*
'*	Se procesaran los datos entre las fecha de inicio y fin, trabajando de a un día por 
'*	vez. Cada día será considerado un segmento de la información que será almacenado en un 
'*	archivo individual hasta que se completen todos los segmentos y los mismos sean
'*	unificados.
'*****************************************************************************************
Dim rsOp,rtrn
fechaDesdeD = GF_PARAMETROS7("fecContableDS", "", 6)
fechaDesdeM = GF_PARAMETROS7("fecContableMS", "", 6)
fechaDesdeA = GF_PARAMETROS7("fecContableAS", "", 6)
fileCode	  = GF_PARAMETROS7("fileCode", "", 6)
maxSegment	  = GF_PARAMETROS7("maxSegment", 0, 6)
accion		  =	 GF_PARAMETROS7("accion", "", 6)
g_strPuerto	  =	 GF_PARAMETROS7("pto", "", 6)
myOperativo = GF_PARAMETROS7("Operativo", "", 6)
myCartaPorte = GF_PARAMETROS7("CartaPorte", "", 6)
myTurno = GF_PARAMETROS7("turno", "", 6)
myIdVagon = GF_PARAMETROS7("nroVagon", "", 6)
myCdCoordinador = GF_PARAMETROS7("cdCoordinador", "", 6)
myCdCoordinado = GF_PARAMETROS7("cdCoordinado", "", 6)
myCdProducto = GF_PARAMETROS7("cdProducto", "", 6)
myCdCorredor = GF_PARAMETROS7("cdCorredor", "", 6)
myCdVendedor = GF_PARAMETROS7("cdVendedor", "", 6)
myCdEntregador = GF_PARAMETROS7("cdEntregador", "", 6)
myEstado = GF_PARAMETROS7("cmbEstado", 0, 6)
myOrder = GF_PARAMETROS7("myOrder", "" ,6)
Call GF_STANDARIZAR_FECHA(fechaDesdeD, fechaDesdeM, fechaDesdeA)
fechaHastaD = fechaDesdeD
fechaHastaM = fechaDesdeM
fechaHastaA = fechaDesdeA
strName =  "../Temp/OPERATIVOS" & fileCode & ".txt"
Call createFileSegment(strName)
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set arch = fs.OpenTextFile(strPath, 8, true)

myFecContableDesde = fechaDesdeA & fechaDesdeM & fechaDesdeD
myFecContableHasta = fechaHastaA & fechaHastaM & fechaHastaD
rtrn = 0
Set rsOp = loadOperativosPuertos()
rtrn = rsOp.RecordCount
if not rsOp.Eof then	
	while not rsOp.eof		
		strCabecera = Trim(rsOp("SQTURNO"))&";"&Trim(rsOp("CDOPERATIVO"))&";"&GF_EDIT_CTAPTE(Trim(rsOp("CARTAPORTE")))&";"&GF_FN2DTE(rsOp("DTINICIO"))&_
					  ";"&Trim(rsOp("DSCLIENTE"))&";"&Trim(rsOp("DSPRODUCTO"))&";"&Trim(rsOp("DSCORREDOR"))&_
					  ";"&Trim(rsOp("DSVENDEDOR"))&";"&Trim(rsOp("DSESTADO"))
		arch.WriteLine(strCabecera)
		rsOp.MoveNext()		
	wend
	'if (Len(strCabecera) > 0) then strCabecera = left(strCabecera,Len(strCabecera)-1)
	
end if
arch.close()
Set arch = Nothing
%>
<HTML>
<HEAD>
<script type="text/javascript">
	parent.generateSegment_callback();
</script>
</HEAD>
<BODY>
<P>&nbsp;</P>
</BODY>
</HTML>