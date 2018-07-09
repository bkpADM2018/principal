<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosAlmacenes.asp"-->
<!--#include file="Includes/procedimientosVales.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosTraducir.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<%
dim accion, fileno
accion = GF_Parametros7("accion","",6)
fileno = GF_Parametros7("fileno",0,6)
select case accion
	case ACCION_ARCHIVO_LECTURA	
		Call ActualizarArchivoLectura(fileno)		
end select 


' Autor: 	Ajaya Nahuel - CNA
' Fecha: 	20/12/2011
' Objetivo:	
'			Guarda en la base de datos la fecha de lectura y usuario que abrio el archivo de un pedido de precio
' Parametros:
'			pIdCotizacion[int]
' Devuelve: 
'			---
Function ActualizarArchivoLectura(pIdCotizacion)
	Dim strSQL

	strSQL = "Update TBLPCTCOTIZACIONES set FECHALECTURA = " & session("MmtoSistema") &", CDUSRLECTURA ='" & session("Usuario") & "' where IDCOTIZACION =" & pIdCotizacion
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "UPDATE", strSQL)
End function
%>