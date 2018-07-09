<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="include/procedimientoProducto.asp"-->
<%
'-------------------------------------------------------------------------------------------------------------------
Dim cdProducto,accion,g_strPuerto

cdProducto = GF_Parametros7("cdProducto",0,6)
g_strPuerto = GF_Parametros7("pto","",6)
accion = GF_Parametros7("accion","",6)

if (accion = ACCION_BORRAR) then 

    Call deleteProducto(cdProducto,g_strPuerto)
    Call deleteAtributo(cdProducto,0,g_strPuerto)			
    Call deleteCosecha(cdProducto,0,g_strPuerto)
    Call EliminarBiotecnologiaDeProducto(0,cdProducto,g_strPuerto)

end if
%>