<!--#include file="procedimientosConexion.asp"-->
<%
'Constantes de estados de camiones
Const CAMIONES_ESTADO_INGRESADO = 1
Const CAMIONES_ESTADO_CALADO    = 2
Const CAMIONES_ESTADO_CALADOCOND= 3
Const CAMIONES_ESTADO_PESADOBRUTO = 5
Const CAMIONES_ESTADO_EGRESADOOK = 6
Const CAMIONES_ESTADO_RECHAZADO = 7
Const CAMIONES_ESTADO_PESADOTARA = 8
Const CAMIONES_ESTADO_CTRLPORTERIA = 9
Const CAMIONES_ESTADO_SINCUPO = 10
Const CAMIONES_ESTADO_DEMORADO = 11
Const CAMIONES_ESTADO_BAJA = 12

Const CIRCUITO_CAMION_TODOS = 0
Const CIRCUITO_CAMION_DESCARGA = 1
Const CIRCUITO_CAMION_CARGA = 2

Const WSCTG_PENDIENTE = 0
Const WSCTG_CONFIRMADO = 1
Const WSCTG_MANUAL = 2
Const WSCTG_EXENTO = 3
Const WSCTG_QUITADO = 4

Const DEVICE_CODE_AFIP = 3

'-------------------------------------------------------------------------------------------------
'Autor: Ezequiel A. Bacarini
'Fecha: 05/01/2007
Function IIF(Expression, TruePart, FalsePart)
		If Expression Then
			If IsObject(TruePart) Then
				Set IIF = TruePart
			Else
				IIF = TruePart
			End If
		Else
			If IsObject(FalsePart) Then
				Set IIF = FalsePart
			Else
				IIF = FalsePart
			End If
		End If
	End Function
'-------------------------------------------------------------------------------------------------
'Autor: Ezequiel A. Bacarini
'Fecha: 05/01/2007
Function VerNull(Dato)
		If IsNull(Dato) then
			VerNull = Empty
		Else
			VerNull = Dato
		End If
End Function
'-------------------------------------------------------------------------------------------------
function GF_BD_Puertos(pDbSite, byref rs, P_oprc, P_strSQL)
    GF_BD_Puertos = executeQueryDb(pDbSite, rs, P_oprc, P_strSQL)	
end function
'-------------------------------------------------------------------------------------------
%>