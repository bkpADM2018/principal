<%
'--------------------------------------------------------------------------------------
'Obtiene el tipo de moneda de un contrato de mercaderia dependiendo de su codigo de operacion
Function obtenerMonedaDeContrato(pCdOperacion)
    select case cint(pCdOperacion)
        case 0,1,2,3,5:
            obtenerMonedaDeContrato = MONEDA_PESO
        case 6,9,10,11,12:
            obtenerMonedaDeContrato = MONEDA_DOLAR
    end select
End Function
'--------------------------------------------------------------------------------------
'Devuelve la carpeta que se encuentra activa para un Usuario
Function getFloderActiveByUser()
    Dim rs
    getFloderActiveByUser = ""
    Set sp_ret = executeSP(rs, "MERFL.TBLCTOGTIACARPETASACTIVAS_GET_BY_USER", Session("Usuario") &"||"& Left(Session("MmtoSistema"),8) )
    if ( not rs.Eof ) then getFloderActiveByUser = rs("FECHA") & rs("IDSUCURSAL") & rs("SUBINDICE")
End Function
'--------------------------------------------------------------------------------------
%>