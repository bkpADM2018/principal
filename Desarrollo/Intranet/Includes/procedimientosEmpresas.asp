<%
'*****************************************************************************************************************
function getEnterpriseCUIT(p_KCVEN)
    dim strSQl, rs

    getEnterpriseCUIT = "00-00000000-0"
    if isnumeric(p_KCVEN) then
        strSQL = "select CUIT from Toepferdb.VWEMPRESAS where IDEMPRESA=" & p_KCVEN
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            getEnterpriseCUIT = rs("CUIT")
        end if
    end if
end function
'*****************************************************************************************************************
function getEnterpriseDIR(p_KCVEN)
    dim strSQl, rs

    getEnterpriseDIR = "#KCVEN invalido#"
    if isnumeric(p_KCVEN) then
        strSQL = "Select DOMICI as Domicilio from MERFL.TCB6A1F1 where NROPRO =" & p_KCVEN
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            getEnterpriseDIR = rs("Domicilio")
        end if
    end if
end function
'*****************************************************************************************************************
function getEnterpriseLoc(p_KCVEN)
    dim strSQl, rs

    getEnterpriseLoc = "#KCVEN invalido#"
    if isnumeric(p_KCVEN) then
        strSQL = "Select LOCALI as Localidad from MERFL.TCB6A1F1 where NROPRO =" & p_KCVEN
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            getEnterpriseLoc = rs("Localidad")
        end if
    end if
end function
'*****************************************************************************************************************
function getEnterpriseCP(p_KCVEN)
    dim strSQl, rs

    getEnterpriseCP = 0
    if isnumeric(p_KCVEN) then
        strSQL = "Select CODPOS as CPostal from MERFL.TCB6A1F1 where NROPRO =" & p_KCVEN
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            if not isnull(rs("CPostal")) then
                getEnterpriseCP = clng(rs("CPostal"))
            else
                getEnterpriseCP = "Dato inexistente"
            end if
        end if
    end if
end function
'*****************************************************************************************************************
function getEnterpriseProv(p_KCVEN)
    dim strSQl, rs

    getEnterpriseProv = "#KCVEN invalido#"
    if isnumeric(p_KCVEN) then
        strSQL = "Select B.DESCPO as Provincia from MERFL.TCB6A1F1 A inner join MERFL.MER1K2F1 B on A.CODPRO = B.CODIPO where A.NROPRO =" & p_KCVEN
        call GF_BD_AS400_2(rs, conn, "OPEN", strSQL)
        if not rs.eof then
            if not isnull(rs("Provincia")) then
                getEnterpriseProv = rs("Provincia")
            else
                getEnterpriseProv = "Dato inexistente"
            end if
        end if
    end if
end function
'*****************************************************************************************************************

%>
