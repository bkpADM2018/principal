function ordenar_onClick(p_pagina,p_strParam, p_valorCampoOrden)
{

            var strSQL = new String(p_pagina);
            
            if (p_strParam.substr(0,1) != "?") strSQL = strSQL.concat("?");
            
            strSQL = strSQL.concat(p_strParam);
            strSQL = strSQL.concat("&campoOrden=");
            strSQL = strSQL.concat(p_valorCampoOrden);
            document.location.href = strSQL;
}

