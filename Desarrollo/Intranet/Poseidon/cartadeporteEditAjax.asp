<!--#include file="cartadeporteEditCommon.asp"-->
<%
'------------------------------------------------------------------------------------------------------
'Dibuja las opciones de los intervinientes de la carta de porte dependiendo del cuit ingresado
'   - Puede ser una sola descripcion o puede que el cuit tenga varias empresas (en este caso dibujara un elemento Select)
Function drawOptionToInterveniente(p_Rs, pInterviniente, pDs, pCd)
    Dim auxSelected,key
    if p_Rs.RecordCount > 1 then %>
        <select id="cmbInterviniente_<%=pInterviniente %>" name="cmbInterviniente_<%=pInterviniente %>" onchange="changeDsCUIT(this,<%=pInterviniente %>)">
            <option value="" >Seleccione...</option>
            <% while not p_Rs.Eof
                  auxSelected = ""
                  key = p_Rs("CD") & "-" & Trim(p_Rs("DS"))
                  if (pInterviniente <> INTERVINIENTE_CHOFER) then
                    if (pCd <> "" and (p_Rs("CD") <> "")and(p_Rs("CD") <> 0)) then
                        if(Cdbl(pCd) = Cdbl(p_Rs("CD"))) then auxSelected = "selected"
                    else
                        if(Ucase(Trim(pDs)) = Ucase(Trim(p_Rs("DS")))) then auxSelected = "selected"
                    end if  
                  end if  %>
                <option value="<%=key%>" <%=auxSelected %> ><%=p_Rs("DS")%></option>
            <%    p_Rs.MoveNext()
               wend %>
        </select>
<%  else
        Response.Write p_Rs("CD") & "|" & p_Rs("DS")
    end if
End Function
'------------------------------------------------------------------------------------------------------
Function searchEmpresaByCUIT(pCuit, pInterviniente, pDs, pCd)
    Dim strSQL,rs1,rs2
    Select Case Cdbl(pInterviniente)
        Case INTERVINIENTE_INTERMEDIARIO, INTERVINIENTE_REMITENTE
            strSQL  = "SELECT CDVENDEDOR AS CD, RTRIM(DSVENDEDOR) AS DS FROM VENDEDORES WHERE RTRIM(NUDOCUMENTO) ='"& pCuit &"'"
        Case INTERVINIENTE_TITULAR
            strSQL  = "SELECT CDVENDEDOR AS CD, RTRIM(DSVENDEDOR) AS DS FROM VENDEDORES WHERE RTRIM(NUDOCUMENTO) ='"& pCuit &"'"
        Case INTERVINIENTE_CORREDOR
            strSQL = "SELECT CDCORREDOR AS CD, RTRIM(DSCORREDOR) AS DS FROM CORREDORES WHERE RTRIM(NUCUIT) = '"& pCuit &"'"
        Case INTERVINIENTE_ENTREGADOR
            strSQL = "SELECT CDENTREGADOR AS CD, RTRIM(DSENTREGADOR) AS DS FROM ENTREGADORES WHERE RTRIM(NUCUIT) = '"& pCuit &"'"
        Case INTERVINIENTE_DESTINATARIO
            strSQL = "SELECT CDCLIENTE AS CD,RTRIM(DSCLIENTE) AS DS FROM CLIENTES WHERE RTRIM(NUCUIT) ='"& pCuit &"'"
        Case INTERVINIENTE_TRANSPORTISTA
            strSQL = "SELECT CDTRANSPORTISTA AS CD,RTRIM(DSTRANSPORTISTA) AS DS FROM TRANSPORTISTAS WHERE RTRIM(NUDOCUMENTO) ='"& pCuit &"'"
        Case INTERVINIENTE_CHOFER
            strSQL ="SELECT '"& pCuit &"' AS CD , RTRIM(DSNOMBRE) + ',' + RTRIM(DSAPELLIDO) AS DS FROM CONDUCTOR WHERE RTRIM(NCUIT) = '" & pCuit &"'"
    end select
    Call GF_BD_Puertos(g_pto, rs1, "OPEN",strSQL)
     if (not rs1.Eof) then        
        Call drawOptionToInterveniente(rs1, pInterviniente, pDs, pCd)
     'else
     '   if (Cdbl(pInterviniente) = INTERVINIENTE_TITULAR) then
     '       Call executeQuery(rs2, "OPEN", strSQL1)
     '       if not rs2.Eof then
     '           Call drawOptionToInterveniente(rs2, pInterviniente, pDs, pCd)
     '       else
                Response.Write ESTADO_BAJA
     '       end if
     '   else
     '       Response.Write ESTADO_BAJA
     '   end if
    end if
End function
'------------------------------------------------------------------------------------------------------
Function drawOptionBiotecnologiaByProducto(p_CdProducto)
    Dim rsBio
    Set rsBio = getBiotecnologiaByProducto(g_pto, p_CdProducto)
    while not rsBio.Eof %>
        <option value="<%=rsBio("IDBIOTECNOLOGIA") %>" <% if(auxBiotecnologia = Cdbl(rsBio("IDBIOTECNOLOGIA")))then %>selected<% end if %>><%=rsBio("IDBIOTECNOLOGIA")&"-"&rsBio("DSBIOTECNOLOGIA") %></option>
   <%   rsBio.MoveNext()
    wend
End Function
'------------------------------------------------------------------------------------------------------
Dim cuit,tipoInterviniente,cdProducto,accion,valorRebaja,cdRubro

accion = GF_Parametros7("accion","",6)
cdProducto = GF_Parametros7("cdProducto",0,6)

select case accion
    case ACCION_CONTROLAR
       'Busca el intervinite correspondiente al CUIT/CUIL
        cuit = GF_Parametros7("cuit","",6)
        tipoInterviniente = GF_Parametros7("tipoInterviniente",0,6)
        dsInterviniente   = GF_Parametros7("ds","",6)
        cdInterviniente   = GF_Parametros7("cd","",6)
        Call searchEmpresaByCUIT(Trim(cuit),tipoInterviniente,dsInterviniente,cdInterviniente)
    case ACCION_VISUALIZAR
       'Esta accion dibuja los elementos del Combo Box de Biotecnologia
        Call drawOptionBiotecnologiaByProducto(cdProducto)
    case ACCION_CALCULAR 
       'Esta accion calcula el valor de la merma
       valorRebaja = GF_Parametros7("valorRebaja","",6)
       cdRubro = GF_Parametros7("cdRubro",0,6)
       'La funcion calcular merma se encuentra en cartadeporteEditCommon.asp
       Response.Write CalcularMerma(cdProducto, cdRubro, valorRebaja, g_pto)
End select

%>
