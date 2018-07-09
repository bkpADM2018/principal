<!--#include file="../includes/procedimientos.asp"-->
<!--#include file="../includes/procedimientosPuertos.asp"-->
<!--#include file="../includes/procedimientosParametros.asp"-->
<!--#include file="../includes/procedimientossql.asp"-->
<!--#include file="../includes/procedimientostraducir.asp"-->
<!--#include file="../includes/procedimientosfechas.asp"-->
<!--#include file="../includes/procedimientosFormato.asp"-->
<!--#include file="../includes/procedimientosUnificador.asp"-->
<%
Function getComboBoxRubros(p_CdRubro)
    Dim rsRubros
    Call GF_BD_Puertos(g_strPuerto, rsRubros, "OPEN", "SELECT * FROM RUBROS") %>
    <select id="cmbRubros" name="cmbRubros"> 
        <option value="0">Seleccione..</option>
<%  if not rsRubros.Eof then
        while not rsRubros.Eof  %>
            <option value="<%=rsRubros("CDRUBRO") %>" <%if(Cdbl(rsRubros("CDRUBRO")) = Cdbl(p_CdRubro))then %> selected <%end if %>><%= rsRubros("CDRUBRO")&"-"&rsRubros("DSRUBRO")%></option>
        <%  rsRubros.MoveNext()
        wend
    end if
    %>
    </select >  
    <%
End Function
'---------------------------------------------------------------------------------------------------------------
Dim g_strPuerto, g_idCamion, g_dtContable, g_ctaPorte,g_sqCalada,accion

g_strPuerto = GF_Parametros7("pto","",6)
g_dtContable = GF_Parametros7("dtContable","",6)
g_dtContable = Left(g_dtContable,4) & "-" &mid(g_dtContable,5,2) &"-"& Right(g_dtContable,2)
g_ctaPorte = GF_Parametros7("ctaPte","",6)
g_idCamion = GF_Parametros7("idCamion","",6)
g_sqCalada = GF_Parametros7("sqCalada",0,6)
g_cdRubro = GF_Parametros7("cdRubro",0,6)
accion = GF_Parametros7("accion","",6)


Call getComboBoxRubros(g_cdRubro)


	

%>