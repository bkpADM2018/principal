<!--#include file="../../includes/procedimientosPuertos.asp"-->
<!--#include file="../../includes/procedimientos.asp"-->
<!--#include file="../../includes/procedimientosParametros.asp"-->
<!--#include file="../../includes/procedimientostraducir.asp"-->
<!--#include file="../../includes/procedimientosFormato.asp"-->
<!--#include file="../../includes/procedimientosFechas.asp"-->
<!--#include file="../../includes/procedimientosUnificador.asp"-->
<!--#include file="../../includes/procedimientosTitulos.asp"-->
<!--#include file="../../includes/procedimientosSQL.asp"-->
<!--#include file="Include/procedimientoProducto.asp"-->
<%

Const TIPO_ENVIO_ACEP_CONF_COND_CAMARA = 1
Const TIPO_ENVIO_COND_CAMARA		   = 2
Const TIPO_PRODUCTO_STD				   = 1
Const TIPO_PRODUCTO_BASE			   = 2
Const HUMEDIMETRO_SI				   = 1
Const HUMEDIMETRO_NO				   = 0
'----------------------------------------------------------------------------------------------------------------------
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'----------------------------------------------------------------------------------------------------------------------
'Controla que los datos de la cabecera esten correctos
Function checkHeaderProducto()
	Dim rsPro	
	checkHeaderProducto = false
	if (g_cdProducto = 0) then 
		Call setError(CODIGO_VACIO)
	else		
		if (not g_IsEdit) then
			Set rsPro = getProductoByCdProducto(g_cdProducto,g_strPuerto)
			if (not rsPro.EoF) then Call setError(CODIGO_EXISTE)
		end if	
	end if		
	if (Trim(g_Descripcion) = "") then Call setError(DESCRIPCION_VACIA)
	if (Trim(g_DescripcionAbr) = "") then Call setError(ABRAVIATURA_VACIA)
	if (not hayError) then checkHeaderProducto = true
End Function
'-----------------------------------------------------------------------------------------------------------------------
'Se obtiene los datos del producto 
Function readProductoDB(pCdProducto)
	Dim rs
	Set rs = getProductoByCdProducto(g_cdProducto,g_strPuerto)	
	g_IsEdit = true
	if not rs.Eof then
		g_DescripcionAbr = rs("DSPRODUCTOABR")
		g_Descripcion	 = rs("DSPRODUCTO")
		g_HumedadRecep   = Cdbl(rs("PRK1"))
		g_HumedadBase	 = Cdbl(rs("PRK2"))
		g_Coeficiente1	 = Cdbl(rs("PRK3"))
		g_Coeficiente2	 = Cdbl(rs("PRK4"))
		g_UltimaBoleta	 = CLng(rs("NUULTBOLETCAMARA"))
		g_BaseTrigo		 = CLng(rs("VLBASETRIGO"))
		g_TipoEnvio		 = CInt(rs("ICTIPOENVIO"))
		g_TipoProducto	 = CInt(rs("ICESTANDARBASE"))				
		g_CodigoCamara	 = CInt(rs("CDPRODUCTOCAMARA"))
		g_Humedimetro	 = CInt(rs("ICHUMEDIMETRO"))
		g_UltimoTurno	 = rs("CDNUMERADORTURNO")			
	end if
End Function
'---------------------------------------------------------------------------------------------------------
Function puedeAgregarAtributo(pCdProducto)
	Dim rtrn
	rtrn = false
	if pCdProducto > 0 then 
		Set rs = getProductoByCdProducto(pCdProducto,g_strPuerto)
		if (not rs.Eof) then rtrn = true
	end	if
	puedeAgregarAtributo = rtrn
End Function
'**********************************************************************************************************************
'********************************************* COMIENZA LA PAGINA *****************************************************
'**********************************************************************************************************************
Dim g_strPuerto,params,g_cdProducto,g_TipoEnvio,g_CodigoCamara,g_UltimoTurno,accion,g_UltimaBoleta,g_IsEdit,flagGrabar,flagAdd, g_Humedimetro
Dim g_DescripcionAbr,g_Descripcion,g_HumedadRecep,g_HumedadBase,g_Coeficiente2,g_Coeficiente1,g_BoletaCamara,g_BaseTrigo, g_TipoProducto
Dim flagPermiso 
g_strPuerto = GF_Parametros7("Pto","",6)
call addParam("Pto", g_strPuerto, params)
flagPermiso = true
if (leerPermisos(g_strPuerto, TASK_PRODUCT_USER) = NO_TIENE_PERMISO) then flagPermiso = false

accion = GF_Parametros7("accion","",6)
g_cdProducto = GF_PARAMETROS7("cdProducto",0,6)
call addParam("cdProducto", g_cdProducto, params)
g_IsEdit = GF_Parametros7("isEdit","",6)
call addParam("isEdit", g_IsEdit, params)

if (not isFormSubmit()) then	
	g_IsEdit = false	
	if Cdbl(g_cdProducto) <> 0 then Call readProductoDB(g_cdProducto)	
else	
	g_DescripcionAbr = GF_Parametros7("descripcionAbr","",6)
	g_Descripcion = GF_Parametros7("descripcion","",6)
	g_HumedadRecep = GF_Parametros7("humedadRecepcion",2,6)
	g_HumedadBase = GF_Parametros7("humedadBase",2,6)
	g_Coeficiente1 = GF_Parametros7("coeficiente1",2,6)
	g_Coeficiente2 = GF_Parametros7("coeficiente2",2,6)
	g_UltimaBoleta = GF_Parametros7("ultimaBoleta",0,6)
	g_BaseTrigo = GF_Parametros7("baseTrigo",0,6)
	g_TipoEnvio = GF_Parametros7("tipoEnvio",0,6)
	if (g_TipoEnvio = 0) then g_TipoEnvio = TIPO_ENVIO_ACEP_CONF_COND_CAMARA
	g_TipoProducto = GF_Parametros7("tipoProducto",0,6)
	if (g_TipoProducto = 0) then g_TipoProducto = TIPO_PRODUCTO_STD		
	g_CodigoCamara = GF_Parametros7("codigoCamara",0,6)
	g_Humedimetro = GF_Parametros7("humedimetro",0,6)
	g_UltimoTurno = GF_Parametros7("cmbUltimoTurno","",6)
	if (accion = ACCION_GRABAR) then			
		if (checkHeaderProducto()) then
			if (g_IsEdit) then
				Call updateProducto(g_cdProducto,g_Descripcion,g_HumedadRecep,g_HumedadBase,g_Coeficiente1,g_Coeficiente2,g_UltimaBoleta,g_BaseTrigo,g_TipoEnvio,g_CodigoCamara,g_UltimoTurno,g_TipoProducto,g_Humedimetro,g_DescripcionAbr,g_strPuerto)				
			else
				Call addProducto(g_cdProducto,g_Descripcion,g_HumedadRecep,g_HumedadBase,g_Coeficiente1,g_Coeficiente2,g_UltimaBoleta,g_BaseTrigo,g_TipoEnvio,g_CodigoCamara,g_UltimoTurno,g_TipoProducto,g_Humedimetro,g_DescripcionAbr,g_strPuerto)
			end if	
			flagGrabar = true
			g_IsEdit = true
		end if
	end if
end if	
flagAdd = puedeAgregarAtributo(g_cdProducto)


%>
<HTML>
<HEAD>
	<TITLE>Poseidon - Administracion de Productos </TITLE>
	<meta http-equiv="x-ua-compatible" content="IE=11">
	<link href="../../css/ActisaIntra-1.css" rel="stylesheet" type="text/css" />
	<link rel="stylesheet" href="../../css/Toolbar.css" type="text/css">		
	<link rel="stylesheet" href="../../css/main.css" type="text/css">		
	<link rel="stylesheet" href="../../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />		
				
	<style type="text/css">
		
	</style>
</HEAD>
<script type="text/javascript" src="../../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../../scripts/channel.js"></script>
<script type="text/javascript" src="../../scripts/controles.js"></script>
<script type="text/javascript" src="../../scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
<script type="text/javascript" src="../../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../../Scripts/jquery/jquery-1.5.1.min.js"></script>
<script language="javascript">
	var ch = new channel();
	var up1;	
	function onLoadPage(){				
		tb = new Toolbar('toolbar', 6,'../../images/');				
		tb.addButton("back-16.png","Volver", "volver()");
		<% if (flagPermiso) then %>
        tb.addButton("save-16.png", "Grabar", "submitInfo('<%=ACCION_GRABAR%>')");		
        <% end if %>
        tb.addButton("refresh-16.png", "Refrescar", "submitInfo('<%=ACCION_SUBMITIR%>')");
		tb.draw();
		document.getElementById("msjGrabar").innerHTML  = "";		
		<%if ((flagAdd)and(g_IsEdit))then %>
            loadAtributo(<%=g_cdProducto%>);
            loadCosecha(<%=g_cdProducto%>);
            loadBiotecnologia(<%=g_cdProducto%>);
        <%end if%>
		<% if(flagGrabar) then %>
				document.getElementById("msjGrabar").className  = "reg_Header_success";
				document.getElementById("msjGrabar").innerHTML  = "Se grab� correctamente."
				document.getElementById("accion").value = "";
				document.getElementById("isEdit").value = "<%=g_IsEdit%>";
		<% end if %>		
	}
	function lightOn(tr) {
		tr.className = "reg_Header_navdosHL";
	}
	function volver(){	
		document.location.href = "productoAdministrar.asp?pto=<%=g_strPuerto%>"
	}
	function reloadPage(prod,edit){
		document.getElementById("cdProducto").value = prod;
		document.getElementById("isEdit").value = edit;
		submitInfo("");
	}
	function lightOff(tr) {
		tr.className = "reg_Header_navdos";
	}
	function submitInfo(acc){		
		document.getElementById("accion").value = acc;
		document.getElementById("form1").submit();		
	}	
	function asignarAceptacion(){
		document.getElementById("cdAceptacion").value = document.getElementById("cmbAceptacion").value;		
	}
	
    /* -------------------------------------------    COSECHA   -----------------------------------------   */
    function loadCosecha(pProducto){
        ch.bind("productoCosechaAjax.asp?pto=<%=g_strPuerto %>&cdProducto="+pProducto+"&accion=<%=ACCION_VISUALIZAR%>&permiso=<%=flagPermiso%>", "verDetalleCosecha_Callback("+pProducto+")");
		ch.send();
    }    
    function verDetalleCosecha_Callback( pCdProducto ) {
	    var ret  = ch.response();
        document.getElementById("loadingCosecha").style.display = "none";
	    document.getElementById("divCosecha").innerHTML = ret;
	}
    function AddCosecha(){
        var index = document.getElementById("maxRowCosecha").value;
		var tr = document.createElement('tr');
        var td_Cab_1 = document.createElement('td');
        var txt_cosecha_1 = document.createElement('input');
            txt_cosecha_1.id = "txtCosecha1_"+index;
            txt_cosecha_1.type = "text";
            txt_cosecha_1.setAttribute("onKeyPress","return controlIngreso (this, event, 'N');");
            txt_cosecha_1.size = 6;            
            txt_cosecha_1.setAttribute("maxlength",4);
        td_Cab_1.appendChild(txt_cosecha_1);
        td_Cab_1.align = "center";
        tr.appendChild(td_Cab_1);
        var td_Cab_2 = document.createElement('td');
        var txt_cosecha_2 = document.createElement('input');
            txt_cosecha_2.id = "txtCosecha2_"+index;
            txt_cosecha_2.type = "text";
            txt_cosecha_2.setAttribute("onKeyPress","return controlIngreso (this, event, 'N');");
            txt_cosecha_2.size = 6;
            txt_cosecha_2.setAttribute("maxlength",4);
        td_Cab_2.appendChild(txt_cosecha_2);
        td_Cab_2.align = "center";
        tr.appendChild(td_Cab_2);
        var td_Cab_3 = document.createElement('td');
        var chk = document.createElement('input');
            chk.id = "chk_"+index;
            chk.type = "checkbox";
            chk.setAttribute("onKeyPress","return controlIngreso (this, event, 'N');");
            chk.size = 6;
            chk.maxlength = 4;
        td_Cab_3.appendChild(chk);
        td_Cab_3.align = "center";
        tr.appendChild(td_Cab_3);
        var td_Cab_4 = document.createElement('td');
        var Img_1 = document.createElement('img');
            Img_1.src = "../../images/save-16.png";
            Img_1.id = "guardarAtributo_"+index;
            Img_1.title = "Guardar";
            Img_1.setAttribute("style","cursor:pointer;");
            Img_1.setAttribute("onclick","guardarCosecha("+ index +");");
            td_Cab_4.appendChild(Img_1);
            td_Cab_4.align = "center";
        tr.appendChild(td_Cab_4);
        var td_Cab_5 = document.createElement('td');
        tr.appendChild(td_Cab_5);
        $("#tblCosechas tbody").append(tr);
        document.getElementById("maxRowCosecha").value  = parseInt(index) + 1;
    }
    function editarCosecha(pIndex){
        document.getElementById("spanCosecha3_"+pIndex).style.display = "none";
        document.getElementById("chk_"+pIndex).style.display = "block";
        document.getElementById("editCosecha_"+pIndex).style.display = "none";
        document.getElementById("guardarCosecha_"+pIndex).style.display = "block";
        document.getElementById("eliminarCosecha_"+pIndex).style.display = "block";
    }
    function guardarCosecha(pIndex){
        if(document.getElementById("hidCosecha1_"+pIndex)){
            var cosechaInicio = document.getElementById("hidCosecha1_"+pIndex).value;
            var cosechaFin    = document.getElementById("hidCosecha2_"+pIndex).value;
            var isEdit = true;
        }
        else{
            var cosechaInicio = document.getElementById("txtCosecha1_"+pIndex).value;
            var cosechaFin    = document.getElementById("txtCosecha2_"+pIndex).value;
            var isEdit = false
        }
        if (cosechaInicio.trim() != "" && cosechaFin.trim() != ""){
            ch.bind("productoCosechaAjax.asp?pto=<%=g_strPuerto %>&cdProducto=<%= g_cdProducto%>&accion=<%=ACCION_GRABAR%>&permiso=<%=flagPermiso%>&cosechaI="+cosechaInicio+"&cosechaF="+cosechaFin+"&habilitado="+convertToBoolean($('#chk_'+pIndex).is(':checked'))+"&isEdit="+isEdit , "guardarCosecha_Callback()");
		    ch.send();
        }
        else{
            dibujarErrorCosecha("Debe completar todos los campos");
        }
    }
    function guardarCosecha_Callback(){
        var ret  = ch.response();
        if (ret != "")
            dibujarErrorCosecha(ret);
        else
            loadCosecha("<%= g_cdProducto%>");
    }
    function dibujarErrorCosecha(pDsError){
        document.getElementById("msgErrorCosecha").style.display = "block";
        document.getElementById("msgErrorCosecha").style.color = '#f40800';
        document.getElementById("msgErrorCosecha").innerHTML = pDsError;
    }
    function convertToBoolean(pVal){
        if(pVal == true)
            return "1";
        else
            return "0";

    }
    function deleteCosecha(pIndex){
        if (confirm("Desea eliminar esta Cosecha?")){
            var cosechaI = document.getElementById("hidCosecha1_"+pIndex).value; 
            var cosechaF = document.getElementById("hidCosecha2_"+pIndex).value;
			ch.bind("productoCosechaAjax.asp?accion=<%=ACCION_BORRAR%>&pto=<%=g_strPuerto%>&cdProducto=<%=g_cdProducto %>&cosechaI="+cosechaI+"&cosechaF="+cosechaF , "deleteCosechas_callback()");	
			ch.send();
		}	
    }
    function deleteCosechas_callback(){
        var resp = ch.response();
        if (resp == "") 
            loadCosecha("<%=g_cdProducto %>");
        else
            dibujarErrorCosecha(resp);

    }
    /* -------------------------------------------    ATRIBUTO   -----------------------------------------   */
    function loadAtributo(pProducto){
        ch.bind("productoAtributoAjax.asp?pto=<%=g_strPuerto %>&cdProducto="+pProducto+"&accion=<%=ACCION_VISUALIZAR%>&permiso=<%=flagPermiso%>", "verDetalleAtributo_Callback("+pProducto+")");
		ch.send();
    }
    function verDetalleAtributo_Callback( pCdProducto ) {
	    var ret  = ch.response();        
        document.getElementById("loadingAtributo").style.display = "none";
	    document.getElementById("divAtributo").innerHTML = ret;
	}
    function editarAtributo( pIndex ){
        document.getElementById("spanSticker_"+pIndex).style.display = "none";
        document.getElementById("divSticker_"+pIndex).style.display = "block";        
        document.getElementById("spanSupervisor_"+pIndex).style.display = "none";
        document.getElementById("divSupervisor_"+pIndex).style.display = "block";        
        document.getElementById("spanCamara_"+pIndex).style.display = "none";
        document.getElementById("divCamara_"+pIndex).style.display = "block";        
        document.getElementById("spanRechazo_"+pIndex).style.display = "none";
        document.getElementById("divRechazo_"+pIndex).style.display = "block";        
        document.getElementById("spanGrado_"+pIndex).style.display = "none";
        document.getElementById("divGrado_"+pIndex).style.display = "block";        
        document.getElementById("spanMerma_"+pIndex).style.display = "none";
        document.getElementById("divMerma_"+pIndex).style.display = "block";
        document.getElementById("spanRubo_"+pIndex).style.display = "none";                                
        document.getElementById("divRubro_"+pIndex).style.display = "block";
        document.getElementById("spanBalde_"+pIndex).style.display = "none";
        document.getElementById("divBalde_"+pIndex).style.display = "block";
        document.getElementById("spanAcon_"+pIndex).style.display = "none";
        document.getElementById("divAcon_"+pIndex).style.display = "block";
        document.getElementById("spanInforme_"+pIndex).style.display = "none";
        document.getElementById("divInforme_"+pIndex).style.display = "block";
        document.getElementById("editarAtributo_"+pIndex).style.display = "none";
        document.getElementById("guardarAtributo_"+pIndex).style.display = "block";
    }	
	function cargaDel(pCdProducto, pCdAceptacion){
		if (confirm("Desea eliminar este Atributo?")){
			ch.bind("productoAtributoAjax.asp?accion=<%=ACCION_BORRAR%>&pto=<%=g_strPuerto%>&cdProducto=" + pCdProducto + "&cdAceptacion=" + pCdAceptacion, "loadAtributo(" + pCdProducto + ")");	
			ch.send();
		}	
	}	
    function getComboBoxAtributo(pIndex){
        ch.bind("productoAtributoAjax.asp?accion=<%=ACCION_PROCESAR%>&pto=<%=g_strPuerto%>&cdProducto=<%=g_cdProducto %>&indice="+pIndex , "getComboBoxAtributo_Callback(" + pIndex + ")");
		ch.send();
	}
    function getComboBoxAtributo_Callback(pIndex){
        var ret  = ch.response();
        document.getElementById("divCmb_"+pIndex).innerHTML = ret;
    }
    var obj_tr;
    function AddAtributo(){
        var index = document.getElementById("maxRowAtributo").value;
		obj_tr = document.createElement('tr');
        var td_Cab = document.createElement('td');
        var div_Combo = document.createElement('div');
            div_Combo.id = "divCmb_"+index;
        td_Cab.appendChild(div_Combo);
        obj_tr.appendChild(td_Cab);
        getComboBoxAtributo(index);
        //creo las celdas de las opciones
        createOptionAtribute( "rdbSticker_" + index );
        createOptionAtribute( "rdbSupervisor_" + index );
        createOptionAtribute( "rdbCamara_" + index );
        createOptionAtribute( "rdbRechazo_" + index );
        createOptionAtribute( "rdbGrado_" + index );
        createOptionAtribute( "rdbMerma_" + index );
        createOptionAtribute( "rdbRubro_" + index );
        createOptionAtribute( "rdbBalde_" + index );
        createOptionAtribute( "rdbAcon_" + index );
        createOptionAtribute( "rdbInforme_" + index );
        var td_Img_1 = document.createElement('td');
            td_Img_1.align = "center";
        var Img_1 = document.createElement('img');
            Img_1.src = "../../images/save-16.png";
            Img_1.id = "guardarAtributo_"+index;
            Img_1.title = "Guardar"
            Img_1.setAttribute("style","cursor:pointer;")
            Img_1.setAttribute("onclick","guardarAtributo("+ index +");");
            td_Img_1.appendChild(Img_1);
        obj_tr.appendChild(td_Img_1);
        var td_Img_2 = document.createElement('td');
        obj_tr.appendChild(td_Img_2);
        $("#tblAtributo tbody").append(obj_tr);
        document.getElementById("maxRowAtributo").value  = parseInt(index) + 1;        
    }    
    // Crea las opciones de los Atributos, se le pasa el nombre que va a tener cada option
    function createOptionAtribute(pNameAtr){
        var td = document.createElement('td');
            td.align = "center" ;
            var div_Si  = document.createElement('div');
                div_Si.setAttribute("style","margin:5px;");
            var rdb_Si = document.createElement('input');
                rdb_Si.type = "radio";
                rdb_Si.name = pNameAtr;
                rdb_Si.value = "<%=VALUE_ATRIBUTE_AFIRMATIVO%>";
                rdb_Si.setAttribute("style","float:left;margin:0;");
                rdb_Si.checked = "true"; //Por defecto aparece checkeado el valor Si
            div_Si.appendChild(rdb_Si);
            var font_Si  = document.createElement('font');
                font_Si.setAttribute("style","margin:0 auto;");
                font_Si.innerHTML = "Si";
            div_Si.appendChild(font_Si);
            td.appendChild(div_Si);
            var div_No  = document.createElement('div');
                div_No.setAttribute("style","margin:5px;");
            var rdb_No  = document.createElement('input');
                rdb_No.type = "radio";
                rdb_No.name = pNameAtr;
                rdb_No.value = "<%=VALUE_ATRIBUTE_NEGATIVO%>";
                rdb_No.setAttribute("style","float:left;margin:0;");
            div_No.appendChild(rdb_No);
            var font_No  = document.createElement('font');
                font_No.setAttribute("style","margin:0 auto;");
                font_No.innerHTML = "No";
            div_No.appendChild(font_No);
            td.appendChild(div_No);
            var div_Opcional  = document.createElement('div');
                div_Opcional.setAttribute("style","margin:5px;");
            var rdb_Opcional  = document.createElement('input');
                rdb_Opcional.type = "radio";
                rdb_Opcional.name = pNameAtr;
                rdb_Opcional.value = "<%=VALUE_ATRIBUTE_OPCIONAL%>";
                rdb_Opcional.setAttribute("style","float:left;margin:0;");
            div_Opcional.appendChild(rdb_Opcional);
            var font_Opcional  = document.createElement('font');
                font_Opcional.setAttribute("style","margin:0 auto;");
                font_Opcional.innerHTML = "Opcional";
            div_Opcional.appendChild(font_Opcional);
            td.appendChild(div_Opcional);        
        obj_tr.appendChild(td);
        }
    function guardarAtributo(pIndex){
        if (document.getElementById("cdAceptacion_"+pIndex)){
            var cdConcepto = document.getElementById("cdAceptacion_"+pIndex).value;
            var isEditAtribute = true;
        }
        else{
            var cdConcepto = document.getElementById("cmbAceptacion_"+pIndex).value;
            var isEditAtribute = false;
        }
        if (cdConcepto != 0){
            var sticke = $('input:radio[name=rdbSticker_'+pIndex+']:checked').val();
            var superv = $('input:radio[name=rdbSupervisor_'+pIndex+']:checked').val();
            var camara = $('input:radio[name=rdbCamara_'+pIndex+']:checked').val();
            var rechaz = $('input:radio[name=rdbRechazo_'+pIndex+']:checked').val();
            var grado  = $('input:radio[name=rdbGrado_'+pIndex+']:checked').val();
            var merma  = $('input:radio[name=rdbMerma_'+pIndex+']:checked').val();
            var rubro  = $('input:radio[name=rdbRubro_'+pIndex+']:checked').val();
            var balde  = $('input:radio[name=rdbBalde_'+pIndex+']:checked').val();
            var acon   = $('input:radio[name=rdbAcon_'+pIndex+']:checked').val();
            var inform = $('input:radio[name=rdbInforme_'+pIndex+']:checked').val();            
            ch.bind("productoAtributoAjax.asp?pto=<%=g_strPuerto %>&cdProducto=<%=g_cdProducto %>&accion=<%=ACCION_GRABAR%>&cdAceptacion="+cdConcepto+"&sticker="+sticke+"&superv="+superv+"&camara="+camara+"&grado="+grado+"&merma="+merma+"&rubro="+rubro+"&balde="+balde+"&acon="+acon+"&inform="+inform+"&rechaz="+rechaz+"&isEdit="+isEditAtribute, "guardarAtributo_Callback()");
		    ch.send();
        }
        else{            
            dibujarErrorAtributo("Debe seleccionar el Concepto");
        }
    }
    function guardarAtributo_Callback(){
        var ret  = ch.response();
        if (ret != "") 
            dibujarErrorAtributo(ret);
        else
            loadAtributo("<%= g_cdProducto%>");
    }

    function dibujarErrorAtributo(pDsError){
        document.getElementById("msgErrorAtributo").style.display = "block";
        document.getElementById("msgErrorAtributo").style.color = '#f40800';
        document.getElementById("msgErrorAtributo").innerHTML = pDsError;
    }
    /* -------------------------------------------    BIOTECNOLOGIA   -----------------------------------------   */
    function loadBiotecnologia(pProducto){
        ch.bind("productoBiotecnologiaAjax.asp?pto=<%=g_strPuerto %>&cdProducto="+pProducto+"&accion=<%=ACCION_VISUALIZAR%>&permiso=<%=flagPermiso%>", "verDetalleBiotecnologia_Callback("+pProducto+")");
		ch.send();
    }
    function verDetalleBiotecnologia_Callback( pCdProducto ) {
	    var ret  = ch.response();
        document.getElementById("loadingBiotecnologia").style.display = "none";
	    document.getElementById("divBiotecnologia").innerHTML = ret;
	}
    function EditBiotecnologia( pIndex ){
        document.getElementById("spanDsBiotecnologia_"+pIndex).style.display = "none";
        document.getElementById("spanDsCliente_"+pIndex).style.display = "none";
        document.getElementById("spanNuSobre_"+pIndex).style.display = "none";
        document.getElementById("editarBiotecnologia_"+pIndex).style.display = "none";
        document.getElementById("guardarBiotecnologia_"+pIndex).style.display = "block";        
        document.getElementById("txtDsBiotecnologia_"+pIndex).style.display = "block";
        document.getElementById("dsCoordinado_"+pIndex).style.display = "block";
        document.getElementById("spanEstado_"+pIndex).style.display = "none";
        document.getElementById("spanDeshabilitado_"+pIndex).style.display = "block";
        document.getElementById("spanHabilitado_"+pIndex).style.display = "block";
        document.getElementById("txtNuSobre_"+pIndex).style.display = "block";        
        autoCompleteCoordinado(pIndex);    
    }
    function AddBiotecnologia(){
        var index = document.getElementById("maxRowBiotecnologia").value;
		var tr_1 = document.createElement('tr');
        var td_Cab_1 = document.createElement('td');
        tr_1.appendChild(td_Cab_1);
        var td_Cab_2 = document.createElement('td');
        td_Cab_2.align = "center";
        var txt_Cab_1  = document.createElement('input');
            txt_Cab_1.type = "text";
            txt_Cab_1.id = "txtDsBiotecnologia_"+index;
            txt_Cab_1.style = "width:100%;text-transform:uppercase;";
            txt_Cab_1.maxlength= "250";
        td_Cab_2.appendChild(txt_Cab_1);
        tr_1.appendChild(td_Cab_2);
        var td_Cab_3 = document.createElement('td');
        td_Cab_3.align = "center";
        var txt_Cab_2  = document.createElement('input');
            txt_Cab_2.type = "text";
            txt_Cab_2.id = "dsCoordinado_"+index;
            txt_Cab_2.name = "dsCoordinado_"+index;
            txt_Cab_2.setAttribute("onblur","controlarProveedor("+index+")");
            txt_Cab_2.style = "width:100%;";
        var hid_Cab_2  = document.createElement('input');
            hid_Cab_2.type = "hidden";
            hid_Cab_2.id = "cdCoordinado_"+index;
            hid_Cab_2.name = "cdCoordinado_"+index;
        td_Cab_3.appendChild(txt_Cab_2);
        td_Cab_3.appendChild(hid_Cab_2);
        tr_1.appendChild(td_Cab_3);
        var td_Cab_4 = document.createElement('td');
        td_Cab_4.align = "center";
        var spa_Cab_1  = document.createElement('span');
            spa_Cab_1.id = "spanHabilitado_"+index;
        td_Cab_4.appendChild(spa_Cab_1);
        var rad_Cab_1  = document.createElement('input');
            rad_Cab_1.type = "radio";
            rad_Cab_1.id = "rdb_"+index;
            rad_Cab_1.name = "rdb_"+index;
            rad_Cab_1.title = "Habilitado";
            rad_Cab_1.value = "V";
            rad_Cab_1.checked = "true";
        td_Cab_4.appendChild(rad_Cab_1);
        var spa_Cab_2  = document.createElement('span');
            spa_Cab_2.id = "spanDeshabilitado_"+index;            
            spa_Cab_2.innerHtml = "F:";
        td_Cab_4.appendChild(spa_Cab_2);
        var rad_Cab_2  = document.createElement('input');
            rad_Cab_2.type = "radio";
            rad_Cab_2.id = "rdb_"+index;
            rad_Cab_2.name = "rdb_"+index;
            rad_Cab_2.value = "F";
            rad_Cab_2.title = "Deshabilitado";
        td_Cab_4.appendChild(rad_Cab_2);
        tr_1.appendChild(td_Cab_4);
        var td_Cab_5 = document.createElement("td");
        td_Cab_5.align = "center";
        var txt_Cab_3 = document.createElement("input");
        txt_Cab_3.id = "txtNuSobre_"+index;
        txt_Cab_3.name = "txtNuSobre_"+index;
        txt_Cab_3.style = "width:100%;";
        txt_Cab_3.setAttribute('onkeypress', "return controlIngreso(this, event, 'N')");
        txt_Cab_3.setAttribute('maxlength', "4");
        td_Cab_5.appendChild(txt_Cab_3);
        tr_1.appendChild(td_Cab_5);
        var td_Cab_6 = document.createElement("td");
        td_Cab_6.align = "center";
        var img_Cab_1 = document.createElement("img");
        img_Cab_1.id = "guardarBiotecnologia_"+index;
        img_Cab_1.name = "guardarBiotecnologia_"+index;
        img_Cab_1.style = "cursor:pointer;";
        img_Cab_1.src = "../../images/save-16.png"
        img_Cab_1.title = "Guardar"
        img_Cab_1.setAttribute("onclick","SaveBiotecnologia("+index+")");
        td_Cab_6.appendChild(img_Cab_1);
        tr_1.appendChild(td_Cab_6);
        var td_Cab_7 = document.createElement("td");
        tr_1.appendChild(td_Cab_7);
        $("#tblBiotecnologia tbody").append(tr_1);
        document.getElementById("spanHabilitado_"+index).innerHTML = "V:";
        document.getElementById("spanDeshabilitado_"+index).innerHTML = "F:";
        autoCompleteCoordinado(index);
        document.getElementById("maxRowBiotecnologia").value  = parseInt(index) + 1;
    }
    
    // Envia los datos por Ayax para guardarlos en caso que no esten vacios
    function SaveBiotecnologia( pIndex ){
        var dsBiotecnologia = document.getElementById("txtDsBiotecnologia_"+pIndex).value;
        var cdCliente = document.getElementById("cdCoordinado_"+pIndex).value;
        
        if (dsBiotecnologia != "" && cdCliente != ""){
            var idBiotecnologia = 0;
            if ( document.getElementById("hidIdBiotecnologia_"+pIndex) ) var idBiotecnologia = document.getElementById("hidIdBiotecnologia_"+pIndex).value;
            ch.bind("productoBiotecnologiaAjax.asp?pto=<%=g_strPuerto %>&cdProducto=<%=g_cdProducto %>&accion=<%=ACCION_GRABAR%>&id="+idBiotecnologia+"&descripcion="+dsBiotecnologia+"&cliente="+cdCliente+"&nuSobre="+document.getElementById("txtNuSobre_"+pIndex).value+"&estado="+$('input:radio[name=rdb_'+pIndex+']:checked').val(), "SaveBiotecnologia_Callback("+pIndex+")");
		    ch.send();
        }
        else{
            dibujarErrorBiotecnologia("Debe completar todos los campos");
        }
    }
    function SaveBiotecnologia_Callback( pIndex ){
        var ret  = ch.response();
        if (ret != "") 
            dibujarErrorBiotecnologia(ret);
        else
            loadBiotecnologia("<%= g_cdProducto%>");

    }
    function DeleteBiotecnologia(pIndex){
        if (confirm("Desea eliminar la Biotecnologia?")){
            ch.bind("productoBiotecnologiaAjax.asp?pto=<%=g_strPuerto %>&cdProducto=<%=g_cdProducto %>&accion=<%=ACCION_BORRAR%>&id="+document.getElementById("hidIdBiotecnologia_"+pIndex).value, "loadBiotecnologia(<%= g_cdProducto%>)");
		    ch.send();
        }
    }    
    function dibujarErrorBiotecnologia( pDsError ){
        document.getElementById("msgErrorBiotecnologia").style.display = "block";
        document.getElementById("msgErrorBiotecnologia").style.color = '#f40800';
        document.getElementById("msgErrorBiotecnologia").innerHTML = pDsError;
    }
    function controlarProveedor(pIndex){
        var auxDs = document.getElementById("dsCoordinado_"+pIndex).value.toString();
        if (auxDs.trim() == "") document.getElementById("idCoordinado_"+pIndex).value = "";
    }
    function autoCompleteCoordinado( pIndex ){
        $( "#dsCoordinado_"+pIndex ).autocomplete({
			minLength: 1,
			source: "../puertosStreamElementos.asp?tipo=JQClientes&pto=<%=g_strPuerto%>",
			focus: function( event, ui ) {
				$( "#dsCoordinado_"+pIndex).val(ui.item.dscliente);
			return false;
			},
			select: function( event, ui ) {
				$( "#dsCoordinado_"+pIndex    ).val (ui.item.dscliente);
				$( "#cdCoordinado_"+pIndex    ).val (ui.item.cdcliente);
				return false;
			},
			change: function( event, ui ) {
				if (!ui.item) {
					$( "#dsCoordinado_"+pIndex).val ("");
					$( "#cdCoordinado_"+pIndex).val ("");
				}
			}
		})
		.data( "autocomplete" )._renderItem = function( ul, item ) {
			return $( "<li></li>" )
				.data( "item.autocomplete", item )
				.append( "<a>" + item.cdcliente + " - <font style='font-size:10;'>" + item.dscliente + "</font></a>" )
				.appendTo( ul );
		};
	}
</script>
<BODY onload="onLoadPage()">
<DIV id="toolbar"></DIV>
<form name="form1" id="form1" method=post>					
<div class="tableaside size100"> 
	<div ><% call showErrors() %></div>
	<div id="msjGrabar"></div>
	<h3><%=GF_Traducir("Datos del Producto")%></h3>
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("C�digo") %> </div>        
        <div class="col26">
			<% If (g_IsEdit) Then 
				Response.Write g_cdProducto %>
				<input type="hidden" id="cdProducto" name="cdProducto" value="<%=g_cdProducto%>">
			<% else %>			
				<input type="text" id="cdProducto" name="cdProducto" <%if(g_cdProducto<>0)then%>value="<%=g_cdProducto%>"<%end if%> onKeyPress="return controlIngreso (this, event, 'N');">
			<% end if %>			
		</div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripci�n Abr.") %> </div>
        <div class="col26"><input type="text" id="descripcionAbr" name="descripcionAbr" value="<%= g_DescripcionAbr %>"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripci�n") %> </div>
        <div class="col26"><input type="text" id="descripcion" name="descripcion" value="<%= g_Descripcion %>"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Humedad recepci�n") %> </div>
        <div class="col26"><input type="text" id="humedadRecepcion" name="humedadRecepcion"  value="<%= g_HumedadRecep %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Humedad base") %> </div>
        <div class="col26"><input type="text" id="humedadBase" name="humedadBase"  value="<%= g_HumedadBase %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Coeficiente 1") %> </div>
        <div class="col26"><input type="text" id="coeficiente1" name="coeficiente1"  value="<%= g_Coeficiente1 %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Coeficiente 2") %> </div>
        <div class="col26"><input type="text" id="coeficiente2" name="coeficiente2"  value="<%= g_Coeficiente2 %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Ultima boleta c�mara") %> </div>
        <div class="col26"><input type="text" id="ultimaBoleta" name="ultimaBoleta"  value="<%= g_UltimaBoleta %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Base Trigo") %> </div>
        <div class="col26"><input type="text" id="baseTrigo" name="baseTrigo"  value="<%= g_BaseTrigo %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("C�digo C�mara") %> </div>
        <div class="col26"><input type="text" id="codigoCamara" name="codigoCamara"  value="<%= g_CodigoCamara %>" onKeyPress="return controlIngreso (this, event, 'N');"></div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Tipo de Producto") %> </div>
        <div class="col26">
			<input name="tipoProducto" id="tipoProducto" type="radio" value="<%=TIPO_PRODUCTO_STD%>"  <% if (g_TipoProducto = TIPO_PRODUCTO_STD) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("Std.")%>
			<input name="tipoProducto" id="tipoProducto" type="radio" value="<%=TIPO_PRODUCTO_BASE%>" <% if (g_TipoProducto = TIPO_PRODUCTO_BASE) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("Base")%>
        </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Humed�metro") %> </div>
        <div class="col26">
			<input name="humedimetro" id="humedimetro" type="radio" value="<%=HUMEDIMETRO_SI%>" <% if (g_Humedimetro = HUMEDIMETRO_SI) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("Si")%>
			<input name="humedimetro" id="humedimetro" type="radio" value="<%=HUMEDIMETRO_NO%>" <% if (g_Humedimetro = HUMEDIMETRO_NO) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("No")%>
        </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Tipo de Env�o") %> </div>
        <div class="col46">
			<input name="tipoEnvio" id="tipoEnvio" type="radio" value="<%=TIPO_ENVIO_ACEP_CONF_COND_CAMARA%>" <% if ((g_TipoEnvio <= TIPO_ENVIO_ACEP_CONF_COND_CAMARA)or(g_TipoEnvio="")) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("Acep. Conf. y Cond. C�mara")%>			
			<input name="tipoEnvio" id="tipoEnvio" type="radio" value="<%=TIPO_ENVIO_COND_CAMARA%>" <% if (g_TipoEnvio = TIPO_ENVIO_COND_CAMARA) then Response.Write "checked" end if %>/><%=GF_TRADUCIR("Cond. C�mara solamente")%>
        </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("�ltimo turno") %> </div>
        <div class="col46">
			<select id="cmbUltimoTurno" name="cmbUltimoTurno">
				<option value="0"><% =GF_TRADUCIR("Seleccione..") %></option>
			<%	strSQL = "SELECT CDNUMERADOR,DSNUMERADOR FROM dbo.CONTADORESNUMERADORES ORDER BY DSNUMERADOR"
				call GF_BD_Puertos (g_strPuerto, rsNum, "OPEN",strSQL)										
				while not rsNum.eof
					if cstr(g_ultimoTurno) = cstr(rsNum("CDNUMERADOR")) then
						mySelected = "SELECTED"
					else
						mySelected = ""
					end if	%>
					<option value="<%=rsNum("CDNUMERADOR")%>" <%=mySelected%>><%=rsNum("DSNUMERADOR")%></option>
				<%	rsNum.movenext
				wend  %>
			</select>
        </div>
	</div>	
</div>
<% if ((flagAdd)and(g_IsEdit))then %>
    <div class="tableaside size100"> 
	    <h3> <%=GF_Traducir("Atributos del Producto")%> </h3>		
	    <div class="tableasidecontent">
            <img src="../../images/Loading4.gif" id="loadingAtributo" name="loadingAtributo" style="display:block;margin:0 auto;" />
		    <div id="divAtributo"></div>
	    </div>
    </div>
    <div class="tableaside size100"> 
	    <h3> <%=GF_Traducir("Cosecha")%> </h3>		
	    <div class="tableasidecontent">
		    <img src="../../images/Loading4.gif" id="loadingCosecha" name="loadingCosecha" style="display:block;margin:0 auto;" />        
            <div id="divCosecha"></div>
	    </div>
    </div>
    <div class="tableaside size100"> 
	    <h3> <%=GF_Traducir("Biotecnolog�as")%> </h3>		
	    <div class="tableasidecontent">
            <img src="../../images/Loading4.gif" id="loadingBiotecnologia" name="loadingBiotecnologia" style="display:block;margin:0 auto;" />
            <div id="divBiotecnologia"></div>
	    </div>
    </div>
<% end if %>
<input type="hidden" name="accion" id="accion" <%=accion%>>
<input type="hidden" name="isEdit" id="isEdit" value="<%=g_IsEdit%>">
</form>
</BODY>
</HTML>