<!--#include file="../Includes/procedimientostraducir.asp"-->	
<!--#include file="../Includes/procedimientosMG.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->		
<!--#include file="../Includes/procedimientosSql.asp"-->
<!--#include file="../Includes/procedimientosUser.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
<% 
'-----------------------------------------------------------------------------------

Dim codParam, admin, accion, dsMensaje,idProveedor,pto,rsPar, esmodificable,strMsj
Dim nomParam,valParam,visitante,rsEdit,rsPue,v_tieneExtra,Edit,Check_Edit
Dim codParam_old, ValParametro_old, Check_Edit_old, idPuesto_old,nomParametro_old, dsMensaje_old

codParam = GF_Parametros7("codParametro","",6)
accion = GF_Parametros7("accion",0,6)
pto = GF_Parametros7("pto","",6)
visitante = GF_Parametros7("visitante",0,6)
dsMensaje = replace (GF_PARAMETROS7("DsParametro","" ,6),chr(13)&CHR(10),"<br>")
valParam = GF_Parametros7("ValParametro",0,6)
idPuesto = GF_Parametros7("idPuesto",0,6)
idPuesto_old = GF_Parametros7("Puesto_old",0,6)
idPuesto_old = 0
codParam_old = GF_Parametros7("codParametro_old","",6)
ValParametro_old = GF_Parametros7("ValParametro_old",0,6)
Check_Edit_old = GF_Parametros7("Check_Edit_old","",6)
dsMensaje_old = replace (GF_PARAMETROS7("DsParametro_old","" ,6),chr(13)&CHR(10),"<br>")


if(accion = ACCION_MODIFICAR_PARAMETRO )then
	v_tieneExtra = tieneParametrosExtra(codParam,pto)
	'en v_tieneExtra dice si el parametro pasado tiene o no parametrosextra(true false) 		
	esmodificable = true
	'si modifico tengo que traer los datos del parametro elegido para q los muestre en los text
	set rsPar = leerParametros(pto, codParam, "", 0, "", true)					
	'cargo en la variable el registro de descripcion de parametro 
	nomParam = rsPar("DSPARAMETRO")
	valParam = rsPar("VLPARAMETRO")
	nomParametro_old = rsPar("DSPARAMETRO")
	ValParametro_old = rsPar("VLPARAMETRO")
	'esto va despues en parametrosajax.asp para que luego de hacer el insert o update 
	'grabe el Log indicando si es mod o alta
end if	
%>
<html>
<head>
	<link rel="stylesheet" href="../css/ActiSAIntra-1.css"	 type="text/css">
	<script type="text/javascript" src="../scripts/channel.js"></script>
	<script defer type="text/javascript" src="../scripts/pngfix.js"></script>	
	<script>
		
	var ch= new channel();	
	var chn= new channel();	
	var chnl= new channel();	
	
	//nota: esta funcion me permite validar que no se ingresen ciertos carracteres como el 
	//&,%,$,+,¿,....como parte del codigo de parametro
	function controlCarracteresEspeciales(campo, evento){	
		var ret = true;	
		var auxText = new String();
		var ascii = (document.all) ? evento.keyCode : evento.which;				
		var caracter = String.fromCharCode(ascii);			
		auxText = campo.value;			
		if (ascii==8 || ascii==0 || ascii==9 || ascii==13 || (ascii > 47 && ascii< 58)){ 
		ret = true;}
		if (ascii==36 || ascii==37 || ascii==38 || ascii==47 || ascii==43 || ascii==34 || ascii==64 || ascii==63 || ascii==61){ 
		ret = false;}
		return ret			
	}
	
	function modificarParametro(acc,pto,cod){
		var pMsj = document.getElementById("DsParametro").value;
		if(pMsj == ""){
			alert("Debe ingresar la descripcion del parametro");
		}
		else{		
			var pVal = document.getElementById("ValParametro").value;					
			var pMsj_old = document.getElementById("DsParametro_old").value;
			var pVal_old = document.getElementById("ValParametro_old").value;		
			pMsj = pMsj.replace(/\n/gi,"<br>");
			pMsj_old = pMsj_old.replace(/\n/gi,"<br>");				
			ch.bind("ParametrosAjax.asp?cdParam="+cod+"&nomParam="+pMsj+"&nomParam_old="+pMsj_old+"&valParam="+pVal+"&valParam_old="+pVal_old+"&accion="+acc+"&pto="+pto,"modificarParametroExtra('"+cod+"','"+pto+"')");
			ch.send();			
		}
	}		
	
	function modificarParametroExtra(pcod, pto){
		<%if(visitante = TASK_PARAM_ADMIN)then%>
		var myEdit = document.getElementById("Check_Edit").checked;
		var myPuesto = document.getElementById("idPuesto").value;
		var myEdit_old = document.getElementById("Check_Edit_old").value;
		var myPue_old = document.getElementById("Puesto_old").value;	
		if(myPuesto > 0){
			/*
			 *	Se guardarn solo aquellos que se selecciono el puesto, por que puede que 
			 *	se quiera guardar un parametro pero no su extra, por lo tanto no se va a 
			 *	seleccionar nada(extra) y no se lo va a guardar(extra)
			 *	
			 */
			chn.bind("ParametrosAjax.asp?cdParam="+pcod+"&editable="+myEdit+"&editable_old="+myEdit_old+"&puesto="+myPuesto+"&puesto_old="+myPue_old+"&accion=<% =ACCION_MODIFICAR_PARAMETRO_EXTRA%>&pto="+pto,"modificacionRealizada()");
			chn.send();	
		}
		else{
			if(myEdit == true){
				alert("Para guardar el campo Editable necesita saleccionar un Puesto");
			}
			else{
				modificacionRealizada();
			}
		}
		<%else%>
			modificacionRealizada();
		<%end if%>
	}
	
	function modificacionRealizada(){
		document.getElementById("avisoModificado").innerHTML="<% =GF_TRADUCIR("Se modifico correctamente") %>";
		document.getElementById("avisoModificado").className = "TDSUCCESS";	
		}
	
	
	function agregarParametro(pto){
			
			var pCod = document.getElementById("codParametro").value;
			if(pCod == ""){
				alert("Debe ingresar el codigo");
			}
			else{			
				var pMsj = document.getElementById("DsParametro").value;
				if(pMsj == ""){
					alert("Debe ingresar la descripcion del parametro");
				}
				else{		
					var pVal = document.getElementById("ValParametro").value;		
					if((pVal == 0)||(pVal == "")){
						alert("Debe ingresar el valor del parametro");
					}
					else{						
						pMsj = pMsj.replace(/\n/gi,"<br>");		
						var myEdit = document.getElementById("Check_Edit").checked;
						var myPuesto = document.getElementById("idPuesto").value;
												
						if((myPuesto > 0)||((myEdit == false)&&(myPuesto <= 0))){
							/*
							*  Primero va a ver si ya existe ese codigo, en caso de que no lo crea ahi nomas
							*	En caso de que exista duplicado lo informa en el callBack
							*/
							ch.bind("ParametrosAjax.asp?cdParam="+pCod+"&nomParam="+pMsj+"&valParam="+pVal+"&accion=<%=ACCION_COMPROBAR_PARAMETRO%>&pto="+pto,"agregarParametroExtra('"+pCod+"','"+pto+"')");	
							ch.send();	
						}
						else{
							if(myEdit == true){
								alert("Para guardar el campo Editable necesita saleccionar un Puesto");
							}						
						}
					}	
				}	
			}
		}
		
	function agregarParametroExtra(pCod,pto){
		var v_parametroNoExistente;
		v_parametroNoExistente = ch.response();
		if(v_parametroNoExistente == <%=PARAMETRO_NO_EXISTENTE%>){	
			var myEdit = document.getElementById("Check_Edit").checked;
			var myPuesto = document.getElementById("idPuesto").value;
			if(myPuesto > 0){
				/* 
				 *	Con estas validaciones evitamos viajar por channel en caso de que
				 * 	no se halla seleccionado solamente el puesto, debido a que se puede
				 *	guardar aunque se halla seleccionado si es editable o no. Es decir
				 * 	que no se guarda los registros EXTRAS que no eligen el puesto
				 */
				chn.bind("ParametrosAjax.asp?cdParam="+pCod+"&editable="+myEdit+"&puesto="+myPuesto+"&accion=<% =ACCION_AGREGAR_PARAMETRO_EXTRA%>&pto="+pto,"agregarParametroExtra_callBack('"+pCod+"')");
				chn.send();									
			}
			else{
				agregarParametroExtra_callBack(pCod);
			}
		}	
		else{
			alert("El codigo de paramtro ya existe");			
		}
	}				
	
	function agregarParametroExtra_callBack(pCod){		
		document.getElementById("avisoModificado").innerHTML="<% =GF_TRADUCIR("Se dio de alta correctamente") %>";
		document.getElementById("avisoModificado").className = "TDSUCCESS";		
		parent.window.cerrarPopUpPar(pCod);
	}
	
	
	
	</script>
</head>
<body>	
	<input name="codParametro_old" id="codParametro_old" type="hidden" value=<%=codParam%>>
	<form id="myForm" name="myForm" action="parametrosPopUp.asp" method="post"  onsubmit="return false;" >
		<input type="hidden" name="accion" id="accion" value=<%=accion%>>		
		<input type="hidden" name="visitante" id="visitante" value=<%=visitante%>>	
		<table width="100%">		
			<tr><td><div id="avisoModificado" align="center" class="TDBAJAS"></div></td></tr>
			<tr>			
				<td>			
					<table class="reg_header" width="100%" border="0" id="Detalle" name="Detalle" cellpadding="1" cellspacing="2">				
						<tr>
							<td class="reg_header_nav">
								<%=GF_Traducir("Cod. Parametro")%>
							</td>
							<td >
								<%if(esmodificable) then %>
									<%=codParam%>	
								<%else%>
									<input name="codParametro" id="codParametro" onKeyPress="return controlCarracteresEspeciales(this, event);" type="text" value="<%=Trim(codParam)%>">
								<%end if%>
								
							</td>
						</tr>
						<tr>
							<td class="reg_header_nav">
								<%=GF_Traducir("Descripcion")%>
							</td>
							<td >
								<textarea rows="4" name="DsParametro" id="DsParametro" onKeyPress="return controlCarracteresEspeciales(this, event);" style="width:310px"><%=replace(Trim(nomParam),"<br>",chr(13)&CHR(10))%></textarea>
								<input name="DsParametro_old" id="DsParametro_old" type="hidden" value='<%=nomParametro_old%>'>	</input>						
							</td>
						</tr>
						<tr>
							<td class="reg_header_nav">
								<%=GF_Traducir("Valor")%>
							</td>
							<td>			
								<input name="ValParametro" id="ValParametro" type="text" onKeyPress="return controlCarracteresEspeciales(this, event);" value="<%=Trim(valParam)%>">
								<input name="ValParametro_old" id="ValParametro_old" type="hidden" value="<%=ValParametro_old%>">
							</td>
						</tr>	
						<%if(visitante = TASK_PARAM_ADMIN)then
							if(esmodificable) then 
							'solo muestra si es editable y el puesto a aquellos que son administradores '
								if(v_tieneExtra)then
									'Si tiene parametros extras los cargo en el textbox y el radio
									set rsEdit = traerParametrosEditables(codParam, pto)
									idPuesto_old = cint(rsEdit("PUESTO"))																		
								else
									'si no tiene parametrosextra tienen que aparecer limpios
									esmodificable = false
									Check_Edit = PARAMETRO_NO_EDITABLE
								end if	
							end if%>					
						<tr>						
							<td class="reg_header_nav">						
								<%=GF_Traducir("Editable")%>
							</td>
							<td>
								<input type="hidden" id="Puesto_old" name="Puesto_old" value=<%=idpuesto_old%>>						
									<%if(esmodificable) then
										if(rsEdit("EDITABLE") = PARAMETRO_EDITABLE)then 
											Check_Edit = PARAMETRO_EDITABLE
										else
											Check_Edit = PARAMETRO_NO_EDITABLE
										end if								
									end if%>
								<input style="border:none;cursor:pointer;" type="checkbox" name="Check_Edit" id="Check_Edit" value="<%=Check_Edit%>" <%if(Check_Edit = PARAMETRO_EDITABLE)then Response.Write "checked"%>><%=GF_Traducir("SI")%>
								<input type="hidden" name="Check_Edit_old" id="Check_Edit_old" value=<%=Check_Edit%>>	
							</td>
						</tr>
						<tr>						
							<td class="reg_header_nav">
								<%=GF_Traducir("Puesto")%>
							</td>
							<td>
							<%Set rsPue = leerPuestos(pto)%>							
							<select style="z-index:-1;" name="idPuesto" id="idPuesto">
								<%if(esmodificable = false)then %>
									<option SELECTED value="0">-<% =GF_TRADUCIR("Seleccione")%>-
								<%end if%>
								<%while (not rsPue.eof)		
									selected = ""		
									if(esmodificable) then 	
										if (Cint(rsPue("IDPUESTO")) = Cint(rsEdit("PUESTO"))) then 
											selected = "selected"										
										end if		
									end if%>
									<option value="<% =rsPue("IDPUESTO")%>" <% =selected %>><% =rsPue("DSPUESTO") %>                                        
								<%rsPue.MoveNext()
								 wend%>
							</select>
							</td>
						</tr>	
						<%end if%>						
						<tr>
							<td colspan="2" align="center">
								<%if(accion = ACCION_MODIFICAR_PARAMETRO )then%>
									<input type="SUBMIT" value="Aceptar" name="Modificar" onclick="modificarParametro(<%=accion%>,'<%=pto%>','<%=codParam%>');">
								<%else%>
									<input type="SUBMIT" value="Aceptar" name="Agregar" onclick="agregarParametro('<%=pto%>');">
								<%end if%>			
							</td>						
						</tr>
					</table>
				</td>
			</tr>			
		</table>
	</form>
</body>
</html>

