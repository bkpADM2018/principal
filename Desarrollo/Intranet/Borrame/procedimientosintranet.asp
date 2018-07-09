<%
'                   ***   PROCEDIMIENTOS  INTRANET   ***
'                   ***   Autor: Javier A. Scalisi   ***

function LF_CONTROL_USUARIO(ByRef Dic,ByRef blnCrearUsuario, ByRef strErrorMsg)
' Este procedimientos es el encargado de controlar los datos de un usuario de la 
' INTRANET.

Dim intKR, strDS, dteFecha, bolControlFecha

'Control de los campos para que esten completos aquellos que son obligatorios.
   if (Dic("strID")="") or (GF_MGC("SG",Dic("strID"),intKR,strDS) and (Dic("hdnOperacion") = "ADD")) then strErrorMsg= "+ " & GF_TRADUCIR("Nombre de Usuario ya existe") & ".<br>"
   if (Dic("strPrimerNombre")="?") or (Dic("strPrimerNombre")="") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("Debe ingresar un Nombre") & ".<br>"
   if (Dic("strApellido")="?") or (Dic("strApellido")="") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("Debe ingresar un Apellido") & ".<br>"
   if (Dic("cmbEmpresa")="0") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("La persona debe pertenecer a una Empresa") & ".<br>"
   if (Dic("cmbSector")="0") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("La persona debe pertenecer a un Sector") & ".<br>"
   if (Dic("cmbCargo")="0") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("La Persona debe poseer un Cargo en la Empresa") & ".<br>"
   'Se controlan los campos optativos.
   if (Dic("dteFechaNac") <> "") and (Dic("dteFechaNac") <> "?") then
      'Solo si se asigno una fecha valida se controla.
      dteFecha =Dic("dteFechaNac")
      bolControlFecha = GF_CONTROL_FECHA_2(dteFecha)
      Dic("dteFechaNac") = dteFecha
      if not(bolControlFecha) then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("Fecha de Nacimiento Incorrecta") & "."
   end if
   Dic("strPwd")= Trim(Dic("strPwd"))
   blnCrearUsuario= false
   if (Dic("strPwd") <> "?") and (Dic("strPwd") <> "") then  
      'Controlo el password si se asigno uno valido.
      if (Dic("strPwd") <> Dic("strCnfPwd")) then 
	     strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("El password es incorrecto") & ".<br>"
	  else
		 blnCrearUsuario=true
	   end if    
   end if	
end function
'--------------------------------------------------------------------------------------------
function LF_GUARDAR_USUARIO(ByRef Dic, byref strErrorMsg,byref clsTD,p_accion)
' Este procedimientos es el encargado de guardar los datos de un usuario de la 
' INTRANET.

  Dim blnCrearUsuario,bolControlFecha
  Dim intKRSG,int3OKR,intKRSR,intKRAux
  Dim strSQL,con, rs

  	
   LF_GUARDAR_USUARIO= false
   'Controlo los campos.
   strErrorMsg=""
   LF_CONTROL_USUARIO Dic,blnCrearUsuario,strErrorMsg
   clsTD="TDERROR"
   if (strErrorMsg = "") then
            'Se guardan los datos.    	    
            if (p_accion = "UPD") and (Dic("usrold") <> Dic("strID")) then
		strSQL="Update MG set MG_KC='" & Dic("strID") & "' where MG_KM='SG' and MG_KC='" & Dic("usrold") & "'"				
		Call GF_BD_CONTROL("",con,"EXEC",strSQL) 
            END IF
		GF_mgADD "SG",Dic("strID"),Dic("strApellido") & ", " & Dic("strPrimerNombre"),intKRSG   
 		'Se da de alta la relacion con el sector.
		strSQL = "Select * from RelacionesConsulta where SRO1KR = " & GF_SESSIONKR("SR","SGSS") & " and SRO2KR= " & GF_SESSIONKR("SG",Dic("strID")) & " and SRValor <> '*'"
		Call GF_BD_CONTROL(rs,con,"OPEN",strSQL)		
    		if(not rs.eof) then									
			GF_MGSRADD  rs("SRO1KR"), rs("SRO2KR"), rs("SRO3KR"),"*",rs("SR3OKR")  
		end if
		GF_MGSRADD  GF_SESSIONKR("SR","SGSS"), GF_SESSIONKR("SG",Dic("strID")), Dic("cmbSector"),1,""
		'Se da de alta la relacion con el cargo.
		strSQL = "Select * from RelacionesConsulta where SRO1KR = " & GF_SESSIONKR("SR","SGUC") & " and SRO2KR= " & GF_SESSIONKR("SG",Dic("strID")) & " and SRValor <> '*'"
		Call GF_BD_CONTROL(rs,con,"OPEN",strSQL)		
		if(not rs.eof) then									
			GF_MGSRADD  rs("SRO1KR"), rs("SRO2KR"), rs("SRO3KR"),"*",rs("SR3OKR")  
		end if
		GF_MGSRADD  GF_SESSIONKR("SR","SGUC"), GF_SESSIONKR("SG",Dic("strID")), Dic("cmbCargo"),1,""							
		GF_DT1W "SGNAME","SG",Dic("strID"),Dic("strPrimerNombre"),""	   
		GF_DT1W "SGAPE","SG",Dic("strID"),Dic("strApellido"),""	
		if (Dic("dteFechaNac") <> "?") and (Dic("dteFechaNac") <> "") then GF_DT1W "SGFCNC","SG",Dic("strID"),Dic("dteFechaNac"),""
	        'Para el alta de la empresa se da de alta la relacion correspondiente.
		GF_MGC "SR","TRABAJAEN",intKRSR,strDS
		GF_MGSRADD intKRSR, intKRSG, Dic("cmbEmpresa"), 9, int3OKR
		if (Dic("strEmail") <> "?") and (Dic("strEmail") <> "") then GF_DT1W "SGEMAIL","SG",Dic("strID"),Dic("strEmail"),""
	  	if (Dic("strTelOf") <> "?") and (Dic("strTelOf") <> "") then GF_DT1W "SGTEOF","SG",Dic("strID"),Dic("strTelOf"),""
     		if (Dic("strTelPar") <> "?") and (Dic("strTelPar") <> "") then GF_DT1W "SGTENR","SG",Dic("strID"),Dic("strTelPar"),""
      		if (Dic("strFax") <> "?") and (Dic("strFax") <> "") then GF_DT1W "SGTEFX","SG",Dic("strID"),Dic("strFax"),""
    		if (Dic("strInterno") <> "?") and (Dic("strInterno") <> "") then GF_DT1W "SGNRIN","SG",Dic("strID"),Dic("strInterno"),""
    		if (Dic("strCelular") <> "?") and (Dic("strCelular") <> "") then GF_DT1W "SGTECL","SG",Dic("strID"),Dic("strCelular"),""
            	if (blnCrearUsuario) then
   			if (p_accion = "UPD") and (Dic("usrold") <> Dic("strID")) then
				strSQL="Update MG set MG_KC='" & Dic("strID") & "' where MG_KM='UP' and MG_KC='" & Dic("usrold") & "'"				
				Call GF_BD_CONTROL("",con,"EXEC",strSQL) 
			END IF
			GF_mgADD "UP",Dic("strID"),Dic("strApellido") & ", " & Dic("strPrimerNombre"),""			   
			GF_DT1W "UPPSWR","UP",Dic("strID"),Dic("strPwd"),""  
   		end if   
		strErrorMsg= "Se concluyo la operacion con EXITO."		
		clsTD="TDNOHAY"
		LF_GUARDAR_USUARIO=true
   end if		

end function
'---------------------------------------------------------------------------------------------
function LF_CONTROL_USUARIO_TEMP(ByRef Dic, ByRef strErrorMsg)

Dim intKR, strDS

LF_CONTROL_USUARIO_TEMP=false
'Control de los campos para que esten completos aquellos que son obligatorios.
   if (Dic("strID")="") or (GF_MGC("SG",UCASE(Dic("strID")),intKR,strDS)) then strErrorMsg= "+ " & GF_TRADUCIR("Nombre de Usuario NO Valido") & ".<br>"
   if (Dic("strPrimerNombre")="") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("Debe ingresar su primer nombre") & ".<br>"
   if (Dic("strApellido")="") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("Debe ingresar su apellido") & ".<br>"
   if not(GF_CONTROL_CUIT(Dic("strCUIT1")& "-" & Dic("strCUIT2")& "-" & Dic("strCUIT3"))) then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("El numero de CUIT no es valido") & ".<br>"
   if (Dic("cmbSector")="?") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("La persona debe pertenecer a un Sector") & ".<br>"
   if (Dic("cmbCargo")="?") then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("La Persona debe poseer un Cargo en la Empresa") & ".<br>"
   if not(GF_CONTROL_EMAIL(Dic("strEmail"))) then strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("El E-Mail ingresado no es valido") & ".<br>"
   Dic("strPwd")= Trim(Dic("strPwd"))
   'Controlo el password si se asigno uno valido.
   if not(GF_CONTROL_PASSWORD(Dic("strPwd"))) or (Dic("strPwd") <> Dic("strCnfPwd")) then 
	  strErrorMsg= strErrorMsg & "+ " & GF_TRADUCIR("El password es incorrecto") & ".<br>"
   end if	
   if (strErrorMsg = "") then LF_CONTROL_USUARIO_TEMP=true
end function
'--------------------------------------------------------------------------------------------
' Este procedimientos es el encargado de guardar los datos de un usuario externo en la 
' tabla de personas temporales.
function LF_GUARDAR_USUARIO_TEMP(ByRef Dic, byref strErrorMsg)

  Dim strSQL,rs,oConn

   LF_GUARDAR_USUARIO_TEMP= false
   'Controlo los campos.
   strErrorMsg=""
   LF_CONTROL_USUARIO_TEMP Dic,strErrorMsg
   if (strErrorMsg = "") then   
	  'Controlo que el usuario elegido tampoco exista en la tabla de personas temporales
      strSQL = "Select strID from PersonasTemp where strID='" & Dic("strID") & "'"
      GF_BD_CONTROL rs,oConn,"OPEN",strSQL
      if not(rs.EOF) then strErrorMsg= "+ El Nombre de Usuario ya Existe"
   end if
   if (strErrorMsg = "") then
      'Se guardan los datos.
      strSQL = "Insert into PersonasTemp(strID,strPrimerNombre,strApellido,cmbSector,cmbCargo"
	  strSQL = strSQL & ",strCUIT,strEmail,strPWD,MmtoRegistro)"
	  strSQL = strSQL & "values('" & UCASE(Dic("strID")) & "'"
	  strSQL = strSQL & ",'" & Dic("strPrimerNombre") & "'"
	  strSQL = strSQL & ",'" & Dic("strApellido") & "'"
	  strSQL = strSQL & ",'" & Dic("cmbSector") & "'"
	  strSQL = strSQL & ",'" & Dic("cmbCargo") & "'"
	  strSQL = strSQL & ",'" & Dic("strCUIT1") & "-" & Dic("strCUIT2") & "-" & Dic("strCUIT3") & "'"
  	  strSQL = strSQL & ",'" & Dic("strEmail") & "'"
	  strSQL = strSQL & ",'" & Dic("strPwd") & "'"
	  strSQL = strSQL & "," & GF_MOMENTOSISTEMA() & ")"
	  GF_BD_CONTROL "",oConn,"EXEC",strSQL
	  strErrorMsg= "Se concluyo la operacion con EXITO."		
	  LF_GUARDAR_USUARIO_TEMP=true
   end if		

end function
'---------------------------------------------------------------------------------------------
function LF_TRANSFERIR_USUARIO(Dic)

Dim strSQL,oConn
Dim strErrorMsg,clsTD
'Se marca el alta del usuario.
strSQL = "UPDATE PersonasTEMP SET MRCALTA=1 where strID='" & Dic("strID") & "'"
GF_BD_CONTROL "",oConn,"EXEC",strSQL
'Se guarda el usuario
'*** REVISAR ***
'LF_GUARDAR_USUARIO Dic,strErrorMsg,clsTD
end function
'---------------------------------------------------------------------------------------------
function LF_ELIMINAR_USUARIO_TEMP(P_ID)

Dim strSQL,oConn

strSQL="Delete from PersonasTemp where strID='" & P_ID & "'"
GF_BD_CONTROL "",oConn,"EXEC",strSQL

end function
'---------------------------------------------------------------------------------------------
function LF_ACTUALIZAR_USUARIO_TEMP(Dic, ByRef strErrorMsg)

   Dim strSQL,oConn
   Dim aux
   
   strErrorMsg=""
   aux=Dic("strID")
   Dic("strID")="@$@" 'Nombre de usuario que no puede existir nunca.
   LF_CONTROL_USUARIO_TEMP Dic,strErrorMsg
   if (strErrorMsg = "") then
      'Se guardan los datos.
      strSQL = "Update PersonasTemp set"
	  strSQL = strSQL & " strPrimerNombre='" & Dic("strPrimerNombre") & "'"
	  strSQL = strSQL & ", strApellido='" & Dic("strApellido") & "'"
	  strSQL = strSQL & ", cmbSector='" & Dic("cmbSector") & "'"
	  strSQL = strSQL & ", cmbCargo='" & Dic("cmbCargo") & "'"
	  strSQL = strSQL & ", strCUIT='" & Dic("strCUIT1") & "-" & Dic("strCUIT2") & "-" & Dic("strCUIT3") & "'"
	  strSQL = strSQL & ", strEmail='" & Dic("strEmail") & "'"
	  strSQL = strSQL & ", strPWD='" & Dic("strPwd") & "'"
	  strSQL = strSQL & " where strID='" & aux & "'"
	  GF_BD_CONTROL "",oConn,"EXEC",strSQL
   end if		
  
end function
'---------------------------------------------------------------------------------------------
sub LP_NOTIFICAR_REGISTRO(P_ID)
    
Dim strTexto,strAsunto
Dim rs,oConn,strSQL

'Obtengo los datos del usuario
strSQL="select * from PersonasTemp where strID='" & P_ID & "'"
GF_BD_CONTROL rs,oConn,"OPEN",strSQL
'Se arma el asunto del Mail.
strAsunto=GF_TRADUCIR("Confirmacion de Registro a ActiSA")
'Se arma el cuerpo del Mail.
strTexto=GF_TRADUCIR("Estimado") & " " & rs("strPrimerNombre") & " " & rs("strApellido") & ":" & vbCrLf & vbCrLf
strTexto=strTexto & GF_TRADUCIR("Tenemos el agrado de informarle que ha sido dado de alta")
strTexto=strTexto & GF_TRADUCIR(" en nuestro sistema y ya puede acceder a todos los servivios de Nuestra Empresa.")& vbCrLf & vbCrLf
strTexto=strTexto & GF_TRADUCIR("Su informacion de acceso es la siguiente") & ":" & vbCrLf
strTexto=strTexto & GF_TRADUCIR("Usuario") & ":" & rs("strID") & vbCrLf
strTexto=strTexto & GF_TRADUCIR("Contraseña") & ":" & rs("strPWD") & vbCrLf & vbCrLf
strTexto=strTexto & GF_TRADUCIR("Lo saluda muy atte.")
'Se envia el mail.
GP_ENVIAR_MAIL strAsunto,strTexto,GF_DT1("READ","SGEMAIL","","","SG",session("Usuario")),rs("strEmail")
'Marco el registro como notificado.
strSQL="Update PersonasTemp set MRCNOTIFICACION=1 where strID='" & P_ID & "'"
GF_BD_CONTROL "",oConn,"EXEC",strSQL

end sub
%>