<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosmail.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="../Includes/procedimientosvalidacion.asp"-->
<% 

Dim idProveedor, accion, mail, strSQL, rs, orden, tarea

idProveedor = GF_PARAMETROS7("idProveedor", "", 6)
mail		= GF_PARAMETROS7("mail", "", 6)
accion		= GF_PARAMETROS7("accion", "", 6)
orden		= GF_PARAMETROS7("orden", "", 6)
Select case accion
	Case ACCION_BORRAR
		Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_DEL_BY_ID", orden)
		Response.Write RESPUESTA_OK
	Case ACCION_PROCESAR
		if (GF_CONTROL_EMAIL(mail)) then
		    Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_UPD_EMAIL_BY_IDTAREAMAIL", orden &"||"& Trim(mail))            
			Response.Write RESPUESTA_OK		
		else			
			Response.Write "La direccion ingresada no es valida"
		end if
	Case ACCION_GRABAR
		if (GF_CONTROL_EMAIL(mail)) then			
			Call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_INS", TASK_FAC_MAIL_ALERT &"||"& idProveedor &"||"& Trim(mail))
			Response.Write RESPUESTA_OK
		else
			Response.Write "La direccion ingresada no es valida"
		end if	
End Select
%>