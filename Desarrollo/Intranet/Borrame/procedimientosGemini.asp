<!--#include file="procedimientosConexion.asp"-->
<%
'CONTANTES PARA MIGRATION TEMPLATE (template 1)
Const TEMPLATE_MIGRATION = 1

'CONSTANTES PARA BUG TRACKING (NUEVO TEMPLATE 13) 
Const TEMPLATE_BUG_TRACKING = 13


const ISSUE_STATUS_UNASSIGNED 	= 31
const ISSUE_STATUS_ASSIGNED   	= 32
const ISSUE_STATUS_IN_PROGRESS 	= 33
const ISSUE_STATUS_CLOSE 		= 34
const ISSUE_STATUS_SIGN 	    = 57
const ISSUE_STATUS_TESTING 		= 59
const ISSUE_STATUS_FAIL_TEST 	= 60
const ISSUE_STATUS_APP_USUARIO  = 58
const ISSUE_STATUS_CANCELED 	= 62

Const ISSUE_SOLU_CANCELED 	= 24
Const ISSUE_SOLU_NO_SOLUTION= 14
Const ISSUE_SOLU_DUPLICATE 	= 23
'Const ISSUE_SOLU_NO_ERROR 	= 4 'NO DEFINIDA
Const ISSUE_SOLU_FINISH 	= 15
Const ISSUE_SOLU_SUSPENDED 	= 25

'CONSTANTES PARA HELP DESK (NUEVO TEMPLATE 12) 
Const TEMPLATE_HELP_DESK = 12

const ISSUE_STATUS_HD_UNASSIGNED 	= 27
const ISSUE_STATUS_HD_ASSIGNED   	= 30
const ISSUE_STATUS_HD_IN_PROGRESS 	= 28
const ISSUE_STATUS_HD_CLOSE 		= 29
const ISSUE_STATUS_HD_APR_TECH      = 65
const ISSUE_STATUS_HD_CANCELED 	    = 46

Const ISSUE_SOLU_HD_UNSOLVED 	= 12
Const ISSUE_SOLU_HD_COMPLETED   = 13
Const ISSUE_SOLU_HD_REJECTED 	= 27

Const CONEXION_GEMINI = "GEMINI"

'----------------------------------------------------------------------------------------
Function GF_BD_GEMINI(byref pRs, pOperacion, pSql)
on error resume next
	GF_BD_GEMINI = false
	session("strSQL") = pSql	
	if(IsEmpty(session("conn" & CONEXION_GEMINI &  "Alias")))then	Call loadConfigFile(CONEXION_GEMINI)		
	select case ucase(pOperacion)
		case "CLOSE"
			executeQuery = true
		case "OPEN" 
			set con = server.CreateObject("ADODB.connection")
			set pRS = server.CreateObject("ADODB.Recordset")
			con.open session("conn" & CONEXION_GEMINI &  "Alias"), session("conn" & CONEXION_GEMINI &  "User"), session("conn" & CONEXION_GEMINI &  "Key")
            pRS.CursorLocation = 3
			pRS.Open pSql, con, 2, 3, 1
			GF_BD_GEMINI = true
		case "EXECUTE", "EXEC" 
			if pSql <> "" then 
				Set con = server.CreateObject("ADODB.connection")
				con.open session("conn" & CONEXION_GEMINI &  "Alias"), session("conn" & CONEXION_GEMINI &  "User"), session("conn" & CONEXION_GEMINI &  "Key")
                con.execute pSql
				con.close
				GF_BD_GEMINI = true
			end if
		case "EXISTS"
			if pSql <> "" then 
				dim mySQL
				set con = server.CreateObject("ADODB.connection")
				set pRS = server.CreateObject("ADODB.Recordset")
				con.open session("conn" & CONEXION_GEMINI &  "Alias"), session("conn" & CONEXION_GEMINI &  "User"), session("conn" & CONEXION_GEMINI &  "Key")
				pRS.CursorLocation = 3
				mySQL = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='" & pSQL & "'"
				pRS.Open mySQL, con, adOpenDynamic, adLockOptimistic, adCmdText
				if not pRS.eof then GF_BD_GEMINI = true
			end if			
	end select	
	if err.number <> 0 and err.number <> 424 then
		Response.Write "<br><b><font color=red>HA OCURRIDO UN ERROR!</font><b><br>POR FAVOR REVISE EL SQL<Hr>" & pSql & "<hr>Error:" & err.number   & err.Description 
		err.Clear 
		response.end 
	end if
end function

'--------------------------------------------------------------------------------------------------
Function getIssue(pId)
	Dim strSQL,rs
		
    strSQL = "select  gi.issueid,gi.projectid,gi.fixedinversionid,gi.reportedby,gu.username, gi.summary,gi.longdesc,gi.issuetypeid,gi.issuepriorityid,gi.issueseverityid,gi.issuestatusid,gi.issueresolutionid,gi.userdata1,gi.userdata2,gi.userdata3,gi.percentcomplete,gi.estimatehours,gi.estimateminutes,gi.startdate,gi.duedate,gi.visibility,gi.parentissueid,gi.created,gi.revised,gi.resolveddate,gi.closeddate,UPPER(SUBSTRING(RTRIM(gi.originatordata),CHARINDEX('|', RTRIM(gi.originatordata),1)+1,LEN(RTRIM(gi.originatordata)))) as mailUsuario,gi.creator, gp.templateid " &_    
	         " from gemini_issues gi " &_ 
	         "  inner join gemini_users gu on gi.reportedby=gu.userid " &_
	         "  inner join gemini_projects gp on gp.projectid = gi.projectid " &_
	         " where issueid = " & pId	         
	Call GF_BD_GEMINI(rs, "OPEN", strSQL)
	
	set getIssue = rs
	
End Function
'--------------------------------------------------------------------------------------------------
Function getGeminiUserFullName(pId)
	Dim usrName,usrLName
	
	call getGeminiUserDs(pId,usrName,usrLName)
	if (usrLName = "" and usrName = "") then
		rtrn = ""
	else
		rtrn = usrLName & "," & usrName
	end if
	
	getGeminiUserFullName =  rtrn
End function
'--------------------------------------------------------------------------------------------------
Function getGeminiUserName(pId)
	Dim rtrn
	
	call getGeminiUserDs(pId,"",rtrn)
	
	getGeminiUserName = rtrn
End function
'--------------------------------------------------------------------------------------------------
'			*************				Funcion:  getGeminiUserDs				*****************
'*************** Le paso la ID del Usuario y me devuelve el Nombre y Apellido ****************
Function getGeminiUserDs(pId, ByRef pName, ByRef pLName)
	Dim strSQL,rs
	
	if (pId <> "") then
		strSQL = "select firstname,surname from gemini_users where userid = " & pId
		Call GF_BD_GEMINI(rs, "OPEN", strSQL)
		
		if (not rs.EoF) then
			pName = trim(rs("firstname"))
			pLName = trim(rs("surname"))
		end if
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function getGeminiUserCD(pId)
	Dim strSQL,rs,rtrn
	
	if (pId <> "") then
		'obtengo el id de usuario'
		strSQL = "SELECT * FROM gemini_issueresources where issueid =" & pId  
		call GF_BD_GEMINI(rs, "OPEN", strSQL)
		rtrn = ""
		if (not rs.EoF) then rtrn = rs("userid")

		if (rtrn <> "") then
			strSQL = "select * from gemini_users where userid = " & rtrn
			Call GF_BD_GEMINI(rs, "OPEN", strSQL)
			rtrn = ""
			if (not rs.EoF) then rtrn = UCase(rs("username"))
		end if
	end if

	getGeminiUserCD = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function getIssueStatusDs(pId)
	Dim strSQL,rs,rtrn
	
	strSQL = "select statusdesc from gemini_issuestatus where statusid = " & pId
	Call GF_BD_GEMINI(rs, "OPEN", strSQL)
	
	rtrn = ""
	if (not rs.EoF) then
		rtrn = rs("statusdesc")
	end if 
	
	getIssueStatusDs = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function getIssueField(pField,pId)
	Dim strSQL,rs,rtrn

	strSQL = "select "&pField&" from gemini_issues where issueid = " & pId
	Call GF_BD_GEMINI(rs, "OPEN", strSQL)
	
	rtrn = ""
	if (not rs.EoF) then
		rtrn = rs(pField)
	end if 
	
	getIssueField = rtrn

End Function 
'--------------------------------------------------------------------------------------------------
Function getIssueStatus(pId)
	Dim aux
	aux = getIssueField("issuestatusid",pId)
	if (aux <> "") then
		getIssueStatus = cdbl(aux)
	else
		getIssueStatus = 0
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function getIssueTitle(pId)
	getIssueTitle = getIssueField("summary",pId)
End Function
'--------------------------------------------------------------------------------------------------
Function getIssueProject(pId)
	getIssueProject = getIssueField("projectid",pId)
End Function
'--------------------------------------------------------------------------------------------------
Function getIssueDesc(pId)
	getIssueDesc = getIssueField("longdesc",pId)
End Function
'--------------------------------------------------------------------------------------------------
Function getIssues(pStatus)
	Dim strSQL,rs
	
	strSQL = "select * from gemini_issues "
	if (pStatus <> 0) then strSQL = strSQL & "where issuestatusid " = pStatus
	Call GF_BD_GEMINI(rs, "OPEN", strSQL)
	
	set getIssues = rs
End Function
'--------------------------------------------------------------------------------------------------
Function updateIssueStatus(pId,pStatus,pSolution)
	Dim strSQL,rs, myFecha
	
	strSQL = "update gemini_issues set issuestatusid="&pStatus
	if (pSolution <> "") then
		strSQL = strSQL & ", issueresolutionid="&pSolution
	end if
	if (CInt(pStatus) = ISSUE_STATUS_CLOSE) then
	    'Se cerro la tarea, pongo la fecha del cierre.	    
	    myFecha = left(session("MmtoDato"),4) & "-" & Mid(session("MmtoDato"), 5, 2) & "-" & Mid(session("MmtoDato"), 7, 2) & " " & Mid(session("MmtoDato"), 9, 2) & ":" & Mid(session("MmtoDato"), 11, 2) & ":" & Right(session("MmtoDato"), 2) & ".0"
        strSQL = strSQL & ", closeddate='" & myFecha & "'"	    
	end if
	strSQL = strSQL & " where issueid = " & pId
	Call GF_BD_GEMINI(rs, "EXEC", strSQL)
End Function
'--------------------------------------------------------------------------------------------------
Function getGeminiTaskCode(pId)
	Dim strSQL,rs,rtrn
	
	strSQL = "select p.projectcode + '-' + cast(i.issueid as varchar) geminitask from gemini_issues i "
	strSQL = strSQL & "inner join gemini_projects p on i.projectid = p.projectid "
	strSQL = strSQl & "where i.issueid = " & pId
	Call GF_BD_GEMINI(rs, "OPEN", strSQL)
	
	rtrn = ""
	if (not rs.EoF) then rtrn = rs("geminitask")
	
	getGeminiTaskCode = rtrn
End Function
'--------------------------------------------------------------------------------------------------
' Autor: Nahuel Ajaya
' Fecha: 03/11/2011
' Objetivo:
'                                   Me devuelve el Nombre y Apellido del usuario de una cierta Tarea
' Parametros:
'                                   [int]      pId  : IdTarea
' Devuelve:
'                                   [int]      
' Modificaciones:
'                                   --/--/-- - XXX
Function getUserAndIssue(pId)
Dim rs, strSQL, rtrn
strSQL = "SELECT * FROM gemini_issueresources where issueid =" & pId  
call GF_BD_GEMINI(rs, "OPEN", strSQL)
rtrn = ""
if (not rs.EoF) then rtrn = rs("userid")
getUserAndIssue = getGeminiUserFullName(rtrn)
end function
'--------------------------------------------------------------------------------------------------
' Autor: Guido Fonticelli - GFG
' Fecha: 15/11/2011
' Objetivo:
'			Obtener el porcentaje de completado de la tarea de gemini
' Parametros:
'			[int]	pIdIssue
' Devuelve:
'			[int]	porcentaje
' Nota:
'			Los porcentajes se estableceran de la siguiente forma
'
'			Estado		 |	Porcentaje
'			--------------------------
'			Sin asignar	 |	  0 %
'			Asignada	 |	  0 %
'			En Proceso	 | 	 50 %
'			Testing		 |	 50 %
'			Testing Fail |	 40 %
'			Testing Ok	 |	 50 %
'			A la firma	 |	 90 %
'			Cerrada		 |	100 %
'
' Modificaciones:
'			--/--/-- - XXX
'--------------------------------------------------------------------------------------------------
Function getIssueProgress(pIdIssue)
	Dim issueStatus,rtrn
	issueStatus = cdbl(getIssueStatus(pIdIssue))
	select case issueStatus 
		case ISSUE_STATUS_UNASSIGNED,ISSUE_STATUS_ASSIGNED
			rtrn = 0
		case ISSUE_STATUS_FAIL_TEST
			rtrn = 40
		case ISSUE_STATUS_IN_PROGRESS,ISSUE_STATUS_TESTING
			rtrn = 50
		case ISSUE_STATUS_SIGN
			rtrn = 90
		case ISSUE_STATUS_CLOSE
			rtrn = 100
	end select
	getIssueProgress = rtrn
End Function
'--------------------------------------------------------------------------------------------------
Function getLastVersionOrder(pIdProject)
	Dim strSQL,rs,rtrn
	strSQL = "select max(versionorder) last from gemini_versions where projectid = " & pIdProject
	call GF_BD_GEMINI(rs, "OPEN", strSQL)
	rtrn = 0
	if (not isnull(rs("last"))) then rtrn = rs("last")
	getLastVersionOrder = cdbl(rtrn)
End Function
'--------------------------------------------------------------------------------------------------
'Archiva la version del requerimiento creada por el Gemini, de esta manera no se utilizará de nuevo esta version
Function archivedGeminiVersion(pNroReq)
	Dim strSQL,rs,auxIdReq
	if (not DESARROLLO) then
		auxIdReq = cstr(pNroReq)
		for i = len(auxIdReq) to 2
			auxIdReq = "0" & auxIdReq
		next
				
		strSQL = "update gemini_versions set versionreleased = 'true',versionarchived = 'true' where versionnumber = 'REQ-" & auxIdReq & "'"
		call GF_BD_GEMINI(rs, "EXEC", strSQL)
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function createVersionToIssue(pIdIssue,pIdReq,pForzar)
	Dim strSQL,rs,issueProject,auxIdReq,version,versionOrder,idVer

	if (not DESARROLLO) then
		'verifico que el issue no posea version ya
		strSQL = "select count(*) cant from gemini_affectedversions where issueid = " & pIdIssue		
		call GF_BD_GEMINI(rs, "OPEN", strSQL)
		
		if (cdbl(rs("cant")) = 0 or pForzar = true) then
			Call GP_ConfigurarMomentos
			
			'creo la version
			issueProject = getIssueProject(pIdIssue)
			
			auxIdReq = cstr(pIdReq)
			for i = len(auxIdReq) to 2
				auxIdReq = "0" & auxIdReq
			next
			
			version = "REQ-"&auxIdReq
			versionOrder = getLastVersionOrder(issueProject) +1
			
			strSQL = "select versionId from gemini_versions where versionnumber = '"&version&"'"
			call GF_BD_GEMINI(rs, "OPEN", strSQL)
				
			idVer = 0
			if (not rs.EoF) then
				idVer = cdbl(rs("versionId"))
			else
				strSQL = "insert into gemini_versions (projectid,versionnumber,versionname,versiondesc,versionreleased,versionorder,versionarchived,created) "
				strSQL = strSQL & " values ("&issueProject&",'"&version&"','"&version&"','"&version&"','false',"&versionOrder&",0,'"&now()&"') "
				call GF_BD_GEMINI(rs, "EXEC", strSQL)
				
				'obtengo el id de version
				strSQL = "select versionid from gemini_versions where versionnumber = '" & version &"'"
				call GF_BD_GEMINI(rs, "OPEN", strSQL)

				if (not rs.EoF) then idVer = cdbl(rs("versionId"))
			end if
			
			if (idVer <> 0 ) then
				'creo la asociacion de la version con la tarea
				strSQL = "insert into gemini_affectedversions (issueid,versionid,created) " 
				strSQL = strSQL & " values ("&pIdIssue&","&idVer&",'"&now()&"')"
				call GF_BD_GEMINI(rs, "EXEC", strSQL)
				
				strSQL = "update gemini_issues set fixedinversionid = " & idVer & " where issueid = " & pIdIssue
				call GF_BD_GEMINI(rs, "EXEC", strSQL)
			end if
		end if
	end if
End Function
'--------------------------------------------------------------------------------------------------
Function deleteVersionToIssue(pIdIssue)
	Dim strSQL,rs
	if (not DESARROLLO) then
		strSQL = "delete from gemini_affectedversions where issueid = " & pIdIssue
		call GF_BD_GEMINI(rs, "EXEC", strSQL)
		
		strSQL = "update gemini_issues set fixedinversionid = null where issueid = " & pIdIssue
		call GF_BD_GEMINI(rs, "EXEC", strSQL)
	end if
End Function
'--------------------------------------------------------------------------------------------------
'Devuelve el mail del solicitante de una tarea
Function getIssueMailApplicant(p_IdTarea)
    Dim str
    'El campo mail del  solicitante es 'originatordata', esta se encuentra concatenado con el nombre del Usuario
    str = getIssueField("originatordata",p_IdTarea)
    if (instr(str, "|") > 0) then
        str = Split(str,"|")
        getIssueMailApplicant = Trim(Ucase(str(1)))
    end if
End function
'--------------------------------------------------------------------------------------------------
Function getDsProjectByIssue(p_IdTarea)
	Call GF_BD_GEMINI(rs, "OPEN", "SELECT PROJECTNAME FROM GEMINI_PROJECTS A INNER JOIN GEMINI_ISSUES B ON A.PROJECTID = B.PROJECTID WHERE B.ISSUEID ="& p_IdTarea)
	getDsProjectByIssue = ""
	if (not rs.EoF) then getDsProjectByIssue = rs("PROJECTNAME")
End function
'**********************************************************************************************************************************
'********************************************** FUNCIONES HEREDADAS DE REQUERIMIENTOS  ********************************************
'**********************************************************************************************************************************
'Devuelve las firmas que fueron registradas para una tarea, ordenado por secuencia
Function getIssueSign(pIdTarea)
	Dim strSQL,rtrn
	strSQL = "select idtarea, secuencia, cdusuario, "&_
             "      case when hkey is null then '' else hkey end as hkey, "&_
             "      case when mmto is null then '' else mmto end as mmto "&_
             "from tblsysfirmas "&_
             "where idTarea = " & pIdTarea &_
             " order by secuencia"
    Call executeQueryDb(DBSITE_SQL_INTRA, rtrn, "OPEN", strSQL)
	set getIssueSign = rtrn
End Function

%>