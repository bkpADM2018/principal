<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosHKEY.asp"-->
<%

'---------------------------------------------------------------------------------------------
Function leerRegistroFirmas()
	Dim conn, strSQL, rs, ret, km, ds
		
	ret = false
	gCdUsuario = ""
	if (HK_isKeyReady()) then
		gMyKey = HK_readKey()
		strSQL = "Select * from TBLREGISTROFIRMAS where HKEY='" & gMyKey & "'"
		Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.eof) then
			ret = True
			gCdUsuario = UCase(rs("CDUSUARIO"))						
		else
			gCdUsuario = ""
		end if
	end if
	leerRegistroFirmas = ret
End Function
'---------------------------------------------------------------------------------------------
'	COMEINZO DE PAGINA
'---------------------------------------------------------------------------------------------

	Dim acc, gCdUsuario, gMyKey, msg
	
	acc = GF_PARAMETROS7("a", "", 6)
	
	if(acc = "F") then	
		'Firma		
		if (leerRegistroFirmas()) then
			msg = "Llave leida: " & gMyKey & ", pertenece al usuario: " & gCdUsuario
		else
			msg = "Llave leida: " & gMyKey & ", no se encuentra usuario relacionado."
		end if
		Call HK_sendResponse(msg)
		response.end
	end if
%>

<html>
	<head>
		<script type="text/javascript" src="scripts/channel.js"></script>
		<script type="text/javascript" src="scripts/hkey.js"></script>
		
		<script type="text/javascript">
			
			function bodyOnLoad() {
				var link = "test5.asp?a=F";	
				var hkey1 = new Hkey('hk1', link, '<% =HKEY() %>', 'check_callback()');
				hkey1.start();
			}
			
			function check_callback(resp) {
				alert(resp);
			}	
		</script>
		
	</head>
	<body onLoad="bodyOnLoad()">				
		<div id='hk1'></div>		
	</body>
</html>