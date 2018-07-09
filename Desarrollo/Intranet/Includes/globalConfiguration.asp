<%
Const SITE_INTRANET = True 'True: Para el Server de la Intranet - False: Para el Server de la Extranet
Const SITE_ROOT = "/Intranet/"

Const INTRANET_FILE_PATH = "http://localhost/ActisaIntra/Temp/"

Const URL_INTRANET_ARROYO 	= "http://VN5SADATOSSQL1/ActisaIntra/"
Const URL_INTRANET_TRANSITO = "http://VN4SADATOSSQL1/ActisaIntra/"
Const URL_INTRANET_BAHIA 	= "http://VN7SADATOSSQL1/ActisaIntra/"

Const PATH_COMPRAS_TEMP =  "Temp"
Const PATH_COMPRAS_FINAL =  "Documentos\ArchivosCompras"

'Constantes para envio de mails.
Const SENDER_MERCADERIAS	= "mercaderias.ar@toepfer.com"
Const SENDER_IMPUESTOS		= "mercaderias.ar@toepfer.com"
Const SENDER_LICITACIONES	= "ArgentinaCompras@adm.com"
Const SENDER_LEGALES		= "ArgentinaLegal@adm.com"
Const SENDER_COTIZACIONES	= "ArgentinaCompras@adm.com"
Const SENDER_TESORERIA      = "Tesoreria.AR@toepfer.com"
Const SENDER_FACTURACION    = "Cobranzas.AR@toepfer.com"
Const MAILTO_INTEGRIDAD     = "Control.Analisis@toepfer.com"
Const MAILTO_COMPRAS		= "ArgentinaCompras@adm.com"
Const SENDER_CUPOS_RESOLUCION_25_13	= "cosechagruesa@santafe.gov.ar"

Const MAILTO_SM_AS     = "VN5Mantenimiento@adm.com"
Const MAILTO_SM_ET     = "VN4Mantenimiento@adm.com"
Const MAILTO_SM_BB     = "VN7Mantenimiento@adm.com"

Const cdoSmtpUseSSL		        = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
Const cdoSendUsingMethod        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSendUsingPort          = 2
Const cdoSMTPServer             = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPServerPort         = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoSMTPConnectionTimeout  = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Const cdoSMTPAuthenticate       = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Const cdoBasic                  = 0 'CDO_SMTP_AUTHENTICATE
Const cdoAdvance                = 1 'CDO_SMTP_AUTHENTICATE
Const cdoSendUserName           = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Const cdoSendPassword           = "http://schemas.microsoft.com/cdo/configuration/sendpassword"


Const CDO_SMTP_USE_SSL = 0
Const CDO_SMTP_SERVER = "smtp.adm.com"
Const CDO_SMTP_PORT = 25
Const CDO_SMTP_AUTHENTICATE = 0 
Const CDO_SEND_USER_NAME = ""
Const CDO_SEND_PASSWORD = ""

Const DBSITE_SQL_INTRA = "SQLCS"
Const DBSITE_SQL_SAF = "SAF"
Const DBSITE_SQL_MAGIC = "MAGIC"
Const DBSITE_ARROYO = "ARROYO"
Const DBSITE_TRANSITO = "TRANSITO"
Const DBSITE_BAHIA = "PIEDRABUENA"
%>
