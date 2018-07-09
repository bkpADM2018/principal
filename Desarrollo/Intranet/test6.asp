<!--#include file="Includes/procedimientosUnificador.asp"-->
<%


strAsunto = "Prueba de Asunto"
strBody = "Prueba de Cuerpo" & vbcrlf & "Prueba de Cuerpo"
att = "'http://bai-vm-intra-1/Intranet/temp/CARTASPORTE_20180118.txt'"
%>

<html>
<head>
<script type="text/javascript">
	function sendEmail(){
          var theApp = new ActiveXObject("Outlook.Application");
       try{
          var objNS = theApp.GetNameSpace('MAPI');
          var theMailItem = theApp.CreateItem(0); // value 0 = MailItem
          theMailItem.to = ('test@gmail.com');
          theMailItem.Subject = ('test');
          theMailItem.Body = ('test');
          theMailItem.Attachments.Add("http://bai-vm-intra-1/Intranet/temp/CARTASPORTE_20180118.txt");
          theMailItem.display();
      }
      catch (err) {
         alert(err.message);
      } 
   }
</script>
</head>
<body>
	<a href="mailto:scalisij@toepfer.com; Ruben.Alessio@adm.com; MariaSoledad.Baccaro@adm.com; Hugo.Cosentino@adm.com; Daniel.Gonzalez@adm.com; FedericoTomas.Salinas@adm.com?subject=<% =strAsunto %>&body=<% =strBody %>&attachment=<% =att %>"> Enviar </a>
	<br><br><br><br><br>
	<a href="javascript:sendEmail()"> Enviar JS </a>
</body>
</html>