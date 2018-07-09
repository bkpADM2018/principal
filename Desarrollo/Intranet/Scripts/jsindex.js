$(document).ready(function() {
    $('#mainMenu span')		
		.click(function() {			
			var source = "public/" + $(this).parent().attr('id') + ".asp";
			$("#mainContent").css('display', 'none');
			$("#mainContent").load(source);
			$("#mainContent").delay(0).fadeIn(900);					   
		});		    
	$("#mainContent").css('display', 'none');
	$("#mainContent").load('public/section_home.asp');
	$("#mainContent").delay(0).fadeIn(900);
});

function checkEnter(e) {
	var ascii = (document.all) ? e.keyCode : e.which;			
	if (ascii == 13) login();		
}

function login(){
		if ($("#Password").length) {			
			var hashPass = MD5(MD5($("#Password").val()) + $("#llave").val());
			var user     = $("#Username").val();			
			$.ajax({
					url: "public/loginAjax.asp?user="+ user +"&pass="+hashPass,
					//Debug - url: "public/loginAjax.asp?user="+ user +"&pass="+hashPass+"&p=" + MD5($("#Password").val()) + "&ll=" + $("#llave").val(),
					context: document.body,
					success: function(data){
						var myJson = eval("("+data+")");
						if (myJson["error"] == ""){
							window.parent.location.href = myJson["url"];						
						}
						else{						
							$("#llave").val(myJson["llave"]);
							$("#msg").addClass("txtMsgError")
							$("#msg").html(myJson["error"]);															
						}					
					}
			});
		}
}