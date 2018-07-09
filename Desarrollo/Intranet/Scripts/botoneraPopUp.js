var btn_botonera_class = "botonera";
var btn_botones_class = "botones";

function botonera(idDiv)
{
	this.idDiv = idDiv;
	this.htmlButtons = "";
	
	this.addbutton = function(pNombre,pAction){
			this.htmlButtons += "&nbsp;<input type='button' value='"+pNombre+"' onclick='"+pAction+"'>";
			
		}
	this.addSwith = function(pNombre,pAction){
			this.htmlButtons += "&nbsp;<input type='checkbox' id='"+pNombre+"' onclick='"+pAction+"' /><label for='"+pNombre+"'>"+pNombre+"</label>";
		}
		
	this.show = function(){
			document.getElementById(this.idDiv).innerHTML = this.createDivs();
			$(function() {
				$( "input,checkbox", ".botones" ).button();
			});
		}
	
	this.createDivs = function(){
			var rtrn =  "<div class='"+btn_botonera_class+"'>"
			rtrn += 		"<div class='"+btn_botones_class+"'>";
			rtrn +=				this.htmlButtons;
			rtrn +=			"</div>";
			rtrn +=		"</div>";
			return rtrn;
		}
		
}