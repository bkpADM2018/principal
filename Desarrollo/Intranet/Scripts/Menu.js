
var MENU_ITEM_NODE = "UC"
var MENU_ITEM_LEAF = "SP"

var MENU_LEVEL_1 = 1
var MENU_LEVEL_2 = 2


//Global System Menu used for internal job.
var $Menu;

var $menu_ridx;
var $menu_lidx;

var $menu_last_active = new Array(); //Vector con el último nodo activo de cada nivel.

//-----------------------------------------------------------------------
// MENU OBJECT
//-----------------------------------------------------------------------
function Menu(p_root)
{
// Atributes
this.Items= new Array();
this.reference= new Array();
this.root = p_root.toLowerCase() || "root";
// Methods
this.AddNode= mnuAddNode;
this.AddLeaf= mnuAddLeaf;
this.Draw= mnuInitDraw;
this.Run= mnuRun;
}
//-----------------------------------------------------------------------
function mnuAddNode(P_Parent,P_Name,P_Text,P_imgExpanded,P_imgCollapsed)
{	
	var level = MENU_LEVEL_1;	
	if (P_Parent != this.root) level = MENU_LEVEL_2;	
	this.Items.push(new Node(P_Parent,P_Name,P_Text, level));
	this.reference[P_Name.toLowerCase()] = this.Items.length-1;
	return this.Items.length-1;
}
//-----------------------------------------------------------------------
function mnuAddLeaf(P_Parent,P_Text,P_Link,P_Target,P_Status,P_Image)
{		
	this.Items.push(new Leaf(P_Parent,P_Text,P_Link,P_Target));
	return this.Items.length-1;
}
//-----------------------------------------------------------------------
function mnuInitDraw()
{
	$("body").append("<DIV id=\"div" + this.root + "\"></DIV>");
	//Se agrega el nodo raiz.
	$menu_ridx = this.AddNode("null",this.root,"#","#","#");
	//Se agrega la hoja que se muestra para la espera de la carga. 
	$menu_lidx  = this.AddLeaf("null", "Cargando...", "javascript:void(0)", "", "", "");
	var html = "<UL class=\"mainmenu\" id=\"" + this.root + "childs\">";	
	
	var htmlChilds = this.Items[$menu_ridx].DrawChilds(this.Items);	
	html += htmlChilds;
	html += "<li class='endmenu'></li>";
	html += "</UL>";
	$("#div" + this.root).html(html);
	//Si no tiene ningún hijo, se crea el hijo por defecto para la espera al momento de la búsqueda.
	if (htmlChilds == "") {
		//No hay nodos para la raiz. Se traen del server.
		loadChilds(this.root);	
	}
}
//-----------------------------------------------------------------------
function mnuRun()
{
    $Menu= new Menu(this.root);
    $Menu.Items = this.Items;
    $Menu.reference = this.reference;
    $Menu.Draw();
}
//-----------------------------------------------------------------------
// NODE OBJECT
//-----------------------------------------------------------------------
function Node(P_Parent,P_Name,P_Text, P_level)
{
// Atributes
this.Parent=P_Parent.toLowerCase();
this.Name=P_Name.toLowerCase();
this.Text=P_Text;
this.level=P_level;
this.Type="node";
// Methods
this.Draw= nodeDraw;
this.DrawChilds= nodeDrawChilds;
}
//-----------------------------------------------------------------------
function nodeDraw(items)
{	
	var htmlChilds = "";	
	var html = "<li>";	
	html += "<a id=\"" + this.Name + "\" href=\"javascript:nodeAction('" + this.Name + "')\""; 
	htmlChilds += this.DrawChilds(items);
	//Si no tiene ningún hijo, se crea el hijo por defecto para la espera al momento de la búsqueda.
	if (htmlChilds == "") {
		html += " class=\"nochild\"";		
	}
	html += ">" + this.Text + "</a>";
	html += "<UL id=\"" + this.Name + "childs\" class=\"mainmenulv" + this.level + "\">";
	html += htmlChilds;
	html += "</UL>";
	return html;
}
//-----------------------------------------------------------------------
function nodeDrawChilds(items)
{		
	var html = "";	
	for(var i in items) {
		if (this.Name == items[i].Parent) html += items[i].Draw(items);		
	}	
	return html;
}
//-----------------------------------------------------------------------
// LEAF OBJECT
//-----------------------------------------------------------------------
function Leaf(P_Parent,P_Text,P_Link,P_Target)
{
//Atributes
this.Parent=P_Parent.toLowerCase();
this.Text=P_Text;
this.Link=P_Link;
this.Target=P_Target;
this.type="leaf";
//Methods
this.Draw= leafDraw;
}
//-----------------------------------------------------------------------
function leafDraw(items)
{	
	var html = "<li ";
       html += " onClick=\"window.open('" + this.Link + "','" + this.Target + "')\" >";
      html += "<a href=\"javascript:void(0)\">" + this.Text + "</a></li>";
	return html;
}
//-----------------------------------------------------------------------
// FUNCIONES GLOBALES
//-----------------------------------------------------------------------
function nodeAction(nodeName) {
	
	var nodeHTML = $("#" + nodeName);
	var node = $Menu.Items[$Menu.reference[nodeName]];
	
	if(!nodeHTML.hasClass('active')) {				
		$("#" + $menu_last_active[node.level]).removeClass('active');
		$("#" + $menu_last_active[node.level] + "childs").filter(':visible').slideUp('normal');
		//$menu_a2.removeClass('active');
		//$menu_ul2.filter(':visible').slideUp('normal');
		nodeHTML.addClass('active').next().stop(true,true).slideDown('normal');
		$menu_last_active[node.level] = nodeName;
		if (nodeHTML.hasClass('nochild')) {
			//Es un nodo y no tiene hijos. Se procede a cargarlos.
			loadChilds(nodeHTML.attr('id'));
		}
	} else {
		$menu_last_active[node.level] = "";
		nodeHTML.removeClass('active');
		nodeHTML.next().stop(true,true).slideUp('normal');
	}
}
//-----------------------------------------------------------------------
//Obtiene y carga los hijos del nodo indicado
function loadChilds(nodeName) {
	//Se dibuja el hijo por default para mostrar la carga en progreso.
	$("#" + nodeName + "childs").html($Menu.Items[$menu_lidx].Draw(""));	
	$.getJSON("menu.asp?action=any&node=" + nodeName,function(data) {
				for(var i in data) {
					var item = data[i];				
					if (item.type == MENU_ITEM_NODE) {
						$Menu.AddNode(item.parent, item.name, item.desc, "#", "#");
					} else {
						$Menu.AddLeaf(item.parent,item.desc,item.link,item.target,"","");
					}
				}				
				//Dibujo los hijos desde el vector de los items.								
				var idx = $Menu.reference[nodeName.toLowerCase()];
				var html = $Menu.Items[idx].DrawChilds($Menu.Items);
				if (html != "") {
					$("#" + nodeName).removeClass('nochild');		
					$("#" + nodeName + "childs").html(html);
				}				
	});
	
}

