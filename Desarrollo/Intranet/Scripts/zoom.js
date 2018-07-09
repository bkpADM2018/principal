/* Este script fue hecho por Bratta en www.bratta.com y puede ser usado
gratuitamente siempre que se deje este mensaje intacto. */

        text=new Array('Depósito','de','JavaScript','Abierto','las','24 hs')

// numero de textos
        var numText=6
        
// colores:
// el primer colos sera el color del texto cuando hace el zoom
        color=new Array('#E8E8E8','#C2C2C2','#8E8E8E','#424242','#202020')
        
// numero de colores
        var numColors=5
        
//tamano del zoom al finalizar
        var endSize=80

//Velocidad del zoom (en milisegundos)
        var Zspeed=30

//Velocidad de cambio en los colores
        var Cspeed=200

//Fuente
        var font='Arial Black'
        
// Esconder el texto cuando acaba el zoom? (true o false)
        var hide=true


var size=10
var gonum=0

if (document.all) {
                n=0
                ie=1
                zoomText='document.all.zoom.innerText=text[num]'
                zoomSize='document.all.zoom.style.fontSize=size'
                closeIt=""
                fadeColor="document.all.zoom.style.color=color[num]"
        }
if (document.layers) {
                n=1;ie=0
                zoomText=""
                zoomSize="document.zoom.document.write('<p align=\"center\" style=\"font-family:'+font+'; font-size:'+size+'px; color:'+color[0]+'\">'+text[num]+'</p>')"
                closeIt="document.zoom.document.close()"
                fadeColor="document.zoom.document.write('<p align=\"center\" style=\"font-family:'+font+'; font-size:'+endSize+'px; color:'+color[num]+'\">'+text[numText-1]+'</p>')"
        }

function zoom(num,fn){
        if (size<endSize){
                eval(zoomText)
                eval(zoomSize)
                eval(closeIt)
                size+=5;
                setTimeout("zoom("+num+",'"+fn+"')",Zspeed)
        }else{
                eval(fn);
        }
}

function fadeIt(num){
        if (num<numColors){
                eval(fadeColor)
                eval(closeIt)
                num+=1;
                setTimeout("fadeIt("+num+")",Cspeed)
        }else{
                hideIt()
        }
}

function hideIt(){
        if(hide){
                if(ie)document.all.zoom.style.visibility="hidden"
                if(n)document.layers.zoom.visibility="hidden"
        }
}

function init(){
        if(ie){
                document.all.zoom.style.color=color[0]
                document.all.zoom.style.fontFamily=font}
        go(0)   

}
function go(num){
        gonum+=1
        size=10
        if(num<numText){
                zoom(num,'go('+gonum+')')
        }else{
                fadeIt(0)
        }
}

setTimeout ('init()', 10);
