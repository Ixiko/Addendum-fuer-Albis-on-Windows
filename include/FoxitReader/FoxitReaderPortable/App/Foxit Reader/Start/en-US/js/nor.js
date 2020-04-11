function clickIE(){if(document.all){return false;}}
function clickNS(e){
if(document.layers||(document.getElementById&&!document.all)){
if(e.which==2||e.which==3){return false;}}}
if(document.layers){
document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;
}else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}
