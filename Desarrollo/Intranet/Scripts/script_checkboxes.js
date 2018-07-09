var check_qnt; //Cantidad de checkboxes dibujadas.

function checkAll(){
   var i,objTmp;
   var obj = document.getElementById("CHKALL");
   if (obj.checked){
      for(i=0; i<check_qnt; i++){
         objTmp = document.getElementById("CHK" + i);
         objTmp.checked = 1;
      }
   }
}
