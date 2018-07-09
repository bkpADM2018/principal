function ControlarNumeros(p_ctrl, p_desde, p_hasta, p_msg)
{
 var n,v;
 var my_msg = new String("");

 v = new String(p_ctrl.value);
 if (p_msg == "")
    p_msg = p_ctrl.name;
  v = v.replace(" ","");

 if (v != "")
 {
    if (!isNaN(v))
    {
       n = eval(v);
       if ((n < p_desde)||(n > p_hasta))
          my_msg = "Con valor de " + n;
       if (v.length < 2)
          v = "0" + v;
    }
    else
    {
       my_msg = "Campo no numerico: (" + v + ")";
    }
 }
 if (my_msg != "")
 {
    alert("Error en el " + p_msg + ". " + my_msg);
    p_ctrl.focus();
 }
 p_ctrl.value = v;
 return true;
}
//**************************************************************************************
function ControlarMes(p_ctrl)
{
   return ControlarNumeros(p_ctrl,1,12,"Mes");
}
//**************************************************************************************
function ControlarDia(p_ctrl)
{
   return ControlarNumeros(p_ctrl,1,31,"Dia");
}
//**************************************************************************************
function ControlarAnio(p_ctrl)
{
 var n;

 if ((!isNaN(p_ctrl.value))&&(p_ctrl.value != ""))
 {
    n = eval(p_ctrl.value);
    if ((n>49)&&(n<100))
       n = n + 1900;
    if (n < 50)
       n = n + 2000;
    p_ctrl.value = n;
 }
 return ControlarNumeros(p_ctrl,1998,2050,"Año");
}

