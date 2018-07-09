<!--#include file="includes/procedimientosMG.asp"-->
<!--#include file="includes/procedimientostraducir.asp"-->
<html>
<head>
  <title> Opciones de Impresion                                </title>
  <link rel="shortcut icon" href="images/alertIcon.gif" />
  <link href="CSS/ActisaIntra-1.css" rel="stylesheet" type="text/css">
  <script language="javascript">
    function fcnImprimir()
    {

        window.returnValue = 'completos';
        if (document.form1.modo[0].checked)
            window.returnValue = 'caratulas'
        else if (document.form1.modo[1].checked)
            window.returnValue = 'descargas';
        window.close();
    }
  </script>
</head>
<body>
<form name="form1" action="cor-CtoGeneratorOpciones.asp" method="post">
    <table width="270" border="0" align="center">
        <tr>
            <td colspan="2">
                <font size="3"><b>
                    <% =GF_Traducir("Ud. Esta a punto de imprimir Contratos.")%><br>
                    <% =GF_Traducir("¿Que desea imprimir?")%>
                </b></font>
            </td>
        </tr>
        <tr>
            <td colspan="2"><input type="radio" checked class="NOBORDER" name="modo" value="caratulas"><font size="3"><% =GF_Traducir("Solamente las caratulas")%></font></td>
        </tr>
        <tr>
            <td colspan="2"><input type="radio" class="NOBORDER" name="modo" value="descargas"><font size="3"><% =GF_Traducir("Solamente las descargas")%></font></td>
        </tr>
        <tr>
            <td colspan="2"><input type="radio" class="NOBORDER" name="modo" value="completo"><font size="3"><% =GF_Traducir("Caratulas y descargas")%></td>
        </tr>
        <tr height="50">
            <td width="50%" align="center"><input type="button" name="imprimir" value="Imprimir" style="cursor:hand;" onClick="javascript:fcnImprimir();"></td>
            <td align="center"><input type="button" name="cancelar" value="Cancelar" style="cursor:hand;" onClick="javascript:window.close();"></td>
        </tr>
    </table>
</form>
</body>
</html>
<script language="javascript">window.returnValue = 'nada';</script>
