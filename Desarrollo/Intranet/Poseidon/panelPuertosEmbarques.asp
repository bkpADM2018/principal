<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosTraducir.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientosUnificador.asp"-->
<%
Const TRANSPORTE_EMBARQUE = 50
'--------------------------------------------------------------------------------------------------------------
Function getUltimoEmbarque(pPto, ByRef pIdRegistroIni, ByRef pIdRegistroFin)
    Dim strSQL, rsEmb, salir, myBuque
    pIdRegistroIni = 0
    pIdRegistroFin = 0
    myBuque = ""
    'Busco el primer registro del ultimo embarque registrado.
    'Se asume que solo se carga un buque por vez en el puerto por lo cual los regsitros del ultimo buque estaran todos juntos ordenados por tiempo.
    'Tambien se asume que el mismo buque no viene 2 veces seguidas sin otro buque en medio de 2 de sus visitas.
    strSQL= " select IDREGISTRO, BUQUE from EMBARQUESREGISTROPESO order by mmto desc "            
    Call GF_BD_Puertos(pPto, rsEmb, "OPEN", strSQL)    
    if (not rsEmb.eof) then
        salir = false
        pIdRegistroFin = CInt(rsEmb("IDREGISTRO"))
        myBuque = Trim(rsEmb("BUQUE")) 
        while ((not rsEmb.eof) and (not salir))        
            if (Trim(rsEmb("BUQUE")) = myBuque) then
                pIdRegistroIni = CInt(rsEmb("IDREGISTRO"))
                rsEmb.MoveNext()
            else
                salir = true                
            end if            
        wend
    end if            
End Function 
'--------------------------------------------------------------------------------------------------------------
Function getTotalEmbarqueActual(pPto, pBza, pIdRegistroIni, pIdRegistroFin)
    Dim rs, strSQL, salir, totalGeneral, myBuque, myProducto, salirProd, totalProducto, myWhere        
      
    'Obtengo los datos a mostrar.
    if (pBza <> "") then myWhere = " and Balanza = '" & pBza & "'"
    strSQL="Select BUQUE, COMMODITY, FECHA, TURNO, BODEGA, Sum(PESO) PESO from (Select BUQUE, CONCAT(COMMODITY, CONCAT('/', Exportador)) COMMODITY, Left(MMTO, 8) Fecha, TURNO, BODEGA, PESO from EMBARQUESREGISTROPESO where IDREGISTRO >= " & pIdRegistroIni & " and IDREGISTRO <=" & pIdRegistroFin & myWhere & ") T group by BUQUE, COMMODITY, FECHA, TURNO, BODEGA order by BUQUE, COMMODITY, FECHA, TURNO, BODEGA"        
    Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)            
    if (not rs.eof) then
        myBuque = Trim(rs("BUQUE"))        
        'Comienzo armando los titulos
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='32'>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='buque' colspan='5'><strong>ULTIMO EMBARQUE: " & myBuque & "</strong></td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='25px'>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='titulo' width='20%'></td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='titulo' width='20%'>" & GF_TRADUCIR("Fecha") &  "</td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='titulo' width='20%'>" & GF_TRADUCIR("Turno") & "</td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='titulo' width='20%'>" & GF_TRADUCIR("Bodega")  & "</td>"    
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='titulo' width='20%'>" & GF_TRADUCIR("Carga (kg)") & "</td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"        
        totalGeneral = 0
        salir = false    
        while ((not rs.eof) and (not salir))
            myProducto = Trim(rs("COMMODITY"))
            salirProd = false   
            totalProducto = 0         
            getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='32'>"
            getTotalEmbarqueActual = getTotalEmbarqueActual & "<td style='text-align:left;' colspan='5'><strong>" & myProducto & "</strong></td>"
            getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"
            while ((not rs.eof) and (not salir) and (not salirProd))
                if (myBuque = Trim(rs("BUQUE"))) then
                    if (myProducto = Trim(rs("COMMODITY"))) then                                    
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='25'>"
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td ></td>"
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='textoComun'>" & GF_FN2DTE(rs("FECHA")) &  "</td>"
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='textoComun'>" & GF_nChars(rs("TURNO"), 10, " ", CHR_AFT) & "</td>"  
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='textoComun'>" & GF_nChars(rs("BODEGA"), 10, " ", CHR_AFT)  & "</td>"    
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='importe'>" & GF_nChars(GF_EDIT_DECIMALS(rs("PESO"), 0), 10, " ", CHR_FWD) & "</td>"
                        getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"                        
                        totalProducto = totalProducto + CLng(rs("PESO"))        
                        rs.MoveNext()
                    else
                        salirProd = true
                    end if            
                else
                    salir = true                
                end if                
            wend                    
            getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='32'>"
            getTotalEmbarqueActual = getTotalEmbarqueActual & "<td colspan='4' class='textoSubTotal '>" & GF_TRADUCIR("SUBTOTAL:") & "</td>"    
            getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='subtotal'>" & GF_nChars(GF_EDIT_DECIMALS(totalProducto, 0), 10, " ", CHR_FWD) & "</td>"
            getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"                                        
            totalGeneral = totalGeneral + totalProducto            
        wend    
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<tr height='32'>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td colspan='4' class='total'>" & GF_TRADUCIR("TOTAL BUQUE:") & "</td>"    
        getTotalEmbarqueActual = getTotalEmbarqueActual & "<td class='total'>" & GF_nChars(GF_EDIT_DECIMALS(totalGeneral, 0), 10, " ", CHR_FWD) & "</td>"
        getTotalEmbarqueActual = getTotalEmbarqueActual & "</tr>"                                                
    end if        
End function
'--------------------------------------------------------------------------------------------------------------
Function controlDeCorteProducto(rsEmb,p_Commodity)
    Dim Commodity
    controlDeCorteProducto = false
    if not rsEmb.eof then
        Commodity= Trim(Ucase(rsEmb("Commodity")))
        if (Cstr(Commodity) = Cstr(p_Commodity)) then controlDeCorteProducto = true
    end if
End Function
'*************************************************************************************
'***************************** COMIENZO DE LA PAGINA *********************************
'*************************************************************************************
Dim g_strPuerto, regIni, regFin, count
   
g_strPuerto = GF_Parametros7("pto","",6)
'regIni = GF_Parametros7("ri",0,6)
'regFin = GF_Parametros7("rf",0,6)

Call getUltimoEmbarque(g_strPuerto, regtIni, regFin)
%>
<html>
<head>        
    <style>
        
        .linkBza {
            cursor: pointer;            
        }
        
        .linkBza:hover {
            cursor: pointer;
            background: rgba(230, 250, 200, 1);
        }
        .buque 
        {
            text-align: center ;
	        border:2px solid #396E8F ;
	        font-size: 18px;
	        font-weight: bold;	
	        background: #FFFFFF;        
        }
        .titulo 
	    {
	        text-align: center ;
	        border:2px solid #396E8F ;
	        font-size: 14px;
	        font-weight:500;
	        color: #2e6b4d;
	        line-height: 1px;
	        background: #FFFFFF;
	    }
	    .textoComun 
	    {
	        border: none; 
	        text-align:center;
	        font-size: 14px;
	    }
	    .textoSubTotal 
	    {
	        border: none; 
	        text-align:right;
	        font-size: 14px;
	        font-weight: bold;	 
			background: #FFFFFF;			
	    }
	    .importe
	    {
	        border: none; 
	        text-align:right;
	        font-size: 14px;
	    }	        
	    .subtotal
	    {	        
	        text-align:right;
	        border-top:2px solid #396E8F ;	        
	        font-size: 14px;
	        font-weight: bold;	        
	        line-height: 1px;
			background: #FFFFFF;
	    }
	    .total
	    {	        
	        text-align:right;
	        border-top:2px solid #396E8F ;
	        border-bottom:2px solid #396E8F ;
	        font-size: 16px;
	        font-weight: bold;	        
	        line-height: 1px;
			background: #FFFFFF;
	    }    
    </style>
    <script type="text/javascript">
    function abrirRegistroBalanzaEmbarque(bza, nbr) {
        var Ancho = screen.width - 950;
        var Alto = screen.height - 750;
        var lft = 10;
        if ((nbr % 2) == 0) lft = Ancho;
        var tp = 10;
        if (nbr > 2) tp = Alto;
        window.open('Embarques/registroBalanzaOnLine.asp?pto=<%=g_strPuerto%>&bza=' + bza, "_blank", "width=950,height=750,left=" + lft + ",top=" + tp + ", toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=no");        
    }
    </script>
</head>
<body>
    
<%  'Se muestran las balanzas disponibles para ver online        
    strSQL="Select distinct BALANZA from EMBARQUESREGISTROBALANZA"
    Call GF_BD_Puertos(g_strPuerto, rs, "OPEN", strSQL)
    if (not rs.eof) then
%>
    <table width="100%">
        <tr height="32px">
            <td class="buque" colspan="5"><strong>BALANZAS ON-LINE</strong></td>
        </tr>
        <tr height='25px'>
<%    
    count = 0
    while (not rs.eof) 
        count=count+1
%> 
            <td class="titulo linkBza" width="25%" title="Ver Online" onClick="abrirRegistroBalanzaEmbarque('<% =rs("BALANZA") %>', <% =count %>)"><% =rs("BALANZA") %></td>
<%      rs.MoveNext()
    wend
%>      
        </tr>  
    </table>
<%  end if %>    
    <table width="100%">
        <% response.write getTotalEmbarqueActual(g_strPuerto, "", regtIni, regFin) %>
    </table>
</body>
</html>



