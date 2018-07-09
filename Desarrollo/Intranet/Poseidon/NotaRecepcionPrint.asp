<!--#include file="../Includes/procedimientosUnificador.asp"-->
<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosPuertos.asp"-->
<!--#include file="../Includes/procedimientosSQL.asp"-->
<!--#include file="../Includes/procedimientos.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosfechas.asp"-->
<!--#include file="../Includes/procedimientosuser.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosPDF.asp"-->
<%
CONST PARAM_CD_RUBRO_HUMEDAD = "CDRUBROHUMEDAD"
Const PARAM_LEYENDA_LUGAR = "LEYENDALUGAR"
'------------------------------------------------------------------------------------------------------------------------
Function getSQLRecepcionVagones(pDtContable, pcartaPorte, pCdVagon)
	Dim strSQL,myWhere,diaHoy,rs
	if (pDtContable <> "")  then auxDtContable = left(pDtContable, 4) & "-" & Mid(pDtContable, 5, 2) & "-" & Right(pDtContable,2)		
	if (pcartaPorte <> "")	then Call mkWhere(myWhere, "T1.CDOPERATIVO", pcartaPorte, "=", 3)
	if (pCdVagon <> "")		then Call mkWhere(myWhere, "T1.CDVAGON", pCdVagon, "=", 3)
	
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)
	strSQL= "SELECT T1.dtcontable,T1.cdoperativo,T1.nucartaporteserie,T1.nucartaporte,T1.cdvagon,T1.vlbrutoorigen,T1.vltaraorigen,T1.vlnetoorigen,T1.nubarras, 0 CTG, T1.NUCUITREM, "&_
            "       T1.dtarribo, T1.dtCalada, T1.dtpesadaBruto, T1.dtpesadaTara, T1.dtfin, "&_
            "       T1.vlpesadaBruto,T1.vlpesadaTara,T1.cdaceptacion,T1.cdtransportista,T1.sqturno, "&_
			"		T1.pcmerma,T1.vlmermakilos,T1.cdproducto,T1.cdcosecha,T1.cdempresa,T1.cdcliente,T1.cdcorredor,T1.cdvendedor,T1.cdentregador,T1.cdprocedencia,T1.sqcalada, "&_
            "       d.dstransportista,d.DSDOMICILIO,d.CDTIPODOC,d.NUDOCUMENTO,g.DSACEPTACION, T1.DSOBSERVACIONES, T1.CDESTADO, "&_
			"		i.dsproducto,j.dsempresa,k.dscliente,l.dscorredor,m.dsvendedor,n.dsentregador,o.dsprocedencia "&_
			"FROM  ((SELECT (YEAR(a.DtContable)*10000 + Month(a.DtContable)*100 + DAY(a.DtContable)) dtcontable,c.cdoperativo,a.nucartaporteserie,a.nucartaporte,a.cdvagon,a.vlbrutoorigen,a.vltaraorigen,a.vlnetoorigen,b.nubarras, "&_			
			"               format(a.dtarribo, 'yyyyMMddHHmmss') AS dtarribo, "&_            
            "               format(audCalada.dtauditoria, 'yyyyMMddHHmmss')  AS dtCalada, "&_
            "               format(audBruto.dtauditoria, 'yyyyMMddHHmmss')  AS dtpesadaBruto, "&_
            "               format(audTara.dtauditoria, 'yyyyMMddHHmmss')  AS dtpesadaTara, "&_
            "               format(a.dtfin, 'yyyyMMddHHmmss') AS dtfin, "&_			
            "               c.NUCUITREM, e.vlpesada as vlpesadaBruto,f.vlpesada as vlpesadaTara,b.cdaceptacion,c.cdtransportista,c.sqturno, "&_
			"			    b.pcmerma,h.vlmermakilos,a.cdproducto,a.cdcosecha,c.cdempresa,c.cdcliente,c.cdcorredor,c.cdvendedor,c.cdentregador,c.cdprocedencia,b.sqcalada "&_
			"		FROM hvagones a "&_
			"		LEFT JOIN (SELECT nubarras,cdvagon,dtcontable,dtcalada,cdoperativo,cdaceptacion,pcmerma,sqcalada FROM hcaladadevagones "&_
			"					where sqcalada = (Select Max(T1.SqCalada) from dbo.HCaladadeVagones T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"') "&_
			"				  ) b  on a.cdvagon = b.cdvagon and b.CDOPERATIVO = a.CDOPERATIVO and a.dtcontable = b.dtcontable "&_
			"		left join Hoperativos c  on a.cdoperativo = c.cdoperativo and c.dtcontable = a.dtcontable "&_			
			"		left join ( SELECT vlpesada,dtpesada,cdoperativo,cdvagon,dtcontable "&_
			"					FROM dbo.hpesadasvagon "&_
			"					WHERE  cdpesada = 1"&_
			"						AND sqpesada = (SELECT Max(T1.sqpesada) "&_
            "										FROM   dbo.hpesadasvagon T1 "& myWhere & " AND dtcontable= '"&auxDtContable&"' AND T1.CDPESADA = 1)) e "&_
			"			on c.cdoperativo = e.cdoperativo and e.dtcontable = a.dtcontable and e.cdvagon = a.cdvagon "&_
			"		left join ( SELECT vlpesada,dtpesada,cdoperativo,cdvagon,dtcontable "&_
			"				    FROM   dbo.hpesadasvagon"&_
			"					WHERE CDPESADA = 2 AND sqpesada = (SELECT Max(T1.sqpesada) "&_
            "										FROM   dbo.hpesadasvagon T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"' AND T1.CDPESADA = 2)) f  "&_
			"			on c.cdoperativo = f.cdoperativo and f.dtcontable = a.dtcontable and f.cdvagon = a.cdvagon "&_
			"		 LEFT JOIN (SELECT VLMERMAKILOS,cdoperativo,cdvagon,dtcontable FROM HMERMASVAGONES "&_
			"					WHERE SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASVAGONES T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"')) h"&_
			"			on h.cdoperativo = c.cdoperativo AND h.dtcontable = a.dtcontable AND h.cdvagon = a.cdvagon  "&_			
			"		left join (select t1.* from hauditoriavagones T1 "& myWhere &" AND T1.dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=54 and T1.cdestadoposterior=2) audCalada "&_
            "			on audCalada.cdoperativo = c.cdoperativo AND audCalada.dtcontable = a.dtcontable and audCalada.cdvagon = a.cdvagon "&_
			"		left join (select t1.* from hauditoriavagones T1 "& myWhere &"  AND dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=7 and T1.cdestadoposterior=5) audBruto "&_
            "	        on audCalada.cdoperativo = c.cdoperativo AND audBruto.dtcontable = a.dtcontable and audBruto.cdvagon = a.cdvagon "&_
            "		left join (select t1.* from hauditoriavagones T1 "& myWhere &"  AND dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=8 and T1.cdestadoposterior=8 ) audTara"&_
            "	        on audCalada.cdoperativo = c.cdoperativo AND audTara.dtcontable = a.dtcontable and audTara.cdvagon = a.cdvagon "&_					
			"	)union	 "&_
			"		 (SELECT "& diaHoy &" AS dtcontable,c.cdoperativo,a.nucartaporteserie,a.nucartaporte,a.cdvagon,a.vlbrutoorigen,a.vltaraorigen,a.vlnetoorigen,b.nubarras, a.cdestado, "&_
            "               format(a.dtarribo, 'yyyyMMddHHmmss') AS dtarribo, "&_            
            "               format(audCalada.dtauditoria, 'yyyyMMddHHmmss')  AS dtCalada, "&_
            "               format(audBruto.dtauditoria, 'yyyyMMddHHmmss')  AS dtpesadaBruto, "&_
            "               format(audTara.dtauditoria, 'yyyyMMddHHmmss')  AS dtpesadaTara, "&_
            "               format(a.dtfin, 'yyyyMMddHHmmss') AS dtfin, b.DSOBSERVACIONES, "&_
            "               c.NUCUITREM, e.vlpesada as vlpesadaBruto,f.vlpesada as vlpesadaTara,b.cdaceptacion,c.cdtransportista,c.sqturno, "&_
			"				b.pcmerma,h.vlmermakilos,a.cdproducto,a.cdcosecha,c.cdempresa,c.cdcliente,c.cdcorredor,c.cdvendedor,c.cdentregador,c.cdprocedencia,b.sqcalada  "&_
			"		  FROM vagones a "&_
			"		  LEFT JOIN (SELECT nubarras,cdvagon,dtcalada,cdoperativo,cdaceptacion,pcmerma,sqcalada, DSOBSERVACIONES,	"&_
			"					 FROM   caladadevagones "&_
			"					 WHERE  sqcalada = (SELECT Max(T1.sqcalada) FROM dbo.caladadevagones T1 "& myWhere &" )) b"&_
			"			ON a.cdvagon = b.cdvagon AND b.cdoperativo = a.cdoperativo "&_
			"		  LEFT JOIN operativos c ON a.cdoperativo = c.cdoperativo  "&_
			"		  LEFT JOIN (SELECT vlpesada,dtpesada,cdoperativo,cdvagon FROM dbo.pesadasvagon "&_ 
			"					 WHERE  cdpesada = 1 AND sqpesada = (SELECT Max(T1.sqpesada) FROM dbo.pesadasvagon T1 "& myWhere &" AND T1.cdpesada = 1)) e"&_
			"			ON c.cdoperativo = e.cdoperativo AND e.cdvagon = a.cdvagon "&_
			"		  LEFT JOIN (SELECT vlpesada,dtpesada,cdoperativo,cdvagon FROM   dbo.pesadasvagon 	"&_
			"					 WHERE  cdpesada = 2 AND sqpesada = (SELECT Max(T1.sqpesada) FROM dbo.pesadasvagon T1 "& myWhere &" AND T1.cdpesada = 2)) f "&_
			"			ON c.cdoperativo = f.cdoperativo AND f.cdvagon = a.cdvagon "&_
			"		  LEFT JOIN (SELECT VLMERMAKILOS,cdoperativo,cdvagon FROM MERMASVAGONES "&_
			"					 WHERE SQPESADA= (SELECT MAX(SQPESADA) FROM MERMASVAGONES T1 "& myWhere &" )) h "&_
			"			ON h.cdoperativo = c.cdoperativo AND h.cdvagon = a.cdvagon 	"&_			
			"		left join (select t1.* from auditoriavagones T1 "& myWhere &" AND t1.cdtransaccion=54 and T1.cdestadoposterior=2) audCalada "&_
            "			on audCalada.cdoperativo = c.cdoperativo AND audCalada.cdvagon = a.cdvagon "&_
			"		left join (select t1.* from auditoriavagones T1 "& myWhere &" AND t1.cdtransaccion=7 and T1.cdestadoposterior=5) audBruto "&_
            "	        on audCalada.cdoperativo = c.cdoperativo AND audBruto.cdvagon = a.cdvagon "&_
            "		left join (select t1.* from auditoriavagones T1 "& myWhere &" AND t1.cdtransaccion=8 and T1.cdestadoposterior=8 ) audTara"&_
            "	        on audCalada.cdoperativo = c.cdoperativo AND audTara.cdvagon = a.cdvagon))AS T1 "&_
			"LEFT JOIN transportistas d ON d.cdtransportista = T1.cdtransportista 	"&_
			"LEFT JOIN aceptacioncalidad g ON g.cdaceptacion = T1.cdaceptacion	"&_ 
			"LEFT JOIN productos i ON i.cdproducto = T1.cdproducto "&_
			"LEFT JOIN empresas j ON j.cdempresa = T1.cdempresa	"&_ 
			"LEFT JOIN clientes k ON k.cdcliente = T1.cdcliente	"&_
			"LEFT JOIN corredores l ON l.cdcorredor = T1.cdcorredor "&_
			"LEFT JOIN vendedores m ON m.cdvendedor = T1.cdvendedor "&_
			"LEFT JOIN entregadores n ON n.cdentregador = T1.cdentregador "&_
			"LEFT JOIN procedencias o ON o.cdprocedencia = T1.cdprocedencia "&_
			myWhere & " AND dtcontable= " & pDtContable
'response.write strSQL
    Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	Set getSQLRecepcionVagones= rs
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getSQLRecepcionCamiones(pDtContable, pcartaPorte, pIdCamion, pTipoCamion)
	Dim strSQL,myWhere,diaHoy,rs	
	if (pDtContable <> "")  then auxDtContable = left(pDtContable, 4) & "-" & Mid(pDtContable, 5, 2) & "-" & Right(pDtContable,2)		
	if (pIdCamion <> "")	then Call mkWhere(myWhere, "T1.IDCAMION", pIdCamion, "=", 3)
	
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)
	strSQL= "SELECT T1.dtcontable,T1.nucartaporte,T1.idcamion,T1.vlbrutoorigen,T1.vltaraorigen,T1.nubarras,T1.vlpesadaBruto,T1.vlpesadaTara,T1.cdaceptacion, T1.CTG, "&_
            "       T1.pcmerma,T1.vlmermakilos,T1.cdproducto,T1.cdcosecha,T1.sqcalada,T1.cdtransportista,T1.sqturno,T1.cdempresa,T1.cdcliente,T1.cdcorredor,T1.cdvendedor,T1.cdentregador,T1.cdprocedencia,T1.CDCHAPACAMION,T1.CDCHAPAACOPLADO, "&_
            "       d.dstransportista,d.DSDOMICILIO,d.CDTIPODOC,d.NUDOCUMENTO,g.DSACEPTACION, T1.cdestado, "&_
            "       T1.DTINGRESO, T1.dtCalada, T1.dtpesadaBruto, T1.dtpesadaTara, T1.DTEGRESO, T1.DSOBSERVACIONES, "&_
			"		i.dsproducto,j.dsempresa,k.dscliente,l.dscorredor,m.dsvendedor,n.dsentregador,o.dsprocedencia, T1.NUCUITREM "&_
			"FROM  (SELECT  format(a.DTCONTABLE, 'yyyyMMdd') dtcontable,c.nucartaporte,a.idcamion,b.nubarras, a.cdestado, "
            if (CInt(pTipoCamion) = CIRCUITO_CAMION_DESCARGA) then
                strSQL = strSQL & "c.vlbrutoorigen,c.vltaraorigen,c.cdprocedencia, "
            else
                strSQL = strSQL & "0 as vlbrutoorigen,0 as vltaraorigen,c.cddestino as cdprocedencia,"
            end if
            strSQL = strSQL & " format(a.dtingreso, 'yyyyMMddHHmmss') AS DTINGRESO, "&_
            "               format(audCalada.dtauditoria, 'yyyyMMddHHmmss') AS dtCalada, "&_
            "               format(audBruto.dtauditoria, 'yyyyMMddHHmmss') AS dtpesadaBruto, "&_
            "               format(audTara.dtauditoria, 'yyyyMMddHHmmss') AS dtpesadaTara, "&_
            "               format(a.dtegreso, 'yyyyMMddHHmmss') AS DTEGRESO, "&_
			"				e.vlpesada as vlpesadaBruto,f.vlpesada as vlpesadaTara,b.cdaceptacion, b.DSOBSERVACIONES,"
			if (CInt(pTipoCamion) = CIRCUITO_CAMION_DESCARGA) then
			    strSQL = strSQL & " C.CTG " 
            else
                strSQL = strSQL & " '' "
            end if                
			strSQL = strSQL & " CTG , a.NUCUITREM,	b.pcmerma,h.vlmermakilos,a.cdproducto,c.cdcosecha,b.sqcalada,a.cdtransportista,a.sqturno,c.cdempresa,c.cdcliente,c.cdcorredor,c.cdvendedor,c.cdentregador,a.CDCHAPACAMION,a.CDCHAPAACOPLADO "&_
			"		FROM hcamiones a "&_
			"		LEFT JOIN (SELECT nubarras,idcamion,dtcontable,dtcalada,cdaceptacion,pcmerma,sqcalada, DSOBSERVACIONES FROM hcaladadecamiones  "&_
			"					where sqcalada = (Select Max(T1.SqCalada) from dbo.HCaladadecamiones T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"') "&_
			"				  ) b  on a.idcamion = b.idcamion and a.dtcontable = b.dtcontable "
            if (CInt(pTipoCamion) = CIRCUITO_CAMION_DESCARGA) then
			    strSQL = strSQL & " inner join hcamionesdescarga c on c.idcamion = a.idcamion and c.dtcontable = a.dtcontable "
            else
                strSQL = strSQL & " inner join hcamionescarga c on c.idcamion = a.idcamion and c.dtcontable = a.dtcontable "
            end if
			strSQL = strSQL & " left join ( SELECT vlpesada,dtpesada,idcamion,dtcontable   "&_
			"					FROM dbo.hpesadascamion "&_
			"					WHERE  cdpesada = 1"&_
			"						AND sqpesada = (SELECT Max(T1.sqpesada) "&_
            "										FROM   dbo.hpesadascamion T1 "& myWhere & " AND dtcontable= '"&auxDtContable&"' AND T1.CDPESADA = 1)) e "&_
			"			on e.dtcontable = a.dtcontable and e.idcamion = a.idcamion "&_
			"		left join ( SELECT vlpesada,dtpesada,idcamion,dtcontable "&_
			"				    FROM   dbo.hpesadascamion"&_
			"					WHERE CDPESADA = 2 AND sqpesada = (SELECT Max(T1.sqpesada) "&_
            "										FROM   dbo.hpesadascamion T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"' AND T1.CDPESADA = 2)) f  "&_
			"			on f.dtcontable = a.dtcontable and f.idcamion = a.idcamion "&_
			"		 LEFT JOIN (SELECT VLMERMAKILOS,idcamion,dtcontable FROM HMERMASCAMIONES  "&_
			"					WHERE SQPESADA= (SELECT MAX(SQPESADA) FROM HMERMASCAMIONES T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"')) h"&_
			"			on h.dtcontable = a.dtcontable AND h.idcamion = a.idcamion   "&_
			"		left join (select t1.* from hauditoriacamiones T1 "& myWhere &" AND dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=2 and t1.cdestadoanterior=1) audCalada "&_
            "			on audCalada.dtcontable = a.dtcontable and audCalada.idcamion = a.idcamion "&_
			"		left join (select t1.* from hauditoriacamiones T1 "& myWhere &"  AND dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=7) audBruto "&_
            "	        on audBruto.dtcontable = a.dtcontable and audBruto.idcamion = a.idcamion "&_
            "		left join (select t1.* from hauditoriacamiones T1 "& myWhere &"  AND dtcontable= '"&auxDtContable&"' and t1.cdtransaccion=8 ) audTara"&_
            "	        on audTara.dtcontable = a.dtcontable and audTara.idcamion = a.idcamion "&_                   			
			"	)AS T1	"&_
			"LEFT JOIN transportistas d ON d.cdtransportista = T1.cdtransportista 	"&_
			"LEFT JOIN aceptacioncalidad g ON g.cdaceptacion = T1.cdaceptacion	"&_ 
			"LEFT JOIN productos i ON i.cdproducto = T1.cdproducto "&_
			"LEFT JOIN empresas j ON j.cdempresa = T1.cdempresa	"&_ 
			"LEFT JOIN clientes k ON k.cdcliente = T1.cdcliente	"&_
			"LEFT JOIN corredores l ON l.cdcorredor = T1.cdcorredor "&_
			"LEFT JOIN vendedores m ON m.cdvendedor = T1.cdvendedor "&_
			"LEFT JOIN entregadores n ON n.cdentregador = T1.cdentregador "&_
			"LEFT JOIN procedencias o ON o.cdprocedencia = T1.cdprocedencia "&_
			myWhere & " AND dtcontable= " & pDtContable & " AND NUCARTAPORTE = '"&pcartaPorte&"'"				
	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)	
	Set getSQLRecepcionCamiones = rs
End Function
'---------------------------------------------------------------------------------------------------------------------
Function armadoPDF(pcartaPorte,pCdVagon,pDtContable,pPto,pTipoCamion,pIdCamion)
	Dim rs		
	if (flagIsCamion) then
		Set rs = getSQLRecepcionCamiones(pDtContable,pcartaPorte,pIdCamion,pTipoCamion)
	else		
		Set rs = getSQLRecepcionVagones(pDtContable,pcartaPorte,pCdVagon)
	end if
	Call drawHeader(pPto,pDtContable, pTipoCamion, rs)
	Call drawTitle(pPto, rs)
	Call drawDetail(rs, pTipoCamion)
	Call drawFooter()
end Function
'------------------------------------------------------------------------------------------------------------------------
Function drawHeader(pPto, dtContable, pTipoCamion, pRs)
    Call GF_squareBox(oPDF,5  ,5,585,140,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)	
    Call GF_verticalLine(oPDF,292,5,140)
	Call GF_writeImage(oPDF, Server.MapPath("..\Images\logo1.jpg"),10, 30, 81, 75, 0)		
	Call drawHeaderLeft(pPto)		
	Call drawHeaderRight(pPto, dtContable, pTipoCamion, pRs)		
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getNroRecepcion(pPto, dtContable)
	Dim strSQL,diaHoy,rtrn,auxDtcontable
	diaHoy = Year(Now()) & "-" & GF_nDigits(Month(Now()),2) & "-" & GF_nDigits(Day(Now()),2)		
	auxDtcontable = left(dtContable, 4) & "-" & Mid(dtContable, 5, 2) & "-" & Right(dtContable,2)		
	if (flagIsCamion) then
		strSQL = "SELECT T.NUAUTSALIDA AS NRORECEPCION "&_
				 "FROM ((SELECT DTCONTABLE,NUAUTSALIDA,IDCAMION FROM HCAMIONES ) UNION "&_
				 "		(SELECT '"& diaHoy &"' AS DTCONTABLE,NUAUTSALIDA,IDCAMION FROM CAMIONES)) T "&_
				 "WHERE T.IDCAMION='"&idCamion&"' AND T.DTCONTABLE='"&auxDtcontable&"'"
	else
		strSQL = "SELECT T.NURECIBO AS NRORECEPCION "&_
				 "FROM ((SELECT DTCONTABLE,CDVAGON,CDOPERATIVO,NURECIBO FROM HVAGONES ) UNION "&_
				 "		(SELECT '"& diaHoy &"' AS DTCONTABLE,CDVAGON,CDOPERATIVO,NURECIBO FROM VAGONES)) T "&_
				 "WHERE T.CDOPERATIVO='"&cartaPorte&"' AND T.CDVAGON='"&cdvagon&"' AND T.DTCONTABLE='"&auxDtcontable&"'"	
	end if
    
	Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
	rtrn = ""
	if not rs.Eof then rtrn = rs("NRORECEPCION")
	getNroRecepcion = rtrn
End Function 
'------------------------------------------------------------------------------------------------------------------------
Function drawHeaderLeft(pPto)
    Call GF_setFont(oPDF,"COURIER", 18 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,100, 40, "ADM AGRO S.R.L." , 200,PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)	
	 select case UCase(pPto)
	    case TERMINAL_ARROYO: 
			Call GF_writeTextAlign(oPDF,100, 70, "Arroyo Seco - Planta N° 20573" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 82, "Ruta 21 Km. 277" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 94, "C.P: (2128) - Santa Fe" , 200,PDF_ALIGN_LEFT)			
	    case TERMINAL_TRANSITO: 	
	        Call GF_writeTextAlign(oPDF,100, 70, "Puerto Gral. San Martín (Rosario)" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 82, "Alem esq. América s/n. - Planta N° 20572" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 94, "C.P. (2202) - Santa Fe" , 200,PDF_ALIGN_LEFT)	
	    case TERMINAL_PIEDRABUENA:		    
			Call GF_writeTextAlign(oPDF,100, 70, "A.Alcorta s/n - Planta N° 20574" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 82, "Puerto Ing. White (Bahia Blanca)" , 200,PDF_ALIGN_LEFT)
			Call GF_writeTextAlign(oPDF,100, 94, "C.P. (8103) - Buenos Aires" , 200,PDF_ALIGN_LEFT)
    end select	
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawHeaderRight(pPto, pDtContable, pTipoCamion, pRs)	
    Dim strTipoDoc, nroRecepcion
    
    
	if (CInt(pRs("CDESTADO")) <> CAMIONES_ESTADO_RECHAZADO) then 	
		Call GF_setFont(oPDF,"COURIER", 14 , FONT_STYLE_BOLD)			
		nroRecepcion = getNroRecepcion(pPto, dtContable)
		strTipoDoc = "NOTA DE RECEPCION Nº  "
		if (pTipoCamion = CIRCUITO_CAMION_CARGA) then strTipoDoc = "REMITO DE CARGA  Nº   "		
		Call GF_writeTextAlign(oPDF,310, 30, strTipoDoc &  nroRecepcion, 200,PDF_ALIGN_LEFT)	
	else
		Call GF_setFont(oPDF,"COURIER", 14 , FONT_STYLE_BOLD)			
		strTipoDoc = "PASE DE SALIDA"
		Call GF_writeTextAlign(oPDF,310, 30, strTipoDoc, 250,PDF_ALIGN_CENTER)	
	end if	
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,310, 52, "FECHA: " & GF_FN2DTE(pDtContable), 200,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,310, 64, "CUIT: " & GF_STR2CUIT(CUIT_TOEPFER), 200,PDF_ALIGN_LEFT)
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_BOLD)
	Call GF_writeTextAlign(oPDF,310, 124, "DOCUMENTO NO VALIDO COMO FACTURA" , 200,PDF_ALIGN_LEFT)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getPartesInvolucradas(pPto, pRs, ByRef pTitular, ByRef pIntermediario, ByRef pRteComercial)
    Dim strSQL, rs
    
    'Obtengo el titular de la carta de porte
    if (Trim(pRS("NUCUITREM")) <> "") then
        strSQL="Select * from VENDEDORES where NUDOCUMENTO='" & Trim(pRS("NUCUITREM")) & "'"        
        Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
        if (not rs.eof) then 
            pTitular = rs("DSVENDEDOR")
        else
            pTitular = "C.U.I.T. " & GF_STR2CUIT(Trim(pRS("NUCUITREM")))
        end if
    end if
 if (flagIsCamion) then            
		'Obtengo el Intermediario y el Remitente Comercial
		strSQL="Select CO.*, V.DSVENDEDOR from HCUENTAYORDENESCAMIONES CO inner join VENDEDORES V on CO.CDVENDEDOR=V.CDVENDEDOR where DTCONTABLE='" & GF_FN2DTCONTABLE(pRs("DTCONTABLE")) & "' and IDCAMION='" & pRs("IDCAMION") & "'"
		Call GF_BD_Puertos(pPto, rs, "OPEN", strSQL)
		while (not rs.eof) 
		    if (CInt(rs("SQORDEN")) = 1) then
		        pIntermediario = rs("DSVENDEDOR")            
		    else
		        pRteComercial  = rs("DSVENDEDOR")
		    end if            
		    rs.MoveNext()
		wend
    else
		pIntermediario = " "
		pRteComercial = Trim(pRs("dsvendedor"))
    end if
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawTitle(pPto, pRs)	    
	Dim auxCosecha, myTitular, myIntermediario, myRteComercial
	
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_BOLD)
	Call GF_squareBox(oPDF,5  ,150,585,140,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_horizontalLine(oPDF,5,180,585)
	Call GF_horizontalLine(oPDF,5,260,585)
	Call GF_verticalLine(oPDF,292,150,30)
	Call GF_verticalLine(oPDF,292,260,30)
	Call GF_writeTextAlign(oPDF,8  ,160, "Destinatario:", 20,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,300,160, "Vendedor:" , 20,PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(oPDF,8,190, "Titular.......:" , 20,PDF_ALIGN_LEFT)			
	Call GF_writeTextAlign(oPDF,8,206, "Rte. Comercial:" , 20,PDF_ALIGN_LEFT)		
	Call GF_writeTextAlign(oPDF,8,222, "Intermediario.:" , 20,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,8,239, "Corredor......:" , 20,PDF_ALIGN_LEFT)	
	
	Call GF_writeTextAlign(oPDF,295,190, "Procedencia:" , 20,PDF_ALIGN_LEFT)					
	Call GF_writeTextAlign(oPDF,295,239, "Entregador.:" , 20,PDF_ALIGN_LEFT)
	
	Call GF_writeTextAlign(oPDF,8,271, "Producto:" , 20,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,295,271, "Cosecha:" , 20,PDF_ALIGN_LEFT)

	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)		
	'if (not IsNull(pRs("DSEMPRESA"))) Then
	'	auxDsEmpresa = Trim(pRs("DSEMPRESA"))
	'	if (Len(auxDsEmpresa) > 55) then auxDsEmpresa = Left(auxDsEmpresa,50) & ".."
	'	Call GF_writeTextAlign(oPDF,65  ,160, auxDsEmpresa , 285,PDF_ALIGN_LEFT)
	'end if	
	Call GF_writeTextAlign(oPDF,80  ,160, recortarDescripcion(pRs("DSCLIENTE")) , 285,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,350  ,160, recortarDescripcion(pRs("DSVENDEDOR")) , 285,PDF_ALIGN_LEFT)
	
	Call getPartesInvolucradas(pPto, pRs, myTitular, myIntermediario, myRteComercial)
	Call GF_writeTextAlign(oPDF,90  ,190, recortarDescripcion(myTitular) , 285,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,90  ,206, recortarDescripcion(myRteComercial) , 285,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,90  ,222, recortarDescripcion(myIntermediario) , 285,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(oPDF,90  ,239, recortarDescripcion(pRs("DSCORREDOR")) , 285,PDF_ALIGN_LEFT)	
	Call GF_writeTextAlign(oPDF,360  ,190, recortarDescripcion(pRs("DSPROCEDENCIA")) , 285,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,360, 239, recortarDescripcion(pRs("DSENTREGADOR")) , 285,PDF_ALIGN_LEFT)
	Call GF_writeTextAlign(oPDF,65  ,271, pRs("DSPRODUCTO") , 285,PDF_ALIGN_LEFT)
	if (not IsNull(pRs("CDCOSECHA"))) Then
		auxCosecha = pRs("CDCOSECHA")
		if (Len(auxCosecha) = 8) then auxCosecha = Left(pRs("CDCOSECHA"),4) &"/"& Right(pRs("CDCOSECHA"),4)
		Call GF_writeTextAlign(oPDF,350  ,271, auxCosecha , 285,PDF_ALIGN_LEFT)		
	end if	
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_BOLD)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function recortarDescripcion(pDs)
    Dim aux 
'response.write "kuku"  & "pipi"
'response.end
    recortarDescripcion = ""
    if (not IsNull(pDs) and pDs <> "") Then
		aux = Trim(pDs)
		if (Len(aux) > 40) then aux = Left(aux,40) & ".."
		recortarDescripcion = aux		    		
	end if	
    'response.write "esta aca"
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawDetail(pRs, pTipoCamion)
	Call GF_squareBox(oPDF,5 ,295,585,400,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)	
	
	Call GF_horizontalLine(oPDF,5,335,585)
	Call GF_horizontalLine(oPDF,5,375,585)
	Call GF_horizontalLine(oPDF,5,415,585)
	Call GF_horizontalLine(oPDF,5,455,585)
	Call GF_horizontalLine(oPDF,5,495,585)
	Call GF_horizontalLine(oPDF,5,535,585)
		
	if (flagIsCamion) then
		Call GF_verticalLine(oPDF,146,295,80)
		Call GF_verticalLine(oPDF,292,295,80)		
		Call GF_verticalLine(oPDF,438,295,80)
		Call GF_verticalLine(oPDF,516,295,40)
		Call GF_writeTextAlign(oPDF,5,300, "N° Interno:" , 140,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,80,300, "Carta Porte:" , 280,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,227,300, "CTG:" , 280,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,438,300, "Camion:" , 75,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,516,300, "Acolplado:" , 75,PDF_ALIGN_CENTER)
	else
		Call GF_verticalLine(oPDF,146,295,80)
		Call GF_verticalLine(oPDF,292,335,40)
		Call GF_verticalLine(oPDF,438,295,80)
		Call GF_writeTextAlign(oPDF,5,300, "N° Interno:" , 140,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,146,300, "Carta Porte:" , 280,PDF_ALIGN_CENTER)
		Call GF_writeTextAlign(oPDF,438,300, "Vagon:" , 140,PDF_ALIGN_CENTER)
	end if	
	
	Call GF_writeTextAlign(oPDF,5  ,340, "Bruto Origen:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,340, "Tara Origen:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,295,340, "Neto Origen:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,440,340, "Oblea N°:" , 145,PDF_ALIGN_CENTER)
	
	Call GF_verticalLine(oPDF,195,375,40)
	Call GF_verticalLine(oPDF,390,375,40)
	
	Call GF_writeTextAlign(oPDF,5  ,380, "Transportista:" , 193,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,198,380, "Domicilio:" , 193,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("CDTIPODOC"))) then 
		Call GF_writeTextAlign(oPDF,391,380, pRs("CDTIPODOC") & ":" , 193,PDF_ALIGN_CENTER)
	else 
		Call GF_writeTextAlign(oPDF,391,380, "C.U.I.T:" , 193,PDF_ALIGN_CENTER)
	end if
	Call GF_verticalLine(oPDF,117,415,40)
	Call GF_verticalLine(oPDF,234,415,40)
	Call GF_verticalLine(oPDF,351,415,40)
	Call GF_verticalLine(oPDF,468,415,40)
	
	Call GF_writeTextAlign(oPDF,5  ,420, "Hs Entrada:" , 117,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,117,420, "Hs Calada:" , 117,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,234,420, "Hs P.Bruto:" , 117,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,351,420, "Hs P.Tara:" , 117,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,468,420, "Hs Salida:" , 117,PDF_ALIGN_CENTER)	
	
	Call GF_verticalLine(oPDF,146,455,80)
	Call GF_verticalLine(oPDF,292,455,80)
	Call GF_verticalLine(oPDF,438,455,80)
	
	Call GF_writeTextAlign(oPDF,5  ,460, "Merma Secado:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,460, "Merma Zarandeo:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,295,460, "Merma Convenida:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,440,460, "Otras Mermas:" , 145,PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(oPDF,5  ,500, "Bruto:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,500, "Tara:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,295,500, "Total Merma:" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,440,500, "Neto:" , 145,PDF_ALIGN_CENTER)
		
	Call GF_writeTextAlign(oPDF,8,540, "Observaciones:" , 580,PDF_ALIGN_LEFT)
	
	if flagIsCamion then
		Call drawDetailCamiones(pRs, pTipoCamion)
	else 
		Call drawDetailVagones(pRs)
	end if
	
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawDetailVagones(pRs)
	Dim auxDsTransportista,auxDsDomicilio,auxNumDocumento,auxMermaConvenida,auxAceptacionRabaja,px,py
	
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,315, pRs("SQTURNO") , 140,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,146,315,GF_EDIT_CTAPTE(pRs("NUCARTAPORTESERIE")&GF_nChars(pRs("NUCARTAPORTE"),12,"0",CHR_AFT)) , 280,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,438,315, pRs("CDVAGON") , 140,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("VLBRUTOORIGEN"))) Then Call GF_writeTextAlign(oPDF,5  ,355, GF_EDIT_DECIMALS(pRs("VLBRUTOORIGEN"),0) & " Kg.", 145,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("VLTARAORIGEN"))) Then Call GF_writeTextAlign(oPDF,150,355, GF_EDIT_DECIMALS(pRs("VLTARAORIGEN"),0) & " Kg." , 145,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("VLNETOORIGEN"))) Then Call GF_writeTextAlign(oPDF,295,355, GF_EDIT_DECIMALS(pRs("VLNETOORIGEN"),0) & " Kg." , 145,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("NUBARRAS"))) Then Call GF_writeTextAlign(oPDF,440,355, pRs("NUBARRAS") , 145,PDF_ALIGN_CENTER)
	
	if (not IsNull(pRs("DSTRANSPORTISTA"))) Then
		auxDsTransportista = Trim(pRs("DSTRANSPORTISTA"))
		if (Len(auxDsTransportista) > 45) then auxDsTransportista = Left(auxDsTransportista,42) & ".."
		Call GF_writeTextAlign(oPDF,5  ,395, auxDsTransportista , 193,PDF_ALIGN_CENTER)
	end if	
	if (not IsNull(pRs("DSDOMICILIO"))) Then
		auxDsDomicilio = Trim(pRs("DSDOMICILIO"))
		if (Len(auxDsDomicilio) > 45) then auxDsDomicilio = Left(auxDsDomicilio,42) & ".."
		Call GF_writeTextAlign(oPDF,198,395, auxDsDomicilio , 193,PDF_ALIGN_CENTER)
	end if
	if (not IsNull(pRs("NUDOCUMENTO"))) Then
		auxNumDocumento = Trim(pRs("NUDOCUMENTO"))
		if (Len(auxNumDocumento) > 45) then auxNumDocumento = Left(auxNumDocumento,42) & ".."
		Call GF_writeTextAlign(oPDF,391,395, auxNumDocumento , 193,PDF_ALIGN_CENTER)
	end if
	
	if (not IsNull(pRs("DTARRIBO"))) Then Call GF_writeTextAlign(oPDF,5  ,435, GF_FN2DTE(pRs("DTARRIBO")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTCALADA"))) Then Call GF_writeTextAlign(oPDF,117,435, GF_FN2DTE(pRs("DTCALADA")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTPESADABRUTO"))) Then Call GF_writeTextAlign(oPDF,234,435, GF_FN2DTE(pRs("DTPESADABRUTO")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTPESADATARA"))) Then Call GF_writeTextAlign(oPDF,351,435, GF_FN2DTE(pRs("DTPESADATARA")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTFIN"))) Then Call GF_writeTextAlign(oPDF,468,435, GF_FN2DTE(pRs("DTFIN")) , 117,PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(oPDF,5  ,475, "0 Kg." , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,475, "0 Kg." , 145,PDF_ALIGN_CENTER)	
	auxMermaConvenida = 0
	if (not IsNull(pRs("PCMERMA"))) Then auxMermaConvenida = pRs("PCMERMA")
	Call GF_writeTextAlign(oPDF,295,475, GF_EDIT_DECIMALS(auxMermaConvenida,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,440,475, "0 Kg." , 145,PDF_ALIGN_CENTER)
	
	auxVlBruto = 0
	if (not IsNull(pRs("VLPESADABRUTO"))) Then auxVlBruto = pRs("VLPESADABRUTO")
	Call GF_writeTextAlign(oPDF,5  ,515, GF_EDIT_DECIMALS(auxVlBruto,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	auxVlTara = 0	
	if (not IsNull(pRs("VLPESADATARA"))) Then auxVlTara = pRs("VLPESADATARA")
	Call GF_writeTextAlign(oPDF,150,515, GF_EDIT_DECIMALS(auxVlTara,0) & " Kg." , 145,PDF_ALIGN_CENTER)	
	auxVlMerma = 0	
	if (not IsNull(pRs("VLMERMAKILOS"))) Then auxVlMerma = pRs("VLMERMAKILOS")
	Call GF_writeTextAlign(oPDF,295,515, GF_EDIT_DECIMALS(auxVlMerma,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	
	auxNeto = Cdbl(auxVlBruto) - Cdbl(auxVlTara) - Cdbl(auxVlMerma)	
	Call GF_writeTextAlign(oPDF,440,515, GF_EDIT_DECIMALS(auxNeto,0) & " Kg.", 145,PDF_ALIGN_CENTER)
	
	'Observaciones	
	if (not isNull(pRs("CDACEPTACION"))) then
		if ((CLng(pRs("CDACEPTACION")) = ACEPTACION_REBAJA_CONVENIDA) or (CLng(pRs("CDACEPTACION")) = ACEPTACION_RECHAZO)) then
			Set rsRub = getRubroVagonBySQCalada(cartaPorte,cdVagon,dtContable,pRs("SQCALADA"))
			if not rsRub.Eof then
				auxParamHumedad = getValueParametro(PARAM_CD_RUBRO_HUMEDAD,pto)
				px = 8
				py = 555
				while not rsRub.Eof
					if (auxParamHumedad <> rsRub("CDRUBRO")) then
						Call GF_writeTextAlign(oPDF,px,py, Trim(rsRub("DSRUBRO")) & " ( " & rsRub("VLBONREBAJA") & " ) " , 580,PDF_ALIGN_LEFT)
						py = py + 15
						if (py > 680) then
							py = 550
							px = 280
						end if	
					end if
					rsRub.MoveNext()
				wend
			end if
		else
			Call GF_writeTextAlign(oPDF,8,550, pRs("DSACEPTACION") , 580,PDF_ALIGN_LEFT)
		end if
	end if	
	if (CInt(pRs("CDESTADO")) = CAMIONES_ESTADO_RECHAZADO) then Call GF_writeTextAlign(oPDF,8, py, "MOTIVO DEL REACHAZO: " & pRs("DSOBSERVACION") , 580,PDF_ALIGN_LEFT)	
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getRubroVagonBySQCalada(pcartaPorte,pCdVagon,pDtContable,pSqCalada)
	Dim strSQL,diaHoy
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)	
	strSQL = "Select a.*,b.dsrubro from ("&_
			 "	Select (YEAR(DtContable)*10000 + Month(DtContable)*100 + DAY(DtContable)) dtcontable,cdoperativo,cdvagon,sqcalada,cdrubro,vlbonrebaja from HRUBROSVISTEOVAGONES "&_
			 "	union "&_
			 "	Select "&diaHoy&" as dtcontable,cdoperativo,cdvagon,sqcalada,cdrubro,vlbonrebaja from RUBROSVISTEOVAGONES) as a "&_
			 "inner join rubros b on a.cdrubro = b.cdrubro "&_
			 "where dtcontable = "&pDtContable&" and cdoperativo = '"&cartaPorte&"' and cdvagon = '"&cdVagon&"' and sqcalada = "&pSqCalada
	
    Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	Set getRubroVagonBySQCalada = rs
End Function
'------------------------------------------------------------------------------------------------------------------------
Function getRubroCamionBySQCalada(pIdCamion,pDtContable,pSqCalada)
	Dim strSQL,diaHoy
	
	diaHoy = Year(Now()) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)		
	strSQL = "Select a.*,b.dsrubro from ("&_
			 "	Select (YEAR(DtContable)*10000 + Month(DtContable)*100 + DAY(DtContable)) dtcontable,idcamion,sqcalada,cdrubro,vlbonrebaja from HRUBROSVISTEOCAMIONES "&_
			 "	union "&_
			 "	Select "&diaHoy&" as dtcontable,idcamion,sqcalada,cdrubro,vlbonrebaja from RUBROSVISTEOCAMIONES) as a "&_
			 "inner join rubros b on a.cdrubro = b.cdrubro "&_
			 "where dtcontable = "& pDtContable &" and idcamion = '"&pIdCamion&"' and sqcalada = "&pSqCalada
	Call GF_BD_Puertos(pto, rs, "OPEN", strSQL)
	Set getRubroCamionBySQCalada = rs
End Function 
'------------------------------------------------------------------------------------------------------------------------
'Dibuja las firmas
Function drawFooter()
	Call GF_squareBox(oPDF,5 ,700,585,140,0,"#FFFFFF",NEGRO,1,PDF_SQUARE_ROUND)
	Call GF_horizontalLine(oPDF,26,815,150)
	Call GF_writeTextAlign(oPDF,26,820, "Firma Conductor" , 150,PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,225,815,150)
	Call GF_writeTextAlign(oPDF,225,820, "Firma Responsable" , 150,PDF_ALIGN_CENTER)
	Call GF_horizontalLine(oPDF,415,815,150)
	Call GF_writeTextAlign(oPDF,415,820, "Firma Responsable" , 150,PDF_ALIGN_CENTER)
End Function
'------------------------------------------------------------------------------------------------------------------------
Function drawDetailCamiones(pRs, pTipoCamion)
	Dim auxDsTransportista,auxDsDomicilio,auxNumDocumento,auxMermaConvenida,auxAceptacionRabaja,px,py
	Dim myBrutoOrigen, myTaraOrigen
	
	Call GF_setFont(oPDF,"COURIER", 8 , FONT_STYLE_NORMAL)
	Call GF_writeTextAlign(oPDF,5,315, pRs("IDCAMION") , 140,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,150,315,GF_EDIT_CTAPTE(GF_nChars(pRs("NUCARTAPORTE"),12,"0",CHR_AFT)), 145,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,295,315,pRs("CTG"), 145,PDF_ALIGN_CENTER)	
	Call GF_writeTextAlign(oPDF,438,315, GF_EDIT_PATENTE(pRs("CDCHAPACAMION")) , 75,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,516,315, GF_EDIT_PATENTE(pRs("CDCHAPAACOPLADO")) , 75,PDF_ALIGN_CENTER)	
	
	myBrutoOrigen = Trim(pRs("VLBRUTOORIGEN"))
	myTaraOrigen = Trim(pRs("VLTARAORIGEN"))
	if (pTipoCamion = CIRCUITO_CAMION_CARGA) then
	    'Si el camión se cargó en el puerto, entonces el origen es lo pesado en el propio puerto.
	    myBrutoOrigen = Trim(pRs("VLPESADABRUTO"))
	    myTaraOrigen = Trim(pRs("VLPESADATARA"))
	end if	
	if (not isNumeric(myBrutoOrigen)) then myBrutoOrigen = 0
	if (not isNumeric(myTaraOrigen)) then myTaraOrigen = 0
	Call GF_writeTextAlign(oPDF,5  ,355, GF_EDIT_DECIMALS(myBrutoOrigen,0) & " Kg.", 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,355, GF_EDIT_DECIMALS(myTaraOrigen,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,295,355, GF_EDIT_DECIMALS(Cdbl(myBrutoOrigen)-Cdbl(myTaraOrigen),0) & " Kg." , 145,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("NUBARRAS"))) Then Call GF_writeTextAlign(oPDF,440,355, pRs("NUBARRAS") , 145,PDF_ALIGN_CENTER)
	
	if (not IsNull(pRs("DSTRANSPORTISTA"))) Then
		auxDsTransportista = Trim(pRs("DSTRANSPORTISTA"))
		if (Len(auxDsTransportista) > 45) then auxDsTransportista = Left(auxDsTransportista,42) & ".."
		Call GF_writeTextAlign(oPDF,5  ,395, auxDsTransportista , 193,PDF_ALIGN_CENTER)
	end if	
	if (not IsNull(pRs("DSDOMICILIO"))) Then
		auxDsDomicilio = Trim(pRs("DSDOMICILIO"))
		if (Len(auxDsDomicilio) > 45) then auxDsDomicilio = Left(auxDsDomicilio,42) & ".."
		Call GF_writeTextAlign(oPDF,198,395, auxDsDomicilio , 193,PDF_ALIGN_CENTER)
	end if
	if (not IsNull(pRs("NUDOCUMENTO"))) Then
		auxNumDocumento = GF_STR2CUIT(Trim(pRs("NUDOCUMENTO")))
		Call GF_writeTextAlign(oPDF,391,395, auxNumDocumento , 193,PDF_ALIGN_CENTER)
	end if
	
	if (not IsNull(pRs("DTINGRESO"))) Then Call GF_writeTextAlign(oPDF,5  ,435, GF_FN2DTE(pRs("DTINGRESO")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTCALADA"))) Then Call GF_writeTextAlign(oPDF,117,435, GF_FN2DTE(pRs("DTCALADA")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTPESADABRUTO"))) Then Call GF_writeTextAlign(oPDF,234,435, GF_FN2DTE(pRs("DTPESADABRUTO")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTPESADATARA"))) Then Call GF_writeTextAlign(oPDF,351,435, GF_FN2DTE(pRs("DTPESADATARA")) , 117,PDF_ALIGN_CENTER)
	if (not IsNull(pRs("DTEGRESO"))) Then Call GF_writeTextAlign(oPDF,468,435, GF_FN2DTE(pRs("DTEGRESO")) , 117,PDF_ALIGN_CENTER)
	
	Call GF_writeTextAlign(oPDF,5  ,475, "0 Kg." , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,150,475, "0 Kg." , 145,PDF_ALIGN_CENTER)	
	auxMermaConvenida = 0
	if (not IsNull(pRs("PCMERMA"))) Then auxMermaConvenida = CDbl(pRs("PCMERMA"))
	Call GF_writeTextAlign(oPDF,295,475, GF_EDIT_DECIMALS(auxMermaConvenida*100,2) & " %" , 145,PDF_ALIGN_CENTER)
	Call GF_writeTextAlign(oPDF,440,475, "0 Kg." , 145,PDF_ALIGN_CENTER)
	
	auxVlBruto = 0
	if (not IsNull(pRs("VLPESADABRUTO"))) Then auxVlBruto = pRs("VLPESADABRUTO")
	Call GF_writeTextAlign(oPDF,5  ,515, GF_EDIT_DECIMALS(auxVlBruto,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	auxVlTara = 0	
	if (not IsNull(pRs("VLPESADATARA"))) Then auxVlTara = pRs("VLPESADATARA")
	Call GF_writeTextAlign(oPDF,150,515, GF_EDIT_DECIMALS(auxVlTara,0) & " Kg." , 145,PDF_ALIGN_CENTER)	
	auxVlMerma = 0	
	if (not IsNull(pRs("VLMERMAKILOS"))) Then auxVlMerma = pRs("VLMERMAKILOS")
	Call GF_writeTextAlign(oPDF,295,515, GF_EDIT_DECIMALS(auxVlMerma,0) & " Kg." , 145,PDF_ALIGN_CENTER)
	
	auxNeto = Cdbl(auxVlBruto) - Cdbl(auxVlTara) - Cdbl(auxVlMerma)	
	Call GF_writeTextAlign(oPDF,440,515, GF_EDIT_DECIMALS(auxNeto,0) & " Kg.", 145,PDF_ALIGN_CENTER)
	
	'Observaciones	 
	if (not isNull(pRs("CDACEPTACION"))) then
		if ((CLng(pRs("CDACEPTACION")) = ACEPTACION_REBAJA_CONVENIDA) or (CLng(pRs("CDACEPTACION")) = ACEPTACION_RECHAZO)) then
			Set rsRub = getRubroCamionBySQCalada(idCamion,dtContable,pRs("SQCALADA"))
			if not rsRub.Eof then
				auxParamHumedad = getValueParametro(PARAM_CD_RUBRO_HUMEDAD,pto)
				px = 8
				py = 555
				while not rsRub.Eof
					if (auxParamHumedad <> rsRub("CDRUBRO")) then
						Call GF_writeTextAlign(oPDF,px,py, Trim(rsRub("DSRUBRO")) & " ( " & rsRub("VLBONREBAJA") & " ) " , 580,PDF_ALIGN_LEFT)
						py = py + 15
						if (py > 680) then
							py = 550
							px = 280
						end if	
					end if
					rsRub.MoveNext()
				wend
			end if
		else
			Call GF_writeTextAlign(oPDF,8,550, pRs("DSACEPTACION") , 580,PDF_ALIGN_LEFT)
		end if
	end if	
	if (CInt(pRs("CDESTADO")) = CAMIONES_ESTADO_RECHAZADO) then Call GF_writeTextAlign(oPDF,8, py, "MOTIVO DEL RECHAZO: " & pRs("DSOBSERVACIONES") , 580,PDF_ALIGN_LEFT)
End Function
'------------------------------------------------------------------------------------------------------------------------
Dim rs, conn, strSQL, oPDF, pto, idCamion, cartaPorte, cdVagon, dtContable, SEPARATION, MARGIN, PAGE_HEIGHT_SIZE, PAGE_TOP_INIT
Dim flagIsCamion,tipoCamion, nroCopias, cantCopiasAux

SEPARATION = 10
MARGIN = 0
PAGE_HEIGHT_SIZE = 800
PAGE_TOP_INIT = 82
nroPagina = 1
cantCopiasAux = 1

pto			 = GF_Parametros7("pto", "", 6)
cartaPorte   = GF_Parametros7("cartaPorte", "", 6)
cdVagon		 = GF_Parametros7("cdVagon", "", 6)
dtContable   = GF_Parametros7("dtContable", "", 6)
idCamion	 = GF_Parametros7("idCamion", "", 6)
nroCopias	 = GF_Parametros7("nroCopias", 0, 6)
tipoCamion   = GF_Parametros7("tipoCamion", 0, 6)
if (tipoCamion = "") then tipoCamion = CIRCUITO_CAMION_DESCARGA



flagIsCamion = false
if (idCamion <> "") then flagIsCamion = true

Set oPDF = GF_createPDF("PDFTemp")
Call GF_setPDFMODE(PDF_STREAM_MODE)
call armadoPDF(cartaPorte,cdVagon,dtContable,pto,tipoCamion,idCamion)
if nroCopias > 1 then
	while cantCopiasAux < nroCopias
		Call GF_newPage (oPDF)
		call armadoPDF(cartaPorte,cdVagon,dtContable,pto,tipoCamion,idCamion)
		cantCopiasAux = cantCopiasAux + 1
	wend
end if
Call GF_closePDF(oPDF)


%>