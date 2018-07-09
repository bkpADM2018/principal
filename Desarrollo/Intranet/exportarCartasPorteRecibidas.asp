<!--#include file="Includes/procedimientosUnificador.asp"-->
<!--#include file="Includes/procedimientosFacturacionCalidad.asp"-->
<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientos.asp"-->
<!--#include file="Includes/procedimientosPuertos.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<!--#include file="Includes/procedimientosMail.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->

<%
'------------------------------------------------------------------------------------------------------------------
Function getSqlCdProductoTercero()
getSqlCdProductoTercero = "(select cdtercero from tblconversiones where nucuitcliente = '" & CONV_KEY_AFIP & "'"&_
			 " and tipodato = '" & CONV_KEY_PRODUCTO & "' and cdpropio = cc.cdproducto)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlCdProductoTerceroVagones()
getSqlCdProductoTerceroVagones = "(select cdtercero from tblconversiones where nucuitcliente = '" & CONV_KEY_AFIP & "'"&_
			 " and tipodato = '" & CONV_KEY_PRODUCTO & "' and cdpropio = ho.cdproducto)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlCuitEntregador()
getSqlCuitEntregador = "(select nucuit from entregadores e where e.cdentregador = cc.cdentregador)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlCuitEntregadorVagones()
getSqlCuitEntregadorVagones = "(select nucuit from entregadores e where e.cdentregador = ho.cdentregador)"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlPesoNeto()
getSqlPesoNeto = "(select (select vlpesada from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " & PESADA_BRUTO & "  " &_
			 " and sqpesada = (select max(sqpesada) from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_BRUTO & ")) "&_             
             " - (select vlpesada from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_TARA & " "&_                          
             " and sqpesada = (select max(sqpesada) from hpesadascamion where dtcontable = cc.dtcontable and idcamion = cc.idcamion and cdpesada = " &  PESADA_TARA & ")))"
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlPesoNetoVagones(pFecha)
		getSqlPesoNetoVagones = "(	" &_
        	" SELECT Sum(vlpesada) " &_
        	" FROM (SELECT dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, dtpesada, vlpesada FROM hpesadasvagon " &_
         	"	UNION " &_
         	"	SELECT Cast(Getdate() AS DATE) AS dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, dtpesada, vlpesada FROM pesadasvagon ) hpv " &_
         	"	INNER JOIN " &_
         	"		(	" &_
         	"			SELECT dtcontable, nucartaporte, cdvagon, Max(sqpesada) AS sqpesada " &_
         	"			FROM ( " &_
         	"				SELECT dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, vlpesada FROM hpesadasvagon where NUCARTAPORTE=ho.nucartaporte and CDVAGON in (Select CDVAGON from HVAGONES where NUCARTAPORTE=ho.nucartaporte and DTCONTABLEVAGON='" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) and cdpesada = " & PESADA_BRUTO &_
         	"				UNION  " &_
            "   			SELECT Cast(Getdate() AS DATE) AS dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, vlpesada FROM pesadasvagon where NUCARTAPORTE=ho.nucartaporte and CDVAGON in (Select CDVAGON from VAGONES where NUCARTAPORTE=ho.nucartaporte and DTCONTABLEVAGON='" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) and cdpesada = " & PESADA_BRUTO  &_
            "   			) t1 " &_
            "   		GROUP BY dtcontable, nucartaporte, cdvagon " &_
            "   	) hpvs " &_
            "   ON hpv.dtcontable = hpvs.dtcontable AND hpv.nucartaporte = hpvs.nucartaporte AND hpv.cdvagon = hpvs.cdvagon AND hpv.sqpesada = hpvs.sqpesada  " &_
            "   GROUP BY hpvs.nucartaporte " &_
          " ) " &_
          " - " &_
          " ( " &_
        	" SELECT Sum(vlpesada) " &_
        	" FROM (SELECT dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, dtpesada, vlpesada FROM hpesadasvagon " &_
         	"	UNION " &_
         	"	SELECT Cast(Getdate() AS DATE) AS dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, dtpesada, vlpesada FROM pesadasvagon ) hpv " &_
         	"	INNER JOIN " &_
         	"		(	" &_
         	"			SELECT dtcontable, nucartaporte, cdvagon, Max(sqpesada) AS sqpesada " &_
         	"			FROM (" &_
         	"				SELECT dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, vlpesada FROM hpesadasvagon where NUCARTAPORTE=ho.nucartaporte and CDVAGON in (Select CDVAGON from HVAGONES where NUCARTAPORTE=ho.nucartaporte and DTCONTABLEVAGON='" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) and cdpesada = " & PESADA_TARA &_
         	"				UNION " &_
            "  			SELECT Cast(Getdate() AS DATE) AS dtcontable, nucartaporte, cdvagon, sqpesada, cdpesada, vlpesada FROM pesadasvagon where NUCARTAPORTE=ho.nucartaporte and CDVAGON in (Select CDVAGON from VAGONES where NUCARTAPORTE=ho.nucartaporte and DTCONTABLEVAGON='" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) and cdpesada = " & PESADA_TARA &_
            "		) t1" &_
            "   		GROUP BY dtcontable, nucartaporte, cdvagon" &_
            "  	) hpvs " &_
            "   ON hpv.dtcontable = hpvs.dtcontable AND hpv.nucartaporte = hpvs.nucartaporte AND hpv.cdvagon = hpvs.cdvagon AND hpv.sqpesada = hpvs.sqpesada " &_
            "   GROUP BY hpvs.nucartaporte " &_
          " ) "
end function
'------------------------------------------------------------------------------------------------------------------
Function getSqlCosechaVagones()
getSqlCosechaVagones = "(select  max(cdcosecha) from (select cdvagon, cdcosecha from hvagones where nucartaporte = ho.nucartaporte union "&_
			 " select cdvagon, cdcosecha from vagones where nucartaporte = ho.nucartaporte) tab)"
end function
'------------------------------------------------------------------------------------------------------------------
Function obtenerCartasPorteRecibidas(pFecha,pPto)
	Dim strSQL,myWhere
	Dim sqlPesoNeto, sqlCdProductoTercero, sqlCuitEntregador
	myWhere = " where cc.CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ") "
	Call mkWhere(myWhere, "cc.DtContable", pFecha, "=", 3)
	Call mkWhere(myWhere, "cc.cdtipocamion", CIRCUITO_CAMION_DESCARGA, "=", 1)
	
	sqlCdProductoTercero = getSqlCdProductoTercero()
	sqlCuitEntregador = getSqlCuitEntregador()
	sqlPesoNeto = getSqlPesoNeto()	
		
	strSQL = "select do.nctipotransporte as TipoTransporte, "&_
			 " cc.cdtipocamion as TipoCartaPorte, "&_
			 " ltrim(rtrim(cc.nucartaporte)) as NroCartaPorte, "&_
			 " ltrim(rtrim(do.ncau)) as NroCEE, "&_
			 " ltrim(rtrim(cc.ctg)) as NroCTG, "&_
			 " CAST( (right(cc.dtcartaporte,2) + substring(cast(cc.dtcartaporte as varchar),6,2) + left(cc.dtcartaporte,4)) as varchar) as FechaCarga, "&_
			 " case when do.ncuitremitente is null then '00000000000' else do.ncuitremitente end as CuitTitular, "&_
			 " case when do.ncuitc1 is null then '00000000000' else do.ncuitc1 end as CuitIntermediario, "&_
			 " case when do.ncuitc2 is null then '00000000000' else do.ncuitc2 end as CuitRteComercial, "&_
             " case when do.ncuitcorre is null then '00000000000' else do.ncuitcorre end as CuitCorredor, " &_
             " case when " & sqlCuitEntregador & " is null then '00000000000'  else ltrim(rtrim(" & sqlCuitEntregador & ")) end as CuitRepteEntregador, "&_                          
             " case when do.ncuitdestinatario is null then '00000000000' else ltrim(rtrim(do.ncuitdestinatario)) end as CuitDestinatario, "&_             
             " case when do.nccuitdest is null then '00000000000' else do.nccuitdest end as CuitEstablDestino, "&_             
             " case when do.ncuittransportista is null then '00000000000' else do.ncuittransportista end as CuitTransportista, "&_                          
             " case when do.ncuitchofer is null then '00000000000' else do.ncuitchofer end as CuitChofer, "&_
			 " substring(cast(cdcosecha as varchar(8)),3,2) + '-' + right(cast(cdcosecha as varchar(8)),2) as Cosecha, "&_             
             " replicate('0', 3 - len(" & sqlCdProductoTercero & ")) + cast (" & sqlCdProductoTercero & " as varchar) as CodigoEspecie, "&_
			 " '00' as TipoGrano, "&_      
             " '00000000000000000001' as Contrato, "&_
			 " cc.cdtipocamion as TipoPesado, " &_
             " replace(replicate('0', 11-len(do.nqpesoc)) + cast(do.nqpesoc as varchar),'.',',') as PesoNetoOrigen, " &_       
             " Right(concat('000000', LTrim(RTrim(do.ncestableproce))), 6) as CdEstableProcedencia, " &_
             " replicate('0', 5-len(do.nclocproce)) + cast(ltrim(rtrim(do.nclocproce)) as varchar) as CdLocalidadProcedencia, "&_             
             " replicate('0', 6-len(do.ncestabledest)) + cast(ltrim(rtrim(do.ncestabledest)) as varchar) as CdEstableDestino, "&_             
             " replicate('0', 5-len(ltrim(rtrim(do.nclocdest)))) + ltrim(rtrim(cast(do.nclocdest as varchar))) as CdLocalidadDestino, "&_             
             " replicate('0', 4-len(ltrim(rtrim(do.kmrecorrer)))) + ltrim(rtrim(cast(do.kmrecorrer as varchar))) as KmRecorrer, "&_             
             " ltrim(rtrim(cc.cdchapacamion)) + replicate(' ', 11-len(ltrim(rtrim(cc.cdchapacamion)))) as PatenteCamion, "&_       
             " ltrim(rtrim(cc.cdchapaacoplado)) + replicate(' ', 11-len(ltrim(rtrim(cc.cdchapaacoplado)))) as PatenteAcoplado, "&_       
             " replace(replicate('0', 8-len(do.fitarifa)) + cast(do.fitarifa as varchar),'.',',') as TarifaTonelada, "&_                                       
			 " CAST( (right(cc.dtcontable,2) + substring(cast(cc.dtcontable as varchar),6,2) + left(cc.dtcontable,4)) as varchar) as FechaDescarga, "&_
			 " CAST( (right(do.NCFECHARRIBODEST,2) + substring(cast(do.NCFECHARRIBODEST as varchar),6,2) + left(do.NCFECHARRIBODEST,4)) as varchar) as FechaArribo, "&_
             " replace(replicate('0', 11-len( " & sqlPesoNeto & ")) + cast(" & sqlPesoNeto & " as varchar),'.',',') as PesoNeto, " &_
             " case when do.nccuitredest is null then '00000000000' else do.nccuitredest end as CuitEstablRedestino, "&_             
             " case when do.nclocredest is null then '00000' else replicate('0', 5-len(do.nclocredest)) + cast(do.nclocredest as varchar) end as CdLocalidadRedestino, "&_             
             " case when do.ncestableredest is null then '000000' else replicate('0', 6-len(do.ncestableredest)) + cast(do.ncestableredest as varchar) end as CdEstableRedestino, "&_             
             " case when replace(replicate('0', 8-len(do.tarifaref)) + cast(do.tarifaref as varchar),'.',',') is null then '00000,00' else replace(replicate('0', 8-len(do.tarifaref)) + cast(do.tarifaref as varchar),'.',',') end as TarifaReferencia "&_             
             " from "&_       
             " (select hc.DTCONTABLE as DTCONTABLE, hc.IDCAMION as IDCAMION,CDCHAPACAMION,CDCHAPAACOPLADO,CDTIPOCAMION,CDTRANSPORTISTA, "&_       
             " DSNOMBRECONDUCTOR,DSAPELLIDOCONDUCTOR,CDTIPODOC,NUDOCUMENTO,DTINGRESO,DTEGRESO,CDESTADO,CDCIRCUITO,CDFILA,CDSILO,CDPLATAFORMA,"&_                          
             " ICTRANSMITIDO,SQCAMION,CDPRODUCTO,NUAUTSALIDA,SQTURNO,ICCUPO,DSCUPO,ICCONTRATOESP,IDCUPOASIGNADO,NUCUITREM,NUCUPO, " &_
             " NUCARTAPORTE,DTCARTAPORTE,CDEMPRESA,CDCLIENTE,CDCORREDOR,CDVENDEDOR,CDCOSECHA,CDPROCEDENCIA,CDENTREGADOR, " &_
             " VLBRUTOORIGEN,VLTARAORIGEN,NUINFOANALISIS,NURECIBO,DTCPVENCIMIENTO,NUCTAPTEDIG,CTG,NUTICKETPLAYA " &_
             " from hcamiones hc inner join hcamionesdescarga hcd  "&_
			 " on hc.idcamion = hcd.idcamion and hc.dtcontable = hcd.dtcontable) cc " &_
			 "	left join datosoncca do on cc.nucartaporte = do.nccartaporte " &_
	     	 myWhere
			 
			 'logMig.info(strSQL)   			 

	Call executeQueryDb(pPto, rs, "OPEN", strSQL)	
	Set obtenerCartasPorteRecibidas = rs
End Function
'---------------------------------------------------------------------------------------------------------
Function obtenerCartasPorteRecibidasVagones(pFecha,pPto)
	Dim strSQL
	Dim sqlPesoNetoVagones, sqlCdProductoTerceroVagones, sqlCuitEntregadorVagones, sqlCosechaVagones
	
	sqlCdProductoTerceroVagones = getSqlCdProductoTerceroVagones()
	sqlCuitEntregadorVagones = getSqlCuitEntregadorVagones()
	sqlPesoNetoVagones = getSqlPesoNetoVagones(pFecha)
	sqlCosechaVagones = getSqlCosechaVagones()
	
	strSQL = "select case when do.nctipotransporte is Null then 0 else do.nctipotransporte end as TipoTransporte, "&_
			 " do.nctipocartaporte as TipoCartaPorte, "&_
			 " left(concat(ho.nucartaporteserie, ho.nucartaporte), 12) as NroCartaPorte, "&_
			 " ltrim(rtrim(do.ncau)) as NroCEE, "&_
			 " '00000000' as NroCTG, "&_
			 " CAST( (right(ho.dtemision,2) + substring(cast(ho.dtemision as varchar),6,2) + left(ho.dtemision,4)) as varchar) as FechaCarga,  "&_
			 " case when do.ncuitremitente is null then '00000000000' else do.ncuitremitente end as CuitTitular, "&_
			 " case when do.ncuitc1 is null then '00000000000' else do.ncuitc1 end as CuitIntermediario,  "&_
			 " case when do.ncuitc2 is null then '00000000000' else do.ncuitc2 end as CuitRteComercial, "&_
             " case when do.ncuitcorre is null then '00000000000' else do.ncuitcorre end as CuitCorredor, " &_
             " case when " & sqlCuitEntregadorVagones & " is null then '00000000000' else ltrim(rtrim(" & sqlCuitEntregadorVagones & ")) end as CuitRepteEntregador, "&_                          
             " case when do.ncuitdestinatario is null then '00000000000' else ltrim(rtrim(do.ncuitdestinatario)) end as CuitDestinatario, "&_             
             " case when do.nccuitdest is null then '00000000000' else do.nccuitdest end as CuitEstablDestino, "&_             
             " case when do.ncuittransportista is null then '00000000000' else do.ncuittransportista end as CuitTransportista, "&_                          
             " case when do.ncuittransportista is null then '00000000000' else do.ncuittransportista end as CuitConductor, "&_
			 " substring(cast(" & sqlCosechaVagones & " as varchar(8)),3,2) + '-' + right(cast(" & sqlCosechaVagones & " as varchar(8)),2) as Cosecha, "&_
             " 	replicate('0', 3 - len(" & sqlCdProductoTerceroVagones & ")) + cast (" & sqlCdProductoTerceroVagones & " as varchar) as CodigoEspecie,"&_
			 " '00' as TipoGrano, "&_      
             " '00000000000000000001' as Contrato, "&_
			 " do.nctipopesado as TipoPesado, " &_
             " replace(replicate('0', 11-len(do.nqpesoc)) + cast(do.nqpesoc as varchar),'.',',') as PesoNetoOrigen, " &_       
             " Right(concat('000000', LTrim(RTrim(do.ncestableproce))), 6) as CdEstableProcedencia, " &_
             " replicate('0', 5-len(do.nclocproce)) + cast(ltrim(rtrim(do.nclocproce)) as varchar) as CdLocalidadProcedencia, "&_             
             " replicate('0', 6-len(do.ncestabledest)) + cast(ltrim(rtrim(do.ncestabledest)) as varchar) as CdEstableDestino, "&_             
             " replicate('0', 5-len(ltrim(rtrim(do.nclocdest)))) + ltrim(rtrim(cast(do.nclocdest as varchar))) as CdLocalidadDestino, "&_             
             " replicate('0', 4-len(ltrim(rtrim(do.kmrecorrer)))) + ltrim(rtrim(cast(do.kmrecorrer as varchar))) as KmRecorrer, "&_             
             " 'SINPATENTE ' as PatenteCamion, "&_
			 " 'SINPATENTE ' as PatenteAcoplado, "&_
             " case when do.fitarifa is null then '00000,00' else replace(replicate('0', 8-len(do.fitarifa)) + cast(do.fitarifa as varchar),'.',',') end as TarifaTonelada, "&_                                       
			 " CAST( (right('" & pFecha & "',2) + substring(cast('" & pFecha & "' as varchar),6,2) + left('" & pFecha & "',4)) as varchar) as FechaDescarga, "&_
			 " CAST( (right(do.NCFECHARRIBODEST,2) + substring(cast(do.NCFECHARRIBODEST as varchar),6,2) + left(do.NCFECHARRIBODEST,4)) as varchar) as FechaArribo, "&_             
			 " replace(replicate ('0', 11-len(" & sqlPesoNetoVagones & ")) + cast( " & sqlPesoNetoVagones & " as varchar), '.',',') as PesoNeto,  "&_			 			 
			 " case when do.nccuitredest is null then '00000000000' else do.nccuitredest end as CuitEstablRedestino, " &_			 
			 " case when do.nclocredest is null then '00000' else replicate('0', 5-len(do.nclocredest)) + cast(do.nclocredest as varchar) end as CdLocalidadRedestino,  "&_             
             " case when do.ncestableredest is null then '000000' else replicate('0', 6-len(do.ncestableredest)) + cast(do.ncestableredest as varchar) end as CdEstableRedestino, "&_                          
             " case when do.tarifaref is null then '00001,00' else replace(replicate('0', 8-len(do.tarifaref)) + cast(do.tarifaref as varchar),'.',',') end as TarifaReferencia "&_             
             " from "&_
			 " ((select cdentregador, nucartaporteserie, nucartaporte, cdproducto, dtemision from hoperativos where nucartaporte in (select nucartaporte from hvagones where dtcontablevagon = "&_ 
			 " '" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ") union select nucartaporte from vagones where dtcontablevagon = '" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) ) "&_ 
			 " union (select cdentregador, nucartaporteserie, nucartaporte,cdproducto, dtemision from operativos where nucartaporte in (select nucartaporte from hvagones where dtcontablevagon = "&_ 
			 " '" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")  union select nucartaporte from vagones where dtcontablevagon = '" & pFecha & "' and CDESTADO in (" & CAMIONES_ESTADO_PESADOTARA & ", " & CAMIONES_ESTADO_EGRESADOOK & ")) )) ho "&_
			 " left join datosoncca do on left(concat(ho.nucartaporteserie, ho.nucartaporte), 12)=do.nccartaporte "
			 'logMig.info(strSQL)   			 
	Call executeQueryDb(pPto, rs, "OPEN", strSQL)	
	
	Set obtenerCartasPorteRecibidasVagones = rs
End Function
'---------------------------------------------------------------------------------------------------------
Function armarregistroDatos(myRs, myFile, pto)
    
	Dim myArrayAux
    
    On Error Resume Next    
	myArrayAux = myRs.GetRows
	registro = ""
	For row = 0 To UBound(myArrayAux, 2)
		For col = 0 To UBound(myArrayAux, 1) 
			registro = registro & CStr(myArrayAux(col, row))	
		Next
		if registro <> "" then myFile.WriteLine registro
		registro = ""			
	Next	
	logMig.info("Finalizado armado de registros para puerto " & pto)	
End Function
'---------------------------------------------------------------------------------------------------------
Function verificarIntegridadCartasPorte(rsCartasPorte, fecha, pto, listaErrores)
Dim sinErrores, myStrErrores
sinErrores = true
myStrErrores = ""
	if (not rsCartasPorte.eof) then
		while (not rsCartasPorte.Eof)
		if (isNull(rsCartasPorte("NroCartaPorte"))) then
			myStrErrores = myStrErrores & "Carta de Porte: sin datos ONCCA <br>"
			sinErrores = false
		else
		
			if (len(Trim(rsCartasPorte("NroCEE"))) <> 14) and (IsNumeric(rsCartasPorte("NroCEE"))) then 
				myStrErrores = myStrErrores & "Numero CEE: " & CStr(rsCartasPorte("NroCEE")) & ", longitud incorrecta <br>"			
				sinErrores = false
			end if			
			if (isNull(rsCartasPorte("CodigoEspecie"))) then
				myStrErrores = myStrErrores & "Codigo Producto: incorrecto o inexistente para datos oncca <br>"
				sinErrores = false
			end if
			if (isNull(rsCartasPorte("CuitEstablDestino"))) or (rsCartasPorte("CuitEstablDestino") = "00000000000") then
				myStrErrores = myStrErrores & "Cuit Establecimiento Destino: incorrecto o inexistente para datos oncca <br>"
				sinErrores = false
			end if
			if (isNull(rsCartasPorte("FechaArribo"))) then
				myStrErrores = myStrErrores & "Fecha de Arribo: incorrecta o no ingresada <br>"
				sinErrores = false
			end if	
			if (isNull(rsCartasPorte("PesoNeto"))) then
				myStrErrores = myStrErrores & "Peso Neto: incorrecto. Notificar al personal de sistemas <br>"
				sinErrores = false
			end if					
			if (rsCartasPorte("CdLocalidadProcedencia") = "") then
				myStrErrores = myStrErrores & "Localidad de Procedencia Incorrecta o la misma no fue especificada. <br>"
				sinErrores = false
			end if					
			'Controles propios de Camiones
			if (CInt(rsCartasPorte("TipoTransporte")) = TIPO_TRANSPORTE_CAMION) then
				if isNull(rsCartasPorte("TarifaReferencia")) then 
					myStrErrores = myStrErrores & "Tarifa de Referencia: valor incorrecto <br>"
					sinErrores = false
				elseif (CDbl(replace(rsCartasPorte("TarifaReferencia"),",",".")) <= 0) or (CDbl(replace(rsCartasPorte("TarifaReferencia"),",",".")) > 5000) then
					myStrErrores = myStrErrores & "Tarifa de Referencia: " & CStr(rsCartasPorte("TarifaReferencia")) & ", valor fuera de rango <br>"
					sinErrores = false
				end if
				if isNull(rsCartasPorte("TarifaTonelada")) then 
					myStrErrores = myStrErrores & "Tarifa por Tonelada: valor incorrecto <br>"
					sinErrores = false
				elseif (CDbl(replace(rsCartasPorte("TarifaTonelada"),",",".")) <= 0) or (CDbl(replace(rsCartasPorte("TarifaTonelada"),",",".")) > 5000) then
					myStrErrores = myStrErrores & " Tarifa por Tonelada: " & CStr(rsCartasPorte("TarifaTonelada")) & ", valor fuera de rango <br>"
					sinErrores = false
				end if
				if isNull(rsCartasPorte("KmRecorrer")) then 
					myStrErrores = myStrErrores & "Kilometros a Recorrer: valor incorrecto <br>"
					sinErrores = false
				elseif (CInt(rsCartasPorte("KmRecorrer")) <= 0)then
					myStrErrores = myStrErrores & " Kilometros a Recorrer: debe ser mayor a cero <br>"
					sinErrores = false
				end if				
				if (CLng(right(rsCartasPorte("FechaCarga"),4) & Mid(rsCartasPorte("FechaCarga"),3,2) & left(rsCartasPorte("FechaCarga"),2)) > CLng(right(rsCartasPorte("FechaDescarga"),4) & Mid(rsCartasPorte("FechaDescarga"),3,2) & left(rsCartasPorte("FechaDescarga"),2))) then
					myStrErrores = myStrErrores & " Fecha de Emision de CCPP: no puede ser mayor a Fecha de Descarga. <br>"
					sinErrores = false
				end if
				if (CLng(right(rsCartasPorte("FechaCarga"),4) & Mid(rsCartasPorte("FechaCarga"),3,2) & left(rsCartasPorte("FechaCarga"),2)) > CLng(right(rsCartasPorte("FechaArribo"),4) & Mid(rsCartasPorte("FechaArribo"),3,2) & left(rsCartasPorte("FechaArribo"),2))) then
					myStrErrores = myStrErrores & " Fecha de Emision de CCPP: no puede ser mayor a Fecha de Arribo. <br>"
					sinErrores = false
				end if
				if (CLng(rsCartasPorte("CdEstableProcedencia")) = 0) then
					myStrErrores = myStrErrores & " Falta especificar Establecimiento Procedencia. <br>"
					sinErrores = false
				end if
			end if
		end if
			
			'logMig.Info("salida de control: " & myStrErrores)
			if myStrErrores <> "" then listaErrores.Add CStr(rsCartasPorte("NroCartaPorte")), myStrErrores			
			myStrErrores = ""
			
			rsCartasPorte.MoveNext()
		wend
	else
		logMig.info ("algo salio mal, recordSet en verificarIntegridadDatos esta vacio")
	end if

verificarIntegridadCartasPorte = sinErrores
End function
'---------------------------------------------------------------------------------------------------------
Function enviarMailsCartasPorte(pFecha, pFileAttachment, listaErroresArroyo, listaErroresTransito,listaErroresPiedraBuena)
    Dim strBody, strSubject,fs
    strBody = ""    
    strSubject = "Cartas de Porte Recibidas de " & GF_FN2DTE(pFecha) 
    	
	Set fsAux = Server.CreateObject("Scripting.FileSystemObject")		
    if (fsAux.FileExists(pFileAttachment) and fsAux.GetFile(pFileAttachment).Size > 0) then    
        strBody = strBody & enviarMailConErroresParaPuerto(DBSITE_ARROYO, pFecha,listaErroresArroyo)
		strBody = strBody & enviarMailConErroresParaPuerto(DBSITE_TRANSITO, pFecha,listaErroresTransito)
		strBody = strBody & enviarMailConErroresParaPuerto(DBSITE_BAHIA, pFecha,listaErroresPiedraBuena)
		if strBody = "" then						
			strBody = "Se envia adjunto el archivo con las Cartas de Porte" & vbCrLf & vbCrLf & strBody 
			logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )
			Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_INFO_LIST, strSubject, strBody, pFileAttachment)
			'solo se actualiza la fecha de ultima ejecucion del archivo .lck, si la fecha no se ha pasado por parametro
			if not flagForzarEjecucion then flagActualizarFechaUltimaEjecucion = true
		end if
    else					
		strSubject = "Sin Cartas de Porte Recibidas de " & GF_FN2DTE(pFecha) 
		strBody = "No se registraron Cartas de Porte Recibidas para la fecha " & GF_FN2DTE(pFecha) & "."
		logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )
		Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_INFO_LIST, strSubject, strBody, "")
		'solo se actualiza la fecha de ultima ejecucion del archivo .lck, si la fecha no se ha pasado por parametro			
		if not flagForzarEjecucion then flagActualizarFechaUltimaEjecucion = true
    end if
	Set fsAux = Nothing
End Function
'---------------------------------------------------------------------------------------------------------
Function enviarMailConErroresParaPuerto(pto, pFecha,listaErrores)
	Dim strBody, strSubject, myLista, letraPuerto
    strBody = ""
	
    if listaErrores.Count > 0 then
		letraPuerto = getLetraPuerto(pto)
		logMig.info(" Enviando mail de la tarea " & TASK_POS_REPO_CCPP & " con fecha " & GF_FN2DTE(pFecha) )    
		strSubject = "Cartas de Porte Recibidas con errores de " & GF_FN2DTE(pFecha) & " " & pto 
		strBody = strBody & GF_FN2DTE(pFecha) & vbCrLf & "Cartas Porte con errores en " & pto & vbCrLf & vbCrLf
		logMig.info(strBody & " " & strSubject)
		for each cartaPorteConError in listaErrores.Keys
			strBody = strBody & vbTab & "Nro CCPP : " & cartaPorteConError & vbCrLf & vbTab & vbTab & listaErrores(cartaPorteConError)
		Next
		strBody = strBody & vbCrLf
		logMig.info(strBody & " " & strSubject)
		Call SendMail(TASK_POS_REPO_CCPP, MAIL_TASK_ERROR_LIST & letraPuerto, strSubject, strBody, "")		
	end if
	enviarMailConErroresParaPuerto = strBody
end function
'---------------------------------------------------------------------------------------------------------
Function getLastExecDate()
    Dim fso, myFile, myFilename, myData
    
    myFilename = server.MapPath(".") & "\servicioEnvioCartasPorteRecibidas.lck"
	'logMig.info("nombre archivo: " & myFilename)
    Set fso = CreateObject("scripting.filesystemobject")        
	if (fso.FileExists(myFilename)) then
		'logMig.info("archivo existe " & myFilename)
	    Set myFile = fso.OpenTextFile(myFilename,1,false)
        if (not myFile.AtEndOfStream) then
	        getLastExecDate = myfile.ReadLine	               
        else
            getLastExecDate = ""
        end if	        
	    myFile.Close
	    Set myFile = Nothing
	else
		logMig.info("archivo NO existe " & myFilename)
	    getLastExecDate = ""
	end if
	Set fso = Nothing
End Function
'---------------------------------------------------------------------------------------------------------
Function updateLastExecDate(pData)
    Dim fso, myFile, myFilename    
    myFilename = server.MapPath(".") & "\servicioEnvioCartasPorteRecibidas.lck"
    Set fso = CreateObject("scripting.filesystemobject")    
	if (fso.FileExists(myFilename)) then fso.DeleteFile(myFilename)
    Set myFile = fso.CreateTextFile(myFilename, true)
    myFile.WriteLine(pData)
    myFile.Close
	Set myFile = Nothing
	Set fso = Nothing	
End Function
'---------------------------------------------------------------------------------------------------------
Function procesarCartasPorte (rsCartasPorte, fecha, pto,listaErrores, pFile)
	if (not rsCartasPorte.eof) then
			if verificarIntegridadCartasPorte(rsCartasPorte, fecha, pto, listaErrores) then
				rsCartasPorte.MoveFirst()
				Call armarRegistroDatos(rsCartasPorte, pFile, pto)
			else
				logMig.info("Cartas Porte con errores en " & pto)
				for each cartaPorteConError in listaErrores.Keys
					logMig.info("Nro CCPP : " & cartaPorteConError & vbCrLf & listaErrores(cartaPorteConError))
				Next
			end if
		else
			logMig.info("No se encontraron Cartas de Porte para exportar para puerto " & pto & " en la fecha " & fechaContable)
		end if
End function
'---------------------------------------------------------------------------------------------------------
'                                   ***** COMIENZA PAGINA *****
'---------------------------------------------------------------------------------------------------------

Dim fecha, fechaContable, logMig, strSQL, registro, myFilename, g_strPuerto
Dim rsCartasPorteArroyo, rsCartasPorteTransito, rsCartasPortePiedraBuena, rsCartasPortePiedraBuenaVagones
Dim listaErroresArroyo,listaErroresTransito,listaErroresPiedraBuena
Dim fs, myFile, errMsg, myCartaPorte
Dim myArrayAux
Dim flagActualizarFechaUltimaEjecucion, flagForzarEjecucion
set listaErroresArroyo = Server.CreateObject("Scripting.Dictionary")
set listaErroresTransito = Server.CreateObject("Scripting.Dictionary")
set listaErroresPiedraBuena = Server.CreateObject("Scripting.Dictionary")

flagActualizarFechaUltimaEjecucion = false
flagForzarEjecucion = false

fecha = GF_PARAMETROS7("fd", "", 6)

Call GP_ConfigurarMomentos()
session("usuario") = "SYNC"


Set logMig = new classLog
Call startLog(HND_VIEW+HND_FILE,MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG)
logMig.fileName = "EXPORTACION_CARTASPORTE_" & GF_nDigits(Year(Now),4) & GF_nDigits(Month(Now()),2) & GF_nDigits(Day(Now()),2)

if (fecha = "") then
	fecha=getLastExecDate()
	if fecha = "" then
		'fecha = GF_DTEADD(Left(session("MmtoDato"), 8), -1, "D")		
		logMig.info("Error en tomar la fecha del archivo auxiliar de fecha servicioEnvioCartasPorteRecibidas.lck")
	else		
		logMig.info("Fecha tomada del archivo auxiliar de fecha servicioEnvioCartasPorteRecibidas.lck " & fecha)
		'fecha = GF_DTEADD(fecha, 1, "D")
	end if
else
	'se ha pasado fecha por parametro, se debe ejecutar proceso sin actualizacion de fecha
	flagForzarEjecucion = true
end if

if flagForzarEjecucion or (fecha < Left(session("MmtoDato"),8)) then
	fechaContable = GF_FN2DTCONTABLE(fecha)
	myFilename = server.MapPath(".\Temp") & "\CARTASPORTE_" & fecha & ".txt"
	logMig.info("archivo temporal: " & myFilename)

	logMig.info("------------ INCIANDO EXPORTACION CARTAS DE PORTE RECIBIDAS------------------")	
	logMig.info(" ---> FECHA : " & fechaContable)
	logMig.info("-----------------------------------------------------------------------------")	
		
		'On Error Resume Next    
		
 
		logMig.info("Inicializando archivo de datos: " & pFilename)
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		If fs.FileExists(pName) Then  Call fs.deleteFile(pName, true)
		Set myFile = fs.OpenTextFile(myFilename, 2, true)
		logMig.info("Archivo listo para trabajar.") 
		

		Set rsCartasPorteArroyo = obtenerCartasPorteRecibidas(fechaContable, DBSITE_ARROYO)
		Set rsCartasPorteTransito = obtenerCartasPorteRecibidas(fechaContable, DBSITE_TRANSITO)
		Set rsCartasPortePiedraBuena = obtenerCartasPorteRecibidas(fechaContable, DBSITE_BAHIA)	
		Set rsCartasPortePiedraBuenaVagones = obtenerCartasPorteRecibidasVagones(fechaContable, DBSITE_BAHIA)
	
		Call procesarCartasPorte(rsCartasPorteArroyo, fecha, DBSITE_ARROYO,listaErroresArroyo, myFile)
		Call procesarCartasPorte(rsCartasPorteTransito, fecha, DBSITE_TRANSITO,listaErroresTransito, myFile)
		Call procesarCartasPorte(rsCartasPortePiedraBuena, fecha, DBSITE_BAHIA,listaErroresPiedraBuena, myFile)
		Call procesarCartasPorte(rsCartasPortePiedraBuenaVagones, fecha, DBSITE_BAHIA,listaErroresPiedraBuena, myFile)
		
		myFile.Close()		
		Set myFile = Nothing
		Set fs = Nothing		
		
		Call enviarMailsCartasPorte(fecha, myFilename, listaErroresArroyo, listaErroresTransito,listaErroresPiedraBuena)
		'solo actualiza la fecha de ultima ejecucion en el archivo auxiliar si:
		'1: la ejecucion es programada y la fecha no se ha pasado por parametro, sino se la ha tomado del archivo auxiliar
		'2. si no hubo errores en la verificacion de datos.
		'de este modo la proxima ejecucion programada se va a ejecutar sobre la misma fecha, hasta que no haya errores.
		if flagActualizarFechaUltimaEjecucion then 
			fecha = GF_DTEADD(fecha, 1, "D")
			logMig.info("Fecha actualizada " & fecha)
			updateLastExecDate(fecha)
		end if

	logMig.info("--------------------------- FIN PROCESO ---------------------------")
else
	logMig.info("No hay mas Cartas de Porte para procesar. La fecha de ultima ejecucion exitosa es igual a la actual")
end if	
%>