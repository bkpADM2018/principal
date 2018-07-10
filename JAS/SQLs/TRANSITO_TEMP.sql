Select * from productos
Select * from parametros where CDPARAMETRO in ('DTULTFACTACONDP', 'DTULTFACTACONDE')
Select * from parametros where CDPARAMETRO like '%ZARANDA%'

Update parametros set VLPARAMETRO='20140730' where CDPARAMETRO ='DTULTFACTACONDP'
Update parametros set VLPARAMETRO='20140715' where CDPARAMETRO ='DTULTFACTACONDE'

Select * from ESTADOSASIGNA where CDPRODUCTO=19 and CDESTADOINGRESO=15 and CDESTADOSALIDA=2
Select * from ACEPTACIONCALIDAD
Select * from TRANSACCIONES

Select * from PRODUCTOS
------------------------------------------------------------------------------------------------------------------------------------------

Select * from CAMIONES where IDCAMION='0000004713'
Select * from CALADADECAMIONES where IDCAMION='0000004713'
Update CALADADECAMIONES set CDACEPTACION=15 where IDCAMION='0000004713' and SQCALADA=4
Select * from AUDCAMIONES where IDCAMION='0000004713'
Delete from AUDCAMIONES where IDCAMION='0000004713' and SQAUDITORIA=5
Update CAMIONES set CDESTADO=15 where IDCAMION='0000004713'
Select * from estados
------------------------------------------------------------------------------------------------------------------------------------------
-- nueva ACEPTACION DE CALIDAD
insert into ACEPTACIONCALIDAD values(15, 'AUTORIZA ENTREGADOR', 1, 'AE')

-- 1° ESTADO INGRESO = 11 DEMORADO
insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 11, 'AUTORIZA ENTREGADOR',  '',  CDESTADOAUXILIAR3,  CDTRANSACCION, 15 
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 11  and CDESTADOAUXILIAR1='AUTORIZA ENTREGADOR' and CDESTADOSALIDA=15 and CDTRANSACCION = 2)
                                                
-- 2° ESTADO INGRESO = 15 Calado Pendiente de Autorizacion Entregador	

insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 15, 'AUTORIZA ENTREGADOR',  '',  CDESTADOAUXILIAR3,  CDTRANSACCION, 15 
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 15  and CDESTADOAUXILIAR1='AUTORIZA ENTREGADOR' and CDTRANSACCION = 2 and CDESTADOSALIDA = 15)
                                                
-- 3° ESTADO INGRESO = 15 Calado Pendiente de Autorizacion Entregador	/ REBAJA CONVENIDA / Entregador

insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 15,  'REBAJA CONVENIDA', 'Entregador', CDESTADOAUXILIAR3,  2, 2
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 15  and CDESTADOAUXILIAR1='REBAJA CONVENIDA' and CDESTADOAUXILIAR2='Entregador' and CDTRANSACCION = 2 and CDESTADOSALIDA = 2)
                                                
-- -- 4° ESTADO INGRESO = 15 Calado Pendiente de Autorizacion Entregador	/ CONDICION CAMARA / Entregador

insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 15,  'CONDICION CAMARA', '', CDESTADOAUXILIAR3,  2, 2
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 15  and CDESTADOAUXILIAR1='CONDICION CAMARA' and CDTRANSACCION = 2 and CDESTADOSALIDA = 2)
                                                

-- -- 5° ESTADO INGRESO = 15 Calado Pendiente de Autorizacion Entregador	/ ESTADO SALIDA = 7 Rechazo

insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 15,  'RECHAZO', '', CDESTADOAUXILIAR3,  2, 7
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 15  and CDESTADOAUXILIAR1='Rechazo' and CDTRANSACCION = 2 and CDESTADOSALIDA = 7)
                                                
-- -- 6° ESTADO INGRESO = 15 Calado Pendiente de Autorizacion Entregador	/  ESTADO SALIDA = 11 DEMORADO
Delete from ESTADOSASIGNA where CDESTADOINGRESO=15 and CDESTADOSALIDA=11
--ELIMINADO!!!!
insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 15,  'ANALISIS', '', CDESTADOAUXILIAR3,  2, 11
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA 
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 15  and CDESTADOAUXILIAR1='' and CDTRANSACCION = 2 and CDESTADOSALIDA = 11)

-- -- 7° ESTADO INGRESO = 1 Ingresado	/  ESTADO SALIDA = 15 Calado Pendiente de Autorizacion Entregador

insert into ESTADOSASIGNA 
select distinct CDCIRCUITO, CDPRODUCTO, 1, 'AUTORIZA ENTREGADOR',  '',  CDESTADOAUXILIAR3,  CDTRANSACCION, 15 
from ESTADOSASIGNA 
where CDCIRCUITO='DESCARGA'  and CDTRANSACCION = 2
and CDPRODUCTO not in(select CDPRODUCTO from ESTADOSASIGNA Select * from TOEPFERDB.TBLMENSAJES where CDMENSAJE = '0216'
Insert into TOEPFERDB.TBLMENSAJES VALUES('0216', 'Se ha seleccionado como firmante un director, esto no esta permitido para este PIC')
                                                where CDCIRCUITO='DESCARGA' and CDESTADOINGRESO = 1  and CDESTADOAUXILIAR1='AUTORIZA ENTREGADOR' and CDTRANSACCION = 2 and CDESTADOSALIDA = 15)
                                                                                                
-- el atributo ICINFORMEINTERNO se carga con el valor 0 (NO IMPRIME, SE DEBERA CONFIGURAR POR PRODUCTO ) -- Confirmar
insert into AtributosDeProducto(CDPRODUCTO, CDACEPTACION, ICSTICKER, ICSUPERVISOR, ICCAMARA, ICMOTIVORECHAZO, ICGRADO, ICMERMA, ICRUBRO, ICENSAYO, ICBALDE, ICACON, ICINFORMEINTERNO)
select CDPRODUCTO,15, ICSTICKER, ICSUPERVISOR, ICCAMARA, 0, ICGRADO, ICMERMA, ICRUBRO, ICENSAYO, ICBALDE, ICACON, 0
from AtributosDeProducto
where  CDACEPTACION in(4) 
and CDPRODUCTO not in(select CDPRODUCTO from ATRIBUTOSDEPRODUCTO where CDACEPTACION =15)


