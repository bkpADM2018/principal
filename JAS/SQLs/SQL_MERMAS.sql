SELECT * FROM hrubrosvisteocamiones WHERE  dtcontable = '2014-05-20' and IDCAMION='0000000600'
--Update hrubrosvisteocamiones set VLMERMA=1.98 where DTCONTABLE='2014-05-20' and IDCAMION='0000004519' and SQCALADA=5 and CDRUBRO=21; --antes 0
--Update hrubrosvisteocamiones set VLMERMA=1.3 where DTCONTABLE='2014-05-20' and IDCAMION='0000004519' and SQCALADA=5 and CDRUBRO=28; --antes 0
--Select * from Rubros
/*
Select A.DTCONTABLE FECHA, A.IDCAMION, NUCARTAPORTE CARTAPORTE, MERMATOTAL,
ROUND(( (SELECT PC.vlpesada 
			   FROM   db2admin.hpesadascamion PC 
			   WHERE  PC.dtcontable = A.dtcontable 
			          AND PC.idcamion = A.idcamion 
			          AND PC.cdpesada = 1 
			          AND PC.sqpesada = (SELECT Max(sqpesada) 
			                             FROM   db2admin.hpesadascamion 
			                             WHERE  PC.dtcontable = dtcontable 
			                                    AND PC.idcamion = idcamion 
			                                    AND cdpesada = 1)) - 
			    (SELECT PC.vlpesada 
				    FROM   db2admin.hpesadascamion PC 
			     WHERE  PC.dtcontable = A.dtcontable 
			            AND PC.idcamion = A.idcamion 
			            AND PC.cdpesada = 2 
			            AND PC.sqpesada = (SELECT Max(sqpesada) 
			                               FROM   db2admin.hpesadascamion 
			                               WHERE  PC.dtcontable = dtcontable 
			                                      AND PC.idcamion = idcamion 
			                                      AND cdpesada = 2)) ) * VISTEOCABECERA/100, 0)
			MERMACALCULADA, VISTEOCABECERA, VISTEODETALLE
			
from
(Select * from HCAMIONES where CDESTADO in (6, 8)) A 
inner join
HCAMIONESDESCARGA A1 on A.DTCONTABLE=A1.DTCONTABLE and A.IDCAMION=A1.IDCAMION
inner join 
(Select DTCONTABLE, IDCAMION, SUM(VLMERMAKILOS) MERMATOTAL from HMERMASCAMIONES X where SQPESADA = (Select MAX(SQPESADA) from HMERMASCAMIONES where DTCONTABLE=X.DTCONTABLE and IDCAMION=X.IDCAMION) and DTCONTABLE>='2014-04-01' group by DTCONTABLE, IDCAMION) MC on MC.DTCONTABLE=A.DTCONTABLE and MC.IDCAMION=A.IDCAMION
inner join
(Select DTCONTABLE, IDCAMION, SUM(VLMERMA) VISTEODETALLE from HRUBROSVISTEOCAMIONES A where SQCALADA = (Select MAX(SQCALADA) from HRUBROSVISTEOCAMIONES where DTCONTABLE=A.DTCONTABLE and IDCAMION=A.IDCAMION) and DTCONTABLE>='2014-04-01' group by DTCONTABLE, IDCAMION) B 
on A.DTCONTABLE=B.DTCONTABLE and A.IDCAMION=B.IDCAMION
inner join 
(Select DTCONTABLE, IDCAMION, SUM(PCMERMA) VISTEOCABECERA from HCALADADECAMIONES A where SQCALADA = (Select MAX(SQCALADA) from HCALADADECAMIONES where DTCONTABLE=A.DTCONTABLE and IDCAMION=A.IDCAMION ) and DTCONTABLE>='2014-04-01' group by DTCONTABLE, IDCAMION) C on B.DTCONTABLE=C.DTCONTABLE and B.IDCAMION=C.IDCAMION
where (MERMATOTAL=0 and (VISTEOCABECERA<>0 or VISTEODETALLE<>0)) or 
(MERMATOTAL>0 and (VISTEOCABECERA=0 or VISTEODETALLE=0)) or
VISTEODETALLE<>VISTEOCABECERA




