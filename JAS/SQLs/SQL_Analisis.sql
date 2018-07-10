--PROCESADOS
/*
Select count(*) from (
Select distinct fanaca, cPORCA from
MERFL.MER591CA
where FANACA >= 20140401
 and FANACA <= 20140531
) T
*/
/*
Select CDESR6, count(*) from merfl.mer311f6
where FECDR6 >= 20140401
and FECDR6 <= 20140430
group by CDESR6
*/
/*
Select COSTR6, count(*) from 
MERFL.MER311F6
where FECDR6 >= 20100401
 and FECDR6 <= 20140530
group by COSTR6
*/
-- Descargas Aplicadas
/*
Select count(*) from 
(Select CPORR6, MIN(FECDR6) FECHA from MERFL.MER311F6 where CDESR6=36  and COSTR6 in (1, 3, 25) group by CPORR6) T
where FECHA >= 20140401
 and FECHA <= 20140430
*/
/*
Select FECHA, count(*) from 
(Select CPORR6, MIN(FECDR6) FECHA from MERFL.MER311F6 where CDESR6=36  and COSTR6 in (1, 3, 25) group by CPORR6) T
where FECHA >= 20140401
 and FECHA <= 20140430
group by FECHA
*/

Select CPORR6 from 
(Select CPORR6, MIN(FECDR6) FECHA from MERFL.MER311F6 where CDESR6=36  and COSTR6 in (1, 3, 25)  group by CPORR6) T
where FECHA >= 20140412
 and FECHA <= 20140412
ORDER BY CPORR6

--537603657
--537749712
--537770167
/*
Select NUCARTAPORTE from hcamionesdescarga D inner join hcamiones C on D.DTCONTABLE=C.DTCONTABLE and D.IDCAMION=C.IDCAMION
where CDESTADO in (6, 8)
and D.DTCONTABLE >= '2014-04-12'
and D.DTCONTABLE <= '2014-04-12'
order by NUCARTAPORTE
*/
--Select * from merfl.mer311f6 where cporr6=537603657
/*
Select count(*) from Hvagones
where CDESTADO in (6, 8)
and DTCONTABLEVAGON >= '2014-04-01'
and DTCONTABLEVAGON <= '2014-04-30'
*/
/*
Select D.DTCONTABLE, count(*) from hcamionesdescarga D inner join hcamiones C on D.DTCONTABLE=C.DTCONTABLE and D.IDCAMION=C.IDCAMION
where CDESTADO in (6, 8)
and D.DTCONTABLE >= '2014-04-01'
and D.DTCONTABLE <= '2014-04-30'
group by D.DTCONTABLE
*/