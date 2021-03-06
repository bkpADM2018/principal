SELECT NRODOC as CUIT, RAZSOC as RazonSocial, CPROR1 as Especie, CFPAR1 as Operacion, 
CPROR1 as Producto, CSUCR1 as Suursal, COPER1 as Operacion, NCTOR1 as Numero, ACOSR1 AS Cosecha, KGCOR1 as Kilos
FROM MERFL.MER311F1 INNER JOIN MERFL.TCB6A1F1 ON NROPRO=CVENR1
WHERE CPROR1 in (15, 19) and COPER1 in (00, 01, 09, 10) and ACOSR1 = 07 and FCCTR1<20070112
order by CPROR1, CSUCR1, COPER1, NCTOR1, ACOSR1