--  Generar SQL 
--  Versi�n:                   	V7R1M0 100423 
--  Generado en:              	19/09/13 10:46:30 
--  Base de datos relacional:       	S10D441R 
--  Opci�n de est�ndares:          	DB2 for i 
SET PATH "QSYS","QSYS2","SYSPROC","SYSIBMADM","SYNC" ; 
  
CREATE OR REPLACE PROCEDURE MERFL.MER5F1F1_GET_BY_FILTERS ( 
	IN @FDESDE VARCHAR(8) , 
	IN @FHASTA VARCHAR(8) , 
	IN @CPRO VARCHAR(2) , 
	IN @CSUC VARCHAR(1) , 
	IN @COPE VARCHAR(2) , 
	IN @NCTO VARCHAR(5) , 
	IN @ACOS VARCHAR(2) , 
	IN @CTOPROV VARCHAR(15) , 
	IN @PTO VARCHAR(2) , 
	IN @NCPOR VARCHAR(12) , 
	IN @PFORM VARCHAR(4) , 
	IN @NFORM VARCHAR(8) , 
	IN @CCOR VARCHAR(5) , 
	IN @CVEN VARCHAR(5) , 
	IN @CPROV VARCHAR(5) , 
	IN @PAGE INTEGER , 
	IN @LPP INTEGER , 
	OUT @SPTOTALRCD BIGINT, 
	OUT @SPERRORCODE INTEGER , 
	OUT @SPERRORDESC VARCHAR(5000) ) 
	DYNAMIC RESULT SETS 1 
	LANGUAGE SQL 
	SPECIFIC MERFL.MER5F1F1_GET_BY_FILTERS 
	NOT DETERMINISTIC 
	MODIFIES SQL DATA 
	CALLED ON NULL INPUT 
	SET OPTION  ALWBLK = *ALLREAD , 
	ALWCPYDTA = *OPTIMIZE , 
	COMMIT = *NONE , 
	DECRESULT = (31, 31, 00) , 
	DFTRDBCOL = *NONE , 
	DYNDFTCOL = *NO , 
	DYNUSRPRF = *USER , 
	SRTSEQ = *HEX   
	BEGIN 
-- Code an IF statement 
	DECLARE SQLTEXT VARCHAR ( 5000 ) ; 
	DECLARE SQLPGN VARCHAR ( 5000 ) ; 
	DECLARE QRYTEXT VARCHAR ( 5000 ) ; 
	DECLARE RINI BIGINT ; 
	DECLARE RFIN BIGINT ; 
	DECLARE CURSOR1 CURSOR FOR DYNSTATEMENT ; 
		  
	SET QRYTEXT = 'SELECT ROWNUMBER() OVER(ORDER BY F16A.FECDFA DESC) RN, F16A.*, CTO.*, COR.RAZSOC DSCOR, VEN.RAZSOC DSVEN, P.DESCPC DSPROCEDENCIA FROM MERFL.MER5F1F1 F16A' ; 
	SET QRYTEXT = QRYTEXT || ' INNER JOIN MERFL.MER311F1 CTO ON CPROR1= CPROFA AND CSUCR1= CSUCFA AND COPER1= COPEFA AND NCTOR1= NCTOFA AND ACOSR1= ACOSFA' ; 
	SET QRYTEXT = QRYTEXT || ' INNER JOIN MERFL.MER142F1 P ON F16A.CPRDfA=P.CODIPC AND F16A.AUXIfA=P.AUXIPC';
	SET QRYTEXT = QRYTEXT || ' LEFT JOIN MERFL.TCB6A1F1 COR ON COR.NROPRO=CTO.CCORR1' ; 
	SET QRYTEXT = QRYTEXT || ' LEFT JOIN MERFL.TCB6A1F1 VEN ON VEN.NROPRO=CTO.CVENR1' ; 
	SET QRYTEXT = QRYTEXT || ' WHERE F16A.FECDFA > 20130101' ; 
	 
	IF @FDESDE <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.FECDFA >=' || @FDESDE ; 
	END IF ; 
	IF @FHASTA <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.FECDFA <=' || @FHASTA ; 
	END IF ; 
	IF @CPRO <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.CPROFA =' || @CPRO ; 
	END IF ; 
	IF @CSUC <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.CSUCFA =' || @CSUC ; 
	END IF ; 
	IF @COPE <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.COPEFA =' || @COPE ; 
	END IF ; 
	IF @NCTO <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.NCTOFA =' || @NCTO ; 
	END IF ; 
	IF @ACOS <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.ACOSFA =' || @ACOS ; 
	END IF ; 
	IF @CTOPROV <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND  CTO.CONCR1 LIKE ''%' || @CTOPROV || '%''' ; 
	END IF ; 
	IF @PTO <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND  CTO.CDESR1=' || @PTO ; 
	END IF ; 
	IF @NCPOR <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.CPORFA LIKE ''%' || @NCPOR || '%''' ; 
	END IF ; 
	IF @PFORM <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.PPIAFA = ' || @PFORM ; 
	END IF ; 
	IF @NFORM <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND F16A.NPIAFA  LIKE ''%' || @NFORM || '%''' ; 
	END IF ; 
	IF @CVEN <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND CTO.CVENR1 = ' || @CVEN ; 
	END IF ; 
	IF @CCOR <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND CTO.CCORR1 = ' || @CCOR ; 
	END IF ;	 
	IF @CPROV <> '' THEN 
		SET QRYTEXT = QRYTEXT || ' AND (CTO.CCORR1 = ' || @CPROV || ' OR CTO.CVENR1 = ' || @CPROV  || ')';
	END IF ;	 
  
--Se arma la consulta para traer el total de registros
	SET SQLPGN = 'SELECT COUNT(*) TOTALREGISTROS FROM (' || QRYTEXT || ') Q'  ; 
	
	PREPARE DYNSTATEMENT FROM SQLPGN ; 
	OPEN CURSOR1 ; 
	FETCH CURSOR1 INTO @SPTOTALRCD;
	CLOSE CURSOR1;

 --Se calculan los valores para la paginacion.
	IF @LPP = 0 THEN --@LPP en cero implica todos los registros!
		SET @LPP = @SPTOTALRCD;
		SET @PAGE = 1;
	END IF;
	SET RFIN = @PAGE * @LPP ; 
	IF RFIN > @SPTOTALRCD THEN
		SET RFIN = @SPTOTALRCD;
	END IF;
	SET RINI = RFIN - @LPP + 1 ; 
	
--Se arma la consulta para leer los datos 
	SET SQLTEXT = 'SELECT * FROM (' || QRYTEXT || ' ORDER BY F16A.FECDFA DESC)  Q WHERE RN BETWEEN ' || RINI || ' AND ' || RFIN ;  
  
	SET @SPERRORCODE = 0 ; 
	SET @SPERRORDESC =QRYTEXT;  --'OK' ; 
  
	PREPARE DYNSTATEMENT FROM SQLTEXT ; 
	OPEN CURSOR1 ; 
  
RETURN ; 
END  ; 
  
GRANT ALTER , EXECUTE   
ON SPECIFIC PROCEDURE MERFL.MER5F1F1_GET_BY_FILTERS 
TO SYNC ;
