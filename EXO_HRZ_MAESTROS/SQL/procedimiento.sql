CREATE PROCEDURE _98_EXO_ACTUALIZATIPOSFAM
(

)
LANGUAGE SQLSCRIPT
AS
-- Return values
vItemCode NVARCHAR(50);
vCampo nvarchar(50);
vPROPIEDAD NVARCHAR(50);
 i INTEGER;	
BEGIN

DECLARE CURSOR c_ART (x nvarchar(2)) FOR
	SELECT T0."ItmsGrpCod" ItmsGrpCod, T0."ItmsGrpNam",T1."U_EXO_PROPIEDAD" U_EXO_PROPIEDAD ,T2."ItemCode" ItemCode
	FROM OITB T0 
	INNER JOIN "@EXO_TIPOFAM"  T1 ON T0."U_EXO_TIPFAM" = T1."Code"
	INNER JOIN "OITM" T2 ON T0."ItmsGrpCod" = T2."ItmsGrpCod" where T1."U_EXO_PROPIEDAD"= :x; --pPropiedad

FOR i IN 3..13 DO
     
  vCampo := 'QryGroup' || CAST(:i as nvarchar(2)); 
	
							
			FOR c_row_ART AS c_ART(CAST(:i as nvarchar(2))) DO
				vPROPIEDAD := c_row_ART.U_EXO_PROPIEDAD;
				vItemCode := c_row_ART.ItemCode;
								
				EXEC 'UPDATE OITM SET "' || :vCampo || '" = ''Y'' WHERE "ItemCode" = ''' || :vItemCode || '''';
			END FOR;
			
	END FOR;
END;