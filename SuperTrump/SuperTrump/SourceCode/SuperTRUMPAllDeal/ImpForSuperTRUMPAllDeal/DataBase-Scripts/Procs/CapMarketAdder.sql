CREATE OR REPLACE PROCEDURE VIC_ROI.CapMarketAdder (
AccountScheduleNumber  IN      VARCHAR2,
TermValue       IN      NUMBER,
Product          IN      VARCHAR2,
ProgramName     IN      VARCHAR2,
CapADDER     OUT        NUMBER
) 
AS 
TotalCost     NUMBER;
ProductType     VARCHAR2(100);
  
BEGIN
 
   SELECT case when SUM(OEC_ON_ASSET) is null then 0 else SUM(OEC_ON_ASSET) end INTO TotalCost FROM TBL_ASSET_DETAIL WHERE ACCOUNT_SCHEDULE_NBR = AccountScheduleNumber;

  SELECT PType INTO ProductType FROM TBL_PRODUCT_MAPPING WHERE PRODUCT = Product and rownum=1;

  SELECT distinct CAP_MARKET_ADDER INTO CapADDER FROM TBL_RESIDUAL_MAPPING WHERE Program = 'Construction' AND  Product = ProductType AND 
  ((Term_Operator='>' AND TermValue >= TERM) OR (Term_Operator='<' AND TermValue < TERM))
  AND
  ((Deal_size_operator='>' AND TotalCost >= Deal_Size) OR (Deal_size_operator='<' AND TotalCost < Deal_Size));  
    	
END;
/


