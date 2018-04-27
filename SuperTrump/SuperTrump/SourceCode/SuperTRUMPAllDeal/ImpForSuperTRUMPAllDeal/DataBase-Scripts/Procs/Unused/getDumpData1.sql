CREATE OR REPLACE PROCEDURE getDumpData1 IS
tmpVar NUMBER;
/******************************************************************************
   NAME:       getDumpData1
   PURPOSE:    

   REVISIONS:
   Ver        Date        Author           Description
   ---------  ----------  ---------------  ------------------------------------
   1.0        9/27/2010          1. Created this procedure.

   NOTES:

   Automatically available Auto Replace Keywords:
      Object Name:     getDumpData1
      Sysdate:         9/27/2010
      Date and Time:   9/27/2010, 2:02:41 PM, and 9/27/2010 2:02:41 PM
      Username:         (set in TOAD Options, Procedure Editor)
      Table Name:       (set in the "New PL/SQL Object" dialog)

******************************************************************************/
BEGIN
   tmpVar := 0;
   
   select * from dw_Data_Dump;
   EXCEPTION
     WHEN NO_DATA_FOUND THEN
       NULL;
     WHEN OTHERS THEN
       -- Consider logging the error and then re-raise
       RAISE;
END getDumpData1;
/
