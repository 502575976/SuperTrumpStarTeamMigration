CREATE OR REPLACE PROCEDURE UpdateStreamTable IS
CURSOR bd_cur IS
select * from TBL_STREAM_DETAIL_EXTRACT ;
--where account_schedule in  ('7709230001','7710545001','7709595002','7708484001');
/******************************************************************************
   NAME:       UpdateStreamTable
   PURPOSE:    

   REVISIONS:
   Ver        Date        Author           Description
   ---------  ----------  ---------------  ------------------------------------
   1.0        3/19/2011          1. Created this procedure.

   NOTES:

   Automatically available Auto Replace Keywords:
      Object Name:     UpdateStreamTable
      Sysdate:         3/19/2011
      Date and Time:   3/19/2011, 8:20:52 PM, and 3/19/2011 8:20:52 PM
      Username:         (set in TOAD Options, Procedure Editor)
      Table Name:       (set in the "New PL/SQL Object" dialog)
varchar2(100)
******************************************************************************/
 AcNumber VARCHAR2(100 BYTE);
 FirstRecCheck VARCHAR2(100 BYTE);
 Totalcount integer;
 CurrYr_Jan VARCHAR2(100 BYTE);
 CurrYr_Feb VARCHAR2(100 BYTE);	
 CurrYr_Mar VARCHAR2(100 BYTE);	
 CurrYr_Apr VARCHAR2(100 BYTE);	
 CurrYr_May VARCHAR2(100 BYTE);	
 CurrYr_Jun VARCHAR2(100 BYTE);	
 CurrYr_Jul VARCHAR2(100 BYTE);	
 CurrYr_Aug VARCHAR2(100 BYTE);	
 CurrYr_Sep VARCHAR2(100 BYTE);	
 CurrYr_Oct VARCHAR2(100 BYTE);	
 CurrYr_Nov VARCHAR2(100 BYTE);	
 CurrYr_Dec VARCHAR2(100 BYTE);	
 Yr1_Jan VARCHAR2(100 BYTE);	
 Yr1_Feb VARCHAR2(100 BYTE);	
 Yr1_Mar VARCHAR2(100 BYTE);	
 Yr1_Apr VARCHAR2(100 BYTE);	
 Yr1_May VARCHAR2(100 BYTE);	
 Yr1_Jun VARCHAR2(100 BYTE);	
 Yr1_Jul VARCHAR2(100 BYTE);	
 Yr1_Aug VARCHAR2(100 BYTE);	
 Yr1_Sep VARCHAR2(100 BYTE);	
 Yr1_Oct VARCHAR2(100 BYTE);	
 Yr1_Nov VARCHAR2(100 BYTE);	
 Yr1_Dec VARCHAR2(100 BYTE);	
 Yr2_Jan VARCHAR2(100 BYTE);	
 Yr2_Feb VARCHAR2(100 BYTE);	
 Yr2_Mar VARCHAR2(100 BYTE);	
 Yr2_Apr VARCHAR2(100 BYTE);	
 Yr2_May VARCHAR2(100 BYTE);	
 Yr2_Jun VARCHAR2(100 BYTE);	
 Yr2_Jul VARCHAR2(100 BYTE);	
 Yr2_Aug VARCHAR2(100 BYTE);	
 Yr2_Sep VARCHAR2(100 BYTE);	
 Yr2_Oct VARCHAR2(100 BYTE);	
 Yr2_Nov VARCHAR2(100 BYTE);	
 Yr2_Dec VARCHAR2(100 BYTE);	
 Yr3_Jan VARCHAR2(100 BYTE);	
 Yr3_Feb VARCHAR2(100 BYTE);	
 Yr3_Mar VARCHAR2(100 BYTE);	
 Yr3_Apr VARCHAR2(100 BYTE);	
 Yr3_May VARCHAR2(100 BYTE);	
 Yr3_Jun VARCHAR2(100 BYTE);	
 Yr3_Jul VARCHAR2(100 BYTE);	
 Yr3_Aug VARCHAR2(100 BYTE);	
 Yr3_Sep VARCHAR2(100 BYTE);	
 Yr3_Oct VARCHAR2(100 BYTE);	
 Yr3_Nov VARCHAR2(100 BYTE);	
 Yr3_Dec VARCHAR2(100 BYTE);	
 Yr4_Jan VARCHAR2(100 BYTE);	
 Yr4_Feb VARCHAR2(100 BYTE);	
 Yr4_Mar VARCHAR2(100 BYTE);	
 Yr4_Apr VARCHAR2(100 BYTE);	
 Yr4_May VARCHAR2(100 BYTE);	
 Yr4_Jun VARCHAR2(100 BYTE);	
 Yr4_Jul VARCHAR2(100 BYTE);	
 Yr4_Aug VARCHAR2(100 BYTE);	
 Yr4_Sep VARCHAR2(100 BYTE);	
 Yr4_Oct VARCHAR2(100 BYTE);	
 Yr4_Nov VARCHAR2(100 BYTE);	
 Yr4_Dec VARCHAR2(100 BYTE);	
 Yr5_Jan VARCHAR2(100 BYTE);	
 Yr5_Feb VARCHAR2(100 BYTE);	
 Yr5_Mar VARCHAR2(100 BYTE);	
 Yr5_Apr VARCHAR2(100 BYTE);	
 Yr5_May VARCHAR2(100 BYTE);	
 Yr5_Jun VARCHAR2(100 BYTE);	
 Yr5_Jul VARCHAR2(100 BYTE);	
 Yr5_Aug VARCHAR2(100 BYTE);	
 Yr5_Sep VARCHAR2(100 BYTE);	
 Yr5_Oct VARCHAR2(100 BYTE);	
 Yr5_Nov VARCHAR2(100 BYTE);	
 Yr5_Dec VARCHAR2(100 BYTE);	
 Yr6_Jan VARCHAR2(100 BYTE);	
 Yr6_Feb VARCHAR2(100 BYTE);	
 Yr6_Mar VARCHAR2(100 BYTE);	
 Yr6_Apr VARCHAR2(100 BYTE);	
 Yr6_May VARCHAR2(100 BYTE);	
 Yr6_Jun VARCHAR2(100 BYTE);	
 Yr6_Jul VARCHAR2(100 BYTE);	
 Yr6_Aug VARCHAR2(100 BYTE);	
 Yr6_Sep VARCHAR2(100 BYTE);	
 Yr6_Oct VARCHAR2(100 BYTE);	
 Yr6_Nov VARCHAR2(100 BYTE);	
 Yr6_Dec VARCHAR2(100 BYTE);	
 DATE_ORIG_BOOKED VARCHAR2(100 BYTE);
 Prev VARCHAR2(100 BYTE);
 Last VARCHAR2(100 BYTE);
TYPE aat_binary IS TABLE OF VARCHAR2(100) INDEX BY PLS_INTEGER;
aa_binary aat_binary;
 	
BEGIN
 --DBMS_OUTPUT.PUT_LINE('Start');
  OPEN bd_cur;    
  	LOOP
      FETCH bd_cur INTO AcNumber,CurrYr_Jan ,CurrYr_Feb ,CurrYr_Mar ,CurrYr_Apr ,CurrYr_May ,CurrYr_Jun ,CurrYr_Jul ,CurrYr_Aug ,CurrYr_Sep ,CurrYr_Oct ,CurrYr_Nov ,CurrYr_Dec ,Yr1_Jan ,Yr1_Feb ,Yr1_Mar ,Yr1_Apr ,Yr1_May ,Yr1_Jun ,Yr1_Jul ,Yr1_Aug ,Yr1_Sep ,Yr1_Oct ,Yr1_Nov ,Yr1_Dec ,Yr2_Jan ,Yr2_Feb ,Yr2_Mar ,Yr2_Apr ,Yr2_May ,Yr2_Jun ,Yr2_Jul ,Yr2_Aug ,Yr2_Sep ,Yr2_Oct ,Yr2_Nov ,Yr2_Dec ,Yr3_Jan ,Yr3_Feb ,Yr3_Mar ,Yr3_Apr ,Yr3_May ,Yr3_Jun ,Yr3_Jul ,Yr3_Aug ,Yr3_Sep ,Yr3_Oct ,Yr3_Nov ,Yr3_Dec ,Yr4_Jan ,Yr4_Feb ,Yr4_Mar ,Yr4_Apr ,Yr4_May ,Yr4_Jun ,Yr4_Jul ,Yr4_Aug ,Yr4_Sep ,Yr4_Oct ,Yr4_Nov ,Yr4_Dec ,Yr5_Jan ,Yr5_Feb ,Yr5_Mar ,Yr5_Apr ,Yr5_May ,Yr5_Jun ,Yr5_Jul ,Yr5_Aug ,Yr5_Sep ,Yr5_Oct ,Yr5_Nov ,Yr5_Dec ,Yr6_Jan ,Yr6_Feb ,Yr6_Mar ,Yr6_Apr ,Yr6_May ,Yr6_Jun ,Yr6_Jul ,Yr6_Aug ,Yr6_Sep ,Yr6_Oct ,Yr6_Nov ,Yr6_Dec ,DATE_ORIG_BOOKED; 
      EXIT WHEN  bd_cur%NOTFOUND;
aa_binary(1):=CurrYr_Jan ;
 aa_binary(2):=CurrYr_Feb ;	
 aa_binary(3):=CurrYr_Mar ;	
 aa_binary(4):=CurrYr_Apr ;	
 aa_binary(5):=CurrYr_May ;	
 aa_binary(6):=CurrYr_Jun ;	
 aa_binary(7):=CurrYr_Jul ;	
 aa_binary(8):=CurrYr_Aug ;	
 aa_binary(9):=CurrYr_Sep ;	
 aa_binary(10):=CurrYr_Oct ;
  aa_binary(11):=CurrYr_Nov ;	
 aa_binary(12):=CurrYr_Dec ;	
 aa_binary(13):=Yr1_Jan ;	
 aa_binary(14):=Yr1_Feb ;	
 aa_binary(15):=Yr1_Mar ;	
 aa_binary(16):=Yr1_Apr ;	
 aa_binary(17):=Yr1_May ;	
 aa_binary(18):=Yr1_Jun ;	
 aa_binary(19):=Yr1_Jul ;	
 aa_binary(20):=Yr1_Aug ;	
 aa_binary(21):=Yr1_Sep ;	
 aa_binary(22):=Yr1_Oct ;	
 aa_binary(23):=Yr1_Nov ;	
 aa_binary(24):=Yr1_Dec ;	
 aa_binary(25):=Yr2_Jan ;	
 aa_binary(26):=Yr2_Feb ;	
 aa_binary(27):=Yr2_Mar ;	
 aa_binary(28):=Yr2_Apr ;	
 aa_binary(29):=Yr2_May ;	
 aa_binary(30):=Yr2_Jun ;	
 aa_binary(31):=Yr2_Jul ;	
 aa_binary(32):=Yr2_Aug ;	
 aa_binary(33):=Yr2_Sep ;	
 aa_binary(34):=Yr2_Oct ;	
 aa_binary(35):=Yr2_Nov ;	
 aa_binary(36):=Yr2_Dec ;	
 aa_binary(37):=Yr3_Jan ;	
 aa_binary(38):=Yr3_Feb ;	
 aa_binary(39):=Yr3_Mar ;	
 aa_binary(40):=Yr3_Apr ;	
 aa_binary(41):=Yr3_May ;	
 aa_binary(42):=Yr3_Jun ;	
 aa_binary(43):=Yr3_Jul ;	
 aa_binary(44):=Yr3_Aug ;	
 aa_binary(45):=Yr3_Sep ;	
 aa_binary(46):=Yr3_Oct ;	
 aa_binary(47):=Yr3_Nov ;	
 aa_binary(48):=Yr3_Dec ;	
 aa_binary(49):=Yr4_Jan ;	
 aa_binary(50):=Yr4_Feb ;	
 aa_binary(51):=Yr4_Mar ;	
 aa_binary(52):=Yr4_Apr ;	
 aa_binary(53):=Yr4_May ;	
 aa_binary(54):=Yr4_Jun ;	
 aa_binary(55):=Yr4_Jul ;	
 aa_binary(56):=Yr4_Aug ;	
 aa_binary(57):=Yr4_Sep ;	
 aa_binary(58):=Yr4_Oct ;	
 aa_binary(59):=Yr4_Nov ;	
 aa_binary(60):=Yr4_Dec ;	
 aa_binary(61):=Yr5_Jan ;	
 aa_binary(62):=Yr5_Feb ;	
 aa_binary(63):=Yr5_Mar ;	
 aa_binary(64):=Yr5_Apr ;	
 aa_binary(65):=Yr5_May ;	
 aa_binary(66):=Yr5_Jun ;	
 aa_binary(67):=Yr5_Jul ;	
 aa_binary(68):=Yr5_Aug ;	
 aa_binary(69):=Yr5_Sep ;	
 aa_binary(70):=Yr5_Oct ;	
 aa_binary(71):=Yr5_Nov ;	
 aa_binary(72):=Yr5_Dec ;	
 aa_binary(73):=Yr6_Jan ;	
 aa_binary(74):=Yr6_Feb ;	
 aa_binary(75):=Yr6_Mar ;	
 aa_binary(76):=Yr6_Apr ;	
 aa_binary(77):=Yr6_May ;	
 aa_binary(78):=Yr6_Jun ;	
 aa_binary(79):=Yr6_Jul ;	
 aa_binary(80):=Yr6_Aug ;	
 aa_binary(81):=Yr6_Sep ;	
 aa_binary(82):=Yr6_Oct ;	
 aa_binary(83):=Yr6_Nov ;	
 aa_binary(84):=Yr6_Dec ;	
	  
			 Prev:=CurrYr_Jan;
			 Last:=CurrYr_Jan;
			 Totalcount:=0;
			 FirstRecCheck:=0;
			 FOR i IN 1 .. 84 LOOP
			  if Prev <> aa_binary(i)  then	
			  	 	  if FirstRecCheck=0 and Prev='0.00' then
					  	 FirstRecCheck:=1;
					  else
					  	  insert into TBL_STREAM_DETAIL  values(AcNumber,'',Prev,Totalcount);	
					  end if;
				       	  		 	  
				 	  Prev:=aa_binary(i);
				 	  Totalcount:=1;					  					 
					  --DBMS_OUTPUT.PUT_LINE(Prev || ' Inserted=' || Totalcount);		 	 
				 else				 		  
				 	 	  Totalcount := Totalcount + 1;						 
						  --DBMS_OUTPUT.PUT_LINE(aa_binary(i)|| ' Total=' || Totalcount); 			 
				 end if;
			 	 
		     END LOOP;  	
			 	 if Prev<>'0.00' then
				 	insert into TBL_STREAM_DETAIL  values(AcNumber,'',Prev,Totalcount);					 
				 end if;		
			 	 		
			      
	END LOOP; 	 	    		  
END UpdateStreamTable;
/
