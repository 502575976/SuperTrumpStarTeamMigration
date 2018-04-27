CREATE OR REPLACE Function FindAdder
   ( name_in IN varchar2 )
   RETURN number
IS
    cnumber number;
    cursor c1 is
    select Term
      from TBL_Residual_mapping
      where Term = name_in;
BEGIN
open c1;
fetch c1 into cnumber;
if c1%notfound then
     cnumber := 9999;
end if;
close c1;
RETURN cnumber;
EXCEPTION
WHEN OTHERS THEN
      raise_application_error(-20001,'An error was encountered - '||SQLCODE||' -ERROR- '||SQLERRM);
END;
/
