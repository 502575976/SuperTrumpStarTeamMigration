CREATE OR REPLACE PROCEDURE myTest (employee_id NUMBER) AS
   tot_emps NUMBER;
   BEGIN
      DELETE FROM tbl_residual_mapping
      WHERE term=1;
   tot_emps := tot_emps - 1;
   END;
/
