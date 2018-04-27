CREATE OR REPLACE PROCEDURE getDumpData ( p_cursor OUT SYS_REFCURSOR)
  IS
     BEGIN
    OPEN p_cursor FOR
       SELECT *
       FROM DW_DATA_DUMP;
     END getDumpData;
/