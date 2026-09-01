# working on sts.server

# Read distinct server

 read  distinct servername from the excel that we used on `prompt.md`  and implement the rule this curemt prompt.
. findthe distinct servername ,
 if it exist from `sts.server` table and create .sql file

 # steps  below implement on `sts_server_comparison.sql`

  1. step 1-- Find servers from Excel that EXIST in sts.server, using ita.fnSnowServer function on 
  2. step 2 the mapped column that ised on ita.snowserver use that and find the corespondent mapping
   columns in a table sts.server and analysis only the corospondet mapping column display no need to save .py file.
  3. step 3-- Cross-check the user's existing stored procedure mapping (alias t -> tr) against step 2/2b findings. 
  4. step 4 -- create temp tables to store  results of ita.snowserver , sts.server.
  5. step 5 -- the update script that you creted on the compare_servers_consolidated.sql on temp table for ita.snowserver
  6. step 6 -- update script the merge to from temp snow table to  temp sts.server 
