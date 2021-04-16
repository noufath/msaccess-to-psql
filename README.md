# MS Access to psql
This script helps me do the database conversion from Ms Access to postgresql. Since this script need Microsoft Access Driver, this scrip only run in OS Windows.

This are some of the references that I use :
1. [pyodbc's doc](https://github.com/mkleehammer/pyodbc/wiki/Cursor)
2. [SQLStatistic function](https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlstatistics-function?redirectedfrom=MSDN&view=sql-server-ver15)
3. [Workbench migration Ms.Access](https://dev.mysql.com/doc/workbench/en/wb-migration-database-access.html)
4. [https://github.com/remoteworkerid/acc2psql](https://github.com/remoteworkerid/acc2psql)

# Install requirement
1. psycopg2 2.8.6
2. psycopg2-binary 2.8.6
3. pyodbc 4.0.30
4. click 7.1.2

# How to run 
1. preparing Msaccess for migration, grant read access to the Admin role for read relationship/foreign key information by executing scripts below in MsAccess Immediate panel. Open Ms Access file, click menu Database tools - Visual Basic - click menu view - Immediate Windows
  ```
  ?CurrentUser
  CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin"
  ``` 
2. Run convert_access_to_psql.py
3. Fill full path MS Access souce file location, ip host PostgreSQL, PosgreSQL Database name, PostgreSQL user, PostgreSQL password. 



