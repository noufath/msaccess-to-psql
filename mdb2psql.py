import os
import pyodbc
import psycopg2
import psycopg2.extras as extras
import sys


""" 
Use references from:
1. pyodbc's doc: https://github.com/mkleehammer/pyodbc/wiki/Cursor
2. SQLStatistic function : https://docs.microsoft.com/en-us/sql/odbc/reference/syntax/sqlstatistics-function?redirectedfrom=MSDN&view=sql-server-ver15
3. Workbench migration Ms.Access : https://dev.mysql.com/doc/workbench/en/wb-migration-database-access.html
4. Ms.Access migration to postgres : https://github.com/remoteworkerid/acc2psql
"""

class mdb2psql:

    def __init__(self, mdb_file, pg_host, pg_db, pg_user, pg_password, print_SQL):
        self.access_cursor = pyodbc.connect(f'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_file};').cursor()
        self.schema_name = self.get_access_dbname()
        self.print_SQL = print_SQL
        self.pg_user = pg_user

        self.param_dic = {
            "host"      : pg_host,
            "database"  : pg_db,
            "user"      : pg_user,
            "password"  : pg_password
        }

        self.pg_conn = self.pg_connect(self.param_dic)
        self.pg_cursor = self.pg_conn.cursor()

    def pg_connect(self, params_dic):
        pg_conn = None
        try:
            print('Connecting to the PostgreSql database...')
            pg_conn = psycopg2.connect(**params_dic)
        except (Exception, psycopg2.DatabaseError) as error:
            print(error)
            sys.exit(1)
        return pg_conn
        
        
    def get_access_dbname(self):
        for table in self.access_cursor.tables():
            return os.path.splitext(os.path.basename(table.table_cat))[0]

    
    def create_schema(self):
        str_SQL = 'DROP SCHEMA IF EXISTS {schema_name} CASCADE; CREATE SCHEMA {schema_name};'.format(schema_name=self.schema_name)
    
        if self.print_SQL:
            print(str_SQL)
        
        self.pg_cursor.execute(str_SQL)
        self.pg_conn.commit()

        self.create_tables()
        

    def create_tables(self):
        table_list = list()
        for table in self.access_cursor.tables(tableType='TABLE'):
            table_list.append(table.table_name)

        psql = ''
        str_table_independent = ''
        str_table_dependent = ''
        table_independent = list()
        table_dependent = list()
        table_order = list()
        for table in table_list:
            str_SQL = ''
            str_SQL = 'DROP TABLE IF EXISTS {schema_name}.{table} CASCADE;\n' \
                .format(schema_name=self.schema_name, table=table)
            str_SQL += 'CREATE TABLE {schema_name}.{table} (\n' \
                .format(schema_name=self.schema_name, table=table)
            str_SQL += self.create_fields(table)

            # collecting str query to create independent tables first
            if 'FOREIGN KEY' not in str_SQL:
                str_table_independent += str_SQL
                table_independent.append(table)
            else:
                str_table_dependent += str_SQL
                table_dependent.append(table)

        psql = str_table_independent + str_table_dependent 
        if self.print_SQL:
            print(psql)

        self.pg_cursor.execute(psql)
        self.pg_conn.commit()
        
        table_order = table_independent + table_dependent
        self.insert_data(table_order)
        
    def create_fields(self, table):
        postgresql_fields = {
            'COUNTER': 'serial PRIMARY KEY', # autoincrement
            'VARCHAR': 'varchar', # text
            'LONGCHAR': 'varchar', # text
            'BYTE': 'smallint',  # byte
            'INTEGER': 'int',  # integer
            'LONG INTEGER': 'bigint',  # long integer
            'REAL': 'real', # single
            'DOUBLE': 'double precision', # double
            'DATETIME': 'timestamp', # date/time
            'CURRENCY': 'money', # currency 
            'BIT': 'boolean', # yes/no
        }

        '''
        get foreignkeys Reference
        
        1. preparing Msaccess for migration, grant read access to the Admin role for read relationship/foreign key information by 
           executing scripts below in MsAccess Immediate panel
                ?CurrentUser
                CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin"
        2. Executing hidden MsAccess system table : SELECT * FROM MSysRelationships
            it will return information : 
            - ccolumn
            - grbit
            - icolumn
            - szColumn
            - szObject # table_name
            - sZReferencedColumn # field_reference
            - sZReferencedObject # table_reference
            - sZRelationship # relationship_name
        '''
        
        foreign_keys= ''
        foreign_keys_exist = {}
        for row in self.access_cursor.execute("select sZObject, sZreferencedColumn, sZReferencedObject, sZRelationship from MSysRelationships where szObject=?", table):
            foreign_key = row[1]

            table_reference = row[2]
            if foreign_keys_exist.get(row[1], None) is None:
                foreign_keys_exist[row[1]] = True
                foreign_keys = f'{foreign_keys} FOREIGN KEY ({foreign_key}) REFERENCES {self.schema_name}.{table_reference} ({foreign_key}) ON DELETE CASCADE,\n'

        foreign_keys = foreign_keys[:-2]

        if foreign_keys != '':
            foreign_keys = f'\n, {foreign_keys}'

        str_SQL =''
        field_list = list()
        for column in self.access_cursor.columns(table=table):
            if column.type_name in postgresql_fields:
                field_list += [column.column_name + " " + postgresql_fields[column.type_name],]
            elif column.type_name == 'DECIMAL':
                field_list += [column.column_name +
                               ' numeric(' + str(column.column_size) + "," +
                               str(column.decimal_digits) + ")", ]
            else:
                print("column " + table + "." + column.column_name +
                " has uncatered for type: " + column.type_name)

        return ','.join(field_list)  + '\n' + foreign_keys + ');\n'

    def get_msaccess_data(self, table):
        str_SQL = 'SELECT * FROM [{table_name}]'.format(table_name=table)
        self.access_cursor.execute(str_SQL)
        rows = self.access_cursor.fetchall()

        data = [tuple(x) for x in rows]
        #for row in rows:
        #    data += [row, ]
        
        return data

    def insert_data(self, table_list):

        for table in table_list:
            data = self.get_msaccess_data(table)

            if data != []:
                format_string = '(' + ','.join(['%s', ]*len(data[0])) + ')\n'
                 # pre-bind the arguments before executing - for speed
                args_string = ','.join(self.pg_cursor.mogrify(format_string, x).decode('utf-8') for x in data)
                column = ','.join(list(self.get_column(table)))
                str_SQL = "INSERT INTO %s(%s) VALUES " % (self.schema_name + '.' + table, column) + args_string

                if self.print_SQL:
                    print('INSERT INTO {schema_name}.{table_name} VALUES {value_list}'.format(schema_name=self.schema_name, table_name=table, value_list=args_string))

               
                try:
                    self.pg_cursor.execute(str_SQL, data)
                    self.pg_conn.commit()
                except (Exception, psycopg2.DatabaseError) as error:
                    print("Error: %s" % error)
                    self.pg_conn.rollback()
                    self.pg_cursor.close()
                    return 1
                print("Execute() done")
        
        self.pg_cursor.close()
                
    
    def get_column(self, table):
        columns = list()
        for column in self.access_cursor.columns(table=table):
            columns += column.column_name, 
   
        return columns

    