import mdb2psql
import click
import sys
import os


@click.command()
@click.option('--mdbfile', 
    prompt="Full path of MSAccess Source file", help="Full path of MSAccess source file")
@click.option('--psql_host', 
    prompt="PostgreSQL IP server", help="PostgreSQL ip server")
@click.option('--psql_db', 
    prompt="PostgreSQL Database Name", help="Database Name")
@click.option('--psql_user', 
    prompt="PostgreSQL user", help="PostgreSQL user")
@click.option('--psql_pass', 
    prompt="PostgreSQL user password", help="PostgreSQL password")
@click.option('--print_query', 
    prompt="View sqldump ", help="show sql dump to sql", is_flag=True)
@click.option('--use_schema', 
    prompt="Use Schema ", help="choose Y if use schema", is_flag=True)


def convert_mdb_to_psql(mdbfile, psql_host, psql_db, psql_user, psql_pass, use_schema, print_query):
    if not os.path.exists(mdbfile):
        click.echo(f'File not exist: {mdbfile}')
        sys.exit(0)
    click.echo(f'Generate {mdbfile}')
    
    convert_data = mdb2psql.mdb2psql(mdbfile, psql_host, psql_db, psql_user, psql_pass, use_schema, print_query)
    convert_data.create_schema()

if __name__ == "__main__":
    convert_mdb_to_psql()