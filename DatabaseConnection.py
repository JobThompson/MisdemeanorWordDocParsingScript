from time import sleep
from sqlalchemy.engine import URL, create_engine
import pandas

DB_NAME = 'GenericCourt'
DB_HOST = 'JUSTICEDEV'

CONNECTION_STRING = (
                    'DRIVER=SQL Server;'
                    f'SERVER={DB_HOST};'
                    f'DATABASE={DB_NAME};'
                    'Trusted_Connection=yes;'
                )

class DatabaseConnection():
    def __init__(self) -> None:
        self.connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": CONNECTION_STRING})
        self.engine = self.create_engine()
        
    def create_engine(self):
        while True:
            try:
                return create_engine(self.connection_url)
            except Exception as e:
                print(e)
                print('Unable to connect to Database, please check the database name and your network connection.')
                sleep(5) # wait 5 seconds before trying again
                
    def get_engine(self):
        if(self.engine):
            return self.engine
        else:
            self.engine = self.create_engine()
            return self.engine
        
    def execute_sql(self, sql_string = '', silence_errors = False):
        try:
            with self.get_engine().connect() as conn:
                conn.execute(f'{sql_string}')
        except Exception as e:
            if not silence_errors:
                print(f'Unable to run SQL: {e}')
        
    def get_query_result(self, sql_string = '', silence_errors = False):
        try:
            df = pandas.read_sql_query(f'{sql_string}', self.get_engine())
            return df
        except Exception as e:
            if not silence_errors:
                print(f'Unable to get query result: {e}')
            return pandas.DataFrame()
        
DBConn = DatabaseConnection()