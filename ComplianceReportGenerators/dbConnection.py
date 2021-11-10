import json
import pandas as pd
#import cx_Oracle
import pyodbc
import sqlalchemy
from .conf import config
class dbConnection():
    def __init__(self,dbtype):
        self.config = config("database",dbtype)
        if dbtype.upper() == "MSSQL":
            return self.mssqlCon()
        if dbtype.upper() == "ORACLE":
            return self.oracleCon()
        else:
            raise Exception("Unknow database type")
    def oracleCon(self):
        dsn = "oracle+cx_oracle://"+self.config.user+":"+self.config.password+"@"+self.config.ip+":"+self.config.port+"/"+self.config.dbinstance
        try:
            oracle_db = sqlalchemy.create_engine(dsn)
            self.con = oracle_db
        except Exception as e:
            return "Database connection error: "+str(e)
    def mssqlCon(self):
        dsn = "mssql+pyodbc://"+self.config.user+":"+self.config.password+"@"+self.config.ip+":"+self.config.port+"/"+self.config.dbinstance+"?driver=SQL+Server+Native+Client+11.0"
        try:
            mssql_db = sqlalchemy.create_engine(dsn)
            self.con = mssql_db
        except Exception as e:
            return "Database connection error: "+str(e)
            
if __name__ == "__main__":
    a = dbConnection("oracle")
    print(a.con)