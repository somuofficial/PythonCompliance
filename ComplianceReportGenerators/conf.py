from configparser import ConfigParser

class config():
    
    def __init__(self,conftype,dbType=None):
        self.config = ConfigParser()
        self.config.read('ComplianceReportGenerators\\conf.ini')
        self.conftype = conftype.upper()
        if dbType != None:
            if conftype.upper() == "DATABASE":
                self.dbType = dbType
                return self.get_db()
        if conftype.upper() == "FILE":
            return self.get_file()
        else:
            raise Exception("Unknow type")
        
    def get_db(self):
        self.DbName = self.config.get(self.dbType+"-db",'name')
        self.port = self.config.get(self.dbType+"-db",'port')
        self.user = self.config.get(self.dbType+"-db",'user')
        self.password = self.config.get(self.dbType+"-db",'pass')
        self.ip = self.config.get(self.dbType+"-db",'ip')
        self.dbinstance = self.config.get(self.dbType+"-db",'instance')
        
    def get_file(self):
        self.basePath = self.config.get('file','fpath')
        self.inpPath = self.config.get('file','inpPath')
        
if __name__ == "__main__":
    a = config("file")
    print(a.basePath)
    
    
    
    
    