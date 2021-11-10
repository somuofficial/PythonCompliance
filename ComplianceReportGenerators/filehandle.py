import os
import datetime
import shutil
from .conf import config

class fhandle:
    
    def __init__(self):
        self.dt = datetime.datetime.now().strftime("%B-%y")
        self.now = datetime.datetime.now()
        self.config = config("file")
        self.basePath = self.config.basePath
        self.inpPath = self.config.inpPath
        self.newDir = "complaince-"+self.dt
        self.fileDate = datetime.datetime.now().strftime("%Y-%m-%d")
        
    # Output dir creation, return filepath
    def createOutputDir(self):
        newPath = os.path.join(self.basePath, self.newDir)
        if os.path.exists(newPath):
            return newPath
        else:
            os.mkdir(newPath) 
            return self.newDir

if __name__ == "__main__":
    from dateutil.relativedelta import relativedelta
    a = fhandle()
    current = a.now
    ad = relativedelta(months=-1)
    ColDate = current + ad
    ColDate = ColDate.strftime("%b-%y")
    print(ColDate)