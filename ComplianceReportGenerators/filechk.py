from .conf import config
from .filehandle import fhandle
import os
from os import path

class FileCheck:
    def __init__(self,file):
        filehandler = fhandle()
        self.inpPath = filehandler.inpPath
        self.filePath = os.path.join(self.inpPath, file)
        self.out = self.fileChk()    
    def fileChk(self):
        chk = path.isfile(self.filePath)
        if chk == True:
            return True
        else:
            o = "File: "+self.filePath+" not present"
            raise Exception(o)

if __name__ == "__main__":
    a=FileCheck("ASCR 2021-08-5.xlsx")
    print(a.out)