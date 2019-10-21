import os
import traceback
from PyQt5.QtWidgets import QFileDialog

def showFilePathDialog(self,filePath,callback):
    path = QFileDialog.getOpenFileName(self, "Open File Dialog", filePath if os.path.exists(filePath)else "/",
                                           "xls files(*.xls);;xlsx files(*.xlsx)")
    callback(path[0])

def showPathDialog(self,filePath,callback):
    path = QFileDialog.getExistingDirectory(self, "Open File Dialog", filePath if os.path.exists(filePath)else "/")
    
    callback(path)

def showSaveFileDialog(self,filePath,callback):
    try:
        path=QFileDialog.getSaveFileName(self,"save license",filePath if os.path.exists(filePath)else "/",
                                               "xls files(*.xls);;xlsx files(*.xlsx)")
        callback(path[0])
    except:
        traceback.print_exc()
