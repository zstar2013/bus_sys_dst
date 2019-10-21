from PyQt5 import QtWidgets
from listview import Ui_MainWindow
from PyQt5.Qt import QAbstractListModel, QModelIndex, QVariant,QAbstractTableModel
from PyQt5.QtCore import Qt
import sys
sys.path.append("..") #把上级目录加入到变量中
from filepath import get_local_dir
from tools import dialogTools
from tools import xlstool as xt
import traceback
import os
class MyWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    # 类初始化
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setupUi(self)
        self.retranslateUi(self)
        self.current_file_path="G://"
        #显示子窗口
        self.actionopen.triggered.connect(self.showChild)
        #显示保存文件，默认"G:\\test.xls"
        self.pb_savefile.clicked.connect(lambda :self.showSaveFile("G:\\test.xls"))
        #写入文件
        self.pb_writeformat.clicked.connect(lambda :self.writeformat("G:\\test.xls"))
        self.pb_openfile.clicked.connect(self.openXlsfile)
        self.checkBox_3.setTristate(True)
        self.checkBox_3.setCheckState(Qt.PartiallyChecked)
        self.checkBox.toggled.connect(lambda: self.btnstate(self.checkBox))
        self.checkBox_2.toggled.connect(lambda: self.btnstate(self.checkBox_2))
        self.checkBox_3.toggled.connect(lambda: self.btnstate(self.checkBox_3))
        self.list_data = get_local_dir(self.current_file_path)
        lm=MyListModel(self.list_data,self)
        self.listView.setModel(lm)
        self.listView.doubleClicked.connect(self.clicked)

    def clicked(self, qModelIndex):
        # 提示信息弹窗，你选择的信息
        #QtWidgets.QMessageBox.information(self, 'ListWidget', '你选择了：'+str(self.list_data[qModelIndex.row()]))
        self.label_select_path.setText(self.current_file_path+str(self.list_data[qModelIndex.row()]))

    def btnstate(self, btn):
        chk1Status = self.checkBox.text() + ", isChecked=" + str(
            self.checkBox.isChecked()) + ', chekState=' + str(
            self.checkBox.checkState()) + "\n"
        chk2Status = self.checkBox_2.text() + ", isChecked=" + str(
            self.checkBox_2.isChecked()) + ', checkState=' + str(
            self.checkBox_2.checkState()) + "\n"
        chk3Status = self.checkBox_3.text() + ", isChecked=" + str(
            self.checkBox_3.isChecked()) + ', checkState=' + str(
            self.checkBox_3.checkState()) + "\n"
        self.bottomText.setText(chk1Status + chk2Status + chk3Status)

    def openXlsfile(self):
        os.startfile("G:\\test.xls")
    def showChild(self):
        QtWidgets.QMessageBox.information(self,"title","helloworld!")
        self.statusbar.showMessage("打开状态栏提醒")


    ###显示保存窗口
    def showSaveFile(self,filePath="D:\\"):
        print(filePath)
        try:
            dialogTools.showSaveFileDialog(self,filePath,lambda x:self.mCallback(x))
        except:
            traceback.print_exc()

    def mCallback(self,path):
        xt.createNewFile(path)
    def writeformat(self,path):
        try:
            xt.write(path,"sheet1")
        except:
            traceback.print_exc()



class MyListModel(QAbstractListModel):
    def __init__(self, datain,parent=None):
        super(MyListModel, self).__init__(parent)
        # Keep track of which object are checked
        self.list_data = datain

    def rowCount(self, QModelIndex):
        return len(self.list_data)


    def data(self, index, role):
        try:
            if index.isValid() and role == Qt.DisplayRole:
                return QVariant(self.list_data[index.row()])
            else:
                return QVariant()
            #return QVariant(self.list_data[index.row()])
        except:
            traceback.print_exc()

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    dlg = MyWindow()
    dlg.show()

    sys.exit(app.exec_())