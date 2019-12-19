
import sys, os
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
#  3 from PyQt5.QtCore import *
#  4 from PyQt5.QtWidgets import QFileDialog, QMessageBox, QDockWidget, QListWidget
#  5 from PyQt5.QtGui import *

from Ui_test import Ui_MainWindow  #导入创建的GUI类
from compareExcel import getInfoBySheetNamesBeginRowIdAndColId

# excelPath 输入Excel的路径
# sheetNames 需要对比的sheet，格式为list
# beginRowId 检索起始的行，用于排除表头等。起始为0
# keyColId 检索的关键字所在列，起始为0， 比如姓名
# valueColId 匹配的字段所在列，起始为0， 比如身份证号
# 【judgeColId】设置一列用于判断这行的信息是否参与比对。比如1班和3班有重名的人，只比对1班的那个。如果不需要这个功能，则judgeColId、judgeIn、judgeOptions都不设。
# 【judgeIn】True:如果judgeColId对应的值在judgeOptions，则加入比对；False:如果judgeColId对应的值在judgeOptions，则不加入比对
# 【judgeOptions】格式为list
class Excel():
    excelPath = 'path'
    sheetNames = []
    beginRowId = 0
    keyColId = 0
    valueColId = 0
    judgeIn = True
    judgeOptions = []

#自己建一个mywindows类，mywindow是自己的类名。QtWidgets.QMainWindow：继承该类方法
class mywindow(QtWidgets.QMainWindow, Ui_MainWindow):
    newExcel = Excel()
    baseExcel = Excel()
    #__init__:析构函数，也就是类被创建后就会预先加载的项目。
    # 马上运行，这个方法可以用来对你的对象做一些你希望的初始化。
    def __init__(self):
        #这里需要重载一下mywindow，同时也包含了QtWidgets.QMainWindow的预加载项。
        super(mywindow, self).__init__()
        self.setupUi(self)
        self.main()

    def main(self):
        #read new Excel path
        self.button_SelectNewFile.clicked.connect(self.selectNewFile)
        self.button_SelectBaseFile.clicked.connect(self.selectBaseFile)
        self.button_BeginProcess.clicked.connect(self.beginProcess)

    def beginProcess(self):
        self.newExcel.sheetNames = self.lineEdit_newSheetNames.text().split(',')
        self.newExcel.beginRowId = int(self.lineEdit_newBeginRowId.text()) - 1
        self.newExcel.keyColId   = int(self.lineEdit_newKeyColId.text()) - 1
        self.newExcel.valueColId = int(self.lineEdit_newValueColId.text()) - 1
        self.newExcel.judgeColId = int(self.lineEdit_newJudgeColId.text()) - 1
        if self.checkBox_newjudgeIn.isChecked():
            self.newExcel.judgeIn = True
        else:
            self.newExcel.judgeIn = False
        # self.newExcel.judgeIn    = int(self.lineEdit_newJudgeIn.text())
        self.newExcel.judgeOptions = self.lineEdit_newJudgeOptions.text().split(',')

        self.baseExcel.sheetNames = self.lineEdit_baseSheetNames.text().split(',')
        self.baseExcel.beginRowId = int(self.lineEdit_baseBeginRowId.text()) - 1
        self.baseExcel.keyColId   = int(self.lineEdit_baseKeyColId.text()) - 1
        self.baseExcel.valueColId = int(self.lineEdit_baseValueColId.text()) - 1
        self.baseExcel.judgeColId = int(self.lineEdit_baseJudgeColId.text()) - 1
        if self.checkBox_basejudgeIn.isChecked():
            self.baseExcel.judgeIn = True
        else:
            self.baseExcel.judgeIn = False
        # self.baseExcel.judgeIn    = int(self.lineEdit_baseJudgeIn.text())
        self.baseExcel.judgeOptions = self.lineEdit_baseJudgeOptions.text().split(',')

        if not self.checkBox_newIfJudgeSpecified.isChecked():
            self.newExcel.judgeColId = -1
        newInfoDict = getInfoBySheetNamesBeginRowIdAndColId(self.newExcel.fileName,
        self.newExcel.sheetNames,
        self.newExcel.beginRowId,
        self.newExcel.keyColId,
        self.newExcel.valueColId,
        self.newExcel.judgeColId,
        self.newExcel.judgeIn,
        self.newExcel.judgeOptions)

        if not self.checkBox_baseIfJudgeSpecified.isChecked():
            self.baseExcel.judgeColId = -1
        baseInfoDict = getInfoBySheetNamesBeginRowIdAndColId(self.baseExcel.fileName,
        self.baseExcel.sheetNames,
        self.baseExcel.beginRowId,
        self.baseExcel.keyColId,
        self.baseExcel.valueColId,
        self.baseExcel.judgeColId,
        self.baseExcel.judgeIn,
        self.baseExcel.judgeOptions)

        compareResult = []
        newKeyNum = 0
        differentNum = 0
        for newKey in newInfoDict:
            newKeyNum += 1
            if baseInfoDict[newKey] != newInfoDict[newKey]:
                differentNum += 1
                result = 'key:' + newKey + ' new value:' + newInfoDict[newKey] + ' base value:' + baseInfoDict[newKey]
                compareResult.append(result)
        compareResult.append('newKeyNum: ' + str(newKeyNum))
        compareResult.append('differentNum: ' + str(differentNum))
        showResult = ''
        for line in compareResult:
            showResult = showResult + '\n' + line
        self.textBrowser_showResult.setText(showResult)
        
    # def begin(self):
    #     beginRowId = self.lineEdit.text()
    #     print(beginRowId)
    #     self.showResultLabel.setText(beginRowId)

    def selectNewFile(self):
        fileName = QFileDialog.getOpenFileName(self, '打开文件','./',("excel (*.xls *xlsx)"))
        self.newExcel.fileName = fileName[0]
        self.Label_ShowNewFlie.setText(fileName[0])

    def selectBaseFile(self):
        fileName = QFileDialog.getOpenFileName(self, '打开文件','./',("excel (*.xls *xlsx)"))
        self.baseExcel.fileName = fileName[0]
        self.Label_ShowBaseFlie.setText(fileName[0])
    # def getline(self):
    #     line = self.lineEdit.text()
    #     print(line)




 
if __name__ == '__main__': #如果整个程序是主程序
      # QApplication相当于main函数，也就是整个程序（很多文件）的主入口函数。
      # 对于GUI程序必须至少有一个这样的实例来让程序运行。
     app = QtWidgets.QApplication(sys.argv)
     #生成 mywindow 类的实例。
     window = mywindow()
     #有了实例，就得让它显示，show()是QWidget的方法，用于显示窗口。
     window.show()
     # 调用sys库的exit退出方法，条件是app.exec_()，也就是整个窗口关闭。
     # 有时候退出程序后，sys.exit(app.exec_())会报错，改用app.exec_()就没事
     # https://stackoverflow.com/questions/25719524/difference-between-sys-exitapp-exec-and-app-exec
     sys.exit(app.exec_())