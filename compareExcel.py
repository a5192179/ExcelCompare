# 功能：对比两个Excel的列1与列2的对应关系是否正确。比如姓名和身份证的对应关系是否一致
import xlrd
# excelPath 输入Excel的路径
# sheetNames 需要对比的sheet，格式为list
# beginRowId 检索起始的行，用于排除表头等。起始为0
# keyColId 检索的关键字所在列，起始为0， 比如姓名
# valueColId 匹配的字段所在列，起始为0， 比如身份证号
# 【judgeColId】设置一列用于判断这行的信息是否参与比对。比如1班和3班有重名的人，只比对1班的那个。如果不需要这个功能，则judgeColId、judgeIn、judgeOptions都不设。
# 【judgeIn】True:如果judgeColId对应的值在judgeOptions，则加入比对；False:如果judgeColId对应的值在judgeOptions，则不加入比对
# 【judgeOptions】与judgeIn
def getInfoBySheetNamesBeginRowIdAndColId(excelPath, sheetNames, beginRowId, keyColId, valueColId, judgeColId = -1, judgeIn = True, judgeOptions = []):
    book = xlrd.open_workbook(excelPath)
    if len(sheetNames) == 0: 
        sheetNames = book.sheet_names()
    keys = []
    values = []
    for sheetName in sheetNames:
        sheet = book.sheet_by_name(sheetName)
        rows = sheet.nrows #获取行数
        # cols = sheet.ncols #获取列数
        for row in range(beginRowId, rows): #读取每一行的数据
            if judgeColId != -1:
                judgeInfo = sheet.cell(row, judgeColId)
                bInJudgeOptions = False
                for judgeOption in judgeOptions:
                    if judgeOption == judgeInfo.value:
                        bInJudgeOptions = True
                        break
                if bInJudgeOptions and not judgeIn:
                    continue
                if not bInJudgeOptions and judgeIn:
                    continue
            key = sheet.cell(row, keyColId) #读取指定单元格的数据
            value = sheet.cell(row, valueColId) #读取指定单元格的数据
            keys.append(key.value)
            values.append(value.value)
    infoDict = dict(zip(keys, values))
    return infoDict


if __name__ == '__main__':
    newExcelPath = '/Users/test/OneDrive/Project/jialing/ExcelCompare/YibaoData/工作簿1.xlsx'
    # sheetNames = ['1', '2', '3', '8', '9']
    sheetNames = ['Sheet1']
    beginRowId = 0
    keyColId = 0
    valueColId = 2
    # judgeColId = 11
    # judgeIn = True
    # judgeOptions = ['191901', '191902', '191903', '191908', '191909']
    newStudentID = getInfoBySheetNamesBeginRowIdAndColId(newExcelPath, sheetNames, beginRowId, keyColId, valueColId)#, judgeColId, judgeIn, judgeOptions)

    baseExcelPath = '/Users/test/OneDrive/Project/jialing/ExcelCompare/YibaoData/大学生参保登记导盘表格式 - 商学院.xls'
    sheetNames = ['Sheet1']
    beginRowId = 1
    keyColId = 1
    valueColId = 2
    judgeColId = 11
    judgeIn = True
    judgeOptions = ['191901', '191902', '191903', '191908', '191909']#191901,191902,191903,191908,191909
    baseStudentID = getInfoBySheetNamesBeginRowIdAndColId(baseExcelPath, sheetNames, beginRowId, keyColId, valueColId, judgeColId, judgeIn, judgeOptions)

    for name in newStudentID:
        if baseStudentID[name] != newStudentID[name]:
            print(name + ' should be ' + baseStudentID[name] + ' not ' + newStudentID[name])
