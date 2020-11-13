import datetime
import sys
from shutil import copyfile

import pandas as pd
import xlwings as xw

import tools

PathDir = "c:/Users/runda/OneDrive - RundaTech/04 工作/0405 艾格威贸易"


def loadCase(app, filepath):  # 抽取选配方案
    cols = ["牛号", "普精1", "普精2", "普精3", "性控1", "性控2", "性控3"]
    cols2 = ["牛号", "1选", "2选", "3选"]
    wb = app.books.open(filepath)  # 打开选配方案
    sht = wb.sheets[0]
    rng = sht.used_range
    res = []  # 存放抽取的选配方案
    if set(cols2).issubset(set(rng.rows(1).value)):
        # 如果1选2选3选这样的格式
        for col in cols2:
            i = rng.rows(1).value.index(col)
            res.append(rng.columns[i].value)
    else:
        for col in cols:
            i = rng.rows(1).value.index(col)
            res.append(rng.columns[i].value)
    wb.close()  # 关闭选配方案
    data = []  # 对抽取的数据进行转换和填充
    for c in range(len(res[0])):
        tmp = []
        colos = [0, 1, 2, 3, 4, 5, 6]
        if len(res) == 4:
            colos = [0, 1, 2, 3, 1, 2, 3]
        for r in colos:
            tmp.append(res[r][c])
        data.append(tmp)
    return data


# 写入4个牧场的选配方案
def writeCaseData():
    app = xw.App(visible=True, add_book=False)
    for farm in farmList2:
        print("开始处理{}牧场".format(farm))
        caseData = []  # 存储1个牧场的所有选配方案
        path = '{}/ALTA/ALTA/12 汇报/现代牧业/首配准确性/{}/'.format(PathDir, farm)
        resFile = '选配方案执行情况 {} 2020.xlsm'.format(farm)
        for file in casesList[farm].strip("\n").split("\n"):
            if file:
                data = loadCase(app, path + file)  # 提取选配方案
                caseData.append([file[:8], data])
            else:
                continue
        # writeData(caseData, app)
        # 将选配方案和配种记录写入统计表
        resFilePath = path + resFile
        wb = app.books.open(resFilePath)
        for i in range(len(caseData)):
            sheetName = caseData[i][0]
            data = caseData[i][1]
            try:
                sht2 = wb.sheets.add(sheetName, before="配种记录")  # 打开统计表
            except Exception:
                print("{}已存在，继续下一个".format(sheetName))
                continue
            rng2 = sht2.range("A1")
            # rng2 = sht2.used_range
            rng2.value = data  # 将选配方案写入统计表
            print("文件：{}处理完成".format(sheetName))
        wb.save()
        wb.close()
        print("{}牧场处理完成。".format(farm))
    app.quit()


# 写入6个牧场的配种记录
def writeBredLog(month, year=2020):  # 写入配种记录
    app = xw.App(visible=True, add_book=False)
    for farm in farmList:
        path = '{}/ALTA/ALTA/12 汇报/现代牧业/首配准确性/{}/'.format(PathDir, farm)
        path2 = '{}/ALTA_Matching/BredLog/'.format(PathDir)
        resFile = '选配方案执行情况 {} 2020.xlsm'.format(farm)
        resFileBack = '选配方案执行情况 {} 2020_backup.xlsm'.format(farm)
        csvFile = '配种记录_{}年{}月_{}_结果.csv'.format(year, month, farm)
        copyfile(path + resFile, path + resFileBack)  # 复制一份作为备份
        wb = app.books.open(path + resFile)
        # 写入配种记录
        wb.sheets("配种记录").activate()
        sht = wb.sheets("配种记录")
        if sht.range(27, 16).value == "耳号":
            rng = sht.range(27, 16).expand('down')
            r = rng.rows.count + rng.row
            rng = sht.range(r, 16)
            # 读取配种记录
            csvData = pd.read_csv(path2 + csvFile, index_col=0)
            rng.value = csvData.values
        else:
            continue
        wb.save()
        wb.close()
        print("{}牧场完成。".format(farm))
    app.quit()


if __name__ == "__main__":
    t1 = datetime.datetime.now()
    print("开始时间：{}".format(t1))
    # '洪雅', '和林', '汶上', '马鞍山', '蚌埠', '宝鸡'
    farmList = ['洪雅', '和林', '马鞍山', '蚌埠', '宝鸡']  # 配种记录的牧场列表
    farmList2 = ['洪雅', '和林', '马鞍山', '蚌埠', '宝鸡']  # 选配方案的牧场列表
    # 各牧场的选配方案列表
    casesList = {
        '洪雅':
        """
20201017_选配方案_洪雅.xlsx
""",
        '和林':
        """
20201009_选配方案_和林.xlsx
""",
        '马鞍山':
        """
20201006_选配方案_马鞍山.xlsx
20201017_选配方案_马鞍山.xlsx
""",
        '蚌埠':
        """
20201010_选配方案_蚌埠_5.xlsx
""",
        '宝鸡':
        """
20201001_选配方案_宝鸡_青年牛.xlsx
20201009_选配方案_宝鸡_青年牛.xlsx
20201019_选配方案_宝鸡_泌乳牛.xlsx
20201023_选配方案_宝鸡_泌乳牛.xlsx
"""
    }
    mth = 10  # 月份
    yer = 2020  # 年份
    # 1写入6个牧场的配种记录
    stat = tools.extract_log("""配种记录_2020年10月_蚌埠.xlsx
配种记录_2020年10月_和林.xlsx
配种记录_2020年10月_马鞍山.xlsx
配种记录_2020年10月_洪雅.xlsx
配种记录_2020年10月_宝鸡.xlsx""")
    if stat[0] == "1":
        print("提取配种记录出错，可能的原因是列字段名称有变动")
        sys.exit(1)
    t2 = datetime.datetime.now()
    print("提取配种记录完成,用时：{}".format(t2 - t1))
    print("1.开始写入牧场的配种记录：{}".format(t2))
    writeBredLog(mth, yer)  # 参数是月份
    t3 = datetime.datetime.now()
    print("1.完成写入牧场的配种记录：{}\n共用时{}".format(t3, t3 - t2))
    # 2写入6个牧场的选配方案（宝鸡和蚌埠需要再手动处理一下）
    t4 = datetime.datetime.now()
    print("2.开始写入牧场的选配方案：{}".format(t4))
    writeCaseData()
    t5 = datetime.datetime.now()
    print("1.完成写入牧场的选配方案：{}\n共用时{}\n\n全部执行完成，用时{}".format(
        t5, t5 - t4, t5 - t1))
