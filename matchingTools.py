"""
Created on Sep 14 2019  @author: Bright Geng

本模块是matching.py的GUI外壳
"""
import base64
import os
import pickle
import tkinter as tk
import tkinter.messagebox
import traceback
from tkinter import ttk

import pinyin

from creatDB import loaddb, writedb
from icon import img
from log import logger, logging
from matching import Farm
from parameters import DB, YMD, DIR
from tools import extract_log, merge_body, merge_files, mger_pos

version = '2.4-20201123'
CHANGLOG = """
version    Date     comments                        todo
2.4        11/23/20 修改明细中牛号有字母的bug
2.3        8/5/20   流产率设置分为了青年和成母牛
2.2        7/1/20   修改推算怀孕率的bug
2.1        4/13/20  预测的第2方法                   第3方法(机器学习)
2.0        4/03/20  存栏预测功能
1.0        3/20/20  主体已完成                      存栏预测模块更新细节

"""

# 2.3 流产率分为青年/成母，导入牛群明细函数增加替换"   - "的语句
# 2.2  self.herds.columns.tolist().index('繁殖代码')   原来是RC,改为繁殖代码并增加else语句
# 如果在GUI环境下运行，程序执行结果通过这个函数弹出窗口


def out_message(title, res):
    global ROOT
    # logger.info(title + ': ' + res)
    if __name__ == '__main__':
        tk.messagebox.showinfo(title, res, parent=ROOT)


# 如果在GUI环境下运行，程序执行出错的话，弹出这个窗口
def out_error(title, res):
    global ROOT
    logger.debug(title + ': ' + res)
    if __name__ == '__main__':
        tk.messagebox.showerror(title, res, parent=ROOT)


# 让logging同时向GUI打印
class LogtoUi(object):
    def __init__(self, weight):
        self.kongjian = weight

    def write(self, mesage, end='\n'):
        if __name__ == "__main__":
            self.kongjian.configure(state='normal')
            self.kongjian.insert('end', mesage)
            self.kongjian.update()
            self.kongjian.see('end')
            self.kongjian.configure(state='disable')

    def flush(self):
        pass


# GUI界面
class APP(object):
    def __init__(self, master):
        self.fe_5_6 = tk.Frame(master)
        self.menu_bar_(master)
        self.forecast_SubWin4(master)  # 存栏预测
        self.merge_body_SubWin1(master)  # 工具1
        self.extract_log_SubWin2(master)  # 工具2
        self.merge_pos_SubWin3(master)  # 工具3
        self.merge_files_SubWin4(master)  # 工具4 合并选配文件
        self.demand_SubWin5(master)  # 维护牧场需求
        self.semen_SubWin6(master)  # 维护冻精关系
        self.mast_window(master)  # 主窗口

        self.farms = ()
        self.loaded = 0

    def changlog(self):
        logger.info(CHANGLOG)

    # 菜单栏
    def menu_bar_(self, master):
        # global ROOT
        menubar = tk.Menu(master)
        # 工具
        toolmenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='工具', menu=toolmenu)
        toolmenu.add_command(label='1_备用')
        toolmenu.add_separator()  # 添加一条分隔线
        # 设置
        cofigMenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='设置', menu=cofigMenu)
        cofigMenu.add_command(label='创建目录', command=self.mkdir)
        # 帮助
        helpmenu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label='帮助', menu=helpmenu)
        helpmenu.add_command(
            label='建议和反馈',
            command=lambda: out_message(
                '联系方式', '如需帮助，或者有建议和反馈请联系:\n\n耿润达\n手机：153 1313 9012\n微信同号'))
        helpmenu.add_command(label='版本历史', command=self.changlog)
        ROOT.config(menu=menubar)

    # 设置1：创建目录
    def mkdir(self):
        path1 = os.path.join(DIR, 'Match_files')
        path2 = os.path.join(DIR, 'Postions')
        path3 = os.path.join(DIR, '待做体型')
        path4 = os.path.join(DIR, '选配文件')
        path5 = os.path.join(DIR, 'BredLog')
        path6 = os.path.join(DIR, 'body_by_month')
        path7 = os.path.join(DIR, 'body_by_month', 'old')
        path8 = os.path.join(DIR, 'AltaGPS_Data')
        path9 = os.path.join(DIR, 'AltaGPS_Reports')
        path10 = os.path.join(DIR, '主要信息')
        path11 = os.path.join(DIR, 'log')
        path_list = [
            path1, path2, path3, path4, path5, path6, path7, path8, path9,
            path10, path11
        ]
        paths = ""
        paths_is = ""
        for i in range(0, len(path_list)):
            isExists = os.path.exists(path_list[i])
            if not isExists:
                os.makedirs(path_list[i])
                paths += '\n' + path_list[i]
            else:
                paths_is += '\n' + path_list[i]
        res = '如下目录创建完成:\n' + paths
        res2 = '如下目录已存在:\n' + paths_is
        if __name__ == '__main__':
            out_message('执行完成', res + '\n\n' + res2)
        else:
            print(res, res2)

    def getFarms(self, widget):
        self.farms = []
        with open(DB, 'rb') as db:
            demand = pickle.load(db)['demand']
        for i in range(len(demand)):
            self.farms.append(demand[i][0])
        widget["values"] = self.farms

    # 主功能：选配工具
    def mast_window(self, master):
        frame100 = tk.Frame(master, padx=5)
        frame0 = tk.Frame(master, padx=10)
        frame1 = tk.Frame(master, padx=80)
        frame2 = tk.Frame(master, padx=80)
        frame3 = tk.Frame(master, padx=80)
        frame4 = tk.Frame(master, padx=10)
        frame100.pack(side='top', anchor='nw')  # tab页标题
        frame4.pack(side='right', anchor='n')  # 日志

        def updateName(*args):
            s = self.farm.get() + '_体型明细_' + str(YMD[:6]) + '.csv'
            self.last_body_list.delete(0, 'end')
            self.last_body_list.insert(0, s)

        def changeTag(tag):
            for fre in [
                    frame0, frame1, frame2, frame3, self.fe5, self.fe6,
                    self.fe7, self.fe8, self.fe9, self.fe10, self.fe11,
                    self.fe_5_6
            ]:
                fre.pack_forget()
            if tag == 0:
                self.fe8.pack(side='bottom', anchor='nw', pady=10)  # 存栏预测
                frame0.pack(side='left', anchor='n')  # 导入牛群明细
                frame1.pack(side='top', anchor='nw')  # 筛选体型
                frame2.pack(side='top', anchor='nw')  # 生成系谱
                frame3.pack(side='top', anchor='nw')  # 生成选配文件

            elif tag == 1:

                self.fe_5_6.pack(side='top', fill='x')

                self.fe5.pack(side='left',
                              anchor='nw',
                              fill='both',
                              expand='yes')  # 工具1：合并体型
                self.fe6.pack(side='right',
                              anchor='ne',
                              fill='both',
                              expand='yes')  # 工具2：提取配种记录
                self.fe7.pack(side='left',
                              anchor='sw',
                              fill='both',
                              expand='yes',
                              pady=50)  # 工具3：提取定位文件
                self.fe11.pack(side='right',
                               anchor='se',
                               fill='both',
                               expand='yes',
                               pady=50)  # 工具4：合并选配方案
            elif tag == 2:
                pass
            elif tag == 3:
                self.fe9.pack(side='left', anchor='n')  # 设置牧场需求
            elif tag == 4:
                self.fe10.pack(side='left', anchor='n')  # 设置冻精对应关系

        tag = tk.IntVar(value=0)
        tagWidth = 10
        tk.Radiobutton(frame100,
                       text="核心功能",
                       width=tagWidth,
                       variable=tag,
                       value=0,
                       bd=1,
                       indicatoron=0,
                       command=lambda: changeTag(0)).grid(column=0, row=0)
        tk.Radiobutton(frame100,
                       text="工具1-4",
                       width=tagWidth,
                       variable=tag,
                       value=1,
                       bd=1,
                       indicatoron=0,
                       command=lambda: changeTag(1)).grid(column=1, row=0)
        tk.Radiobutton(frame100,
                       text="维护牧场需求",
                       width=tagWidth,
                       variable=tag,
                       value=3,
                       bd=1,
                       indicatoron=0,
                       command=lambda: changeTag(3)).grid(column=2, row=0)
        tk.Radiobutton(frame100,
                       text="维护冻精前缀",
                       width=tagWidth,
                       variable=tag,
                       value=4,
                       bd=1,
                       indicatoron=0,
                       command=lambda: changeTag(4)).grid(column=3, row=0)

        # 第一步：导入牛群明细
        self.lb = tk.Label(frame0, text="\n第一步：导入牛群明细", padx=10, fg='blue')
        self.lb.grid(row=0, column=1, columnspan=2, sticky='w')
        self.lb = tk.Label(frame0, text="牧场名：", padx=10).grid(row=1, column=1)
        self.farm = ttk.Combobox(frame0,
                                 width=67,
                                 state='readonly',
                                 postcommand=lambda: self.getFarms(self.farm))
        self.farm.bind("<<ComboboxSelected>>", updateName)
        self.farm.grid(row=1, column=2)
        self.lb = tk.Label(frame0, text="牛群明细：", padx=10).grid(row=2, column=1)
        self.herds = tk.Entry(frame0, width=70)
        self.herds.grid(row=2, column=2)
        self.re_load = tk.IntVar()
        self.overLoad = tk.Checkbutton(frame0,
                                       variable=self.re_load,
                                       text='从牛群明细文件中读取，即不从pickle文件中读取')
        self.overLoad.grid(row=3, column=2, sticky='e')
        w = tk.Label(frame0,
                     text="下方为对应关系：每一个项目中找到一个即停止，按从左到右的顺序。" + "一般情况下不需要修改",
                     padx=10)
        w.grid(row=4, column=1, columnspan=2, stick='w')
        with open(DB, 'rb') as db:
            col = pickle.load(db)['colums']
        self.lb = tk.Label(frame0, text="牛号", padx=10).grid(row=5, column=1)
        val1 = tk.StringVar(value=','.join(col[0]))
        self.id = tk.Entry(frame0, width=70, textvariable=val1)
        self.id.grid(row=5, column=2)
        self.lb = tk.Label(frame0, text="月龄", padx=10).grid(row=6, column=1)
        val2 = tk.StringVar(value=','.join(col[1]))
        self.age = tk.Entry(frame0, width=70, textvariable=val2)
        self.age.grid(row=6, column=2)
        self.lb = tk.Label(frame0, text="胎次", padx=10).grid(row=7, column=1)
        val3 = tk.StringVar(value=','.join(col[2]))
        self.lact = tk.Entry(frame0, width=70, textvariable=val3)
        self.lact.grid(row=7, column=2)
        self.lb = tk.Label(frame0, text="繁殖状态", padx=10).grid(row=8, column=1)
        val4 = tk.StringVar(value=','.join(col[3]))
        self.rpro = tk.Entry(frame0, width=70, textvariable=val4)
        self.rpro.grid(row=8, column=2)
        self.lb = tk.Label(frame0, text="泌乳天数", padx=10).grid(row=9, column=1)
        val5 = tk.StringVar(value=','.join(col[4]))
        self.dim = tk.Entry(frame0, width=70, textvariable=val5)
        self.dim.grid(row=9, column=2)
        self.lb = tk.Label(frame0, text="产后天数", padx=10).grid(row=10, column=1)
        val6 = tk.StringVar(value=','.join(col[5]))
        self.dsfrh = tk.Entry(frame0, width=70, textvariable=val6)
        self.dsfrh.grid(row=10, column=2)
        self.lb = tk.Label(frame0, text="空怀天数", padx=10).grid(row=11, column=1)
        val7 = tk.StringVar(value=','.join(col[6]))
        self.dopn = tk.Entry(frame0, width=70, textvariable=val7)
        self.dopn.grid(row=11, column=2)
        self.lb = tk.Label(frame0, text="配种次数", padx=10).grid(row=12, column=1)
        val8 = tk.StringVar(value=','.join(col[7]))
        self.tbrd = tk.Entry(frame0, width=70, textvariable=val8)
        self.tbrd.grid(row=12, column=2)
        self.lb = tk.Label(frame0, text="奶量1", padx=10).grid(row=13, column=1)
        val9 = tk.StringVar(value=','.join(col[8]))
        self.milk1 = tk.Entry(frame0, width=70, textvariable=val9)
        self.milk1.grid(row=13, column=2)
        self.lb = tk.Label(frame0, text="奶量2", padx=10).grid(row=14, column=1)
        val10 = tk.StringVar(value=','.join(col[9]))
        self.milk2 = tk.Entry(frame0, width=70, textvariable=val10)
        self.milk2.grid(row=14, column=2)
        self.lb = tk.Label(frame0, text="牛舍名称", padx=10).grid(row=15, column=1)
        val11 = tk.StringVar(value=','.join(col[10]))
        self.pen = tk.Entry(frame0, width=70, textvariable=val11)
        self.pen.grid(row=15, column=2)
        self.lb = tk.Label(frame0, text="出生日期", padx=10).grid(row=16, column=1)
        val12 = tk.StringVar(value=','.join(col[11]))
        self.bdat = tk.Entry(frame0, width=70, textvariable=val12)
        self.bdat.grid(row=16, column=2)
        self.lb = tk.Label(frame0, text="父亲牛号", padx=10).grid(row=17, column=1)
        val13 = tk.StringVar(value=','.join(col[12]))
        self.sid = tk.Entry(frame0, width=70, textvariable=val13)
        self.sid.grid(row=17, column=2)
        self.lb = tk.Label(frame0, text="外祖父号", padx=10).grid(row=18, column=1)
        val14 = tk.StringVar(value=','.join(col[13]))
        self.mgsid = tk.Entry(frame0, width=70, textvariable=val14)
        self.mgsid.grid(row=18, column=2)
        f_rc = tk.StringVar(value=','.join(col[14]))
        tk.Label(
            frame0,
            text="繁殖代码",
        ).grid(row=20, column=1)
        self.f_rc = tk.Entry(frame0, width=70, textvariable=f_rc)
        self.f_rc.grid(row=20, column=2)

        f_cdat = tk.StringVar(value=','.join(col[15]))
        tk.Label(
            frame0,
            text="怀孕日期",
        ).grid(row=60, column=1)
        self.f_cdat = tk.Entry(frame0, width=70, textvariable=f_cdat)
        self.f_cdat.grid(row=60, column=2)

        f_fdat = tk.StringVar(value=','.join(col[16]))
        tk.Label(
            frame0,
            text="产犊日期",
        ).grid(row=70, column=1)
        self.f_fdat = tk.Entry(frame0, width=70, textvariable=f_fdat)
        self.f_fdat.grid(row=70, column=2)

        f_ddat = tk.StringVar(value=','.join(col[17]))
        tk.Label(
            frame0,
            text="干奶日期",
        ).grid(row=80, column=1)
        self.f_ddat = tk.Entry(frame0, width=70, textvariable=f_ddat)
        self.f_ddat.grid(row=80, column=2)

        f_bday = tk.StringVar(value=','.join(col[18]))
        tk.Label(
            frame0,
            text="配种日期",
        ).grid(row=90, column=1)
        self.f_bday = tk.Entry(frame0, width=70, textvariable=f_bday)
        self.f_bday.grid(row=90, column=2)

        f_abdat = tk.StringVar(value=','.join(col[19]))
        tk.Label(
            frame0,
            text="流产日期",
        ).grid(row=100, column=1)
        self.f_abdat = tk.Entry(frame0, width=70, textvariable=f_abdat)
        self.f_abdat.grid(row=100, column=2)

        f_lsir = tk.StringVar(value=','.join(col[20]))
        tk.Label(
            frame0,
            text="与配公牛",
        ).grid(row=110, column=1)
        self.f_lsir = tk.Entry(frame0, width=70, textvariable=f_lsir)
        self.f_lsir.grid(row=110, column=2)

        self.upload = tk.Button(frame0,
                                text="导入牛群明细",
                                command=self.load_btnEvent)
        self.upload.grid(row=200, column=2, sticky='e', pady=10)
        self.save = tk.Button(frame0,
                              text="保存对应关系",
                              command=self.save_btnEvent)
        self.save.grid(row=200, column=1, sticky='e', pady=10)
        # 功能1：筛选体型
        self.lb = tk.Label(frame1, text="\n功能1：筛选体型", fg='blue')
        self.lb.grid(row=0, column=4, columnspan=2, sticky='w')
        w = tk.Label(frame1, text="最新体型明细", width=16).grid(row=10, column=4)
        self.last_body_list = tk.Entry(frame1, width=40)
        self.last_body_list.grid(row=10, column=5)
        mindim = tk.StringVar(value=30)
        w = tk.Label(frame1, text="最小DIM", width=16).grid(row=20, column=4)
        self.minDIM = tk.Entry(frame1, width=40, textvariable=mindim)
        self.minDIM.grid(row=20, column=5)
        maxdim = tk.StringVar(value=150)
        w = tk.Label(frame1, text="最大DIM", width=16).grid(row=30, column=4)
        self.maxDIM = tk.Entry(frame1, width=40, textvariable=maxdim)
        self.maxDIM.grid(row=30, column=5)

        tk.Label(frame1, text="").grid(row=40, column=5)
        self.sel_body = tk.Button(frame1,
                                  text="筛选体型",
                                  command=self.body_btnEvent)
        self.sel_body.grid(row=40, column=5, sticky='e')
        # 功能2：生成系谱
        w = tk.Label(frame2, text="功能2：生成系谱", fg='blue')
        w.grid(row=0, column=4, columnspan=2, sticky='w')

        tk.Label(frame2, text='筛选条件(>=月龄)', width=16).grid(row=10, column=4)
        valage = tk.StringVar(value=11)
        self.bredage = tk.Entry(frame2, width=40, textvariable=valage)
        self.bredage.grid(row=10, column=5)

        tk.Label(frame2, text='提取哪些列', width=16).grid(row=15, column=4)
        pgCol = tk.StringVar(value='牛号, 出生日期, 父亲牛号, 外祖父号, 胎次')
        self.pgCol = tk.Entry(frame2, width=40, textvariable=pgCol)
        self.pgCol.grid(row=15, column=5)

        self.has_title = tk.IntVar()
        self.needTitle = tk.Checkbutton(frame2,
                                        variable=self.has_title,
                                        text='保留标题行（第一行）')
        self.needTitle.grid(row=17, column=5, sticky='e')

        tk.Label(frame2, text="").grid(row=25, column=5)
        self.sel_pg = tk.Button(frame2, text="生成系谱", command=self.pg_btnEvent)
        self.sel_pg.grid(row=30, column=5, sticky='e')
        # 功能3：导出选配文件
        tk.Label(frame2, text="", width=10).grid(row=0, column=6)
        self.lb = tk.Label(frame3, text="功能3：导出选配文件", fg='blue')
        self.lb.grid(row=0, column=7, columnspan=2, sticky='w')
        tk.Label(frame3, text="常规冻精选配文件", width=16).grid(row=1, column=7)
        self.v41 = tk.StringVar(value="")
        self.com_file = tk.Entry(frame3, width=40, textvariable=self.v41)
        self.com_file.grid(row=1, column=8)
        tk.Label(frame3, text="常规冻精", width=16).grid(row=2, column=7)
        self.v42 = tk.StringVar(value="")
        self.com_sirs = tk.Entry(frame3, width=40, textvariable=self.v42)
        self.com_sirs.grid(row=2, column=8)
        tk.Label(frame3, text="常规冻精比例", width=16).grid(row=3, column=7)
        self.v43 = tk.StringVar(value="")
        self.com_rate = tk.Entry(frame3, width=40, textvariable=self.v43)
        self.com_rate.grid(row=3, column=8)
        tk.Label(frame3, text="常规用肉牛的牛列表", width=16).grid(row=4, column=7)
        self.v44 = tk.StringVar(value="")
        self.beef_list = tk.Entry(frame3, width=40, textvariable=self.v44)
        self.beef_list.grid(row=4, column=8)
        tk.Label(frame3, text="肉牛冻精", width=16).grid(row=5, column=7)
        self.v45 = tk.StringVar(value="")
        self.beef_sirs = tk.Entry(frame3, width=40, textvariable=self.v45)
        self.beef_sirs.grid(row=5, column=8)

        tk.Label(frame3, text='').grid(row=6, column=7)

        tk.Label(frame3, text="性控冻精选配文件", width=16).grid(row=8, column=7)
        self.v48 = tk.StringVar(value="")
        self.sex_file = tk.Entry(frame3, width=40, textvariable=self.v48)
        self.sex_file.grid(row=8, column=8)
        tk.Label(frame3, text="性控冻精", width=16).grid(row=9, column=7)
        self.v49 = tk.StringVar(value="")
        self.sex_sirs = tk.Entry(frame3, width=40, textvariable=self.v49)
        self.sex_sirs.grid(row=9, column=8)
        tk.Label(frame3, text="性控冻精比例", width=16).grid(row=10, column=7)
        self.v410 = tk.StringVar(value="")
        self.sex_rate = tk.Entry(frame3, width=40, textvariable=self.v410)
        self.sex_rate.grid(row=10, column=8)
        tk.Label(frame3, text="性控用常规的牛列表", width=16).grid(row=11, column=7)
        self.v411 = tk.StringVar(value="")
        self.com_list = tk.Entry(frame3, width=40, textvariable=self.v411)
        self.com_list.grid(row=11, column=8)

        tk.Label(frame3, text="配种月龄", width=16).grid(row=13, column=7)
        self.v413 = tk.StringVar(value=11)
        self.bred_age = tk.Entry(frame3, width=40, textvariable=self.v413)
        self.bred_age.grid(row=13, column=8)

        self.isAdjust = tk.IntVar()
        self.is_adjust = tk.Checkbutton(frame3,
                                        variable=self.isAdjust,
                                        text='不调整冻精的比例')
        self.is_adjust.grid(row=14, column=8, sticky='e')
        self.isFill = tk.IntVar()
        self.is_fill = tk.Checkbutton(frame3,
                                      variable=self.isFill,
                                      text='不填充匹配不上的')
        self.is_fill.grid(row=15, column=8, sticky='e')

        tk.Label(frame3, text="").grid(row=14, column=7)
        self.matchfile = tk.Button(frame3,
                                   text="生成选配文件",
                                   command=self.matchFile_btnEvent)
        self.matchfile.grid(row=20, column=8, sticky='e')
        tk.Label(frame3, text="").grid(row=20, column=6)
        self.lastData = tk.Button(
            frame3,
            text="上次数据",
            command=lambda: self.lastInfo_btnEvent(master))
        self.lastData.grid(row=20, column=7, sticky='e')
        # 功能4：日志输出
        w = tk.Label(frame4,
                     text="日志: \n(同时输出到ALTA_tool.log)",
                     padx=10,
                     fg='blue',
                     justify='left')
        w.grid(row=20, column=1, columnspan=2, sticky='w')
        s1 = tk.Scrollbar(frame4)
        s1.grid(row=21, column=3, sticky='ns')
        s2 = tk.Scrollbar(frame4, orient='horizontal')
        s2.grid(row=22, column=1, sticky='we', columnspan=2)
        self.outflow = tk.Text(frame4,
                               width=100,
                               height=70,
                               wrap='none',
                               yscrollcommand=s1.set,
                               xscrollcommand=s2.set)
        self.outflow.grid(row=21, column=1, columnspan=2)
        self.outflow.configure(state='disabled')
        s1.config(command=self.outflow.yview)
        s2.config(command=self.outflow.xview)

        changeTag(0)

    # 导入牛群明细按钮事件
    def load_btnEvent(self):
        if not self.farm.get():
            out_message('警告', '请选择牧场')
        elif not self.herds.get():
            out_message('警告', '请录入牧场全群明细文件的全名，包括后缀')
        elif self.herds.get().split('.')[-1].upper() not in ['XLSX', 'CSV']:
            out_message('警告', '只能导入XLSX文件或CSV文件,不区分大小')
        else:
            try:
                self.upload.config(state=tk.DISABLED, text='运行中...')
                self.farm_code = pinyin.get_initial(self.farm.get(),
                                                    delimiter="").upper()
                logger.debug('牧场名的拼音首字母'.format(self.farm_code))
                logger.debug('执行的命令是：{}=Farm({},{},{})'.format(
                    self.farm_code, self.farm.get(), self.herds.get(),
                    int(self.re_load.get())))
                exec('self.farm_code=Farm(self.farm.get(),self.herds.get(),' +
                     'int(self.re_load.get()))')
                self.loaded = 1  # 标识是否导入了牛群明细
            except Exception:
                out_error('错误', traceback.format_exc())
                self.loaded = 0
            else:
                out_message("完成", "牛群明细导入成功！")
            finally:
                self.upload.config(state=tk.NORMAL, text='导入牛群明细')

    # 保存对应关系按钮事件
    def save_btnEvent(self):
        items = [
            self.id, self.age, self.lact, self.rpro, self.dim, self.dsfrh,
            self.dopn, self.tbrd, self.milk1, self.milk2, self.pen, self.bdat,
            self.sid, self.mgsid, self.f_rc, self.f_cdat, self.f_fdat,
            self.f_ddat, self.f_bday, self.f_abdat, self.f_lsir
        ]
        co = []
        for it in range(len(items)):
            co.append(items[it].get().replace(' ', '').strip(',').split(','))
            logger.debug(it)
            logger.debug(items[it].get().split(','))
        with open(DB, 'rb') as db:
            dbDic = pickle.load(db)
        dbDic['colums'] = co
        with open(DB, 'wb') as db:
            pickle.dump(dbDic, db)
        logger.info('保存对应关系成功!')
        out_message('成功', '对应关系保存成功')

    # 筛选体型按钮事件
    def body_btnEvent(self):
        if self.loaded == 0:
            out_message('警告', '请先导入体型明细')
        elif self.loaded == 1:
            if not self.last_body_list.get():
                out_message('警告', '请填写最新体型明细文件名')
            elif self.last_body_list.get().split('.')[-1].upper() != 'CSV':
                out_message('警告', '体型明细只能是CSV(csv)文件')
            else:
                try:
                    self.sel_body.config(state=tk.DISABLED, text='运行中...')
                    logger.debug("执行的命令是：{}.body_select({},{},{})".format(
                        self.farm_code, self.last_body_list.get(),
                        self.minDIM.get(), self.maxDIM.get()))
                    exec("self.farm_code.body_select(" +
                         "self.last_body_list.get()," +
                         "self.minDIM.get(), self.maxDIM.get())")
                except Exception:
                    out_error('错误', traceback.format_exc())
                else:
                    out_message("完成", "体型筛选成功！")
                finally:
                    self.sel_body.config(state=tk.NORMAL, text='筛选体型')

    # 生成系谱按钮事件
    def pg_btnEvent(self):
        if self.loaded == 0:
            out_message('警告', '请先导入体型明细')
        elif self.loaded == 1:
            if not self.bredage:
                out_message('警告', '请输入筛选条件：月龄')
            else:
                try:
                    self.sel_pg.config(state=tk.DISABLED, text='运行中...')
                    logger.debug("执行的命令是：{}.produce_pedigree({},{},{})".format(
                        self.farm_code, int(self.bredage.get()),
                        self.pgCol.get(), self.has_title.get()))
                    exec("self.farm_code.produce_pedigree" +
                         "(int(self.bredage.get()), self.pgCol.get()" +
                         ", self.has_title.get())")
                except Exception:
                    out_error('错误', traceback.format_exc())
                else:
                    out_message("完成", "系谱提取完成！")
                finally:
                    self.sel_pg.config(state=tk.NORMAL, text='生成系谱')

    # 生成选配文件按钮事件
    def matchFile_btnEvent(self):
        if self.loaded == 0:
            out_message('警告', '请先导入体型明细')
        elif self.loaded == 1:
            if not self.bredage:
                out_message('警告', '请输入筛选条件：月龄')
            else:
                self.matchFile_btnEvent2()

    def matchFile_btnEvent2(self):
        self.matchfile.config(state=tk.DISABLED, text='运行中...')
        com_sirs = self.com_sirs.get().replace(' ', '').strip(',').split(',')
        sex_sirs = self.sex_sirs.get().replace(' ', '').strip(',').split(',')
        logger.debug("sex_sirs:{}".format(sex_sirs))
        if len(com_sirs) > 1:
            for i in range(len(com_sirs)):
                if len(com_sirs[i]) != 10:
                    out_message("校验失败", "冻精号{}位数错误".format(com_sirs[i]))
                    self.matchfile.config(state=tk.NORMAL, text='生成选配文件')
                    return
        if len(sex_sirs) > 1:
            for ix in range(len(sex_sirs)):
                if len(sex_sirs[ix]) != 10:
                    out_message("校验失败", "冻精号{}位数错误".format(sex_sirs[ix]))
                    self.matchfile.config(state=tk.NORMAL, text='生成选配文件')
                    return
        farm = {
            'bred_age':
            int(self.bred_age.get()),
            'sirs':
            com_sirs,
            'sirs_rate':
            (self.com_rate.get().replace(' ', '').strip(',').split(',')),
            'beef_list_file':
            self.beef_list.get(),
            'beef_sirs':
            tuple(self.beef_sirs.get().replace(' ', '').strip(',').split(',')),
            'DC_sexFile_heifer':
            self.sex_file.get(),
            'sex_sirs':
            sex_sirs,
            'sex_sirs_rate':
            (self.sex_rate.get().replace(' ', '').strip(',').split(',')),
            'com_list_file':
            self.com_list.get(),
            'isAdjust':
            self.isAdjust.get(),
            'isFill':
            self.isFill.get()
        }
        logger.debug("执行的命令是：{}.creat_match_file({}, **{})".format(
            self.farm_code, self.com_file.get(), farm))
        try:
            exec(
                "self.farm_code.creat_match_file(self.com_file.get(), **farm)")
        except Exception:
            out_error('错误', traceback.format_exc())
            return
        else:
            logger.info('功能3执行完成。\n')
            out_message("完成", "选配文件已生成！\n具体细节请查看日志")
        finally:
            self.matchfile.config(state=tk.NORMAL, text='生成选配文件')
            args = [
                self.farm.get(),
                self.com_file.get(),
                self.com_sirs.get(),
                self.com_rate.get(),
                self.beef_list.get(),
                self.beef_sirs.get(),
                self.sex_file.get(),
                self.sex_sirs.get(),
                self.sex_rate.get(),
                self.com_list.get(),
                self.bred_age.get()
            ]  # 获取参数
            with open(DB, 'rb') as db:
                dbDic = pickle.load(db)
            matchlog = dbDic['matchlog']  # 取出参数历史
            farms = []
            for ind in range(len(matchlog)):
                farms.append(matchlog[ind][0])
            logger.debug('牧场列表'.format(farms))
            if self.farm.get() in farms:  # 如果牧场存在，则更新记录
                i = farms.index(self.farm.get())
                logger.debug('此牧场在列表中的索引'.format(i))
                matchlog[i] = args
            else:  # 如果牧场不存在，则增加新记录
                matchlog.append(args)
            with open(DB, 'wb') as db:
                pickle.dump(dbDic, db)
            logger.info('本次选配参数已保存（更新）到数据库中')

    # 上次数据按钮事件
    def lastInfo_btnEvent(self, master):
        farms = []
        w = [
            self.v41, self.v42, self.v43, self.v44, self.v45, self.v48,
            self.v49, self.v410, self.v411, self.v413
        ]
        # num = [1, 2, 3, 4, 5, 7, 8, 9, 10, 11]
        with open(DB, 'rb') as db:
            matchlog = pickle.load(db)['matchlog']
        for ind in range(len(matchlog)):
            farms.append(matchlog[ind][0])
        if not self.farm.get():
            tk.messagebox.showinfo('提示', '未选中牧场', parent=master)
        elif self.farm.get() in farms:
            i = farms.index(self.farm.get())
            for v in range(len(w)):
                w[v].set(matchlog[i][v + 1])
        else:
            tk.messagebox.showinfo('提示',
                                   self.farm.get() + '没有上次数据',
                                   parent=master)
            for v in range(len(w)):
                w[v].set("")

    # 工具1: GUI 合并体型
    def merge_body_SubWin1(self, master):
        self.fe5 = tk.Frame(self.fe_5_6)
        self.fe5.pack(side='left', anchor='s', padx=10)

        w = tk.Label(self.fe5, text="\n工具1：合并体型文件", fg='blue', padx=10)
        w.grid(row=0, column=1, columnspan=3, sticky='w')
        w = tk.Label(self.fe5,
                     text="\n  把一个牧场的体型数据文件合并为一个excel文件，" +
                     "同时把新数据追加到以前的数据中。\n")
        w.grid(row=1, column=1, columnspan=3, sticky='w')

        self.lb = tk.Label(self.fe5, text="牧场名：", width=16)
        self.lb.grid(row=2, column=1)
        self.mb_farm1 = ttk.Combobox(
            self.fe5,
            width=55,
            state='readonly',
            postcommand=lambda: self.getFarms(self.mb_farm1))
        self.mb_farm1.grid(row=2, column=2)

        self.lb = tk.Label(self.fe5, text="历史体型明细：", width=16)
        self.lb.grid(row=3, column=1)
        self.last_file = tk.Entry(self.fe5, width=58)
        self.last_file.grid(row=3, column=2)

        self.lb = tk.Label(self.fe5, text="这次做体型日期：", width=16)
        self.lb.grid(row=4, column=1)
        self.this_date = tk.Entry(self.fe5, width=58)
        self.this_date.grid(row=4, column=2)

        # 设置垂直和水平滚动条
        s1 = tk.Scrollbar(self.fe5)
        s1.grid(row=5, column=3, sticky='ns')
        s2 = tk.Scrollbar(self.fe5, orient='horizontal')
        s2.grid(row=6, column=2, sticky='we')

        self.lb = tk.Label(self.fe5, text="这次体型文件名：", width=16)
        self.lb.grid(row=5, column=1)
        self.new_file = tk.Text(self.fe5,
                                width=50,
                                height=6,
                                yscrollcommand=s1.set,
                                xscrollcommand=s2.set,
                                wrap='none')
        self.new_file.grid(row=5, column=2)
        # 激活滚动条
        s1.config(command=self.new_file.yview)
        s2.config(command=self.new_file.xview)

        self.lb = tk.Label(self.fe5,
                           text="""
        使用说明：
          1. 最新体型明细：格式为'牧场_体型明细_202001.csv'(不带引号)，
             保存在本程序目录下的Match_files目录下；
          2. 这次做体型日期：格式为'YYYYMMDD',如'20200115'(不带引号)；
          3. 这次的体型文件名：包括后缀，如有多个文件，按行填写，
             保存在本程序目录下的body_by_month下的old目录下；
          4. 本月汇总数据导出到本程序目录下的body_by_month下；
             （追加新数据后的）最新体型数据导出到Match_files目录下。""",
                           justify='left',
                           fg='blue')
        self.lb.grid(row=7, column=1, columnspan=3)

        self.merge = tk.Button(self.fe5,
                               text="处理体型文件",
                               command=self.mergeBodyBtn)
        self.merge.grid(row=8, column=2, sticky='e')

    def mergeBodyBtn(self):
        r = merge_body(self.mb_farm1.get(), self.last_file.get(),
                       self.this_date.get(), self.new_file.get('0.0', 'end'))
        if r[0] == "0":
            out_message("执行成功", r[1:])
        else:
            out_error("错误", r[1:])

    # 工具2: GUI 提取配种记录
    def extract_log_SubWin2(self, master):
        self.fe6 = tk.Frame(self.fe_5_6)
        self.fe6.pack(side='left', anchor='s')
        # 设置垂直和水平滚动条
        s1 = tk.Scrollbar(self.fe6)
        s1.grid(row=2, column=3, sticky='ns')
        s2 = tk.Scrollbar(self.fe6, orient='horizontal')
        s2.grid(row=3, column=2, sticky='we')
        w = tk.Label(self.fe6, text="\n工具2：提取配种记录", fg='blue', padx=10)
        w.grid(row=0, column=1, columnspan=3, sticky='w')
        w = tk.Label(self.fe6,
                     text="\n此工具用按照特定顺序来从的配种记录文件中提取需要的列，" + "并导出以备后续使用\n",
                     padx=10)
        w.grid(row=1, column=1, columnspan=3)

        self.lb = tk.Label(self.fe6, text="配种记录文件：", width=16)
        self.lb.grid(row=2, column=1)
        self.files2 = tk.Text(self.fe6,
                              width=50,
                              height=6,
                              yscrollcommand=s1.set,
                              xscrollcommand=s2.set,
                              wrap='none')
        self.files2.grid(row=2, column=2)
        # 激活滚动条
        s1.config(command=self.files2.yview)
        s2.config(command=self.files2.xview)

        self.lb = tk.Label(self.fe6,
                           text="""
        使用说明：
            1. 配种记录文件：包括后缀，如有多个文件，按行写；
            2. 配种记录文件必须放到BredLog文件夹下，
               导出的文件也在BredLog文件夹下。""",
                           padx=10,
                           justify='left',
                           fg='blue')
        self.lb.grid(row=6, column=1, columnspan=3)

        self.merge = tk.Button(self.fe6,
                               text="导出配种记录",
                               command=self.extractLogBtn)
        self.merge.grid(row=7, column=2, sticky='e')

    def extractLogBtn(self):
        r = extract_log(self.files2.get('0.0', 'end'))
        if r[0] == "0":
            out_message("执行成功", r[1:])
        else:
            out_error("错误", r[1:])

    # 工具3: GUI 提取定位文件
    def merge_pos_SubWin3(self, master):
        self.fe7 = tk.Frame(master)
        # 设置垂直和水平滚动条
        s1 = tk.Scrollbar(self.fe7)
        s1.grid(row=2, column=3, sticky='ns')
        s2 = tk.Scrollbar(self.fe7, orient='horizontal')
        s2.grid(row=3, column=2, sticky='we')

        w = tk.Label(self.fe7, text="工具3：提取定位文件", fg='blue')
        w.grid(row=0, column=1, columnspan=3, sticky='w')
        self.lb = tk.Label(self.fe7,
                           text="\n此工具用批量来提取GPS的定们文件并按牧场和年份汇总到一起" +
                           "并导出以备后续使用\n",
                           padx=10)
        self.lb.grid(row=1, column=1, columnspan=3)

        self.lb = tk.Label(self.fe7, text="定位文件：", width=16)
        self.lb.grid(row=2, column=1)
        self.files3 = tk.Text(self.fe7,
                              width=50,
                              height=6,
                              yscrollcommand=s1.set,
                              xscrollcommand=s2.set,
                              wrap='none')
        self.files3.grid(row=2, column=2)
        # 激活滚动条
        s1.config(command=self.files3.yview)
        s2.config(command=self.files3.xview)

        self.lb = tk.Label(self.fe7,
                           text="""
        使用说明：
            1. 定位文件：包括后缀，如有多个文件，按行写；
            2. 配种记录文件必须放到Postions文件夹下，
               导出的文件也在Postions文件夹下；
            3. 为了准确识别牧场和年份(分组标记)，文件命名规则如下
                规则：XXX_牧场名_年份.xlsx
                例：GeneticPositioning_bj_2014.xlsx
                （XXX可以随意但不能含有点.和下划线_）""",
                           padx=10,
                           justify='left',
                           fg='blue')
        self.lb.grid(row=6, column=1, columnspan=3)

        self.merge = tk.Button(self.fe7,
                               text="提取定位文件",
                               command=self.mgerPosBtn)
        self.merge.grid(row=7, column=2, sticky='e')

    def mgerPosBtn(self):
        res = mger_pos(self.files3.get('0.0', 'end'))
        if res[0] == "0":
            out_message("执行成功", res[1:])
        else:
            out_error("错误", res[1:])

    # 工具4： 合并选配方案文件
    def merge_files_SubWin4(self, master):
        self.fe11 = tk.Frame(master)

        w = tk.Label(self.fe11, text="工具4：合并选配方案文件", fg='blue')
        w.grid(row=0, column=1, columnspan=3, sticky='w')
        w = tk.Label(self.fe11, text="\n  把一个牧场的选配文件合并为一个excel文件。\n")
        w.grid(row=1, column=1, columnspan=3, sticky='w')

        self.lb = tk.Label(self.fe11, text="文件名前缀：", width=16)
        self.lb.grid(row=3, column=1)
        self.file_name_pre = tk.Entry(self.fe11, width=58)
        self.file_name_pre.grid(row=3, column=2)

        self.lb = tk.Label(self.fe11, text="文件名后缀列表：", width=16)
        self.lb.grid(row=4, column=1)
        self.files_list = tk.Entry(self.fe11, width=58)
        self.files_list.grid(row=4, column=2)

        self.lb = tk.Label(self.fe11,
                           text="""
        使用说明：
          1. 此功能仅用于蚌埠牧场；
          2. 文件名前缀为几个文件中前面固定的内容，到_为止；
          3. 文件名后缀列表为_后面的数字或文字，以逗号（,）分割；
          4. 结果文件保存到原目录，_后为'汇总'两字。""",
                           justify='left',
                           fg='blue')
        self.lb.grid(row=7, column=1, columnspan=3)

        self.mergeFile = tk.Button(self.fe11,
                                   text="合并选配方案",
                                   command=self.mergeFileBtn)
        self.mergeFile.grid(row=8, column=2, sticky='e')

    def mergeFileBtn(self):
        self.mergeFile.configure(state='disable', text='进行中...')
        r = merge_files(
            self.file_name_pre.get(),
            self.files_list.get().replace(' ', '').replace(
                '，', ',').strip(',').split(','))
        if r[0] == "0":
            out_message("执行成功", "合并完成")
        else:
            out_error("错误", r[1:])
        self.mergeFile.configure(state='normal', text='合并选配方案')

    # 核心功能4：GUI 牛群存栏预测
    def forecast_SubWin4(self, master):
        # "sunken"，"raised"，"groove" 或 "ridge"
        # , borderwidth=2, relief="groove"
        # TODO 禁配条件8
        # TODO 流产/淘汰牛的占比，进一步细化流产和淘汰比例
        self.fe8 = tk.Frame(master)  # 主框架
        fe1 = tk.Frame(self.fe8)  # 参配条件
        fe2 = tk.Frame(self.fe8)  # 发情揭发率
        fe3 = tk.Frame(self.fe8)  # 怀孕率
        fe4 = tk.Frame(self.fe8)  # 流产率
        fe5 = tk.Frame(self.fe8)  # 母犊率
        fe6 = tk.Frame(self.fe8)  # 淘汰率
        fe7 = tk.Frame(self.fe8)  # 配种方案
        fe8 = tk.Frame(self.fe8)  # 预测参数

        # 排列次级框架
        w = tk.Label(self.fe8, text="功能4：存栏预测", fg='blue')
        w.grid(row=0, column=10, columnspan=30, sticky='w')
        fe1.grid(row=10, column=10, sticky='n', padx=10)
        fe2.grid(row=10, column=20, sticky='n', padx=10)
        fe3.grid(row=10, column=30, sticky='n', padx=10)
        fe4.grid(row=10, column=40, sticky='n', padx=10)
        fe5.grid(row=10, column=50, sticky='n', padx=10)
        fe6.grid(row=20, column=10, sticky='n', padx=10)
        fe7.grid(row=20, column=20, sticky='n', padx=10)
        fe8.grid(row=20, column=40, columnspan=20, sticky='n', padx=10)

        # fe1 参配条件
        agefb = tk.StringVar(value=13)
        wwtp = tk.StringVar(value=63)
        w = tk.Label(fe1, text="1 参配条件：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcbred = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe1, variable=self.bcbred, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        cbt.configure(state='disable')
        tk.Label(fe1, text="首配月龄：", width=8).grid(row=20, column=10)
        tk.Label(fe1, text="主动停配期：").grid(row=30, column=10)
        self.f_agefb = tk.Entry(fe1, width=14, textvariable=agefb)
        self.f_agefb.grid(row=20, column=20)
        self.f_wwtp = tk.Entry(fe1, width=14, textvariable=wwtp)
        self.f_wwtp.grid(row=30, column=20)

        # fe2 发情揭发率
        hdr0 = tk.StringVar(value=1)  # 0.525  0.75
        hdr1 = tk.StringVar(value=0.65)  # 0.455  0.65
        w = tk.Label(fe2, text="2 发情揭发率：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcHDR = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe2, variable=self.bcHDR, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        cbt.configure(state='disable')
        tk.Label(fe2, text="青年牛：", width=8).grid(row=20, column=10)
        tk.Label(fe2, text="成母牛：").grid(row=30, column=10)
        self.hdrL0 = tk.Entry(fe2, width=14, textvariable=hdr0)
        self.hdrL0.grid(row=20, column=20)
        self.hdrL1 = tk.Entry(fe2, width=14, textvariable=hdr1)
        self.hdrL1.grid(row=30, column=20)

        # fe3 怀孕率
        pr0 = tk.StringVar(value=0.6)
        pr1 = tk.StringVar(value=0.4)
        w = tk.Label(fe3, text="3 怀孕率：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.is_bcPR = tk.IntVar(value=2)  # 1为使用给定参数，2为使用平均配次推算
        cbt = tk.Radiobutton(fe3,
                             text='用以下给定值',
                             variable=self.is_bcPR,
                             value=1)
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        tk.Label(fe3, text="青年牛：", width=8).grid(row=20, column=10)
        tk.Label(fe3, text="成母牛：").grid(row=30, column=10)
        self.prHeifer = tk.Entry(fe3, width=14, textvariable=pr0)
        self.prHeifer.grid(row=20, column=20)
        self.prCow = tk.Entry(fe3, width=14, textvariable=pr1)
        self.prCow.grid(row=30, column=20)

        lact = tk.StringVar(value=1)  # 使用推算参数
        tbrd = tk.StringVar(value=9)
        cbt = tk.Radiobutton(fe3,
                             text='用平均配次推算',
                             variable=self.is_bcPR,
                             value=2)
        cbt.grid(row=40, column=10, sticky='w', columnspan=20)
        tk.Label(fe3, text="最大胎次：", width=10).grid(row=50, column=10)
        tk.Label(fe3, text="最大配次：").grid(row=60, column=10)
        self.useLact = tk.Entry(fe3, width=14, textvariable=lact)
        self.useLact.grid(row=50, column=20)
        self.useTbrd = tk.Entry(fe3, width=14, textvariable=tbrd)
        self.useTbrd.grid(row=60, column=20)
        self.caPRbtn = tk.Button(fe3,
                                 text="查看推算的怀孕率",
                                 command=self.PR_btnEvent)
        self.caPRbtn.grid(row=70,
                          column=10,
                          columnspan=20,
                          sticky='e',
                          pady=10)

        # fe4 流产率
        aboRate = tk.StringVar(value=0.05)
        aboRateC = tk.StringVar(value=0.15)
        w = tk.Label(fe4, text="4 年流产率：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcABR = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe4, variable=self.bcABR, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        cbt.configure(state='disable')
        tk.Label(fe4, text="青年牛：", width=10).grid(row=20, column=10)
        tk.Label(fe4, text="成母牛：", width=10).grid(row=30, column=10)
        self.aboRate = tk.Entry(fe4, width=14, textvariable=aboRate)
        self.aboRate.grid(row=20, column=20)
        self.aboRateC = tk.Entry(fe4, width=14, textvariable=aboRateC)
        self.aboRateC.grid(row=30, column=20)

        # fe5 母犊率
        fmC = tk.StringVar(value=0.47)
        fmS = tk.StringVar(value=0.9)  # 0.88
        fmB = tk.StringVar(value=0.45)
        w = tk.Label(fe5, text="5 母犊率：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcFemalR = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe5, variable=self.bcFemalR, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        tk.Label(fe5, text="常规冻精：", width=8).grid(row=20, column=10)
        tk.Label(fe5, text="性控冻精：").grid(row=30, column=10)
        tk.Label(fe5, text="肉牛冻精：").grid(row=40, column=10)
        self.femalRateC = tk.Entry(fe5, width=14, textvariable=fmC)
        self.femalRateC.grid(row=20, column=20)
        self.femalRateS = tk.Entry(fe5, width=14, textvariable=fmS)
        self.femalRateS.grid(row=30, column=20)
        self.femalRateB = tk.Entry(fe5, width=14, textvariable=fmB)
        self.femalRateB.grid(row=40, column=20)

        # fe6 淘汰率
        calfCullRate = tk.StringVar(value=0.08)
        yongCullRate = tk.StringVar(value=0.1)
        cowCullRate = tk.StringVar(value=0.22)
        w = tk.Label(fe6, text="6 年淘汰率：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcCUR = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe6, variable=self.bcABR, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        cbt.configure(state='disable')
        tk.Label(fe6, text="<首配月龄：", width=10).grid(row=20, column=10)
        tk.Label(fe6, text="可配青年牛：", width=10).grid(row=30, column=10)
        tk.Label(fe6, text="成母牛：", width=10).grid(row=40, column=10)
        self.calfCulRate = tk.Entry(fe6, width=14, textvariable=calfCullRate)
        self.calfCulRate.grid(row=20, column=20)
        self.yongCulRate = tk.Entry(fe6, width=14, textvariable=yongCullRate)
        self.yongCulRate.grid(row=30, column=20)
        self.cowCulRate = tk.Entry(fe6, width=14, textvariable=cowCullRate)
        self.cowCulRate.grid(row=40, column=20)

        # fe7 配种方案
        sexLact = tk.StringVar(value=0)
        sexTbrd = tk.StringVar(value=2)
        beefLact = tk.StringVar(value=1)
        beefRate = tk.StringVar(value=0.3)
        w = tk.Label(fe7, text="7 配种方案：", fg='blue')
        w.grid(row=10, column=10, sticky='w', columnspan=20)
        self.bcBP = tk.IntVar(value=1)
        cbt = tk.Checkbutton(fe7, variable=self.bcBP, text='每个月相同')
        cbt.grid(row=15, column=10, sticky='w', columnspan=20)
        tk.Label(fe7, text="用性控胎次：", width=16).grid(row=20, column=10)
        tk.Label(fe7, text="用性控配次：", width=16).grid(row=30, column=10)
        tk.Label(fe7, text="用肉牛胎次：", width=16).grid(row=40, column=10)
        tk.Label(fe7, text="用肉牛比例：", width=16).grid(row=50, column=10)
        self.sexLact = tk.Entry(fe7, width=8, textvariable=sexLact)
        self.sexLact.grid(row=20, column=20)
        self.sexTbrd = tk.Entry(fe7, width=8, textvariable=sexTbrd)
        self.sexTbrd.grid(row=30, column=20)
        self.beefLact = tk.Entry(fe7, width=8, textvariable=beefLact)
        self.beefLact.grid(row=40, column=20)
        self.beefRate = tk.Entry(fe7, width=8, textvariable=beefRate)
        self.beefRate.grid(row=50, column=20)

        # fe8 预测参数
        fileSuffix = tk.StringVar(value='FARM')
        w = tk.Label(fe8, text="8预测参数：", fg='blue')
        w.grid(row=1, column=1, sticky='w', columnspan=2)
        w = tk.Label(fe8, text="牛群明细日期：", padx=10)
        w.grid(row=2, column=1)
        w = tk.Label(fe8, text="预测结束日期：", padx=10)
        w.grid(row=3, column=1)
        w = tk.Label(fe8, text="参数文件的后缀：", padx=10)
        w.grid(row=4, column=1)

        # 文本框
        self.start = tk.Entry(fe8, width=40)
        self.start.grid(row=2, column=2)
        self.end = tk.Entry(fe8, width=40)
        self.end.grid(row=3, column=2)
        self.suffix = tk.Entry(fe8, width=40, textvariable=fileSuffix)
        self.suffix.grid(row=4, column=2)
        self.forecast_btn = tk.Button(fe8,
                                      text="开始预测(1)",
                                      command=self.startForecast_btnEvent)
        self.forecast_btn.grid(row=6, column=1, sticky='e', pady=10)
        self.forecast2_btn = tk.Button(
            fe8,
            text="开始预测(2)",
            command=lambda: self.startForecast_btnEvent(2))
        self.forecast2_btn.grid(row=6, column=2, sticky='e', pady=10)

    # 用配次推算怀孕率
    def PR_btnEvent(self):
        self.caPRbtn.configure(state='disable', text='推算进行中...')
        lact = self.useLact.get().replace(" ", "")
        tbrd = self.useTbrd.get().replace(" ", "")
        logger.debug("self.is_bcPR.get():".format(self.is_bcPR.get()))
        # if self.is_bcPR.get() != 2:
        #    out_message('警告', '未选定使用推算的方法计算怀孕率')
        #    self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
        #    return
        if self.loaded == 0:
            out_message('警告', '请先导入体型明细')
            self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
            return
        elif lact == "":
            out_message('警告', '最大胎次不能为空')
            self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
            return
        elif not self.start.get().replace(" ", ""):
            out_message('警告', '牛群明细开始日期不可不空')
            self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
            return
        elif lact != "":
            try:
                lt = int(lact)
            except ValueError:
                out_message('警告', '最大胎次只能是数字')
                self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
                return
            else:
                if lt not in range(1, 9):
                    out_message('警告', '最大胎次的范围是1-8')
                    self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')
                    return
        else:
            logger.debug('else')
        try:
            try:
                tbrd = int(tbrd)
            except Exception:
                tbrd = 0
            finally:
                exec("self.farm_code.caPR(lact, tbrd, self.start.get())")
        except Exception:
            out_error("错误", traceback.format_exc())
            return
        finally:
            self.caPRbtn.configure(state='normal', text='查看推算的怀孕率')

    # 设置2：设置牧场需求
    def demand_SubWin5(self, master):
        # 定义窗口
        self.fe9 = tk.Frame(master)
        entryedit = ""
        okb = ""
        off = 10
        columns = ('牧场', '排序', '用肉牛否', '青年用性控否', '前几配次用性控', '成母用性控否', '前几胎用性控',
                   '前几配用性控', '选配文件格式', '配种组文件')
        colWidth = [50, 50, 60, 80, 80, 80, 80, 80, 80, 180]
        treeview = ttk.Treeview(self.fe9,
                                height=40,
                                show="headings",
                                columns=columns)  # 表格 height=40,
        for i in range(len(columns)):  # 显示表头
            treeview.column(columns[i], width=colWidth[i], anchor='center')
            treeview.heading(columns[i], text=columns[i])  # 显示表头
        ttk.Label(self.fe9, text="").pack(side='top')
        treeview.pack(side='left', anchor='nw', padx=off)

        with open(DB, 'rb') as db:
            demand = pickle.load(db)['demand']
        for i in range(len(demand)):  # 写入数据
            treeview.insert('', i, values=tuple(demand[i]))
        rowInt = len(treeview.get_children())

        def treeview_sort_column(tv, col, reverse):  # Treeview、列名、排列方式
            sm = [(tv.set(k, col), k) for k in tv.get_children('')]
            sm.sort(reverse=reverse)  # 排序方式
            # rearrange items in sorted positions
            for index, (val, k) in enumerate(sm):  # 根据排序后索引移动
                tv.move(k, '', index)
            tv.heading(col,
                       command=lambda: treeview_sort_column(
                           tv, col, not reverse))  # 重写标题，使之成为再点倒序的标题

        def set_cell_value(event):  # 双击进入编辑状态
            nonlocal entryedit, okb
            if entryedit:
                entryedit.destroy()
                okb.destroy()
            for item in treeview.selection():
                item_text = treeview.item(item, "values")
            column = treeview.identify_column(event.x)  # 列
            row = treeview.identify_row(event.y)  # 行
            x, y, w, h = treeview.bbox(row, column)
            cn = int(str(column).replace('#', ''))
            entryedit = tk.Text(self.fe9,
                                width=colWidth[cn - 1] // 7,
                                height=1)
            try:
                entryedit.insert('end', item_text[cn - 1])
            except IndexError:
                pass
            entryedit.place(x=x + off, y=y + h)
            entryedit.tag_add("tag1", "0.0", "end")
            entryedit.tag_config("tag1", justify='center')

            def saveedit():
                treeview.set(item,
                             column=column,
                             value=entryedit.get(0.0, "end").split('\n')[0])
                entryedit.destroy()
                okb.destroy()

            okb = ttk.Button(self.fe9, text='确认', width=4, command=saveedit)
            okb.place(x=x + colWidth[cn - 1] - 5 + off, y=y + h - 3)

        def newrow():  # 新建牧场
            t = str(len(treeview.get_children()))
            nt = str(int(t) + 1)
            newdata = ('牧场' + nt, 100 + int(nt), 0, 0, 0, 0, 0, 0, 0,
                       'nothing')
            treeview.insert('', int(t), values=newdata)
            treeview.update()
            newbState(int(nt))

        def save():
            newDemand = []  # 存放修改后的demand数据
            for i in treeview.get_children():
                line = list(treeview.item(i, 'values'))
                newDemand.append(line)
            with open(DB, 'rb') as db:
                dbDic = pickle.load(db)  # db数据字典
            dbDic['demand'] = newDemand  # 把修改后的数据写进db的字典的demand中
            with open(DB, 'wb') as db:
                pickle.dump(dbDic, db)  # 把修改后的数据写进数据库
            tk.messagebox.showinfo('保存成功', '数据保存成功', parent=self.fe9)

        def delete():
            selected_item = treeview.selection()  # get selected item
            if selected_item:
                treeview.delete(selected_item)
            treeview.update()
            newbState(len(treeview.get_children()))

        def newbState(num):
            t = 39
            if num >= t:
                newb.config(state='disable', text='最多可以有' + str(t) + '个牧场')
            else:
                newb.config(state='normal', text='新建牧场')

        treeview.bind('<Double-1>', set_cell_value)  # 双击左键进入编辑
        newb = ttk.Button(self.fe9, text='新建牧场', width=20, command=newrow)
        newb.pack(side='top', anchor='w')
        # newb.grid(row=1, column=2)
        # newb.place(x=sum(colWidth)+20, y=45)
        newbState(rowInt)
        deleBt = ttk.Button(self.fe9, text='删除牧场', width=20, command=delete)
        deleBt.pack(side='top', anchor='w')
        # deleBt.grid(row=2, column=2)
        # deleBt.place(x=sum(colWidth)+20, y=85)
        saveBt = ttk.Button(self.fe9, text='保存数据', width=20, command=save)
        saveBt.pack(side='top', anchor='w')
        # saveBt.grid(row=3, column=2)
        # saveBt.place(x=sum(colWidth)+20, y=125)

        # newb.place(x=120, y=(len(name)-1)*20+45)

        for col in columns:  # 绑定函数，使表头可排序
            treeview.heading(col,
                             text=col,
                             command=lambda _col=col: treeview_sort_column(
                                 treeview, _col, False))

    # 设置3： 维护冻精前缀对应关系
    def semen_SubWin6(self, master):
        self.fe10 = tk.Frame(master)
        entryedit = ""
        okb = ""
        off = 10
        columns = ('常规冻精NAAB号前缀', '性控冻精NAAB号前缀')
        colWidth = [100, 100]
        treeview = ttk.Treeview(self.fe10,
                                height=40,
                                show="headings",
                                columns=columns)  # 表格
        for i in range(len(columns)):  # 显示表头
            treeview.column(columns[i], width=colWidth[i], anchor='center')
            treeview.heading(columns[i], text=columns[i])  # 显示表头
        ttk.Label(self.fe10, text="").pack(side='top')
        treeview.pack(side='left', anchor='nw', padx=off)

        with open(DB, 'rb') as db:
            com2sex = pickle.load(db)['com2sex']
        for i in range(len(com2sex)):  # 写入数据
            treeview.insert('', i, values=tuple(com2sex[i]))
        rowInt = len(treeview.get_children())

        def treeview_sort_column(tv, col, reverse):  # Treeview、列名、排列方式
            sm = [(tv.set(k, col), k) for k in tv.get_children('')]
            sm.sort(reverse=reverse)  # 排序方式
            # rearrange items in sorted positions
            for index, (val, k) in enumerate(sm):  # 根据排序后索引移动
                tv.move(k, '', index)
            tv.heading(col,
                       command=lambda: treeview_sort_column(
                           tv, col, not reverse))  # 重写标题，使之成为再点倒序的标题

        def set_cell_value(event):  # 双击进入编辑状态
            nonlocal entryedit, okb
            if entryedit:
                entryedit.destroy()
                okb.destroy()
            for item in treeview.selection():
                item_text = treeview.item(item, "values")
            column = treeview.identify_column(event.x)  # 列
            row = treeview.identify_row(event.y)  # 行
            x, y, w, h = treeview.bbox(row, column)
            cn = int(str(column).replace('#', ''))
            entryedit = tk.Text(self.fe10,
                                width=colWidth[cn - 1] // 7,
                                height=1)
            try:
                entryedit.insert('end', item_text[cn - 1])
            except IndexError:
                pass
            entryedit.place(x=x + off, y=y + h)
            entryedit.tag_add("tag1", "0.0", "end")
            entryedit.tag_config("tag1", justify='center')

            def saveedit():
                treeview.set(item,
                             column=column,
                             value=entryedit.get(0.0, "end").split('\n')[0])
                entryedit.destroy()
                okb.destroy()

            okb = ttk.Button(self.fe10, text='确认', width=4, command=saveedit)
            okb.place(x=x + colWidth[cn - 1] - 5 + off, y=y + h - 3)

        def newrow():  # 新建牧场
            t = str(len(treeview.get_children()))
            nt = str(int(t) + 1)
            newLine = ('HO', 'HO')
            treeview.insert('', int(t), values=newLine)
            treeview.update()
            newbState(int(nt))

        def save():
            newData = []  # 存放修改后的demand数据
            for i in treeview.get_children():
                line = list(treeview.item(i, 'values'))
                newData.append(line)
            with open(DB, 'rb') as db:
                dbDic = pickle.load(db)  # db数据字典
            dbDic['com2sex'] = newData  # 把修改后的数据写进db的字典的demand中
            with open(DB, 'wb') as db:
                pickle.dump(dbDic, db)  # 把修改后的数据写进数据库
            tk.messagebox.showinfo('保存成功', '数据保存成功', parent=self.fe10)

        def delete():
            selected_item = treeview.selection()  # get selected item
            if selected_item:
                treeview.delete(selected_item)
            treeview.update()
            newbState(len(treeview.get_children()))

        def newbState(num):
            t = 39
            if num >= t:
                newb.config(state='disable', text='最多可以有' + str(t) + '条记录')
            else:
                newb.config(state='normal', text='新增')

        treeview.bind('<Double-1>', set_cell_value)  # 双击左键进入编辑
        newb = tk.Button(self.fe10, text='新增', width=20, command=newrow)
        newb.pack(side='top')
        # newb.place(x=sum(colWidth)+50, y=45)
        newbState(rowInt)
        deleBt = tk.Button(self.fe10, text='删除', width=20, command=delete)
        deleBt.pack(side='top')
        # deleBt.place(x=sum(colWidth)+50, y=85)
        saveBt = tk.Button(self.fe10, text='保存数据', width=20, command=save)
        saveBt.pack(side='top')
        # saveBt.place(x=sum(colWidth)+50, y=125)

        # newb.place(x=120, y=(len(name)-1)*20+45)

        for col in columns:  # 绑定函数，使表头可排序
            treeview.heading(col,
                             text=col,
                             command=lambda _col=col: treeview_sort_column(
                                 treeview, _col, False))

    # 开始预测按钮事件
    def startForecast_btnEvent(self, method=1):
        if self.loaded == 0:
            out_message('警告', '请先导入体型明细')
            return
        if not self.end.get() or not self.start.get():
            out_message("提示", "预测结开始日期或束日期不可为空")
            return
        elif not self.suffix.get():
            out_message("提示", "参数后辍不可为空")
            return
        if method == 1:
            self.forecast_btn.config(state=tk.DISABLED, text='运行中...')
            self.forecast2_btn.config(state=tk.DISABLED)
        elif method == 2:
            self.forecast_btn.config(state=tk.DISABLED)
            self.forecast2_btn.config(state=tk.DISABLED, text='运行中...')
        onOff = [
            self.bcbred.get(),
            self.bcHDR.get(),
            self.is_bcPR.get(),
            self.bcABR.get(),
            self.bcFemalR.get(),
            self.bcCUR.get(),
            self.bcBP.get()
        ]
        try:
            args = {
                'fbage': self.f_agefb.get(),
                'wwtp': self.f_wwtp.get(),
                'hdrH': self.hdrL0.get(),
                'hdrC': self.hdrL1.get(),
                'prH': self.prHeifer.get(),
                'prC': self.prCow.get(),
                'aboRate': self.aboRate.get(),
                'aboRateC': self.aboRateC.get(),
                'femalRateC': self.femalRateC.get(),
                'femalRateS': self.femalRateS.get(),
                'femalRateB': self.femalRateB.get(),
                'calfCulRate': self.calfCulRate.get(),
                'yongCulRate': self.yongCulRate.get(),
                'cowCulRate': self.cowCulRate.get(),
                'sexLact': self.sexLact.get(),
                'sexTbrd': self.sexTbrd.get(),
                'beefLact': self.beefLact.get(),
                'beefRate': self.beefRate.get()
            }
        except Exception:
            out_message("警告", "参数获取错误，请先推算怀孕率")
            logger.debug(traceback.format_exc())
            self.forecast_btn.config(state=tk.NORMAL, text='开始预测(1)')
            self.forecast2_btn.config(state=tk.NORMAL, text='开始预测(2)')
            return
        logger.info("执行的命令是：{}.forecast({},{},{},{},{}, {})".format(
            self.farm_code, self.start.get(), self.end.get(),
            self.suffix.get(), onOff, args, method))
        self.PR_btnEvent()  # 计算怀孕率参数
        try:
            exec("self.farm_code.forecast(self.start.get()," +
                 "self.end.get(), self.suffix.get(), onOff, args, method)")
        except Exception:
            self.forecast_btn.config(state=tk.NORMAL, text='开始预测(1)')
            self.forecast2_btn.config(state=tk.NORMAL, text='开始预测(2)')
            out_message("警告", traceback.format_exc())
            logger.debug(traceback.format_exc())
            return
        else:
            self.forecast_btn.config(state=tk.NORMAL, text='开始预测(1)')
            self.forecast2_btn.config(state=tk.NORMAL, text='开始预测(2)')


def quite():
    ans = tk.messagebox.askyesno(title='确认', message='确定要退出？')
    if ans:
        logger.info('程序退出：')
        ROOT.destroy()


# 如果直接执行，则启动GUI主窗口
if __name__ == '__main__':
    if os.path.exists(DB):
        logger.debug("Yes")
    else:
        logger.debug("NO")
        os.mkdir("Match_files")
        writedb(DB, loaddb())
        logger.debug("DONE")

    # 主窗口
    ROOT = tk.Tk()
    ROOT.title("艾格威技术支持工具包")
    # 设置程序的图标
    picture = open("picture.ico", "wb+")
    picture.write(base64.b64decode(img))
    picture.close()
    ROOT.iconbitmap('picture.ico')
    os.remove("picture.ico")
    ROOT.state("zoomed")  # 窗口最大化
    # ROOT.attributes("-fullscreen", 'True')  # 窗口最大化，但没有任务栏
    app = APP(ROOT)
    # 如果GUI启动，则同时向UI写入日志
    togui = LogtoUi(app.outflow)
    ui = logging.StreamHandler(stream=togui)
    # ui.setFormatter(formatter)
    ui.setLevel(logging.INFO)  # 日志级别: UI中的级别
    logger.addHandler(ui)
    # logger.debug('debug message')
    # logger.warning('warning message')
    # logger.error('error message')
    # logger.critical('critical message')
    logger.info('程序启动成功!')
    logger.info('版本：{}'.format(version))
    ROOT.protocol('WM_DELETE_WINDOW', quite)
    ROOT.tk.mainloop()
