from log import logger
from error import MyException
import datetime
import os
import sys
import time
import traceback
import pickle
import calendar
import math
import platform
from fractions import Fraction
from shutil import copyfile
import numpy as np
import pandas as pd
import pinyin
import xlwings as xw
from random import random
from random import normalvariate
from random import randint

DATE0 = datetime.datetime(1970, 1, 1)
DIR = r"c:\Users\runda\OneDrive - RundaTech\04 工作\0405 艾格威贸易\ALTA_Matching"
YMD = time.strftime("%Y%m%d", time.localtime())
DOUT = sys.stdout
DB = os.path.join(DIR, 'Match_files', 'AltaTools.db')


# 主功能（筛选体型，生成系谱，生成选配文件）
class Farm(object):
    """
    牧场类，处理母牛文件、DC选配文件，根据条件选择用普精和用性控冻精的牛并进行匹配，最后调整冻精比例。

    牧场的配置放到一个文件中去，包括：是否用肉牛，青年牛性控用几配次，成母牛是否用性控，几胎几配次；选配文件需要几列什么类别的。
    """
    def __init__(self, farm_name, cows_file, overload=0):
        self.name = farm_name  # 牧场名字
        self.cows_file = cows_file  # 牛群明细文件名
        self.overload = overload  # 是否重新导入牛群明细
        self.demand_file = DB
        self.file_type = self.cows_file.split('.')[-1].upper()
        logger.debug('牛群明细文件的后缀是：' + self.file_type)
        self.name2 = ".csv"
        if len(cows_file.split("_")) > 3:
            self.name2 = '_' + cows_file.split("_")[3].replace("xlsx", "csv")
        with open(DB, 'rb') as db:
            demand = pickle.load(db)['demand']
        for i in range(len(demand)):
            if demand[i][0] == self.name:
                self.demand = demand[i]
                self.use_beef = int(demand[i][2])
                self.use_sex_heifer = int(demand[i][3])
                self.max_tbrd_heifer = int(demand[i][4])
                self.use_sex_cow = int(demand[i][5])
                self.max_lact_cow = int(demand[i][6])
                self.max_tbrd_cow = int(demand[i][7])
                self.sirs_num = int(demand[i][8])
                self.bred_group_file = demand[i][9]
        logger.debug('牧场的选配参数为：\n  {}'.format(self.demand))
        self.herds = self.import_cowsFile()  # 导入牛群明细
        if self.bred_group_file != 'nothing':
            self.split_file(self.id)

    def __str__(self):
        return pinyin.get_initial(self.name, delimiter="").upper()

    # TODO: 已配牛中有一部分配后天数很大，实际上是空怀牛
    def transRPRO(self, rpro):
        if rpro in ["未配", "尚未配种", '后备牛未配', '出生', '发情未配']:
            return 0
        elif rpro == "禁配":
            return 1
        elif rpro in ["产犊", '产后未配']:
            return 2
        elif rpro in [
                "空怀", "初检未孕", "复检无胎", '妊检（-）', '流产', '初检-', '复检-', '流产未配'
        ]:
            return 3
        elif rpro in ["已配", "已配未检", '输精']:
            return 4
        elif rpro in ["怀孕", "初检已孕", "复检有胎", "初检+", "复检+"]:
            return 5
        elif rpro == "干奶":
            return 6
        elif rpro == "出售":
            return 7
        elif rpro == "公牛":
            return 8
        else:
            return 9

    def creat_rc(self, tofile=0):
        try:
            self.herds.columns.tolist().index('繁殖代码')
        except ValueError:
            logger.info('原文件中没有 繁殖代码(RC) 列')
            try:
                self.herds.columns.tolist().index('繁殖状态')
            except ValueError:
                logger.info("原文件中没有【繁殖状态】列，程序退出！")
                return
            # TODO: 怀孕牛中的干奶牛目前没有修正
            else:
                self.herds['RC'] = self.herds.apply(
                    lambda x: self.transRPRO(x['繁殖状态']), axis=1)
                logger.info("增加【RC】列并从繁殖状态转换")
        else:
            self.herds['RC'] = self.herds['繁殖代码']
        if tofile:
            file = os.path.join(DIR, self.name + '_主要信息_含RC_' + YMD + '.csv')
            self.herds.to_csv(file,
                              index=True,
                              header=True,
                              encoding='utf_8_sig',
                              line_terminator='\r\n')

    def import_cowsFile(self):
        """
        导入牧场发来的牛群明细中的指定列。目前包括一牧云、DC305、丰顿(新希望)/奶业之星(一部分现代牧业)。
        """
        global DIR
        herds = pd.DataFrame()
        herds_file = os.path.join(DIR, 'Match_files', self.cows_file)
        pickleFile = os.path.join(DIR, 'Match_files',
                                  self.cows_file.split('.')[0] + '.pickle')
        self.pickleFile = pickleFile
        # 读取牛群明细文件到 herds
        logger.info('\n\n开始执行第一步：导入牛群明细:{}'.format(self.cows_file))
        # 如果存在pickle文件且未指定重新导入牛群文件，则读取pickle文件
        logger.debug('overload: {}'.format(self.overload))
        if not self.overload and os.path.isfile(pickleFile):
            with open(pickleFile, 'rb') as db:
                herds = pickle.load(db)
            logger.info('读取已存在的pickle文件:{}'.format(pickleFile))
        else:  # 如果不存在pickle文件或指定了要重新导入牛群文件，则导入牛群文件
            logger.debug("文件后缀是：{}".format(self.file_type))
            if self.file_type == 'XLSX':
                logger.debug('文件后缀是XLSX。当前操作系统是：{}'.format(platform.system()))
                if platform.system() == 'Windows':
                    app = xw.App(visible=False, add_book=False)
                    wb = app.books.open(herds_file)
                    sht = wb.sheets[0]
                    rng = sht.used_range  # sht.range('A1').expand()
                    herds = rng.options(pd.DataFrame, header=1,
                                        index=False).value
                    app.quit()
                    logger.debug("使用xlwings读取牛群明细, 有记录：{}条".format(len(herds)))
                else:
                    herds = pd.DataFrame(pd.read_excel(herds_file))
                    logger.debug("使用pandas读取牛群明细, 有记录：{}条".format(len(herds)))
            elif self.file_type == 'CSV':
                logger.debug('文件后缀是: CSV')
                try:
                    herds = pd.DataFrame(
                        pd.read_csv(herds_file, encoding='gbk'))
                    logger.debug('encoding为gbk')
                except UnicodeDecodeError:
                    herds = pd.DataFrame(
                        pd.read_csv(herds_file, encoding='utf_8_sig'))
            logger.info('从原始文件中导入数据,原始文件是：{}'.format(herds_file))
            #  保存pickle数据
            with open(pickleFile, 'wb') as db:
                pickle.dump(herds, db)
                logger.info('pickle文件已保存：{}'.format(pickleFile))
        logger.info('已导入所有列')

        # 选择需要的列
        with open(DB, 'rb') as db:
            colums = pickle.load(db)['colums']
        herds_colum = herds.columns.tolist()
        cls = [
            '牛号', '月龄', '胎次', '繁殖状态', '泌乳天数', '产后天数', '空怀天数', '配种次数', '奶量1',
            '奶量2', '牛舍名称', '出生日期', '父亲牛号', '外祖父号', '繁殖代码', '怀孕日期', '产犊日期',
            '干奶日期', '配种日期', '流产日期', '与配公牛'
        ]
        herds1 = pd.DataFrame()
        logger.info('开始提取列...\n  希望提取的列：{}'.format(str(cls)))
        for index in range(len(colums)):
            for it in colums[index]:
                if it in herds_colum:
                    logger.info('  已提取到列:{}，文件中是：{}'.format(cls[index], it))
                    if it == "AID":
                        herds1[cls[index]] = herds[it].apply(
                            lambda x: x / 30.5)
                        logger.info("    日龄AID转换为月龄")
                    else:
                        herds1[cls[index]] = herds[it]
                        if index == 0:
                            self.id = it
                    break
            else:
                logger.info('未提取到列:{}, 文件中未找到{}'.format(
                    cls[index], colums[index]))
                herds1[cls[index]+"_"] = ""
        logger.info('提取完成。')

        # 去除重复耳号及空耳号
        herds1.replace('', -1, inplace=True)
        herds1.replace('-         ', -1, inplace=True)
        herds1.replace('       - ', -1, inplace=True)
        herds1.fillna(-1, inplace=True)
        logger.debug(herds1.dtypes)
        # logger.debug(herds1)
        logger.debug("herds1 中有多少头{}".format(len(herds1)))
        dup_num = len(herds1) - len(herds1.drop_duplicates('牛号', 'first'))
        blk_num = len(herds1[herds1['牛号'] == -1])
        if dup_num > 0:
            logger.info("去除重复耳号记录有：{}条".format(dup_num))
            herds1.drop_duplicates('牛号', 'first', inplace=True)
        if blk_num > 0:
            herds1 = herds1[herds1['牛号'] != -1]
            logger.info('去除重空耳号记录有：{}条'.format(blk_num))

        # 转换日期格式
        dateCols0 = ['出生日期', '怀孕日期', '产犊日期', '干奶日期', '配种日期', '流产日期']
        dateCols = [it for it in dateCols0 if it in herds1.columns]
        logger.info('开始处理日期格式...转换日期格式为"【月/日/年】')
        for i in dateCols:
            herds1[i].replace(-1, DATE0, inplace=True)

            logger.info('  【{}】当前格式是：{},{}'.format(i, herds1[i][0],
                                                   type(herds1[i][0])))
            try:
                herds1[i] = pd.to_datetime(herds1[i])
                logger.info('    --第1步：转换为日期类型。')
            except ValueError:
                herds1[i] = pd.to_datetime(herds1[i], errors="coerce")
                herds1.loc[herds1[i].isnull(), i] = DATE0
                logger.debug(traceback.format_exc())
            except TypeError:
                logger.error('    --出现 TypeError 错误')
                logger.debug(traceback.format_exc())
            except Exception:
                logger.error('    --出现其他错误')
                logger.debug(traceback.format_exc())
            finally:
                herds1[i] = herds1[i].apply(lambda x: x.strftime('%m/%d/%Y'))
                logger.info('    --{}格式处理完成。现在的格式为：【{}】，{}'.format(
                    i, herds1[i][0], type(herds1[i][0])))
        logger.debug(herds1.dtypes)

        # 配种组处理, 目前只有蚌埠需要，根据牛棚号区分配种组
        logger.info('添加配种组列, 默认全为0')
        herds1['配种组'] = 0
        if self.bred_group_file != 'nothing':
            logger.debug('配种组设置文件为：{}'.format(self.bred_group_file))
            herds1['大舍'] = herds1.apply(lambda x: int(x['牛舍名称'][0:2]), axis=1)
            pen_group_file = os.path.join(DIR, 'Match_files',
                                          self.bred_group_file)
            pen_group = pd.DataFrame(pd.read_csv(pen_group_file))
            herds1 = pd.merge(herds1, pen_group, how='left', on='大舍')
            herds1['配种组'] = herds1.apply(lambda x: 5
                                         if x['胎次'] == 0 else x['配种组1'],
                                         axis=1)
            logger.info('配种组处理完成，配种组列下的数字为组别。')
        herds1 = herds1.fillna(0)  # 缺失值替换为0
        herds1['胎次'] = herds1['胎次'].astype('int')
        herds1['配种组'] = herds1['配种组'].astype('int')

        result_file = os.path.join(DIR, self.name + '_主要信息_' + YMD + '.csv')
        logger.info('文件中共有牛头数：【{}】'.format(len(herds1)))
        logger.info('数据保存中...')
        herds1.to_csv(result_file,
                      index=True,
                      header=True,
                      encoding='utf_8_sig',
                      line_terminator='\r\n')

        logger.info('各列数据类型转换中...')
        co = ['牛号', '月龄', '胎次', '产后天数', '配种次数', '泌乳天数', '空怀天数', '奶量1', '奶量2']
        co = [item for item in co if item in herds1.columns]
        for i in co:
            try:
                herds1[i] = herds1[i].apply(int)
            except Exception:
                try:
                    herds1[i] = herds1[i].apply(float)
                except Exception:
                    herds1[i] = herds1[i].apply(str)
                    logger.debug(i)
                continue
        logger.debug(herds1.dtypes)
        logger.info('第一步执行完成，提取的数据保存在：{}\n'.format(result_file))
        return herds1

    def split_file(self, id='牛号'):
        cats = self.herds['配种组'].unique()
        cats.sort()
        with open(self.pickleFile, 'rb') as db:
            herds = pickle.load(db)
        try:
            herds[id] = herds[id].astype(int)
            logger.info("转换为int")
        except Exception:
            herds[id] = herds[id].astype(str)
        data = pd.merge(herds,
                        self.herds[['牛号', '配种组']],
                        how='left',
                        left_on=id,
                        right_on='牛号')
        logger.info("cats:{}".format(cats))
        for cat in cats:
            name = self.cows_file.replace('.', '_{}.'.format(cat))
            toFile = os.path.join(DIR, 'Match_files', name)
            df = data[data['配种组'] == cat]
            df.to_excel(toFile, index=False, encoding='utf8')
            logger.info("保存完成：{}".format(toFile))
        logger.info("处理完成")

    def creat_match_file(self, dc_com_file, **kwargs):
        """
        生成选配方案并调整公牛比例的函数。

        Args:
            dc_com_file: 常规冻精的DC305选配文件，来源于GPS生成。
            **kwargs:
                self.use_sex_cow：=0 表示成母牛不用性控，=1表示成母牛用性控。
                    self.max_lact_cow：成母牛中用性控的最大胎次。
                    self.max_tbrd_cow：成母牛中用性控的最大配次。
                DC_sexFile_cow：成母牛用性控的单独的DC305选配文件，来源于GPS生成，如果没有单独的，可以把用性控的生成统一的
                                一个文件和青年牛放一起。
                DC_sexFile_heifer: 表年牛用性控的DC305选配文件，来源于GPS生成。
                self.use_sex_heifer: 青年年是否用性控。
                    self.max_tbrd_heifer：青年牛用性控的最大配次
                beefFile: 用肉牛列表，来源于excel模型挑选。【后期考虑用Python来实现】
                beefSirs：元组；肉牛公牛列表，可以有多个。
                bred_age： 配种月龄，大于此月龄的牛才匹配公牛。
                sirs: 常规冻精列表，按比例从小到到给定。
                sirs_rate: 常规冻精期望的使用比例，一般为库存的比例，接受小数和分数。
                sex_sirs: 性控冻精列表，按比例从小到到给定。
                sex_sirs_rate: 性控冻精期望的使用比例，一般为库存的比例，接受小数和分数。

        Returns:
            None: 生成配选文件并保存。
        """
        logger.info('\n\n开始执行功能3：生成选配文件')
        logger.info('参数是\ndc_com_file:{}\n{}'.format(dc_com_file, kwargs))
        global DIR
        global YMD
        try:
            self.mf = self.herds[[
                '牛号', '月龄', '胎次', '繁殖状态', '产后天数', '配种次数', '配种组'
            ]].copy()
            logger.info('3.1 从导入的【牛群明细】文件中获取需要的列:\n  {}'.format(
                '牛号,月龄,胎次,繁殖状态,产后天数,配种次数,配种组'))
            logger.debug('牛号等列转换为int后的类型')
            logger.debug(self.mf.dtypes)
        except Exception:
            logger.error('错误:\n'.format(traceback.format_exc()))
            return
        sex_sirs_cow = pd.DataFrame({
            '牛号': [1],
            '性控1': [''],
            '性控2': [''],
            '性控3': ['']
        })
        sex_sirs_heifer = pd.DataFrame({
            '牛号': [1],
            '性控1': [''],
            '性控2': [''],
            '性控3': ['']
        })
        beef_list = pd.DataFrame({'牛号': [1], '用肉牛': [0]})
        sirs = ['普精1', '普精2', '普精3']

        # 计算是否用普精的函数
        def use_com(age, beef, bred_age=11):
            """计算是否用普精的函数"""
            if int(age) >= bred_age and beef == 0:
                return 1
            else:
                return 0

        # 计算是否用性控的函数（青年牛）
        def h_use_sex(lact, tbrd, age, beef, h_sex_tbrd, bred_age=11):
            """计算是否用性控的函数（青年牛）"""
            if (lact == 0 and tbrd <= (h_sex_tbrd - 1) and age >= bred_age
                    and beef == 0):
                return 1
            else:
                return 0

        # 计算是否用性控的函数（成母牛）
        def c_use_sex(lact, tbrd, beef, c_sex_tbrd, c_sex_lact):
            """计算是否用性控的函数（成母牛）"""
            if (0 < lact <= c_sex_lact and tbrd <= (c_sex_tbrd - 1)
                    and beef == 0):
                return 1
            else:
                return 0

        # 修正2选
        def adjust2(oldsirs, sir1, code):
            old = list(oldsirs)
            old.remove(sir1)
            if code > 4:
                return old[0]

        # 修正3选
        def adjust3(oldsirs, sir1, code):
            old = oldsirs.copy()
            old.remove(sir1)
            if code > 4:
                return old[1]

        # 为牛只分组
        def sirs_group(use_for, lact, age, days_after_fresh):  # bredgroup
            # use_for 传入 用普精或 用性控 列。
            """用普精的分组的函数"""
            if use_for == 1:
                if lact == 0:
                    return int(1000 * age / 0.25)  # 值大于 40000 的为青年牛。
                elif lact > 0:
                    return int(days_after_fresh / 7)  # 值小于40000 的为成母牛
                else:
                    return -1
            else:
                return -1

        # 计算可用冻精位置
        def available_sirs_position(sir_1, sir_2, sir_3, one_sir):
            """标记可用公牛位置的函数：1表示这头公牛是1选，2是2选，3是3选。0表示不可用"""
            try:
                return [sir_1, sir_2, sir_3].index(one_sir) + 1
            except Exception:
                return 0

        # 计算可用冻精代码
        def available_sirs_code(st_sir, nd_sir, rd_sir, *sir_list):
            """标记母牛可用公牛识别码：0表示没有任何冻精，2表示只有第1个冻精，3表示只有第2个，
                4表示只有第3个冻精。5表示有第1和第2个冻精，以此类推"""
            v = [0, 0, 0]
            for j in list(range(len(sir_list))):
                if sir_list[j] in [st_sir, nd_sir, rd_sir]:
                    v[j] = j + 2
            return sum(v)

        # 公牛和比例对应，且按比例从小到大给定公牛顺序
        # 调整冻精比例的主函数
        def adjustRate(sirs, sirsRate, sirsKind):
            """调整冻精比例的主函数"""
            sirsNum = len(sirs)
            if sirsNum == 1:
                logger.info('----冻精数量为1, 不调整')
                return
            sir1, sir2, sir3, groupName = '', '', '', ''
            if sirsKind == '用普精':
                sir1, sir2, sir3, groupName = '普精1', '普精2', '普精3', '普精分组'
            elif sirsKind == '用性控':
                sir1, sir2, sir3, groupName = '性控1', '性控2', '性控3', '性控分组'

            # 处理没有匹配上冻精的牛
            if not kwargs['isFill']:
                logger.info('----1.1填充匹配不上的牛的1选')
                con = (self.mf[sirsKind] == 1) & (~self.mf[sir1].isin(sirs))
                sir_zero_num = self.mf[con]['牛号'].count()
                logger.info('------1计算未匹配上{}的1选的牛头数:【{}】'.format(
                    sirsKind, sir_zero_num))
                if sir_zero_num > 0:
                    self.mf.loc[self.mf.loc[con][sir1].index.values.tolist(
                    ), '未匹配_' + sir1] = 1
                    logger.info('------2标记未匹配上{}的1选的牛，用列:未匹配_{}'.format(
                        sirsKind, sir1))
                    s1_fill_num = int(sir_zero_num / sirsNum)
                    yu = sir_zero_num % sirsNum
                    logger.debug('yu:{}, s1_fill_num:{}'.format(
                        yu, s1_fill_num))
                    for s in list(range(sirsNum)):
                        con = ((self.mf[sirsKind] == 1) &
                               (~self.mf[sir1].isin(sirs)))
                        self.mf.loc[self.mf.loc[con][sir1].sample(
                            frac=1)[:s1_fill_num +
                                    yu].index.values.tolist(), sir1] = sirs[s]
                        yu = 0
                    logger.info('------3平均分配给可用冻精')
            else:
                logger.info('----1.0不填充匹配不上的牛的1选')
                logger.debug('isFill:{}'.format(kwargs['isFill']))

            # 如果不调整比例，直接退出
            if kwargs['isAdjust']:
                logger.info('----2.0不调整比例')
                logger.debug('isAdjust:{}'.format(kwargs['isAdjust']))
                return

            logger.info('----2.1开始调整冻精比例')
            self.mf[sirsKind + '_原冻精'] = self.mf.apply(
                lambda x: [x[sir1], x[sir2], x[sir3]], axis=1)
            logger.info('------1增加 "{}_原冻精" 列'.format(sirsKind))

            # 增加可用冻精列，并设置冻精位置，1表示这头公牛是1选，2是2选，3是3选
            for s in sirs:
                self.mf[s] = self.mf.apply(lambda x: available_sirs_position(
                    x[sir1], x[sir2], x[sir3], s),
                                           axis=1)  # 标记可用冻精的位置
                logger.info('------2增加{}列, 其下的数字表示其是几选'.format(s))
            self.mf['available_sirs_code_' + sir1] = self.mf.apply(
                lambda x: available_sirs_code(x[sir1], x[sir2], x[sir3],
                                              *tuple(sirs)),
                axis=1)
            des = ("标记母牛可用公牛识别码：0表示没有任何冻精，2表示只有第1个冻精，" +
                   "3表示只有第2个, 4表示只有第3个冻精。5表示有第1和第2个冻精，以此类推")
            logger.info('------3增加列：available_sirs_code_{}'.format(sir1))

            # 调整1选冻精比例并修正2选及3选
            self.mf['已调整' + sir1] = 0
            fenmu = self.mf[self.mf[sirsKind] == 1][sir1].count()
            for i in range(len(sirs)):
                fenzi = self.mf[(self.mf[sirsKind] == 1)
                                & (self.mf[sir1] == sirs[i])][sir1].count()
                rate = format(fenzi / fenmu, '.2%')
                logger.info('------4调整冻精前:{}的占比为：{}'.format(sirs[i], rate))

            self.mf[groupName] = self.mf.apply(
                lambda x: sirs_group(x[sirsKind], x.胎次, x.月龄, x.产后天数),
                axis=1)  # 普精分组
            sm = '青年牛按月龄,每1/4月为一组; 成母牛按产后天数，每7天为一组。'
            logger.info('------5给【{}】分组。说明:{}'.format(sirsKind, sm))
            group_list = self.mf[groupName].unique().tolist()  # 获取分组后的列表
            logger.info('------6获取分组后的组别列表')
            logger.info('------7按组别开始调整冻精比例......')
            groupNum = len(group_list)
            logger.debug('组别数:{}'.format(groupNum))
            count = 0
            limit = 25
            jiange = 25
            for item in group_list:
                # 按设定的间隔打印进度
                count += 1
                proce = int(100 * count / groupNum)
                if proce >= limit:
                    logger.info('-------已处理{}'.format(str(proce) + '%'))
                    limit += jiange

                if item == -1:
                    continue
                # logger.debug('-------7.1开始处理{}组...'.format(item))
                # 一共需要匹配普精1的牛头数
                total_sirs_num = self.mf[self.mf[groupName] ==
                                         item][sir1].count()
                if total_sirs_num < sirsNum:
                    # logger.debug('-----------7.1.1本组牛头数({})小于冻精个数{}，不调整'.
                    #             format(total_sirs_num, sirsNum))
                    continue
                sirs_num_hope = [None, None, None]  # 本组中，给定各冻精期望匹配1选的牛头数
                # 本组中，原本只有1选是给定冻精且没有2选和3选的牛头数
                sirs_num_only1 = [None, None, None]
                hope_only1 = [None, None, None]  # 前2者的差，是需要调整的牛头数
                for s in list(range(sirsNum)):
                    sirs_num_hope[s] = int(total_sirs_num *
                                           Fraction(sirsRate[s]))
                    # logger.debug('-----------本组中，{}期望匹配1选的牛头数:{}'.
                    #             format(sirsRate[s], sirs_num_hope[s]))
                    sirs_num_only1[s] = (
                        self.mf[(self.mf[groupName] == item)
                                & (self.mf[sir1] == sirs[s]) &
                                (self.mf['available_sirs_code_' + sir1] <
                                 5)][sir1].count())
                    # logger.debug('-----------本组中，只有1选且1选是{}的牛头数：{}'.
                    #             format(sirs[s], sirs_num_only1[s]))
                    hope_only1[s] = sirs_num_hope[s] - sirs_num_only1[s]
                    # logger.debug('-----------本组中，{}需要调整的头牛头数是{}'.
                    #             format(sirs[s], hope_only1[s]))

                    if hope_only1[s] < 0:
                        self.mf.loc[self.mf[(self.mf[groupName] == item)
                                            & (self.mf[sir1] == sirs[s]) &
                                            (self.mf['available_sirs_code_' +
                                                     sir1] < 5)].index.values.
                                    tolist(), 'error_' + sirsKind] = 1
                        # logger.debug('-----------本组中，{}无法调整,增加列：error_{}, ' +
                        #             '值记为1以标记'.format(sirs[s],
                        #                              sirsKind))
                        continue
                    elif hope_only1[s] == 0:
                        # logger.debug('-----------本组中，{}不需要调整'.
                        #             format(sirs[s]))
                        continue
                    elif hope_only1[s] > 0:
                        # 设置需要的数量的公牛,已调整的做上标记
                        if s == sirsNum - 1:
                            self.mf.loc[self.mf[
                                (self.mf[sirs[s]] > 0)
                                & (self.mf[groupName] == item) &
                                (self.mf['available_sirs_code_' + sir1] > 4) &
                                (self.mf['已调整' + sir1] == 0)].index.values.
                                        tolist()[:hope_only1[s] +
                                                 3], [sir1, '已调整' +
                                                      sir1]] = [sirs[s], 1]
                            # logger.debug('-----------本组中{}头牛的{}调整为{}'.
                            #             format(hope_only1[s] + 3, sir1,
                            #                    sirs[s]))
                        else:
                            self.mf.loc[self.mf[
                                (self.mf[sirs[s]] > 0)
                                & (self.mf[groupName] == item) &
                                (self.mf['available_sirs_code_' + sir1] > 4) &
                                (self.mf['已调整' + sir1] == 0
                                 )].index.values.tolist(
                                 )[:hope_only1[s]], [sir1, '已调整' +
                                                     sir1]] = [sirs[s], 1]
                            # logger.debug('-----------本组中{}头牛的{}调整为{}'.
                            #             format(hope_only1[s], sir1,
                            #                    sirs[s]))

            # 修正2选和3选冻精
            self.mf[sir2] = self.mf.apply(
                lambda x: adjust2(x[sirsKind + '_原冻精'], x[sir1], x[
                    'available_sirs_code_' + sir1]),
                axis=1)
            logger.info('------8修正2选冻精'.format(sir2))
            self.mf[sir3] = self.mf.apply(
                lambda x: adjust3(x[sirsKind + '_原冻精'], x[sir1], x[
                    'available_sirs_code_' + sir1]),
                axis=1)
            logger.info('------8修正3选冻精'.format(sir3))
            for i in range(len(sirs)):
                r1 = (self.mf[(self.mf[sirsKind] == 1)
                              & (self.mf[sir1] == sirs[i])][sirs[i]].count() /
                      self.mf[self.mf[sirsKind] == 1][sirs[i]].count())
                ra1 = str('%.2f%%' % (r1 * 100))
                r2 = 0.0 + Fraction(sirsRate[i])
                ra2 = str('%.2f%%' % (r2 * 100))
                dif = str('%.2f%%' % (abs(r1 - r2) * 100))
                logger.info('------9调整之后 {}的占比为: {}. 期望的占比为：{}. 差异为：{}'.format(
                    sirs[i], ra1, ra2, dif))

        # 导入要用性控冻精、常规冻精的选配文件
        sex_sirs = pd.DataFrame()
        kwargs['DC_sexFile_cow'] = ''
        if self.use_sex_cow and kwargs['DC_sexFile_cow']:
            sex_sirs_cow = pd.DataFrame(
                pd.read_csv(os.path.join(DIR, 'AltaGPS_Reports',
                                         kwargs['DC_sexFile_cow']),
                            header=None,
                            names=['牛号', '性控1', '性控2', '性控3']))
            logger.info('导入性控冻精选配文件-成母牛（GPS产生的）')
        if self.use_sex_heifer and kwargs['DC_sexFile_heifer']:
            sex_sirs_heifer = pd.DataFrame(
                pd.read_csv(os.path.join(DIR, 'AltaGPS_Reports',
                                         kwargs['DC_sexFile_heifer']),
                            header=None,
                            names=['牛号', '性控1', '性控2', '性控3']))
            logger.info('导入性控冻精选配文件-青年牛（GPS产生的）')
        sex_sirs = sex_sirs_heifer.append(sex_sirs_cow)
        logger.info('3.2 导入【性控】冻精选配文件（GPS产生的）')
        common_sirs = pd.DataFrame(
            pd.read_csv(os.path.join(DIR, 'AltaGPS_Reports', dc_com_file),
                        header=None,
                        names=['牛号', '普精1', '普精2', '普精3']))
        logger.info('3.3 导入【常规】冻精选配文件（GPS产生的）')

        # 导入常规用肉牛列表，性控用常规冻精列表
        if self.use_beef and kwargs['beef_list_file']:
            beef_list = pd.DataFrame(
                pd.read_csv(
                    os.path.join(DIR, 'Match_files',
                                 kwargs['beef_list_file'])))
            logger.info('3.4 导入常规用【肉牛】列表文件')
        try:
            self.mf = pd.merge(self.mf, beef_list, how='left')  # 连接肉牛
        except ValueError:
            logger.error('错误:\n' + '运行出错。可能的原因是牛群列表中的耳号列有空值或含有字母\n{}'.format(
                traceback.format_exc()))
            return
        if kwargs['com_list_file']:  # 性控用常规
            com_list = pd.DataFrame(
                pd.read_csv(
                    os.path.join(DIR, 'Match_files', kwargs['com_list_file'])))
            self.mf = pd.merge(self.mf, com_list, how='left')
            logger.info('3.5 导入性控用【常规】列表文件')

        # 添加辅助列
        self.mf = self.mf.fillna(0)
        logger.debug(self.mf.dtypes)
        self.mf['用普精'] = self.mf.apply(
            lambda x: use_com(x.月龄, x.用肉牛, kwargs['bred_age']), axis=1)
        logger.info('3.6 添加"用普精"辅助列')
        self.mf['用性控_青年牛'] = self.mf.apply(
            lambda x: h_use_sex(x.胎次, x.配种次数, x.月龄, x.用肉牛, self.
                                max_tbrd_heifer, kwargs['bred_age']),
            axis=1)
        logger.info('3.7 添加"用性控_青年牛"辅助列')
        self.mf['用性控_成母牛'] = self.mf.apply(lambda x: c_use_sex(
            x.胎次, x.配种次数, x.用肉牛, self.max_tbrd_cow, self.max_lact_cow),
                                           axis=1)
        logger.info('3.8 添加"用性控_成母牛"辅助列')
        self.mf['用性控'] = (self.mf['用性控_青年牛'] + self.mf['用性控_成母牛'])
        logger.info('3.9 添加"用性控"辅助列(青年牛和成母牛合并)')
        if kwargs['com_list_file']:
            self.mf['用性控'] = self.mf.apply(lambda x: 0
                                           if x['用常规'] == 1 else x['用性控'],
                                           axis=1)
            logger.info("3.10 修改【性控】用【常规】的冻精")

        # logger.debug(common_sirs)
        self.mf = pd.merge(self.mf, common_sirs, how='left')
        logger.info('3.11 链接【常规选配方案】')
        self.mf = pd.merge(self.mf, sex_sirs, how='left')
        logger.info('3.12 链接【性控选配方案】')

        if self.use_beef and kwargs['beef_list_file']:  # 填充肉牛冻精号
            self.mf.loc[self.mf.用肉牛 == 1, ['普精2', '普精3']] = ''  # 清空用肉牛冻精的公牛号
            for i in list(range(len(kwargs['beef_sirs']))):
                self.mf.loc[self.mf.用肉牛 == 1, sirs[i]] = kwargs['beef_sirs'][i]
            logger.info('3.13 修改【常规选配方案】：用肉牛冻精的牛的2选和3选清空，1选填充为肉牛冻精号')

        # 替换性控冻精的前缀
        with open(DB, 'rb') as db:
            com2sex = pickle.load(db)['com2sex']
        for i in range(len(com2sex)):
            self.mf['性控1'] = self.mf['性控1'].str.replace(
                com2sex[i][0], com2sex[i][1])
            self.mf['性控2'] = self.mf['性控2'].str.replace(
                com2sex[i][0], com2sex[i][1])
            self.mf['性控3'] = self.mf['性控3'].str.replace(
                com2sex[i][0], com2sex[i][1])
        logger.info('3.14 替换性控冻精的前缀（常规/性控冻精对照表见设置菜单栏）')
        # 去掉用肉牛且不用性控的牛的性控冻精
        try:
            self.mf.loc[(self.mf.用性控 == 0) &
                        (self.mf.用肉牛 == 0), ['性控1', '性控2', '性控3']] = ''
        except Exception:
            logger.info('错误:\n{}'.format(traceback.format_exc()))
            return
        logger.info('3.15 去掉不用肉牛且不用性控的牛的性控方案')
        self.mf.loc[(self.mf.用普精 == 0) &
                    (self.mf.用肉牛 == 0), ['普精1', '普精2', '普精3']] = ''
        logger.info('3.16 去掉不用肉牛且不用普精的牛的普精方案')
        # 处理空格冻精 为 None
        for i in ['普精1', '普精2', '普精3', '性控1', '性控2', '性控3']:
            self.mf[i] = self.mf[i].apply(lambda x: ''
                                          if str(x).isspace() else x)
        self.mf = self.mf.where(self.mf.notna(), '')
        logger.info('3.17 处理【空格冻精】为空')

        # 常规冻精和性控冻精分别处理
        logger.info('3.18 开始处理冻精：')
        sirsKind = ['用普精']
        ss = [kwargs['sirs']]
        sr = [kwargs['sirs_rate']]
        if kwargs['sex_sirs']:
            sirsKind.append('用性控')
            ss.append(kwargs['sex_sirs'])
            sr.append(kwargs['sex_sirs_rate'])
        for i in range(len(sirsKind)):
            logger.info('  开始处理【{}】冻精'.format(sirsKind[i]))
            adjustRate(ss[i], sr[i], sirsKind[i])
            logger.info('  调整【{}】冻精完成'.format(sirsKind[i]))
        logger.info('3.18 冻精比例调整完成')

        # 6选处理为3选
        if self.sirs_num == 30:  # 30 用性控的1选和2选用性控，3选为普精1选
            self.mf['1选'] = self.mf.apply(lambda x: x['普精1']
                                          if x['用性控'] == 0 else x['性控1'],
                                          axis=1)
            self.mf['2选'] = self.mf.apply(lambda x: x['普精2']
                                          if x['用性控'] == 0 else x['性控2'],
                                          axis=1)
            self.mf['3选'] = self.mf.apply(lambda x: x['普精3']
                                          if x['用性控'] == 0 else x['普精1'],
                                          axis=1)
            logger.info('3.19：处理完成：性控的1选和2选用性控，3选为普精1选。')
        elif self.sirs_num == 31:  # 31 用性控的，配次0的，1、2选用性迭，配次1的，1选用性控
            self.mf['1选'] = self.mf.apply(lambda x: x['普精1']
                                          if x['用性控'] == 0 else x['性控1'],
                                          axis=1)
            self.mf['2选'] = self.mf.apply(lambda x: x['普精2']
                                          if x['用性控'] == 0 else x['性控2']
                                          if x['配种次数'] == 0 else x['普精1'],
                                          axis=1)
            self.mf['3选'] = self.mf.apply(lambda x: x['普精3']
                                          if x['用性控'] == 0 else x['普精1']
                                          if x['配种次数'] == 0 else x['普精2'],
                                          axis=1)
            logger.info('3.19：处理完成：用性控的，配次0的，1、2选用性迭，配次1的，1选用性控。')
        elif self.sirs_num == 6:  # 6  不处理为3个选择
            logger.info('3.19：不处理为3个冻精选择')
        else:
            logger.info('3.19：参数错误!!!!!!!!!!!!!!!!!!!!!!!')
            logger.info(self.sirs_num)

        # 保存到excel
        logger.info('3.20：正在保存选配文件...')
        result_file = os.path.join(DIR, '选配文件',
                                   self.name + '_选配文件_' + YMD + '.xlsx')
        self.mf.to_excel(result_file, index=True, encoding='utf8')
        rate = self.mf['用肉牛'].sum() / self.mf['用肉牛'].count()
        logger.info('3.21 肉牛比例为：{}'.format(str('%.2f%%' % (rate * 100))))
        logger.info('3.22冻精匹配处理完成，【选配文件】保存在：{}'.format(result_file))

        # 将选配方案追加到原文件中
        if self.file_type == 'XLSX':
            logger.info('3.23 开始将冻精追加到原文件')
            herdsFile = os.path.join(DIR, 'Match_files', self.cows_file)
            logger.info('----1全群明细复制到 "选配文件" 文件夹下')
            matchFile = os.path.join(DIR, '选配文件',
                                     YMD + '_选配方案_' + self.name + '.xlsx')
            copyfile(herdsFile, matchFile)
            if self.sirs_num == 6:
                self.case = self.mf[['普精1', '普精2', '普精3', '性控1', '性控2',
                                     '性控3']].copy()
            else:
                self.case = self.mf[['1选', '2选', '3选']].copy()
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(matchFile)
            sht = wb.sheets[0]
            logger.info('----2选配方案追加到原牛群明细文件中...')
            colNum = sht.range('A1').expand().columns.count
            sht.range(1, colNum + 1).value = self.case
            sht.range(1, colNum + 1).api.EntireColumn.Delete()
            wb.save()
            wb.close()
            app.quit()
            logger.info('3.24【选配方案】追加完成，并保存在：{}'.format(matchFile))
        else:
            logger.info("3.23 开始生成CSV选配文件")
            matchFile = os.path.join(DIR, '选配文件',
                                     YMD + '_选配方案_' + self.name + '.csv')
            self.mf[['牛号', '1选', '2选', '3选']].to_csv(matchFile,
                                                     index=False,
                                                     header=False,
                                                     encoding='utf_8_sig',
                                                     line_terminator='\r\n')
            logger.info('3.24【CSV选配方案】，保存在：{}'.format(matchFile))

    def produce_pedigree(self, age=0, pgCol='牛号, 出生日期, 父亲牛号, 外祖父号, 胎次', t=0):
        """生成系谱并保存的函数"""
        global YMD
        global DIR
        pgCol = pgCol.replace(' ', '').replace('，', ',').strip(',').split(',')
        cols = pgCol + ['月龄', '配种组']
        logger.info("\n\n开始执行功能2: 提取系谱信息")
        try:
            pdg = self.herds[self.herds['月龄'] >= age][cols]
            pdg.replace(-1, "", inplace=True)
            logger.info('从牛群明细中导入需要的列并筛选符合月龄条件的牛, 需要的列是: \n  ' +
                        '{}'.format(cols))
        except Exception:
            logger.error('错误:\n{}'.format(traceback.format_exc()))
            return

        if self.bred_group_file != 'nothing':
            logger.info('此牧场有配种组信息')
            try:
                i = pgCol.index('胎次')
            except ValueError:
                pgCol.append('胎次')
                logger.info('【最后一列】下是配种组的组别号')
            else:
                pgCol[i] = '配种组'
                logger.info('【原胎次列】下是配种组的组别号')
        else:
            logger.info('此牧场没有配种组信息，提取的列为指定的列')
        pdg = pdg[pgCol]
        logger.info('系谱提取完成')
        logger.debug(pdg.dtypes)  # '各列数据类型\n' +

        # 保存到csv
        gps_ata_file = os.path.join(DIR, 'AltaGPS_Data')
        result_file = os.path.join(
            gps_ata_file, YMD + '_' + self.name + '_pedigree' + self.name2)
        logger.debug('t' + str(t))
        pdg.to_csv(result_file,
                   index=False,
                   header=int(t),
                   encoding='utf_8_sig',
                   line_terminator='\r\n')
        cows_num = len(pdg)
        logger.info('共有牛头数：' + str(cows_num))
        logger.info('功能2执行完成，系谱结果文件已保存在：{}\n'.format(result_file))

    def produce_pedigree_by_years(self, age=0, bred_group_file=None):
        """生成系谱并保存的函数"""
        global YMD
        global DIR

        print('执行时间：', time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        pdg1 = self.herds[['牛号', '出生日期', '父亲牛号', '外祖父号', '胎次']].copy()
        # pdg1 = pdg1['牛号', '出生日期', '父亲牛号', '外祖父号', '胎次','配种公牛号']
        pdg_unborn = pdg1[['配种公牛号', '父亲牛号']].copy()
        pdg_unborn.columns = ['父亲牛号', '外祖父号']
        pdg_unborn['出生日期'] = '3/1/2020'
        pdg_unborn['胎次'] = 0
        pdg_unborn['牛号'] = range(9999999, len(pdg_unborn) + 9999999)
        pdg_unborn = pdg_unborn[['牛号', '出生日期', '父亲牛号', '外祖父号', '胎次']]

        pdg = pdg1.append(pdg_unborn)
        # pdg = pdg.fillna(0)  # 缺失值替换为0
        pdg['胎次'] = pdg['胎次'].astype('int')

        # 保存到csv
        gps_ata_file = os.path.join(DIR, 'AltaGPS_Data')
        result_file = os.path.join(gps_ata_file,
                                   YMD + '_' + self.name + '_pedigree.csv')
        pdg.to_csv(result_file,
                   index=False,
                   header=False,
                   encoding='utf_8_sig',
                   line_terminator='\r\n')
        cows_num = len(pdg)
        print('共有牛头数：', cows_num)
        print('按年系谱结果文件已保存：', result_file)
        print('执行结束：', time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    def body_select(self, body_file, mindim, maxdim):
        global YMD
        global DIR

        def todo_or_not(lact, dim, has_body):
            if (lact > 0 and has_body == 0
                    and int(mindim) <= dim <= int(maxdim)):
                return 1
            else:
                return 0

        logger.info('\n\n开始执行功能1：筛选要做体型的牛')
        logger.info('参数是：\n  牧场：{}\n  最新体型明细：{}\n  DIM范围：{} - {}'.format(
            self, body_file, mindim, maxdim))
        try:
            herds_body = self.herds[['牛号', '牛舍名称', '泌乳天数', '产后天数',
                                     '胎次']].copy()
            logger.info('从牛群明细中导入需要的列：牛号，牛舍名称，泌乳天数（产后天数），胎次')
        except Exception:
            logger.error('错误:\n{}'.format(traceback.format_exc()))
            return
        bodyFile = pd.DataFrame(
            pd.read_csv(os.path.join(DIR, 'Match_files', body_file)))
        if self.name in [
                '靖远',
                '吴忠',
        ]:  # 这2牧场的牛号有字母
            # zfill(width) --width 指定字符串的长度。原字符串右对齐，前面填充0
            bodyFile['牛号'] = bodyFile['牛号'].apply(lambda x: x.zfill(6))
            bodyFile['牛号'] = bodyFile['牛号'].apply(lambda x: x.upper())
        logger.info('体型明细，已导入')

        # 处理体型表
        bodyFile_s = bodyFile[['牛号']].copy()  # 体型明细只选取需要的列
        bodyFile_s.loc[:, 'has_body'] = 1  # 体型明细增加辅助列

        # 将体型表连接到全群明细表上
        bodyFile_m = pd.merge(herds_body, bodyFile_s, how='left')
        bodyFile_m = bodyFile_m.fillna(0)  # 缺失值替换为0

        # 添加待做体型标记
        bodyFile_m['todo'] = bodyFile_m.apply(
            lambda x: todo_or_not(x.胎次, x.产后天数, x.has_body), axis=1)

        # 筛选需要做体型的牛并列出牛舍、泌乳天数及胎次信息
        File_todo_body = bodyFile_m[bodyFile_m['todo'] > 0].iloc[:, 0:5]
        num = File_todo_body['牛号'].count()
        logger.info('筛选完毕,本月需要做体型的有{}头'.format(num))

        # 保存到excel
        resultfile = os.path.join(
            DIR, '待做体型',
            '待做体型_' + self.name + '_' + YMD + '_' + str(num) + '_' + '头.xlsx')
        File_todo_body.to_excel(resultfile, index=False, encoding='utf8')
        logger.info('功能1执行完成，待做体型数据已保存在：{}\n'.format(resultfile))

    def caPR(self, lact, tbrd, start):
        global pd
        self.creat_rc(tofile=0)
        try:
            self.herds.columns.tolist().index('RC')
        except ValueError:
            logger.info("明细中没有繁殖代码或繁殖状态，程序退出")
            return
        # 显示所有列
        pd.set_option('display.max_columns', None)
        # 显示所有行
        pd.set_option('display.max_rows', None)
        # 设置value的显示长度为100，默认为50
        pd.set_option('max_colwidth', 40)
        # pd.set_option('display.height', 1000)
        # 设置宽度，防止换行显示
        pd.set_option('display.width', 1000)

        with open(DB, 'rb') as db:
            sexSirs = pickle.load(db)['com2sex']
        sexSirs = [i[1] for i in sexSirs]

        startDay = datetime.datetime.strptime(start, "%m/%d/%Y")
        logger.info('\n\n\n开始推算怀孕率')
        lact = int(lact)
        tbrd = int(tbrd)

        def caSemen(lsir):
            beefSirs = ['AN', 'AG', 'SM']
            for sir in sexSirs:
                if sir[:-1] in str(lsir):
                    return 1
            for sir in beefSirs:
                if sir in str(lsir):
                    return -1
            else:
                return 0

        try:
            self.herds.columns.tolist().index('semen')
        except ValueError:
            logger.info('原文件中没有semen列')
            try:
                self.herds.columns.tolist().index('与配公牛')
            except ValueError:
                logger.info("原文件中没有【与配公牛】列，程序退出！")
                return
            self.herds['semen'] = self.herds.apply(
                lambda x: caSemen(x['与配公牛']), axis=1)
            logger.info("增加【semen】列并从与配公牛计算：0常规1性控-1肉牛")
        self.herdsFC = self.herds[[
            '牛号', 'RC', '胎次', '配种次数', '出生日期', '怀孕日期', '产犊日期', '干奶日期', '配种日期',
            '流产日期', '与配公牛', 'semen'
        ]].copy()
        self.herdsFC.rename(columns={
            '牛号': 'id',
            'RC': 'rc',
            '胎次': 'lact',
            '配种次数': 'tbrd',
            '出生日期': 'bdat',
            '怀孕日期': 'cdat',
            '产犊日期': 'fdat',
            '干奶日期': 'ddat',
            '配种日期': 'lsbd',
            '流产日期': 'abdat',
            '与配公牛': 'lsir'
        },
                            inplace=True)
        self.cows = self.herdsFC.copy()  # 将处理后的牛群明细保存备用
        pickleFile2 = os.path.join(DIR, 'Match_files', 'selfCows.pickle')
        with open(pickleFile2, 'wb') as db:
            pickle.dump(self.cows, db)
            logger.info('pickle文件已保存：{}'.format(pickleFile2))
        logger.debug(self.herdsFC.columns.tolist())
        logger.debug('lact:{}, tbrd:{}'.format(lact, tbrd))
        logger.debug('lact:{}, tbrd:{}'.format(type(lact), type(tbrd)))
        logger.debug('self.herdsFC.head():\n{}'.format(self.herdsFC.head()))

        # 0计算禁配占比，实际禁配的
        dnbList = []
        age11 = startDay - datetime.timedelta(days=397)
        age = startDay - datetime.timedelta(days=35)
        self.herdsFC['lsbd'] = pd.to_datetime(self.herdsFC['lsbd'])
        self.herdsFC = self.herdsFC.loc[self.herdsFC.lsbd <= age]
        self.herdsFC['bdat'] = pd.to_datetime(self.herdsFC['bdat'])
        sum_heifer_dnb = self.herdsFC.loc[(self.herdsFC.rc == 1) & (
            self.herdsFC.lact == 0) & (self.herdsFC.bdat <= age11)].id.count()
        sum_heifer_all = self.herdsFC.loc[(self.herdsFC.lact == 0) & (
            self.herdsFC.bdat <= age11)].id.count()
        dnbRate0 = "%.2f%%" % (100 * sum_heifer_dnb / sum_heifer_all)
        dnbList.append([0, dnbRate0])
        for i in range(1, 8):
            sum_cows_dnb = self.herdsFC.loc[
                (self.herdsFC.rc == 1) & (self.herdsFC.lact == i)].id.count()
            sum_cows_all = self.herdsFC.loc[(
                self.herdsFC.lact == i)].id.count()
            dnbRate1 = "%.2f%%" % (100 * sum_cows_dnb / sum_cows_all)
            dnbList.append([i, dnbRate1])
        df_dnb = pd.DataFrame(columns=['胎次', '禁配牛占比'], data=dnbList)
        herdsFC_back = self.herdsFC.copy()
        self.herdsFC = self.herdsFC.drop(
            self.herdsFC[self.herdsFC.tbrd < 1].index)  # 去除没有配次的牛
        # 1.整体怀孕率
        self.herdsFC['lact2'] = self.herdsFC.apply(
            lambda x: "青年牛" if x['lact'] == 0 else "成母牛", axis=1)
        df = pd.pivot_table(self.herdsFC,
                            index=['lact2'],
                            values='tbrd',
                            margins=True,
                            margins_name="合计")
        df['怀孕率'] = df.apply(lambda x: 1 / x.tbrd, axis=1)
        logger.info("=" * 30 + "整体怀孕率")
        logger.info("青年牛怀孕率：{}".format("%.2f%%" % (100 / df['tbrd']['青年牛'])))
        logger.info("成母牛怀孕率：{}".format("%.2f%%" % (100 / df['tbrd']['成母牛'])))
        logger.info("全群整体怀孕率：{}".format("%.2f%%" % (100 / df['tbrd']['合计'])))
        # 2.各胎次怀孕率
        self.herdsFC.loc[self.herdsFC['lact'] > 7, 'lact'] = 7
        df2 = pd.pivot_table(self.herdsFC,
                             index=['lact'],
                             values=['tbrd'],
                             aggfunc='mean',
                             margins=True,
                             margins_name="合计")
        logger.info("=" * 30 + "各胎次怀孕率")
        df2['怀孕率'] = df2.apply(lambda x: 1 / x.tbrd, axis=1)
        df2['怀孕率'] = df2['怀孕率'].apply(lambda x: format(x, '.2%'))
        df2.columns = ['平均配次', '怀孕率']
        df2.rename(columns={'lact': '胎次'}, inplace=True)
        df2.index.name = '胎次'
        logger.info(df2)

        # 3.各胎次各配次 配种头数 ##################################
        self.herdsFC.loc[self.herdsFC['tbrd'] > 9, 'tbrd'] = 9
        df3 = pd.pivot_table(self.herdsFC,
                             index=['lact'],
                             values='id',
                             columns=['tbrd'],
                             aggfunc='count',
                             margins=True,
                             margins_name="合计")
        df3.fillna(0, inplace=True)
        logger.info("=" * 30 + "各胎次各配次怀孕率")

        # 配种头数合计
        col = ["合计" + str(i) for i in range(1, 10)]
        df4 = pd.DataFrame(data=0, index=df3.index.tolist(), columns=col)
        for i in range(1, 10):
            df4["合计" + str(i)] = 0
            for j in range(i, 10):
                df4["合计" + str(i)] += df3[j]
        df4["胎次总计"] = df4.sum(axis=1)
        self.herdsFC = self.herdsFC.drop(
            self.herdsFC[self.herdsFC.rc != 5].index)  # 去除没有怀孕的
        df6 = pd.pivot_table(self.herdsFC,
                             index=['lact'],
                             values='id',
                             columns=['tbrd'],
                             aggfunc='count',
                             margins=True,
                             margins_name="合计")  # 怀孕的各胎次各配次头数
        df6.fillna(0, inplace=True)

        df5 = pd.DataFrame()
        ltbrd = min(self.herdsFC['tbrd'].max(), tbrd)
        for i in range(1, ltbrd + 1):  # 怀孕率
            df5['tbrd' + str(i)] = df6[i] / df4['合计' + str(i)]
        df5 = df5.round(3)
        df5.fillna(0, inplace=True)
        df5.index.name = '胎次'
        df5.columns = [str(i) + "配" for i in range(1, ltbrd + 1)]
        df5["合计"] = df6["合计"] / df4["胎次总计"]
        df10 = df5.copy()
        for i in df5.columns:
            df10[i] = df10[i].apply(lambda x: format(x, '.2%'))
        logger.info(df10)

        # 禁配比例
        import pandas as pd
        dfTemp = pd.DataFrame(data=1, index=list(range(8)), columns=['1'])
        df7 = pd.DataFrame(data=0,
                           index=list(range(8)),
                           columns=list(range(1, 10)))
        for i in range(1, ltbrd + 1):
            df7[i] = dfTemp['1'] - df5[str(i) + "配"]
        # df7.replace(1,,inplace=True)
        logger.info("=" * 30 + "X配后未孕比例")
        df7['X配后未孕比例'] = df7.prod(axis=1)
        df7['X配后未孕比例'] = df7['X配后未孕比例'].apply(lambda x: format(x, '.2%'))
        logger.info(df7['X配后未孕比例'])
        logger.info("=" * 30 + "禁配比例")
        logger.info(df_dnb)
        logger.info('怀孕率推算结束\n\n\n')

        # 本次预测将使用的参数:
        logger.info('本次预测将使用的参数:')
        herdsFC = herdsFC_back.drop(
            herdsFC_back[herdsFC_back.tbrd < 1].index)  # 去除没有配次的牛
        herdsFC.loc[herdsFC['lact'] > lact, 'lact'] = lact
        herdsFC.loc[herdsFC['tbrd'] > tbrd, 'tbrd'] = tbrd
        df3 = pd.pivot_table(herdsFC,
                             index=['lact'],
                             values='id',
                             columns=['tbrd'],
                             aggfunc='count',
                             margins=True,
                             margins_name="合计")
        df3.fillna(0, inplace=True)
        # 配种头数合计
        col = ["合计" + str(i) for i in range(1, tbrd + 1)]
        df4 = pd.DataFrame(data=0, index=df3.index.tolist(), columns=col)
        for i in range(1, tbrd + 1):
            df4["合计" + str(i)] = 0
            for j in range(i, tbrd + 1):
                df4["合计" + str(i)] += df3[j]
        df4["胎次总计"] = df4.sum(axis=1)
        herdsFC = herdsFC.drop(herdsFC[herdsFC.rc != 5].index)  # 去除没有怀孕的
        df6 = pd.pivot_table(herdsFC,
                             index=['lact'],
                             values='id',
                             columns=['tbrd'],
                             aggfunc='count',
                             margins=True,
                             margins_name="合计")  # 怀孕的各胎次各配次头数
        df6.fillna(0, inplace=True)

        df5 = pd.DataFrame()
        for i in range(1, ltbrd + 1):  # 怀孕率
            df5['tbrd' + str(i)] = df6[i] / df4['合计' + str(i)]
        df5 = df5.round(3)
        df5.fillna(0, inplace=True)
        df5.index.name = '胎次'
        df5.columns = [str(i) + "配" for i in range(1, ltbrd + 1)]
        df5["合计"] = df6["合计"] / df4["胎次总计"]
        for ltb in range(ltbrd, 9):
            df5[str(i + 1) + "配"] = df5[str(ltbrd) + "配"]
        self.PR = df5.copy()  # 将怀孕率添加到类上备用
        df10 = df5.copy()
        for i in df5.columns:
            df10[i] = df10[i].apply(lambda x: format(x, '.2%'))
        logger.info(df10)

    def forecast(self, start, end, parameter_suffix, onOff, args, method=1):
        # TODO: 犊牛死胎率
        # TODO: 已怀犊牛是否留养
        # TODO: 怀孕率调低（天镇，商都。56%-->40%, 性控/常规，经产/头胎）
        # TODO: 死亡、淘汰率，流产率
        # &SHOW ID RC LACT TBRD BDAT CDAT FDAT DDAT %70.BRED.101 ABDAT LSIR
        # farms,
        if int(onOff[2]) == 1:  # 使用指定怀孕率
            lactList = [0]
            tbrdList = list(range(1, 10))
        elif int(onOff[2]) == 2:  # 使用推算怀孕率
            lactList = self.PR.index.tolist()[:-1]
            tbrdList = self.PR.columns.tolist()[:-1]

        AGEFB = 0
        WWTP = 0
        LIVESTOCK = pd.DataFrame()
        LIVESTOCK = pd.DataFrame(columns=('Date', 'Total', 'Yong', 'Y_preg',
                                          'Adult', 'Milking', 'Adult_preg'))
        CATE_LIVESTOCK = pd.DataFrame()
        HERDS = pd.DataFrame()
        semen_usage_ = pd.DataFrame(
            columns=('1配', '2配', '3配', '4配', '5配', '6配', '7配', '8配', '9配'),
            index=['0胎', '1胎', '2胎', '3胎', '4胎', '5胎', '6胎', '7胎']).fillna(0)
        lt = []
        for lact in range(8):
            for tbrd in range(1, 10):
                lt.append("{}胎{}配".format(lact, tbrd))
        semen_usage_0 = pd.DataFrame(columns=lt)
        semen_usage = semen_usage_0.copy()
        semen_usage_1 = semen_usage_0.copy()
        semen_usage_2 = semen_usage_0.copy()
        DATE0 = datetime.datetime(1970, 1, 1)
        FRESH_COWS = 0
        FRESH_COWS_KEEP_COM = 0
        FRESH_COWS_KEEP_SEX = 0
        FRESH_COWS_SOLD_BEEF = 0
        LEFT_COWS_C = 0
        LEFT_COWS_H = 0
        LEFT_COWS_B = 0
        SEMEN_USAGE = 0
        cos = [
            "期初", "产犊", "留养", "围产", "干奶", "流产", "怀孕", "配种", "死淘", "期末", "检查"
        ]
        SUMMARY = pd.DataFrame(columns=cos)
        FEMAL = pd.DataFrame()  # 母犊率
        BRED_PLAN = pd.DataFrame()  # 配种方案
        PREG_RATE = pd.DataFrame()  # 怀孕率
        HDR = pd.DataFrame()  # 发情揭发率
        LEAVE = pd.DataFrame()  # 死淘率
        ABORT_RATE = (0.05, 0.15)
        # TODO: 最后2个应该是0.075，不是0.75
        ABORT_STR = [0, 0.05, 0.45, 0.15, 0.1, 0.05, 0.05, 0.075, 0.075]

        # os.path.join(PATH, herds_file)
        snap_date = datetime.datetime.strptime(start,
                                               "%m/%d/%Y")  # 快照日期 记录初始日期，不变化
        node_date = datetime.datetime.strptime(start,
                                               "%m/%d/%Y")  # 节点日期，生长过程中的
        end = datetime.datetime.strptime(end, "%m/%d/%Y")
        monthNum = ((end.year - snap_date.year) * 12 +
                    (end.month - snap_date.month))
        first = datetime.datetime(snap_date.year, snap_date.month, 1)
        first = datetime.datetime.strftime(first, "%Y%m")
        firstMonth = datetime.datetime(snap_date.year, snap_date.month, 1)
        temp = firstMonth
        firstMonth = datetime.datetime.strftime(firstMonth, "%Y%m")
        monthList = []  # 预测开始日期到结束日期之间的月份列表
        monthList.append(int(firstMonth))
        for i in range(monthNum):
            if temp.month + 1 <= 12:
                item = datetime.datetime(temp.year, temp.month + 1, 1)
            elif temp.month + 1 == 13:
                item = datetime.datetime(temp.year + 1, 1, 1)
            temp = item
            item = datetime.datetime.strftime(item, "%Y%m")
            monthList.append(int(item))

        def parameter(suffix):
            nonlocal HDR, PREG_RATE, ABORT_RATE, FEMAL, LEAVE, BRED_PLAN
            nonlocal onOff, args, monthList, AGEFB, WWTP
            femal_csv = os.path.join(DIR, "femal_rate_" + suffix + ".csv")
            # preg_csv = os.path.join(DIR, "preg_rate_" + suffix + ".csv")
            bred_csv = os.path.join(DIR, "bred_plan_" + suffix + ".csv")
            hdr_csv = os.path.join(DIR, "hdr_" + suffix + ".csv")
            leave_csv = os.path.join(DIR, "leave_" + suffix + ".csv")

            logger.info("1. 开始加载预测参数...")
            # 1参配条件
            if int(onOff[0]) == 1:  # 1为每月相同
                AGEFB = float(args['fbage'])
                WWTP = int(args['wwtp'])
                logger.info("1.1 加载参配条件，每月相同")

            # 2发情揭发率
            if int(onOff[1]) == 1:  # 1为每月相同
                hdrH = float(args['hdrH'])
                hdrC = float(args['hdrC'])
                r = []
                for i in range(9):
                    if i == 0:
                        r.append(hdrH)
                    else:
                        r.append(hdrC)
                HDR = pd.DataFrame()
                HDR['LACT'] = list(range(9))
                HDR['r'] = r
                HDR.set_index('LACT', inplace=True)
                logger.info("1.2 加载指定的发情揭发率参数，每月相同")
            elif int(onOff[1]) == 0:
                if os.path.isfile(hdr_csv):
                    HDR = pd.read_csv(hdr_csv, index_col='LACT')
                    logger.info("1.2 从文件加载指定的发情揭发率，【不】每月相同")
                else:
                    HDR = pd.read_csv('hdr_default.csv', index_col='LACT')
                    logger.info("1.2 从文件加载【默认】发情揭发率，每月相同")

            # 3怀孕率
            LT = []
            if int(onOff[2]) == 1:  # 1为给定值
                LT = []
                pr = [float(args['prH']), float(args['prC'])]
                data = []
                for lact in [0, 1]:
                    for tbrd in list(range(1, 10)):
                        LT.append("L" + str(lact) + "T" + str(tbrd))
                        data.append(pr[lact])
                PREG_RATE = pd.DataFrame(columns=[0])
                PREG_RATE["LT"] = LT
                PREG_RATE[0] = data
                PREG_RATE[1] = PREG_RATE[0] * 0.9
                PREG_RATE[-1] = PREG_RATE[0] * 0.8
                PREG_RATE['LT'] = LT
                PREG_RATE.set_index("LT", inplace=True)
                logger.info("1.3 加载指定的怀孕率参数")
            elif int(onOff[2]) == 2:  # 2为推算
                self.PR.columns = list(range(1, len(self.PR.columns) + 1))
                LT = []
                data = []
                for lact in self.PR.index.tolist()[:-1]:
                    for tbrd in self.PR.columns.tolist()[:-1]:
                        LT.append("L" + str(lact) + "T" + str(tbrd))
                        data.append(self.PR.loc[lact].values[tbrd - 1])
                PREG_RATE = pd.DataFrame(data=data, columns=[0])
                PREG_RATE[1] = PREG_RATE[0] * 0.9
                PREG_RATE[-1] = PREG_RATE[0] * 0.8
                PREG_RATE['LT'] = LT
                PREG_RATE.set_index("LT", inplace=True)
                PREG_RATE = PREG_RATE * 0.9  # 怀孕率*90%是调整
                # 将0值填充为前一个值
                PREG_RATE.replace(0, 0.2, inplace=True)  # method='ffill'
                logger.info("1.3 加载推算的怀孕率参数")

            # 4流产率
            if int(onOff[3]) == 1:  # 选中每月相同
                ABORT_RATE = (float(args['aboRate']), float(args['aboRateC']))
                logger.info("1.4 加载流产率，每月相同")
            elif int(onOff[3]) == 0:
                pass

            # 5母犊率
            if int(onOff[4]) == 1:  # 选中每月相同
                com = float(args['femalRateC'])
                sex = float(args['femalRateS'])
                beef = float(args['femalRateB'])
                FEMAL = pd.DataFrame(index=list(range(monthNum + 1)),
                                     columns=['com', 'sex', 'beef'])
                FEMAL['com'] = com
                FEMAL['sex'] = sex
                FEMAL['beef'] = beef
                logger.debug("FEMAL:{}".format(FEMAL))
                logger.debug("monthList:{}".format(monthList))
                FEMAL['month'] = monthList
                FEMAL.set_index("month", inplace=True)
                logger.info("1.5 加载指定参数，每月相同")
            elif int(onOff[4]) == 0:
                if os.path.isfile(femal_csv):
                    FEMAL = pd.read_csv(femal_csv, index_col='month')
                    logger.info("1.5 加载牧场自己的参数，按月")
                else:
                    FEMAL = pd.read_csv("femal_rate_default.csv",
                                        index_col='month')
                    logger.info("1.5 加载默认参数")

            # 6死淘率
            if int(onOff[5]) == 1:
                calfCulRate = float(args['calfCulRate'])
                yongCulRate = float(args['yongCulRate'])
                cowCulRate = float(args['cowCulRate'])
                LEAVE = pd.DataFrame()
                LEAVE['month'] = monthList
                LEAVE['calfCulRate'] = calfCulRate
                LEAVE['yongCulRate'] = yongCulRate
                LEAVE['cowCulRate'] = cowCulRate
                LEAVE.set_index("month", inplace=True)
                logger.info("1.6 加载指定的死淘率参数，每月相同")
            elif int(onOff[5]) == 0:
                if os.path.isfile(hdr_csv):
                    LEAVE = pd.read_csv(leave_csv, index_col='LACT')
                    logger.info("1.6 从文件加载指定的死淘率，【不】每月相同")
                else:
                    LEAVE = pd.read_csv('leave_default.csv', index_col='LACT')
                    logger.info("1.6 从文件加载【默认】死淘率，每月相同")

            # 7配种方案
            if int(onOff[6]) == 1:  # 1为每月相同
                sex_lact = args['sexLact']
                sex_tbrd = args['sexTbrd']
                beef_lact = args['beefLact']
                beef_rate = args['beefRate']
                BRED_PLAN = pd.DataFrame()
                BRED_PLAN['month'] = monthList
                BRED_PLAN['sex_lact'] = sex_lact
                BRED_PLAN['sex_tbrd'] = sex_tbrd
                BRED_PLAN['beef_lact'] = beef_lact
                BRED_PLAN['beef_rate'] = beef_rate
                BRED_PLAN.set_index('month', inplace=True)
                logger.info("1.7 加载指定的配种方案，每月相同")
            elif int(onOff[6]) == 0:
                if os.path.isfile(bred_csv):
                    BRED_PLAN = pd.read_csv(bred_csv, index_col='month')
                    logger.info("1.7 从文件加载指定的配种方案，【不】每月相同")
                else:
                    BRED_PLAN = pd.read_csv("bred_plan_default.csv",
                                            index_col='month')
                    logger.info("1.7 从文件加载【默认】配种方案，每月相同")

        def cal_closing(node_date):
            """根据节点时间统计牛群中各种牛的数量"""
            logger.info("4. 开始统计当前牛群结构")
            nonlocal LIVESTOCK
            total = HERDS.bdat.count()
            yong = HERDS[HERDS.lact == 0].bdat.count()
            adult = HERDS[HERDS.lact > 0].bdat.count()
            preg = HERDS[(HERDS.lact == 0) & (HERDS.cdat > DATE0) &
                         (HERDS.dslb > 30)].bdat.count()  # 配后天数大于28天是为了避免孕检天数
            milking = HERDS[(HERDS.fdat > DATE0)
                            & (HERDS.ddat <= DATE0)].bdat.count()
            adult_preg = HERDS[(HERDS.lact > 0) & (HERDS.dslb > 30) &
                               (HERDS.cdat > DATE0)].bdat.count()
            tmp = pd.DataFrame(
                {
                    'Date': node_date,
                    'Total': total,
                    'Yong': yong,
                    'Y_preg': preg,
                    'Adult': adult,
                    'Milking': milking,
                    'Adult_preg': adult_preg
                },
                index=[0])
            LIVESTOCK = LIVESTOCK.append(tmp, ignore_index=True, sort=False)

        def update_category(lact, age, dim, dcc, ddat, cdat, now_date):
            category = ""
            if lact > 0:
                if ddat > DATE0 and cdat <= DATE0:  # 未孕干奶牛
                    category = 'DD' + \
                        str(min(math.floor((now_date - ddat).days / 30), 6))
                elif cdat > DATE0:  # 怀孕牛
                    preg_mth = math.floor(dcc / 30) + 1
                    if dcc >= 250:
                        category = 'T9'
                    elif ddat > DATE0:  # 正常干奶牛
                        dry_mth = math.floor((now_date - ddat).days / 30) + 1
                        category = 'D' + str(min(dry_mth, 8))
                    elif ddat <= DATE0:  # 未干奶怀孕牛
                        category = 'MP' + str(min(preg_mth, 8))
                    else:
                        raise MyException("Error")
                elif cdat <= DATE0:  # 未孕牛
                    milk_mth = math.floor(dim / 30) + 1
                    category = 'M' + str(min(milk_mth, 30))
            elif lact == 0:
                if cdat > DATE0:
                    preg_mth = math.floor(dcc / 30) + 1
                    if dcc >= 250:
                        category = 'T9'
                    elif dcc < 250:
                        category = 'P' + str(min(preg_mth, 8))
                    else:
                        raise MyException("Error")
                else:
                    yong_mth = min(math.floor(age / 30.5), 30)
                    category = 'Y' + str(yong_mth)
            return category

        def pre_treat(cows_file, snap_date):
            logger.info("2. 开始预处理...")
            nonlocal HERDS, lactList, tbrdList
            HERDS = cows_file.copy()
            HERDS.columns = [
                'id', 'rc', 'lact', 'tbrd', 'bdat', 'cdat', 'fdat', 'ddat',
                'lsbd', 'abdat', 'lsir', 'semen'
            ]
            HERDS.loc[HERDS.tbrd < 1, 'tbrd'] = 0  # 把文件中的胎次为-1的改为0
            # 初步处理，增加推算需要的信息
            logger.info("2.1 增加需要的列并设置初始值")
            HERDS['due'] = '1/1/1970'
            HERDS['age_day'] = -1
            HERDS['dcc'] = -1
            HERDS['dim'] = -1
            HERDS['dslb'] = -1  # 配后天数
            # HERDS['semen'] = 0  # 0：common常规，1：sex性控，-1：beef肉牛。
            HERDS['tag_bred'] = 0  # 配种标记，为1表示在生长期间内配过种
            HERDS['lact_grp2'] = np.stack(
                (HERDS.lact, np.full(len(HERDS), max(lactList)))).min(axis=0)
            HERDS['tbrd_grp'] = np.stack(
                (HERDS.tbrd, np.full(len(HERDS),
                                     int(tbrdList[-1][0])))).min(axis=0)
            HERDS['type_bred'] = 0
            HERDS['tag_frsh'] = 0  # 产犊标记
            HERDS['tag_dry'] = 0
            HERDS['tag_preg'] = 0
            HERDS['abrt'] = 0  # 流产次数

            logger.info("2.2 填充空日期，并转换日期格式")
            HERDS.cdat.fillna('1/1/1970', inplace=True)
            HERDS.fdat.fillna('1/1/1970', inplace=True)
            HERDS.ddat.fillna('1/1/1970', inplace=True)
            HERDS.lsbd.fillna('1/1/1970', inplace=True)
            HERDS.abdat.fillna('1/1/1970', inplace=True)
            HERDS['bdat'] = pd.to_datetime(HERDS['bdat'], format='%m/%d/%Y')
            HERDS['cdat'] = pd.to_datetime(HERDS['cdat'], format='%m/%d/%Y')
            HERDS['fdat'] = pd.to_datetime(HERDS['fdat'], format='%m/%d/%Y')
            HERDS['ddat'] = pd.to_datetime(HERDS['ddat'], format='%m/%d/%Y')
            HERDS['lsbd'] = pd.to_datetime(HERDS['lsbd'], format='%m/%d/%Y')
            HERDS['due'] = pd.to_datetime(HERDS['due'], format='%m/%d/%Y')
            HERDS['abdat'] = pd.to_datetime(HERDS['abdat'], format='%m/%d/%Y')

            logger.info("2.3 设置胎次最大为7，配次最大为9")
            HERDS.loc[HERDS.lact > 7, 'lact'] = 7
            HERDS.loc[HERDS.tbrd > 9, 'tbrd'] = 9

            # 计算 计算列
            logger.info("2.3 计算可推算信息：日龄，怀孕天数，预产期，泌乳天数，配后天数")
            _num = len(HERDS.loc[HERDS.cdat > DATE0, 'due'])
            HERDS['age_day'] = (
                snap_date - HERDS.bdat).astype('timedelta64[D]').astype(int)
            HERDS.loc[HERDS.cdat > DATE0, 'dcc'] = (
                snap_date - HERDS.loc[HERDS.cdat > DATE0, 'cdat']
            ).astype('timedelta64[D]').astype(int)
            HERDS.loc[HERDS.cdat > DATE0, 'due'] = pd.to_datetime(
                HERDS.loc[HERDS.cdat > DATE0, 'cdat'].values +
                np.random.normal(273, 9, _num).astype('timedelta64[D]'),
                format='%m/%d/%Y')  # 怀孕牛的预产期
            num_f = len(HERDS.loc[(HERDS.due > DATE0) &
                                  (HERDS.due <= snap_date), 'due'])
            if num_f:  # 怀孕牛且预测期在年群明细日期之前的牛的预产期
                HERDS.loc[(HERDS.due > DATE0) &
                          (HERDS.due <= snap_date), 'due'] = (pd.to_datetime(
                              np.full(num_f, snap_date).astype('datetime64') +
                              np.random.randint(1, 7, size=num_f).astype(
                                  'timedelta64[D]')))

            HERDS.loc[HERDS.fdat > DATE0, 'dim'] = (
                snap_date - HERDS.loc[HERDS.fdat > DATE0, 'fdat']
            ).astype('timedelta64[D]').astype(int)
            # 配次大于0的牛的配后天数
            HERDS.loc[HERDS.tbrd > 0, 'dslb'] = (
                snap_date - HERDS.loc[HERDS.tbrd > 0, 'lsbd']
            ).astype('timedelta64[D]').astype(int)
            cond_abdat = (HERDS.cdat <= DATE0) & (HERDS.abdat >
                                                  DATE0) & (HERDS.rc == 3)
            # 流产空怀牛的配后天数 = 流产后天数
            HERDS.loc[cond_abdat, 'dslb'] = (
                snap_date - HERDS.loc[cond_abdat, 'abdat']
            ).astype('timedelta64[D]').astype(int)
            logger.info("2.4 更新牛只类别")
            HERDS['category'] = HERDS.apply(lambda x: update_category(
                x.lact, x.age_day, x.dim, x.dcc, x.ddat, x.cdat, snap_date),
                                            axis=1)
            HERDS['category'] = HERDS['category'].astype(str)

        # 0：common，1：sex，-1：beef。
        def set_bred_plan(sex_lact_list, sex_tbrd_list, beef_lact, beef_rate):
            nonlocal HERDS
            # logger.debug("type(HERDS):{}".format(HERDS.dtypes))
            # logger.debug("type(beef_lact):{}".format(type(beef_lact)))
            # logger.debug("type(beef_rate):{}".format(type(beef_rate)))
            HERDS['semen'] = 0
            if sex_lact_list and sex_tbrd_list:
                for i in range(len(sex_lact_list)):
                    condition = ((HERDS.lact == int(sex_lact_list[i])) &
                                 (HERDS.tbrd <= (int(sex_tbrd_list[i]) - 1)) &
                                 (HERDS.rc < 5))
                    HERDS.loc[condition, 'semen'] = 1
            elif beef_lact and beef_rate:
                condition2 = HERDS[HERDS.lact >= int(beef_lact)].sample(
                    frac=float(beef_rate)).index
                HERDS.loc[condition2, 'semen'] = -1
            else:
                HERDS.loc[HERDS.rc < 5, 'semen'] = 0

        def grow(node_date, end, parameter_suffix):
            # 生长函数：按指定天数生长。范围在【16，20】
            grow_day = 16
            logger.info("5. 牛群开始生长，每次生长最多{}天".format(grow_day))
            nonlocal CATE_LIVESTOCK, HERDS, SEMEN_USAGE, lactList, tbrdList
            nonlocal FRESH_COWS, LEFT_COWS_C, LEFT_COWS_H, LEFT_COWS_B
            nonlocal FRESH_COWS_KEEP_SEX, FRESH_COWS_SOLD_BEEF
            nonlocal semen_usage_0, semen_usage_1, semen_usage_2
            nonlocal FRESH_COWS_KEEP_COM, AGEFB
            ababy = pd.DataFrame(
                {
                    'id': 0,
                    'rc': 0,
                    'lact': 0,
                    'tbrd': 0,
                    'bdat': DATE0,
                    'cdat': DATE0,
                    'fdat': DATE0,
                    'ddat': DATE0,
                    'lsbd': DATE0,
                    'abdat': DATE0,
                    'due': DATE0,
                    'age_day': 0,
                    'dcc': -1,
                    'dim': -1,
                    'dslb': -1,
                    'semen': 0,
                    'category': 'Y0'
                },
                index=[0])

            flag = True
            while flag:
                opening = len(HERDS)
                HERDS[[
                    '_rc', '_lact', '_tbrd', '_cdat', '_fdat', '_ddat',
                    '_lsbd', '_due', '_age_day', '_dcc', '_dim', '_dslb'
                ]] = HERDS[[
                    'rc', 'lact', 'tbrd', 'cdat', 'fdat', 'ddat', 'lsbd',
                    'due', 'age_day', 'dcc', 'dim', 'dslb'
                ]]
                that_day = node_date + datetime.timedelta(days=grow_day)
                if that_day >= end:
                    that_day = end
                    flag = False
                elif that_day.day < grow_day:  # 节点不是月底，且生长grow_day天后和节点不在同一个月
                    that_day = datetime.datetime(
                        node_date.year, node_date.month,
                        calendar.monthrange(node_date.year,
                                            node_date.month)[1])
                day = (that_day - node_date).days
                mth = int(datetime.datetime.strftime(that_day, "%Y%m"))
                if day == 0:
                    logger.info('5. 生长0天，循环直接退出')
                    break

                # 生长
                logger.info("5.1 当前日期：{}。生长了{}天".format(
                    that_day.strftime('%Y-%m-%d'), day))
                logger.debug("5.1.1 更新日龄，怀孕天数，泌乳天数，配后天数")
                HERDS['age_day'] = HERDS['age_day'].values + day
                HERDS.loc[HERDS.dcc >= 0, 'dcc'] = HERDS.loc[
                    HERDS.dcc >= 0, 'dcc'].values + day
                HERDS.loc[(HERDS.fdat > DATE0) &
                          (HERDS.ddat <= DATE0), 'dim'] = HERDS.loc[
                              (HERDS.fdat > DATE0) &
                              (HERDS.ddat <= DATE0), 'dim'].values + day
                HERDS.loc[HERDS.dslb >= 0, 'dslb'] = HERDS.loc[
                    HERDS.dslb >= 0, 'dslb'].values + day
                #  设置配种方案
                logger.debug("5.1.2 设置配种方案")
                # logger.debug("BRED_PLAN:{}".format(BRED_PLAN))
                # logger.debug("mth:{}".format(mth))
                # logger.debug(type(BRED_PLAN['sex_lact'][mth]))
                logger.debug("BREDPLAN{},{},{},{}".format(
                    [BRED_PLAN['sex_lact'][mth]], [BRED_PLAN['sex_tbrd'][mth]],
                    BRED_PLAN['beef_lact'][mth], BRED_PLAN['beef_rate'][mth]))
                set_bred_plan([BRED_PLAN['sex_lact'][mth]],
                              [BRED_PLAN['sex_tbrd'][mth]],
                              BRED_PLAN['beef_lact'][mth],
                              BRED_PLAN['beef_rate'][mth])

                # 围产——>产犊——>出生
                condition_fresh = (HERDS.due > node_date) & (HERDS.due <=
                                                             that_day)
                born_cows = len(HERDS.loc[condition_fresh, 'bdat'])
                born_cows_c = len(
                    HERDS.loc[condition_fresh &
                              (HERDS.semen == 0), 'bdat'])  # 常规牛犊
                born_cows_s = len(
                    HERDS.loc[condition_fresh &
                              (HERDS.semen == 1), 'bdat'])  # 性控牛犊
                born_cows_b = len(
                    HERDS.loc[condition_fresh &
                              (HERDS.semen == -1), 'bdat'])  # 肉牛犊
                babys_num = 0
                if born_cows:
                    # 牛在预产期那一天产犊
                    HERDS.loc[condition_fresh,
                              'fdat'] = HERDS.loc[condition_fresh, 'due']
                    # .astype('timedelta64[D]')).astype('datetime64').astype(np.datetime64)
                    bdat_fdat = HERDS.loc[condition_fresh, 'fdat'].values
                    # print(HERDS.dtypes)
                    HERDS.loc[condition_fresh, 'lact'] = HERDS.loc[
                        condition_fresh, 'lact'].values + 1
                    HERDS.loc[condition_fresh & (HERDS.lact > 7), 'lact'] = 7
                    HERDS.loc[condition_fresh, 'dim'] = (
                        that_day - HERDS.loc[condition_fresh, 'fdat']
                    ).astype('timedelta64[D]').astype(int)
                    HERDS.loc[condition_fresh, [
                        'rc', 'dcc', 'dslb', 'cdat', 'due', 'ddat', 'lsbd',
                        'tbrd', 'tag_frsh'
                    ]] = [2, -1, -1, DATE0, DATE0, DATE0, DATE0, 0, 1]
                    FRESH_COWS += born_cows

                    # 母犊分类统计
                    babys_num_c = np.digitize(np.random.random(born_cows_c),
                                              [FEMAL['com'][mth], 0]).sum()
                    babys_num_s = np.digitize(np.random.random(born_cows_s),
                                              [FEMAL['sex'][mth], 0]).sum()
                    babys_num_b = np.digitize(np.random.random(born_cows_b),
                                              [FEMAL['beef'][mth], 0]).sum()
                    babys_num = babys_num_c + babys_num_s
                    if babys_num_c > 0:
                        FRESH_COWS_KEEP_COM += babys_num_c
                    if babys_num_s > 0:
                        FRESH_COWS_KEEP_SEX += babys_num_s
                    if born_cows_b > 0:
                        FRESH_COWS_SOLD_BEEF += born_cows_b

                    # 母犊入群
                    if babys_num > 0:
                        babys = pd.DataFrame()
                        babys = babys.append([ababy] * babys_num,
                                             ignore_index=True)
                        babys['bdat'] = pd.to_datetime(
                            np.random.choice(a=bdat_fdat,
                                             size=babys_num,
                                             replace=False,
                                             p=None))
                        babys['age_day'] = (
                            that_day -
                            babys.bdat).astype('timedelta64[D]').astype(int)
                        HERDS = HERDS.append(babys,
                                             ignore_index=True,
                                             sort=False)
                    logger.debug(
                        "5.1.3 繁殖事件-1产犊. 事件数：{}, 得母犊(常规性控肉牛):{} {} {}".format(
                            born_cows, babys_num_c, babys_num_s, babys_num_b))

                # 干奶牛——>围产牛
                cows_to_clo = len(HERDS.loc[(HERDS.dcc >= 250), 'dcc'])
                logger.debug("5.1.3 繁殖事件-2围产. 牛头数：{}".format(cows_to_clo))

                # 怀孕——>干奶牛
                condition0 = (HERDS.dcc >= 220) & (HERDS.lact >
                                                   0) & (HERDS.rc == 5)
                cows_to_dry = len(HERDS.loc[condition0, 'dcc'])
                if cows_to_dry:
                    HERDS.loc[condition0, ['rc', 'tag_dry']] = [6, 1]
                    # HERDS['cdat'] = pd.to_datetime(HERDS['cdat'],
                    # format='%m/%d/%Y')
                    HERDS.loc[condition0, 'ddat'] = pd.to_datetime(
                        (HERDS.loc[condition0, 'cdat'].values +
                         np.full(cows_to_dry, 220).astype('timedelta64[D]')
                         ).astype('datetime64'))
                    # HERDS['ddat'] = pd.to_datetime(HERDS['ddat'],
                    # format='%m/%d/%Y')
                logger.debug("5.1.3 繁殖事件-3干奶. 牛头数：{}".format(cows_to_dry))

                # 可配——>已配牛
                # 青年和成母牛配种,2配及以上
                semen_usage.loc[str(that_day.strftime("%m/%d/%Y"))] = 0
                for t in range(1, 10):
                    for l in range(0, 8):
                        con3_ = ((HERDS.lact == l) & (HERDS.tbrd == t) &
                                 (HERDS.dslb >= 21) & (HERDS.rc != 1) &
                                 (HERDS.cdat <= DATE0))
                        condition3 = HERDS[con3_].sample(
                            frac=HDR['r'][l]).index
                        numtl = len(HERDS.loc[condition3, 'tbrd'])
                        if numtl > 0:
                            HERDS.loc[condition3, 'tbrd'] = HERDS.loc[
                                condition3, 'tbrd'].values + 1
                            HERDS.loc[con3_ & (HERDS.tbrd > 9), 'tbrd'] = 9
                            semen_usage.loc[str(that_day.strftime("%m/%d/%Y")),
                                            str(l) + '胎' + str(min(t + 1, 8)) +
                                            '配'] += numtl
                            semen_usage_[str(min(t + 1, 8)) +
                                         '配'][str(l) + '胎'] += (numtl)
                            SEMEN_USAGE += numtl
                            HERDS.loc[condition3,
                                      ['rc', 'tag_bred', 'type_bred']] = [
                                          4, 1, 1
                                      ]
                            HERDS.loc[condition3, 'dslb'] = np.stack(
                                (HERDS.loc[condition3, 'dslb'] - 21,
                                 np.full(numtl, day))).min(axis=0)
                # 青年牛配种,1配
                con1_ = (HERDS.age_day >= 395) & (HERDS.lact == 0) & (
                    HERDS.tbrd <= 0) & (HERDS.rc == 0)  # tbrd 为空
                condition1 = HERDS[con1_].sample(frac=HDR['r'][l]).index
                num10 = len(HERDS.loc[condition1, 'tbrd'])
                if num10 > 0:
                    HERDS.loc[condition1, 'tbrd'] = 1
                    semen_usage.loc[str(that_day.strftime("%m/%d/%Y")
                                        ), '0胎1配'] += num10
                    semen_usage_['1配']['0胎'] += num10
                    SEMEN_USAGE += num10
                    HERDS.loc[condition1, ['rc', 'tag_bred', 'type_bred']] = [
                        4, 1, 2
                    ]
                    HERDS.loc[condition1, 'dslb'] = np.stack(
                        (HERDS.loc[condition1, 'age_day'] - 395,
                         np.full(num10, day))).min(axis=0)
                # 成母牛配种,1配
                for l in range(1, 8):
                    con2_ = ((HERDS.lact == l) & (HERDS.tbrd <= 0) &
                             (HERDS.dim >= 50) & (HERDS.rc != 1) &
                             (HERDS.cdat <= DATE0))
                    condition2 = HERDS[con2_].sample(frac=HDR['r'][l]).index
                    num11 = len(HERDS.loc[condition2, 'tbrd'])
                    if num11 > 0:
                        HERDS.loc[condition2, 'tbrd'] = 1
                        semen_usage.loc[str(that_day.strftime("%m/%d/%Y")),
                                        str(l) + '胎1配'] += num11
                        semen_usage_['1配'][str(l) + '胎'] += num11
                        SEMEN_USAGE += num11
                        HERDS.loc[condition2, ['rc', 'tag_bred', 'type_bred'
                                               ]] = [4, 1, 3]
                        HERDS.loc[condition2, 'dslb'] = np.stack(
                            (HERDS.loc[condition2, 'dim'] - 50,
                             np.full(num11, day))).min(axis=0)
                # 分类统计冻精使用量
                for s in [-1, 0, 1]:
                    locals()['semen_usage_' + str(s + 1)].loc[str(
                        that_day.strftime("%m/%d/%Y"))] = 0
                    for t in range(1, 10):
                        for l in range(0, 8):
                            usage_num = len(
                                HERDS.loc[(HERDS.tag_bred == 1) &
                                          (HERDS.semen == s) &
                                          (HERDS.tbrd == t) &
                                          (HERDS.lact == l), 'tbrd'])
                            locals()['semen_usage_' + str(
                                s + 1)].loc[str(that_day.strftime("%m/%d/%Y")),
                                            str(l) + '胎' + str(t) +
                                            '配'] += usage_num
                # 设置配种日期
                bred_num = len(HERDS.loc[HERDS.tag_bred == 1, 'rc'])
                HERDS.loc[HERDS.tag_bred == 1, 'lsbd'] = pd.to_datetime(
                    np.full(bred_num, that_day).astype('datetime64') -
                    HERDS.loc[HERDS.tag_bred == 1, 'dslb'].values.astype(
                        int).astype('timedelta64[D]'))
                logger.debug("5.1.3 繁殖事件-4配种. 牛头数：{}".format(bred_num))
                HERDS['tag_bred'] = 0

                # 已配 --> 怀孕
                # 用性控冻精的牛怀孕
                HERDS['lact_grp2'] = np.stack(
                    (HERDS.lact, np.full(len(HERDS),
                                         max(lactList)))).min(axis=0)
                HERDS['tbrd_grp'] = np.stack(
                    (HERDS.tbrd, np.full(len(HERDS),
                                         int(tbrdList[-1][0])))).min(axis=0)
                preg_num_this = 0
                logger.debug("tbrdList:{}".format(tbrdList))
                for i in [-1, 0, 1]:
                    for l in lactList:
                        for t in tbrdList:
                            _cond = ((HERDS.cdat <= DATE0) &
                                     (HERDS.lsbd > DATE0) &
                                     (HERDS.lact_grp2 == l) &
                                     (HERDS.tbrd_grp == int(t[0])) &
                                     (HERDS.semen == i))
                            _bred_num = len(HERDS.loc[_cond, 'bdat'])
                            _preg_num = np.digitize(
                                np.random.random(_bred_num), [
                                    PREG_RATE[i]['L' + str(l) + 'T' +
                                                 str(t[0])], 0
                                ]).sum()
                            if i == 0 and l == 0:
                                logger.debug((
                                    "i:{},l:{},t:{},brednum:{},preg:{},peg_rate[i]:{}"
                                ).format(
                                    i, l, t, _bred_num, _preg_num,
                                    PREG_RATE[i]['L' + str(l) + 'T' +
                                                 str(t[0])]))
                            if _preg_num:
                                tmp_index = HERDS[_cond].sample(
                                    n=_preg_num).index
                                HERDS.loc[tmp_index, 'cdat'] = pd.to_datetime(
                                    HERDS.loc[tmp_index, 'lsbd'].values)
                                # np.random.randint(1,22, size=preg_b_num).
                                # astype('timedelta64[D]')
                                HERDS.loc[tmp_index, ['rc', 'tag_preg']] = [
                                    5, 1
                                ]
                                HERDS.loc[tmp_index,
                                          'dcc'] = HERDS.loc[tmp_index, 'dslb']
                                HERDS.loc[tmp_index, 'due'] = (
                                    pd.to_datetime(
                                        HERDS.loc[tmp_index, 'cdat'].values) +
                                    np.random.normal(
                                        273, 9,
                                        _preg_num).astype('timedelta64[D]'))
                                preg_num_this += _preg_num
                    logger.debug("5.1.3 繁殖事件-5孕检怀孕. {}牛头数：{}".format(
                        i, preg_num_this))
                    HERDS['tag_preg'] = 0

                # 月底进行流产和死淘
                left_num = 0
                left_num_c = 0
                left_num_h = 0
                left_num_b = 0
                abort_cows_mth = 0
                if that_day.day == calendar.monthrange(that_day.year,
                                                       that_day.month)[1]:
                    # 怀孕——>流产
                    for d in range(1, 9):
                        _cond2ab = (HERDS.cdat > DATE0) & (HERDS.dcc < 250) & (
                            HERDS.abrt <= 0) & ((
                                (HERDS.dcc + 20) / 30).astype(int) == d)
                        _preg2ab_num = len(HERDS.loc[_cond2ab, 'bdat'])
                        abort_num = np.digitize(
                            np.random.random(_preg2ab_num),
                            [ABORT_RATE[1] * ABORT_STR[d], 0
                             ]).sum()  # 这里简单的只使用了成母牛的流产率
                        if abort_num:
                            abort_cows = HERDS[_cond2ab].sample(
                                n=abort_num).index
                            HERDS.loc[abort_cows, [
                                'rc', 'dcc', 'cdat', 'due', 'ddat', 'abrt'
                            ]] = [3, -1, DATE0, DATE0, DATE0, 1]
                            HERDS.loc[abort_cows, 'dslb'] = np.random.randint(
                                0, day, size=abort_num)
                            # TODO 流产日期可能不是最后的配种日期，这里有可能超过了节点日期
                            HERDS.loc[abort_cows, 'abdat'] = pd.to_datetime(
                                np.full(abort_num, that_day).astype(
                                    'datetime64') -
                                HERDS.loc[abort_cows, 'dslb'].values.astype(
                                    int).astype('timedelta64[D]'))
                            abort_cows_mth += abort_num
                    logger.debug(
                        "5.1.3 繁殖事件-6流产. 牛头数：{}".format(abort_cows_mth))

                    # 6死亡淘汰
                    rateC = LEAVE['cowCulRate'][mth] / 12
                    rateY = LEAVE['yongCulRate'][mth] / 12
                    rateBaby = LEAVE['calfCulRate'][mth] / 12
                    left_c = HERDS[HERDS.lact > 0].sample(frac=rateC).index
                    left_h = HERDS[(HERDS.lact == 0)
                                   & (HERDS.age_day >= AGEFB * 30.5)].sample(
                                       frac=rateY).index
                    left_b = HERDS[(HERDS.lact == 0)
                                   & (HERDS.age_day < AGEFB * 30.5)].sample(
                                       frac=rateBaby).index
                    HERDS.drop(left_c, inplace=True)
                    HERDS.drop(left_h, inplace=True)
                    HERDS.drop(left_b, inplace=True)
                    left_num_c = len(left_c)
                    left_num_h = len(left_h)
                    left_num_b = len(left_b)
                    logger.debug(
                        "5.1.3 繁殖事件-7死淘. 牛头数(成母/青年/犊牛)：{} {} {}".format(
                            left_num_c, left_num_h, left_num_b))
                    LEFT_COWS_C += left_num_c
                    LEFT_COWS_H += left_num_h
                    LEFT_COWS_B += left_num_b
                    cal_closing(that_day)
                    left_num = left_num_c + left_num_h + left_num_b
                elif not flag:
                    cal_closing(that_day)
                closing = len(HERDS)

                SUMMARY.loc[str(that_day.strftime("%m/%d/%Y"))] = [
                    opening, born_cows, babys_num, cows_to_clo, cows_to_dry,
                    abort_cows_mth, preg_num_this, bred_num, left_num, closing,
                    opening - left_num + babys_num - closing
                ]

                node_date = that_day

            else:
                logger.info("5. {}牧场牛群预测结束".format(parameter_suffix))

            # 结果保存
            SUMMARY.to_csv(parameter_suffix + '_SUMMARY.csv',
                           index=True,
                           encoding='utf_8_sig')
            HERDS['category'] = HERDS.apply(lambda x: update_category(
                x.lact, x.age_day, x.dim, x.dcc, x.ddat, x.cdat, end),
                                            axis=1)
            HERDS['category'] = HERDS['category'].astype(str)
            HERDS.to_csv(parameter_suffix + '_HERDS.csv',
                         index=True,
                         encoding='utf_8_sig')
            for s in [-1, 0, 1]:
                locals()['semen_usage_' + str(s + 1)].to_csv(
                    parameter_suffix + '_semen_usage_' + str(s + 1) + '.csv',
                    index=True,
                    encoding='utf_8_sig')
            LIVESTOCK.to_csv(parameter_suffix + '_LIVESTOCK.csv',
                             index=True,
                             encoding='utf_8_sig')
            logger.info("""6. 预测结果已保存.\n{0}_SUMMARY.csv:是各个生长节点的事件数;
                        {0}_HERDS.csv:是预测结束后的牛群明细；
                        {0}_semen_usage_0.csv:是冻精使用量，0是肉牛，1是常规，2是性控；
                        {0}_LIVESTOCK.csv:是牛群结构""".format(parameter_suffix))

        # LIVESTOCK
        LIVESTOCK.drop(LIVESTOCK.index[:], inplace=True)
        parameter(parameter_suffix)
        if method == 1:
            pre_treat(self.cows, snap_date)
            cal_closing(node_date)
            grow(snap_date, end, parameter_suffix)
        if method == 2:
            forecast2(self.cows, snap_date, end, parameter_suffix, AGEFB, WWTP,
                      HDR, PREG_RATE, ABORT_RATE, FEMAL, LEAVE, BRED_PLAN,
                      ABORT_STR)


#  存栏预测方法2
class Cattle():
    def __init__(self, id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                 lsbd, abdat, lsir, semen, isNewBorn=0):
        self.id = id
        self.rc = rc
        self.lact = min(lact, 7)
        self.tbrd = min(tbrd, 9)
        # logger.debug("bdat:{},cdat:{},fdat:{},ddat:{},lsbd:{},abdat:{}".format(type(bdat),type(cdat),type(fdat),type(ddat),type(lsbd),type(abdat)))
        # logger.debug("bdat:{},cdat:{},fdat:{},ddat:{},lsbd:{},abdat:{}".format((bdat),(cdat),(fdat),(ddat),(lsbd),(abdat)))
        if self.tbrd < 0:
            self.tbrd = 0
        if isinstance(bdat, datetime.datetime):
            self.bdat = bdat.date()
        elif isinstance(bdat, datetime.date):
            self.bdat = bdat

        if isinstance(cdat, datetime.datetime):
            self.cdat = cdat.date()
        elif isinstance(cdat, datetime.date):
            self.cdat = cdat

        if isinstance(fdat, datetime.datetime):
            self.fdat = fdat.date()
        elif isinstance(fdat, datetime.date):
            self.fdat = fdat

        if isinstance(ddat, datetime.datetime):
            self.ddat = ddat.date()
        elif isinstance(ddat, datetime.date):
            self.ddat = ddat

        if isinstance(lsbd, datetime.datetime):
            self.lsbd = lsbd.date()
        elif isinstance(lsbd, datetime.date):
            self.lsbd = lsbd

        if isinstance(abdat, datetime.datetime):
            self.abdat = abdat.date()
        elif isinstance(abdat, datetime.date):
            self.abdat = abdat

        self.closeDay = Args.DATE0
        self.abortTag = 0  # 流产标记 0未流产，1流产过
        self.cullTag = 0   # 流产标记 0未淘汰，1淘汰
        self.lsir = lsir
        self.semen = semen


class Summary():
    @classmethod
    def clear(cls):
        # 事件数
        cls.FRESH_NUM = 0
        cls.CLOSE_NUM = 0
        cls.DRY_NUM = 0
        cls.BRED_NUM = 0
        cls.PREG_NUM = 0
        cls.HEAT_NUM = 0
        cls.ABORT_NUM = 0
        cls.CULL_NUM = 0
        # 冻精
        cls.SEMEN_NUM_S = 0
        cls.SEMEN_NUM_B = 0
        cls.SEMEN_NUM_C = 0
        # 牛头数
        cls.AuditNum = 0
        cls.YongNum = 0
        cls.YongPregNum = 0
        cls.AuditPregNum = 0
        cls.AuditMilkNum = 0
        # 留养母牛
        cls.CalfNum = 0
        cls.CalfSet = set()

    @classmethod
    def init(cls):
        cls.CowNum = 0
        cls.BullNum = 0
        cls.CowSet = set()
        cls.BullSet = set()
        cls.clear()


class Args():
    @classmethod
    def init(cls):
        cls.DATE0 = datetime.date(1970, 1, 1)
        cls.HERDSDATE = datetime.date(1970, 1, 1)
        cls.BABY = [0, 0, 0, 0, cls.DATE0, cls.DATE0,
                    cls.DATE0, cls.DATE0, cls.DATE0, cls.DATE0, -1, 0]
        cls.CHK1 = 28
        cls.CHK2 = 60
        cls.CHK3 = 120
        cls.CHK4 = 210
        cls.mth = 0

    @classmethod
    def proInit(cls, AGEFB, WWTP, HDR, PREG_RATE, ABORT_RATE,
                FEMAL, LEAVE, BRED_PLAN, ABORT_STR):
        cls.AGEFB = AGEFB
        cls.WWTP = WWTP
        cls.HDR = HDR
        cls.PREG_RATE = PREG_RATE
        cls.ABORT_RATE = ABORT_RATE
        cls.FEMAL = FEMAL
        cls.LEAVE = LEAVE
        cls.BRED_PLAN = BRED_PLAN
        cls.ABORT_STR = [0.01, 0.23, 0.32, 0.12,
                         0.08, 0.05, 0.05, 0.06, 0.06, 0.02]


class Bull(Cattle):
    def __init__(self, id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                 lsbd, abdat, lsir, semen):
        super().__init__(id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                         lsbd, abdat, lsir, semen)
        Summary.BullNum += 1
        Summary.BullSet.add(self)


class Cow(Cattle):
    def __init__(self, id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                 lsbd, abdat, lsir, semen, isNewBorn=0):
        super().__init__(id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                         lsbd, abdat, lsir, semen, isNewBorn)
        self.preTreat()
        self.updateRC()
        if isNewBorn:
            Summary.CalfNum += 1
            Summary.CalfSet.add(self)
        else:
            Summary.CowNum += 1
            Summary.CowSet.add(self)

    def preTreat(self):  # ageDay, dslb, dcc/due/tchk, dim
        # tchk 孕检次数
        if self.bdat <= Args.DATE0:  # 无出生日期的
            self.ageDay = 0
        elif self.bdat > Args.DATE0:
            self.ageDay = (Args.HERDSDATE - self.bdat).days

        if self.tbrd > 0:  # 已配牛
            if self.rc != 3:  # 怀孕牛，已配牛，禁配牛
                self.dslb = (Args.HERDSDATE - self.lsbd).days
                if self.rc == 4 and self.dslb >= 42:  # 42=最晚孕检天数35+7
                    self.rc = 3  # 已配牛中实际是空怀或未孕检的,当做空怀牛处理

            if self.rc == 3:
                if self.abdat > Args.DATE0:  # 流产空怀牛
                    self.dslb = min((Args.HERDSDATE - self.abdat).days,
                                    (Args.HERDSDATE - self.lsbd).days)
                elif self.abdat <= Args.DATE0:   # 孕检空怀牛
                    dcc = (Args.HERDSDATE - self.lsbd).days
                    if dcc >= Args.CHK4:
                        self.dslb = dcc - Args.CHK4
                    elif dcc >= Args.CHK3:
                        self.dslb = dcc - Args.CHK3
                    elif dcc >= Args.CHK2:
                        self.dslb = dcc - Args.CHK2
                    elif dcc >= Args.CHK1:
                        self.dslb = dcc - Args.CHK1
        else:  # 未配牛
            self.dslb = -1

        if self.cdat > Args.DATE0:  # 怀孕牛
            self.dcc = (Args.HERDSDATE - self.cdat).days
            self.mcc = min(int((self.dcc+9)/30), 9)
            dcc = int(normalvariate(273, 3))
            self.due = self.cdat + datetime.timedelta(dcc)
            if self.due <= Args.HERDSDATE:
                days = randint(0, 9)
                self.due = Args.HERDSDATE + datetime.timedelta(days)
        elif self.cdat <= Args.DATE0:  # 未孕牛
            self.dcc = -1
            self.mcc = -1
            self.due = Args.DATE0

        if self.fdat > Args.DATE0:  # 产过犊的牛
            if self.ddat > Args.DATE0:  # 干奶牛
                self.dim = (self.ddat - self.fdat).days
            elif self.ddat <= Args.DATE0:  # 未干奶的牛
                self.dim = (Args.HERDSDATE - self.fdat).days
        elif self.fdat <= Args.DATE0:
            self.dim = -1

    def updateRC(self):
        if self.rc == 0:
            if self.fdat > Args.DATE0 and self.lsbd == Args.DATE0:
                self.rc = 2
        elif self.rc == 5:
            if self.ddat > Args.DATE0:
                self.rc = 6

    def grow(self, day, today, isMonthEnd=0):
        # today = Args.HERDSDATE + datetime.timedelta(day)
        self.ageDay += day

        if self.ageDay < Args.AGEFB:
            pass
        elif self.ageDay >= Args.AGEFB:
            self.setBredPlan()
            self.cowGrow(today, day, isMonthEnd)
        if isMonthEnd:
            self.cull()

    def setBredPlan(self):
        tbrdList = list(Args.BRED_PLAN['sex_tbrd'][Args.mth])
        lactList = list(Args.BRED_PLAN['sex_lact'][Args.mth])
        self.semen = 0
        for lact in lactList:
            if self.lact == int(lact):
                i = lactList.index(str(self.lact))
                if self.tbrd < int(tbrdList[i]):  # 性控冻精
                    self.semen = 1
        if self.lact >= int(Args.BRED_PLAN['beef_lact'][Args.mth]):
            beefRate = float(Args.BRED_PLAN['beef_rate'][Args.mth])
            if random() <= beefRate:
                self.semen = -1

    def cowGrow(self, today, day, isMonthEnd):
        if self.dcc >= 0:
            self.dcc += day
            self.mcc = min(int((self.dcc+9)/30), 9)
        if self.dim >= 0:
            self.dim += day
        if self.dslb >= 0:
            self.dslb += day

        if self.cdat > Args.DATE0:  # 怀孕牛成长
            if Args.HERDSDATE <= self.due <= today:
                self.fresh(today)
            elif self.dcc >= 250 and self.closeDay <= Args.DATE0:
                self.close(today - datetime.timedelta(self.dcc - 250))
            elif self.dcc >= 220 and self.lact > 0 and self.ddat <= Args.DATE0:
                self.dry(today - datetime.timedelta(self.dcc - 220))
            if isMonthEnd and self.abortTag == 0:  # 月底流产
                if self.lact == 0:
                    if random() <= Args.ABORT_RATE[0]*Args.ABORT_STR[self.mcc]:
                        self.abort(today)
                elif self.lact > 0:
                    if random() <= Args.ABORT_RATE[1]*Args.ABORT_STR[self.mcc]:
                        self.abort(today)
        elif self.rc == 2:  # 新产牛
            if self.dim >= Args.WWTP:
                self.heat(
                    today - datetime.timedelta(self.dim - Args.WWTP), today)
        elif self.rc == 3:  # 空怀牛
            if self.dslb >= 21:
                self.heat(today - datetime.timedelta(self.dslb - 21), today)
        elif self.rc == 4:  # 已配牛
            LT = "L{}T{}".format(min(self.lact, 1), self.tbrd)
            if random() <= Args.PREG_RATE[self.semen][LT]:
                self.preg(pregDay=today -
                          datetime.timedelta(self.dslb - Args.CHK1))
            else:  # 首次孕检未孕牛，置为空怀, 配后天数置为0
                self.rc = 3
                self.dslb = 0
        elif self.rc == 0:  # 未配青年牛
            if self.ageDay >= Args.AGEFB:
                self.heat(
                    today - datetime.timedelta(self.ageDay - Args.AGEFB),
                    today)

    def heat(self, heatDay, today):
        if heatDay < Args.HERDSDATE:
            heatDay = Args.HERDSDATE
        if random() <= Args.HDR['r'][self.lact]:
            self.bred(heatDay, today)
        Summary.HEAT_NUM += 1

    def bred(self, heatDay, today):
        self.tbrd += 1
        self.lsbd = heatDay
        self.dslb = (today - heatDay).days
        self.rc = 4
        Summary.BRED_NUM += 1
        if self.semen == -1:
            Summary.SEMEN_NUM_B += 1
        elif self.semen == 0:
            Summary.SEMEN_NUM_C += 1
        elif self.semen == 1:
            Summary.SEMEN_NUM_S += 1
        if self.tbrd > 9:
            self.dnb()
        self.tbrd = min(self.tbrd + 1, 9)

    def preg(self, pregDay):
        if pregDay < Args.HERDSDATE:
            pregDay = Args.HERDSDATE
        dcc = int(normalvariate(273, 3))
        self.cdat = pregDay
        self.due = pregDay + datetime.timedelta(dcc)
        self.dcc = self.dslb
        self.mcc = min(int((self.dcc+9)/30), 9)
        self.rc = 5
        Summary.PREG_NUM += 1

    def abort(self, aboDay):
        self.abdat = aboDay
        self.rc = 3
        self.dcc = -1
        self.due = Args.DATE0
        self.mcc = -1
        self.dslb = 0  # 流产牛配后天数重新置为0
        self.abortTag = 1  # 流产牛做上标记，一个胎次中最多流产1次
        Summary.ABORT_NUM += 1

    def dry(self, dryDay):
        if dryDay < Args.HERDSDATE:
            dryDay = Args.HERDSDATE
        self.ddat = dryDay
        self.rc = 6
        Summary.DRY_NUM += 1

    def close(self, closeDay):
        if closeDay < Args.HERDSDATE:
            closeDay = Args.HERDSDATE
        self.closeDay = closeDay
        Summary.CLOSE_NUM += 1

    def fresh(self, today):
        BABY = Args.BABY
        self.fdat = self.due
        self.dim = (today - self.due).days
        self.due = Args.DATE0
        self.dcc = -1
        self.cdat = Args.DATE0
        self.ddat = Args.DATE0
        self.dslb = -1
        self.rc = 2
        self.lact = min(self.lact+1, 7)
        self.tbrd = 0
        self.abortTag = 0   # 流产标记复位
        Summary.FRESH_NUM += 1
        if self.semen == 0:  # 常规
            femal = Args.FEMAL['com'][Args.mth]
        elif self.semen == 1:  # 性控
            femal = Args.FEMAL['sex'][Args.mth]
        elif self.semen == -1:  # 肉牛
            femal = Args.FEMAL['beef'][Args.mth]

        if random() <= femal:
            if self.semen > -1:
                Calf(BABY[0], BABY[1], BABY[2], BABY[3], self.fdat, BABY[5],
                     BABY[6], BABY[7], BABY[8], BABY[9], BABY[10], BABY[11], 1)
            else:
                Bull(BABY[0], BABY[1], BABY[2], BABY[3], self.fdat, BABY[5],
                     BABY[6], BABY[7], BABY[8], BABY[9], BABY[10], "HO")
        else:
            Bull(BABY[0], BABY[1], BABY[2], BABY[3], self.fdat, BABY[5],
                 BABY[6], BABY[7], BABY[8], BABY[9], BABY[10], BABY[11])

    def dnb(self):
        self.rc = 1

    def cull(self):
        if self.lact == 0:
            if self.ageDay < Args.AGEFB:
                cullRate = Args.LEAVE['calfCulRate'][Args.mth]/12
            else:
                cullRate = Args.LEAVE['yongCulRate'][Args.mth]/12
        else:  #
            cullRate = Args.LEAVE['cowCulRate'][Args.mth]/12

        if self.rc == 1 or random() <= cullRate:
            Summary.CULL_NUM += 1
            self.cullTag = 1


class Calf(Cow):
    def __init__(self, id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                 lsbd, abdat, lsir, semen, isNewBorn):
        super().__init__(id, rc, lact, tbrd, bdat, cdat, fdat, ddat,
                         lsbd, abdat, lsir, semen, isNewBorn)


def forecast2(herds, start, end, parameter_suffix,  AGEFB, WWTP, HDR,
              PREG_RATE, ABORT_RATE, FEMAL, LEAVE, BRED_PLAN, ABORT_STR):
    logger.info(datetime.datetime.now())
    logger.info("0 方法2初始化中")
    logger.info("怀孕率参数为：{}".format(PREG_RATE))
    Summary.init()    # 存储中间数据的类属性初始化
    Args.init()   # 存储常数的类属性初始化
    # 繁殖参数初始化
    Args.proInit(AGEFB*30.5, WWTP, HDR, PREG_RATE, ABORT_RATE,
                 FEMAL, LEAVE, BRED_PLAN, ABORT_STR)

    # 日期格式转换
    colu = ['bdat', 'cdat', 'fdat', 'ddat', 'lsbd', 'abdat']
    for dat in colu:
        herds[dat] = pd.to_datetime(herds[dat])

    # 预测参数
    start = start.date()
    end = end.date()

    cos = ["期初", "留养", "产犊", "围产", "干奶", "流产", "怀孕", "配种",
           "发情", "死淘", "期末", "检查", "成母", "青年", "青年怀孕", "成母怀孕",
           "成母泌乳", "常规冻精用量", "性控冻精用量", "肉牛冻精用量"]
    SUMMARY = pd.DataFrame(columns=cos)

    # 开始生长循环
    flag = True
    Args.HERDSDATE = start
    logger.info("1 初始化每头牛的属性")
    for item in herds.iterrows():
        Cow(item[1][0], item[1][1], item[1][2], item[1][3], item[1][4],
            item[1][5], item[1][6], item[1][7], item[1][8], item[1][9],
            item[1][10], item[1][11])
    logger.info("2 开始生长")
    while flag:
        opening = Summary.CowNum
        growDay = 16
        today = Args.HERDSDATE + datetime.timedelta(growDay)
        isMonthEnd = 0
        if today >= end:
            today = end
            if (today + datetime.timedelta(1)).day == 1:
                isMonthEnd = 1
            flag = False
        else:
            if today.day < growDay:  # 节点不是月底，且生长growDay天后和节点不在同一个月
                today = (datetime.date(today.year, today.month, 1) -
                         datetime.timedelta(1))
                isMonthEnd = 1
        passDays = (today - Args.HERDSDATE).days
        logger.info("----生长到{}，生长了{}天".format(today, passDays))
        Args.mth = int(datetime.datetime.strftime(today, "%Y%m"))
        if passDays == 0:
            break
        # 生长和统计
        columns = ["ID", "BDAT", "RC", "LACT", "TBRD", "CDAT", "FDAT", "DDAT",
                   "LSBD", "ABDAT", "DUE", "DIM", "DCC", "DSLB"]
        cowsList = []
        for cow in Summary.CowSet:
            cow.grow(passDays, today, isMonthEnd)  # 生长
            if cow.cullTag == 0:
                if cow.lact == 0:
                    Summary.YongNum += 1
                    if cow.cdat > Args.DATE0:
                        Summary.YongPregNum += 1
                else:
                    Summary.AuditNum += 1
                    if cow.cdat > Args.DATE0:
                        Summary.AuditPregNum += 1
                    if cow.fdat > Args.DATE0 and cow.ddat <= Args.DATE0:
                        Summary.AuditMilkNum += 1
                if not flag:
                    cowsList.append([cow.id, cow.bdat, cow.rc, cow.lact,
                                     cow.tbrd, cow.cdat, cow.fdat, cow.ddat,
                                     cow.lsbd, cow.abdat, cow.due, cow.dim,
                                     cow.dcc, cow.dslb])
        cowsFile = pd.DataFrame(columns=columns, data=cowsList)
        # 处理淘汰牛
        for cow in list(Summary.CowSet):
            if cow.cullTag == 1:
                Summary.CowSet.remove(cow)
        Summary.CowNum -= Summary.CULL_NUM
        # 处理留养母犊
        for calf in Summary.CalfSet:
            Summary.CowSet.add(calf)
        Summary.CalfSet.clear()
        Summary.CowNum += Summary.CalfNum
        Summary.YongNum += Summary.CalfNum

        # 统计事件数及牛头数
        SUMMARY.loc[today.strftime("%m/%d/%Y")] = [
            opening, Summary.CalfNum, Summary.FRESH_NUM, Summary.CLOSE_NUM,
            Summary.DRY_NUM, Summary.ABORT_NUM, Summary.PREG_NUM,
            Summary.BRED_NUM, Summary.HEAT_NUM, Summary.CULL_NUM,
            Summary.CowNum,
            opening - Summary.CULL_NUM - Summary.CowNum + Summary.CalfNum,
            Summary.AuditNum, Summary.YongNum, Summary.YongPregNum,
            Summary.AuditPregNum, Summary.AuditMilkNum, Summary.SEMEN_NUM_C,
            Summary.SEMEN_NUM_S, Summary.SEMEN_NUM_B]

        Summary.clear()

        Args.HERDSDATE = today
    logger.info("3 各时间节点的概况：\n{}".format(SUMMARY))
    cowsFile.to_csv("COWS.csv", header=True)
    logger.info(datetime.datetime.now())
