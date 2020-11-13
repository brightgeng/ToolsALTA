import os
import pandas as pd
from log import logger
import traceback
from matching import YMD, DIR


# 工具1：合并体型
def merge_body(farm_name, last_month_file, body_date, new_result):
    def act(farm_name, last_month_file, body_date, new_result):
        global DIR
        ym = body_date[0:6]
        new_body_file = os.path.join(DIR, 'Match_files',
                                     farm_name + '_体型明细_' + ym + '.csv')
        this_month_data = pd.DataFrame()
        count = 1
        for file in new_result.strip('\n').split('\n'):
            if file:
                f_str = os.path.join(DIR, 'body_by_month', 'old', file)
                try:
                    f = pd.DataFrame(pd.read_csv(f_str, encoding='gbk'))
                except UnicodeDecodeError:
                    f = pd.DataFrame(pd.read_csv(f_str, encoding='utf_8_sig'))
                this_month_data = this_month_data.append(f)
                logger.info('已处理第{}个csv文件，包括{}头牛'.format(count, len(f)))
                count += 1
        logger.info('所有csv文件合并完成')
        this_month_data.drop(['TableID'], axis=1, inplace=True)
        this_month_data.drop_duplicates('牛号', keep='last', inplace=True)
        logger.info('合并后的文件去重完成')
        this_month_data['日期'] = body_date
        logger.info('合并后的文件添加日期')
        cow_num = len(this_month_data)
        this_month_file = os.path.join(
            DIR, 'body_by_month',
            body_date + '_' + farm_name + '_体型数据_' + str(cow_num) + '头.xlsx')
        this_month_data.to_excel(this_month_file, index=False, encoding='utf8')
        logger.info('本月数据已单独保存在:\n  {}'.format(this_month_file))

        last_month_data = pd.DataFrame(
            pd.read_csv(os.path.join(DIR, 'Match_files', last_month_file),
                        encoding='utf_8_sig'))
        new_body_data = last_month_data.append(this_month_data)
        logger.info('将本月数据追加到历史数据中')
        new_body_data.drop_duplicates('牛号', keep='last', inplace=True)
        logger.info('累积数据去重完成')
        new_body_data.to_csv(new_body_file, index=False, encoding='utf_8_sig')
        logger.info('累积数据已单独保存在:\n  {}'.format(new_body_file))

        result = ('0\n本月体型汇总数据已保存到：\n  ' + this_month_file + '\n' +
                  '最新体型数据保存到：\n  ' + new_body_file + '\n')
        return result

    if farm_name == last_month_file.split('_')[0]:
        try:
            logger.info('\n\n开始执行程序1：合并体型数据')
            logger.info('参数为:\n  牧场: ' + farm_name + '\n  历史体型明细: ' +
                        last_month_file + '\n  做体型的日期: ' + body_date +
                        '\n  本月数据文件: ' +
                        str(new_result.strip('\n').split('\n')))
            r = act(farm_name, last_month_file, body_date, new_result)
            logger.info('执行程序1完成：合并体型数据！')
        except Exception:
            logger.error("错误:\n{}".format(traceback.format_exc()))
            return "1错误:\n{}".format(traceback.format_exc())
        else:
            return r
    else:
        logger.error('所选牧场和历史体型明细文件不匹配，请检查后再重新开始')
        return '1错误:\n所选牧场和历史体型明细文件不匹配，请检查后再重新开始'


# 工具2：提取配种记录
def extract_log(files):
    def main(file):
        full_file_name_i = os.path.join(DIR, 'BredLog', file)
        bred_log = pd.DataFrame(pd.read_excel(full_file_name_i))
        try:
            bred_log['日龄'] = None
            bred_log['备注'] = None
            bred_log = bred_log[[
                '耳号', '胎次', '胎次', '配次', '配后状态', '舍组床', '配种员', '事件发生日期',
                '事件发生日期', '参配时产后天数', '日龄', '配前所属类别', '与配公牛', '配种模式', '配后状态',
                '上次配种', '备注'
            ]]
            bred_log = bred_log.sort_values(by='事件发生日期', ascending=True)
        except KeyError:
            bred_log = bred_log[[
                '牛号', '当前胎次', '事件胎次', '配次', '繁育状态', '当前牛舍', '工作人员', '事件日期',
                '录入日期', '泌乳天数', '日龄', '牛只类别', '配种公牛号', '配种方式', '配种结果',
                '上次配种日期', '备注'
            ]]
            bred_log = bred_log.sort_values(by='录入日期', ascending=True)
        res = os.path.join(DIR, 'BredLog', file[:-5] + '_结果.csv')
        bred_log.to_csv(res, encoding='utf_8_sig')
        return res

    count = 1
    res = ""
    logger.info('\n\n开始执行程序2：提取配种记录')
    logger.info('参数为：\n  配种记录文件' + str(files.strip('\n').split("\n")))
    for f in files.strip('\n').split("\n"):
        if f:
            try:
                logger.info('开始处理第{}个配种记录'.format(count))
                act = main(f)
                logger.info('已处理第{}个配种记录, 结果保存在：{}'.format(count, act))
                res = res + '\n' + act
                count += 1
            except Exception:
                logger.error('错误:\n{}'.format(traceback.format_exc()))
                return '1错误:\n{}'.format(traceback.format_exc())
    logger.info('执行程序2完成：提取配种记录！')
    return '0执行完成, 结果保存在：\n{}'.format(res)


# 工具3：提取定位文件
def mger_pos(files):
    col_str = ("牧场", "分组", "产奶量", "脂肪", "蛋白质", "体高", "胸宽", "体深", "乳用特征", "尻角度",
               "尻宽", "后肢侧视", "后肢后视", "蹄角度", "前乳房", "后乳房高度", "后乳房宽度", "悬韧带",
               "乳房深度", "乳头位置", "乳头长度", "后乳头位置", "生产寿命", "女儿怀孕率", "体细胞计数",
               "公牛产犊难易度", "女儿产犊难易度", "公牛死胎率", "女儿死胎")
    summary_pd = pd.DataFrame(columns=col_str)

    def act(file):
        nonlocal summary_pd
        filename = file.split(".")[0].split("_")
        path = os.path.join(DIR, "Postions", file)
        data = pd.DataFrame(
            pd.read_excel(path,
                          header=None,
                          skiprows=14,
                          skipfooter=0,
                          usecols='E:F'))
        tmp = pd.DataFrame(
            {
                "牧场": filename[1],
                "分组": filename[2],
                "产奶量": data[4][0],
                "脂肪": data[4][1],
                "蛋白质": data[4][2],
                "体高": data[5][29],
                "胸宽": data[5][30],
                "体深": data[5][31],
                "乳用特征": data[5][32],
                "尻角度": data[5][33],
                "尻宽": data[5][34],
                "后肢侧视": data[5][35],
                "后肢后视": data[5][36],
                "蹄角度": data[5][37],
                "前乳房": data[5][38],
                "后乳房高度": data[5][39],
                "后乳房宽度": data[5][40],
                "悬韧带": data[5][41],
                "乳房深度": data[5][42],
                "乳头位置": data[5][43],
                "乳头长度": data[5][44],
                "后乳头位置": data[5][45],
                "生产寿命": data[4][12],
                "女儿怀孕率": data[4][13],
                "体细胞计数": data[4][14],
                "公牛产犊难易度": data[4][15],
                "女儿产犊难易度": data[4][16],
                "公牛死胎率": data[4][17],
                "女儿死胎": data[4][18]
            },
            index=[0])
        summary_pd = summary_pd.append(tmp, ignore_index=True, sort=False)

    def main(files):
        nonlocal summary_pd
        count = 1
        for f in files.strip('\n').split("\n"):
            if f:
                act(f)
                logger.info('已处理第{}个定位文件, 文件名：{}'.format(count, f))
                count += 1
        summary_pd.sort_values(by=['牧场', '分组'], inplace=True)
        # 保存
        if len(summary_pd) > 0:
            path = os.path.join(DIR, 'Postions')
            result_file = os.path.join(path, YMD + '_定位汇总' + '_postions.csv')
            summary_pd.to_csv(result_file,
                              index=False,
                              header=True,
                              encoding='utf_8_sig',
                              line_terminator='\r\n')
            res = '0定位汇总文件已保存：{}'.format(result_file)
            logger.info(res)
            logger.info('执行程序3完成：提取定位文件！')
            return res

    logger.info('\n\n开始执行程序3：提取定位文件')
    logger.info('参数为：\n  定位文件：' + str(files.strip('\n').split("\n")))
    try:
        r = main(files)
    except Exception:
        logger.error('1错误:\n{}'.format(traceback.format_exc()))
        return "1错误:\n{}".format(traceback.format_exc())
    else:
        return r


# 工具4：合并选配文件
def merge_files(file_name_pre='', file_list=[1, 2, 5]):
    logger.info("\n\n开始执行程序4：合并选配文件")
    try:
        DIR = (r'c:\Users\runda\OneDrive - RundaTech\04 工作\0405 艾格威贸易\ALTA'
               r'\ALTA\12 汇报\现代牧业\首配准确性\蚌埠')
        herds_df = pd.DataFrame()
        file_name = file_name_pre
        herds_file = os.path.join(DIR, file_name + '汇总' + '.xlsx')
        for i in file_list:
            logger.info("开始读取文件：{}".format(i))
            herds_file_xlsx = os.path.join(DIR, file_name + str(i) + '.xlsx')
            df = pd.DataFrame(pd.read_excel(herds_file_xlsx))
            herds_df = herds_df.append(df, ignore_index=True, sort=False)
        logger.info("开始保存合并后的文件")
        # utf_8_sig
        herds_df.to_excel(herds_file, index=False, encoding='utf-8')
        print("保存完成!")
    except Exception:
        logger.error('1错误:\n{}'.format(traceback.format_exc()))
        return "1错误:\n{}".format(traceback.format_exc())
    else:
        return "0文件已保存在：{}".format(herds_file)


# 工具5：
