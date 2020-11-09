import logging
from logging.handlers import TimedRotatingFileHandler as trf


# 创建log
f = '%(asctime)s - %(name)s - %(levelname)s -%(lineno)d- %(message)s'
formatter = logging.Formatter(f)
# 如果注释下面这一行，则在命令行和UI都不打印日志
logging.basicConfig(level=logging.DEBUG, format=f)  # format=f
logger = logging.getLogger(__name__)
# logger.setLevel(logging.INFO)   # 日志级别: CMD中的级别 及UI中的
# when = 'D' 不会滚动
fh = trf('ALTA_tool.log', when='MIDNIGHT', interval=1, backupCount=30)
fh.suffix = "%Y-%m-%d.log"   # log每天建立，保留30个
fh.setFormatter(formatter)
fh.setLevel(logging.DEBUG)    # 日志级别: 日志中的级别
logger.addHandler(fh)