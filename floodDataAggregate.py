# encoding: utf-8

"""
@author: 'liuyuefeng'
@file: floodDataAggregate.py
@time: 2017/9/10 21:37
"""
import os
import xlrd
import xlwt
import sys
import logging
ROOTPATH = ""
#固定的excel文件名
EXCELNAME = "flood forecast.xls"
#读取开始列
BEGINCOL = 7
#读取结束列
ENDCOL = 12
#日期所在行
DATEROW = 10
#写入开始列
COL_WT = 2
#写入开始行
ROW_WT = 6
#excel日期格式
style = xlwt.XFStyle()
style.num_format_str = 'yyyy/mm/dd'
def writeline(sheet, row, col, line):
    """
    写入一行数据
    :param sheet:
    :param row:
    :param col:开始列
    :param line:要写入的列表
    :return:
    """
    for i in range(len(line)):
        sheet.write(row, col + i, line[i])

if __name__ == "__main__":
    try:
        #所有数据的字典
        d = {}
        #传入参数是文件路径
        oldpath = sys.argv[1]
        # oldpath = 'C:/Users/liuyuefeng/Desktop/201501/201501'
        path = oldpath.replace("/", "\\")
        # path = 'C:\\Users\\liuyuefeng\\Desktop\\201501\\201501'
        logging.basicConfig(level=logging.DEBUG,
                            format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                            datefmt='%a, %d %b %Y %H:%M:%S',
                            filename=os.path.join(path, 'info.log'),
                            filemode='w')
        for item in os.listdir(path):
            itempath = os.path.join(path, item)
            if(os.path.isdir(itempath)
                   and os.path.exists(os.path.join(itempath, EXCELNAME))):
                try:
                    sheet = xlrd.open_workbook(os.path.join(itempath, EXCELNAME)).sheet_by_index(0)
                except Exception as e:
                    logging.error("failed to process " + os.path.join(itempath, EXCELNAME))
                    print("failed to process " + os.path.join(itempath, EXCELNAME) )
                    continue
                logging.info("processed " + os.path.join(itempath, EXCELNAME))
                for i in range(BEGINCOL, ENDCOL + 1):
                    if( not d.get(sheet.cell(10, i).value, False)):
                        d[sheet.cell(10, i).value] = sheet.col_values(i)
        l = sorted(d.items(), key=lambda x:x[0])
        wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheetwt = wbk.add_sheet('sheet 1', cell_overwrite_ok=False)
        mindate = l[0][0]
        for item in l:
            rownum = ROW_WT + int(item[0] - mindate)
            sheetwt.write(rownum, COL_WT, xlrd.xldate.xldate_as_datetime(item[0], 0), style)
            writeline(sheetwt, rownum, COL_WT + 1, item[1][11:])
        wbk.save(os.path.join(path, "result.xls"))
    except Exception as e:
        logging.error(e.with_traceback())
        print(e.with_traceback())





