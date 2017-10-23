# encoding: utf-8

import os
import xlrd
import xlwt
import sys
import logging, logging.handlers
#固定的excel文件名
EXCELNAME = "flood forecast.xls"
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
def excelProcess(beginCol, endCol, daterRow, colWt, rowWt, path, resultFile, logFile):
    '''

    :param beginCol:读取开始行
    :param endCol:读取开始列
    :param daterRow:日期所在行
    :param colWt:开始写入列
    :param rowWt:开始写入行
    :param oldpath:文件夹路径
    :param resultFile:结果excel文件名
    :param logFile:日志文件名
    :return:
    '''
    d = {}
    log = logging.getLogger(logFile)
    loghandlers = logging.handlers.RotatingFileHandler(os.path.join(path, logFile), 'w', 0, 1)
    loghandlers.setFormatter(logging.Formatter('%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s'))
    log.addHandler(loghandlers)
    log.setLevel(logging.DEBUG)
    for item in os.listdir(path):
        itempath = os.path.join(path, item)
        if (os.path.isdir(itempath)
            and os.path.exists(os.path.join(itempath, EXCELNAME))):
            try:
                sheet = xlrd.open_workbook(os.path.join(itempath, EXCELNAME)).sheet_by_index(0)
            except Exception as e:
                log.error("failed to process " + os.path.join(itempath, EXCELNAME))
                print("failed to process " + os.path.join(itempath, EXCELNAME))
                continue
            log.info("processed " + os.path.join(itempath, EXCELNAME))
            for i in range(beginCol, endCol + 1):
                if (not d.get(sheet.cell(daterRow, i).value, False)):
                    d[sheet.cell(daterRow, i).value] = sheet.col_values(i)
    l = sorted(d.items(), key=lambda x: x[0])
    wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheetwt = wbk.add_sheet('sheet 1', cell_overwrite_ok=False)
    mindate = l[0][0]
    for item in l:
        rownum = rowWt + int(item[0] - mindate)
        sheetwt.write(rownum, colWt, xlrd.xldate.xldate_as_datetime(item[0], 0), style)
        writeline(sheetwt, rownum, colWt + 1, item[1][daterRow + 1:])
    wbk.save(os.path.join(path, resultFile))
if __name__ == "__main__":
    oldpath = sys.argv[1]
    path = oldpath.replace("/", "\\")
    try:
        excelProcess(7, 12, 50, 2, 6, path, 'discharges.xls', 'discharges.log')
        print("discharges aggregate  success!")
    except Exception as e:
        print(e)
    try:
        excelProcess(7, 12, 10, 2, 6, path, 'flood.xls', 'flood.log')
        print("flood aggregate  success!")
    except Exception as e:
        print(e)
    try:
        excelProcess(5, 5, 50, 2, 6, path, 'rainfall.xls', 'rainfall.log')
        print("rainfall aggregate  success!")
    except Exception as e:
        print(e)
