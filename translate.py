# -*- coding: utf-8 -*-
import xlrd
import json
def read_excel():
    # 打开文件
    try:
        workbook = xlrd.open_workbook(r'./translate.xlsx')
        # workbook = xlrd.open_workbook(r'https://shimo.im/sheets/gGTHQrTqwqQGXyRK/MODOC')
        # 获取所有sheet
        print ('所有的sheet:{}'.format(workbook.sheet_names())) # [u'sheet1', u'sheet2']
        #获取第一个sheet
        sheet_name= workbook.sheet_names()[0]
        print ('第一个sheet:{}'.format(sheet_name))
        # 根据sheet索引或者名称获取sheet内容
        sheet = workbook.sheet_by_name('工作表1')
        # sheet的名称，行数，列数
        nrows = sheet.nrows
        ncols = sheet.ncols
        print ('sheet的名称:{}，行数:{}，列数:{}'.format(sheet.name,nrows,ncols))
        langList = [{},{},{},{},{},{},{},{}]
        for x in range(nrows):
            for y in range(ncols):
                if x > 0 and y > 0:
                    key = sheet.cell(x,0).value
                    if (type(key) == float or type(key) == int):
                        key = str(int(key))
                    if(len(key)==0):
                        key = ('keyword{}'.format(x))
                    value = sheet.cell(x,y).value
                    item = langList[y-1]
                    item[key] = value
        # print ('最终数据:{}'.format(langList))
        nameArr = ['_lang/cn.js','_lang/en.js','_lang/pt.js','_lang/es.js','_lang/vi.js','_lang/it.js','_lang/cr.js','_lang/pl.js']
        for index in range(len(langList)):
            with open(nameArr[index],'w') as file_obj :
                print ('文件名称：{},文件内容：{}'.format(nameArr[index],langList[index]))
                json_string = json.dumps({'auto':langList[index]}, ensure_ascii=False, sort_keys=False, indent=4, separators=(',', ':'))
                file_obj.write('export default')
                file_obj.write(json_string)
                file_obj.close()
        print ('build finish')
    except IOError:
        print ('file is not exist')
    else:
        print ('finish')
if __name__ == '__main__':
    read_excel()