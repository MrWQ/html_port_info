# -*- coding: utf-8 -*-

import re,os,xlwt,xlrd
#
# dir_path = './host'
# file_list = os.listdir(dir_path)
file_list = os.listdir()

try:
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('Sheet1')
    worksheet.write(0, 0, 'ip')
    worksheet.write(0, 1, '端口')
    worksheet.write(0, 2, '协议')
    worksheet.write(0, 3, '服务')
    worksheet.write(0, 4, '状态')
    workbook.save('results.xls')
except Exception as e:
    print(e)
    print('创建文件失败')
try:
    for html_file in file_list:
        if '.html' in html_file:
            # file_path = dir_path +'/' + html_file
            file_path =  html_file
            if os.path.isfile(file_path):
                # # excel 写入
                # workbook = xlwt.Workbook(encoding='ascii')
                # worksheet = workbook.add_sheet('Sheet1')
                # worksheet.write(0, 0, 'ip')
                # worksheet.write(0, 1, '端口')
                # worksheet.write(0, 2, '协议')
                # worksheet.write(0, 3, '服务')
                # worksheet.write(0, 4, '状态')

                print('read file:' + html_file)
                with open(file_path,'r', encoding='UTF-8') as file:
                    file_text = file.read()
                    pattern = re.compile(r' 端口信息</div>(.*?)</table>', re.S)
                    text = re.findall(pattern, file_text)
                    if text != []:
                        text = text[0].replace('\t','').replace('\n','')
                        pattern = re.compile(r'<tbody>(.*?)</tbody>', re.S)
                        text = re.findall(pattern, text)
                        text = text[0]
                        # print(text)
                        pattern = re.compile(r'<tr(.*?)</tr>', re.S)
                        text_list_row = re.findall(pattern, text)
                        data = xlrd.open_workbook('results.xls')
                        table = data.sheet_by_name(u'Sheet1')  # 通过名称获取
                        row = table.nrows
                        # print(row)
                        for row_count in range(0,len(text_list_row)):
                            text = text_list_row[row_count].replace(' ','')
                            worksheet.write(row + row_count, 0, html_file.replace('.html',''))
                            pattern = re.compile(r'<td>(.*?)</td>', re.S)
                            text_list_cln = re.findall(pattern, text)
                            for count in range(0,len(text_list_cln)):
                                worksheet.write(row + row_count , count + 1, text_list_cln[count])
                                # print(text_list_cln[count])
                    else:
                        pass
                workbook.save('results.xls')
except Exception as e:
    print(e)
    print('new error')