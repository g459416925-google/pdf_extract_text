from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
import pymupdf, re, os, openpyxl

path = './8.25/'
# path = './模板/'
files = os.listdir(path)

wb = Workbook()
ws = wb.active
ws.append([
    '文件名',
    '姓名',
    '身份证号',
    '起始日期',
    '结束日期'
    ])

write_data = []
un_write_data = []

for file in files:
    print('读取文件：' + file)
    doc = pymupdf.open(path + file)
    try:
        text = str(doc[3].get_text()).replace('\n',' ')

        # 新入职
        if text.startswith('劳动合同书 甲方（用人单位） 名 称：'):
            content = str(re.findall('联系电话： .*', text)[0]).removeprefix('联系电话： ').split(' ')
            name = content[0]
            id_num = content[-2 if len(content[-2]) != 11 else 2]

            # 提取固定期限
            text = str(doc[5].get_text()).replace('\n',' ')
            content = str(re.findall('信誉的事情； .*', text)[0]).removeprefix('信誉的事情； .*').split(' ')
            offset = 0 if len(content) ==11 else 1
            
            start_date = '-'.join(content[(2-offset):(5-offset)])
            end_date = '-'.join(content[(5-offset):(8-offset)])

            write_data.append([
                file,
                name,
                id_num,
                start_date,
                end_date
            ])

        # 实习
        elif text.startswith('贵州一品药业连锁有限公司 - 4 - 实习意向协议 甲'):
            content = str(re.findall('服从实习单位调派及管理，定期与实.*', text)[0]).removeprefix('服从实习单位调派及管理，定期与实').split(' ')
            offset = 0 if 1<len(content[8]) and len(content[8])<4 else 1
            name = content[8-offset]
            id_num = content[9-offset] 

            # 提取实习期限
            start_date = '-'.join(content[(2-offset):(5-offset)])
            end_date = '-'.join(content[(5-offset):(8-offset)])

            write_data.append([
                file,
                name,
                id_num,
                start_date,
                end_date
            ])
        
        # 续签
        elif text.startswith('（6）非经甲方事先书面批准，不得携出或使用甲方的钱款或财产作非职责用途 ；'):
            text = str(doc[0].get_text()).replace('\n',' ')
            content = str(re.findall('联系电话：.*', text)[0]).removeprefix('联系电话：').split(' ')
            name = content[1]
            id_num = content[-2]

            # 提取实习期限
            text = str(doc[2].get_text()).replace('\n',' ')
            content = str(re.findall('信誉的事情；.*', text)[0]).removeprefix('信誉的事情；').split(' ')
            start_date = '-'.join(content[2:5])
            end_date = '-'.join(content[5:8])

            write_data.append([
                file,
                name,
                id_num,
                start_date,
                end_date
            ])
        
        # 退休
        elif text.startswith('第4 页共10 页 一、合同期限 本合同的期限按以下二种情形中的第'):
            text = str(doc[0].get_text()).replace('\n',' ')
            content = str(re.findall('联系地址：.*', text)[0]).removeprefix('联系地址：').split(' ')
            name = content[1]
            id_num = content[3]

            # 提取实习期限
            text = str(doc[3].get_text()).replace('\n',' ')
            content = str(re.findall('高效完成甲方交办的各项事务；.*', text)[0]).removeprefix('高效完成甲方交办的各项事务；').split(' ')
            start_date = '-'.join(content[2:5])
            end_date = '-'.join(content[5:8])

            write_data.append([
                file,
                name,
                id_num,
                start_date,
                end_date
            ])
    except:
        print('-----------未识别文件：' + file)
        un_write_data.append([
            file,
            '未识别'
        ])

for i in write_data:
    ws.append(i)

for n in un_write_data:
    ws.append(n)

wb.save('数据读取.xlsx')
    
    
