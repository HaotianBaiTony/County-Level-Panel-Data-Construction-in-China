# coding:utf8
'''03_merge.py
【合并数据（自动）】

根据process_data，生成result.xlsx
'''
import pandas as pd
from pathlib import Path
import sys
import json

dir_path = Path(sys.path[0])
process_dir = dir_path / 'process_data'


data = []
for year in range(2008, 2021):
    json_dir = process_dir / f'{year}'
    files = json_dir.rglob('*.json')
    for file in files:
        with open(file, 'r', encoding='utf8') as f:
            s = f.read()
        jss = json.loads(s)
        if '规模以上工业企业 ' in jss:
            jss['规模以上工业企业'] = jss['规模以上工业企业 ']
        if '普通中学在村学生' in jss:
            jss['普通中学在校学生'] = jss['普通中学在村学生']

        data.append({
            '省份': file.stem.split('-')[0],
            '县市': file.stem.split('-')[1],
            '年份': year,
            '地区生产总值': jss.get('地区生产总值'),
            '火车站数量': jss.get('火车站数量'),
            '高铁动车站数量': jss.get('高铁动车站数量'),
            '车站总数': jss.get('车站总数'),
            '地方一般公共预算收入': jss.get('地方一般公共预算收入'),
            '地方一般公共预算支出': jss.get('地方一般公共预算支出'),
            '住户存款余额': jss.get('住户存款余额'),
            '规模以上工业企业': jss.get('规模以上工业企业'),
            '固定电话用户': jss.get('固定电话用户'),
            '普通中学在校学生': jss.get('普通中学在校学生'),
            '小学在校学生': jss.get('小学在校学生'),
            '医疗卫生机构床位': jss.get('医疗卫生机构床位'),
            '提供住宿的民政服务机构': jss.get('提供住宿的民政服务机构'),
            '提供住宿的民政服务机构床位数': jss.get('提供住宿的民政服务机构床位数'),
        })
        print(year, file.stem)
df = pd.DataFrame(data)
df.to_excel(dir_path / 'result.xlsx', index=None)

print('--- end ---')
print()
