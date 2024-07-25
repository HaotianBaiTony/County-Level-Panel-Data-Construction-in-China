# coding:utf8
'''
01_search.py
【查找所有涉及的表（半自动）】

先将用到的包全部解压，不能嵌套压缩包

运行程序，如果程序没整理到的，手动复制一份！！
'''
import pandas as pd
from pathlib import Path
import sys
import shutil
import json

dir_path = Path(sys.path[0])
start_dir = dir_path / 'start_data'
if start_dir.exists():
    shutil.rmtree(start_dir)


df = pd.read_excel(dir_path / 'county.xls')
df = df[df['sample'] == 1]

shenfen_list = list(set(df['省'].tolist()))

with open(dir_path / 'config.json', 'r', encoding='utf8') as f:
    s = f.read()
conf = json.loads(s)
print(conf)


def copy_file(src_path, dst_path):
    content = src_path.read_bytes()
    dst_path.write_bytes(content)


def add_start_year_file(year, file):
    start_dir = dir_path / 'start_data' / str(year)
    start_dir.mkdir(exist_ok=True, parents=True)
    copy_file(file, start_dir / file.name)
    print(file.name)


for year in range(2008, 2021):
    first_dir = f'{year}年中国县域统计年鉴'
    start_dir = Path(conf.get('excels_path')) / first_dir
    excel_files = start_dir.rglob('*.xls')
    for file in excel_files:
        if file.is_file():
            for sub_str in conf.get('xianshijuan_list'):
                if str(file).find(sub_str) > -1:
                    break
            else:
                continue
            for sub_str in shenfen_list:
                if str(file).find(sub_str) > -1:
                    break
            else:
                continue
            add_start_year_file(year, file)

print('--- end ---')
print()
