# coding:utf8
'''02_tiqu.py
【提取数据（半自动）】

如取不到的地区，需检查原因。更新【county2.xlsx】并重新运行本文件。
'''
import xlwings as xw
import pandas as pd
from pathlib import Path
import sys
import json
import threading

dir_path = Path(sys.path[0])
start_dir = dir_path / 'start_data'


df = pd.read_excel(dir_path / 'county2.xlsx')
county_df = df.copy()
df = df[df['sample'] == 2]
sx_df = df[['省', 'NAME']]
# print('郧西县' in df['NAME'].values)
shenfen_list = list(set(df['省'].tolist()))


def save_and_read(fpath, df):
    df.to_excel(fpath, index=None)
    df = pd.read_excel(fpath, header=None)
    return df


def col_number_to_name(col_number):
    string = ""
    while col_number > 0:
        col_number, remainder = divmod(col_number - 1, 26)
        string = chr(65 + remainder) + string
    return string


def find_value(df, value):
    positions = []
    for idx, row in df.iterrows():
        for col in df.columns:
            if not isinstance(df.loc[idx, col], str):
                continue
            if (df.loc[idx, col].strip() == value) or (len(value) > 3 and value in df.loc[idx, col]):
                positions.append((idx, col))
    return positions


def get_positions_from_df(df):
    p1 = find_value(df, '地区生产总值')
    p5 = find_value(df, '地方财政一般预算收入')
    if p5 == []:
        p5 = find_value(df, '公共预算收入')
    if p5 == []:
        p5 = find_value(df, '公共财政收入')
    p6 = find_value(df, '地方财政一般预算支出')
    if p6 == []:
        p6 = find_value(df, '公共预算支出')
    if p6 == []:
        p6 = find_value(df, '公共财政支出')
    p7 = find_value(df, '城乡居民储蓄存款余额')
    if p7 == []:
        p7 = find_value(df, '储蓄存款余额')
    p8 = find_value(df, '规模以上工业企业')
    if p8 == []:
        p8 = find_value(df, '工业企业单位数')
    p9 = find_value(df, '本地电话年末用户')
    if p9 == []:
        p9 = find_value(df, '固定电话用户')
    p10 = find_value(df, '普通中学在校学生')
    p11 = find_value(df, '小学在校学生')
    p12 = find_value(df, '卫生院床位数')
    if p12 == []:
        p12 = find_value(df, '卫生机构床位')
    p13 = find_value(df, '社会福利院数')
    if p13 == []:
        p13 = find_value(df, '收养性单位数')
    if p13 == []:
        p13 = find_value(df, '社会福利院数')
    if p13 == []:
        p13 = find_value(df, '社会工作机构')
    p14 = find_value(df, '社会福利院床位数')
    if p14 == []:
        p14 = find_value(df, '收养性单位床位数')
    if p14 == []:
        p14 = find_value(df, '社会福利院床位数')
    if p14 == []:
        p14 = find_value(df, '社会工作机构床位')
    return p1, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14


def da_kuang_biao(name, year, sheet, max_row, max_column):
    df_path = dir_path / 'temp' / f'{year}-{name}-{sheet.name}.xlsx'
    if df_path.exists():
        df = pd.read_excel(df_path, header=None)
    else:
        col_str = col_number_to_name(max_column)
        range_data = sheet.range(f'A1:{col_str}{max_row}').value
        df = pd.DataFrame(range_data[1:], columns=range_data[0])
        df = save_and_read(df_path, df)

    df.columns = [i for i in range(df.shape[1])]
    if df.shape[1] < 7 or df.shape[0] < 35:
        return
    df[0] = df[0].str.strip()
    p1, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14 = get_positions_from_df(df)

    for shenfen in shenfen_list:
        if shenfen in name:
            t = sx_df[sx_df['省'] == shenfen]
            for xianshi in t['NAME'].values:
                json_path = json_dir / f'{shenfen}-{xianshi}.json'
                if json_path.exists():
                    continue
                positions = find_value(df, xianshi)
                # print(xianshi, positions)
                if len(positions) != 1:
                    positions = find_value(df, xianshi[:4])
                    if len(positions) != 1:
                        county_df.loc[(county_df['省'] == shenfen) & (
                            county_df['NAME'] == xianshi), f'drop-{year}'] = 1
                        continue
                county_df.loc[(county_df['省'] == shenfen) & (
                    county_df['NAME'] == xianshi), f'drop-{year}'] = ''
                py = 0
                for idx, p in enumerate(p5):
                    if positions[0][1] - p[1] > 0 and positions[0][1] - p[1] < 8:
                        py = idx
                        break
                try:
                    dqsczz = df.loc[p1[py][0], positions[0][1]]
                except:
                    dqsczz = ''
                try:
                    gddyyh = df.loc[p9[py][0], positions[0][1]]
                except:
                    gddyyh = ''
                data = {
                    '地区生产总值': dqsczz,
                    '火车站数量': '',
                    '高铁动车站数量': '',
                    '车站总数': '',
                    '地方一般公共预算收入': df.loc[p5[py][0], positions[0][1]],
                    '地方一般公共预算支出': df.loc[p6[py][0], positions[0][1]],
                    '住户存款余额': df.loc[p7[py][0], positions[0][1]],
                    '规模以上工业企业': df.loc[p8[py][0], positions[0][1]],
                    '固定电话用户': gddyyh,
                    '普通中学在校学生': df.loc[p10[py][0], positions[0][1]],
                    '小学在校学生': df.loc[p11[py][0], positions[0][1]],
                    '医疗卫生机构床位': df.loc[p12[py][0], positions[0][1]],
                    '提供住宿的民政服务机构': df.loc[p13[py][0], positions[0][1]],
                    '提供住宿的民政服务机构床位数': df.loc[p14[py][0], positions[0][1]],
                }
                with open(json_path, 'w', encoding='utf8') as f:
                    f.write(json.dumps(data))
                print(year, f'{xianshi}.json')
    county_df.to_excel(dir_path / 'county2.xlsx', index=None)


def da_shu_biao(name, year, sheet, max_row, max_column):
    df_path = dir_path / 'temp' / f'{year}-{name}-{sheet.name}.xlsx'
    if df_path.exists():
        df = pd.read_excel(df_path, header=None)
    else:
        col_str = col_number_to_name(max_column)
        range_data = sheet.range(f'A1:{col_str}{max_row}').value
        df = pd.DataFrame(range_data[1:], columns=range_data[0])
        df = save_and_read(df_path, df)
    df.columns = [i for i in range(df.shape[1])]
    if df.shape[1] < 7 or df.shape[0] < 35:
        return
    df[0] = df[0].str.strip()
    p1, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14 = get_positions_from_df(df)

    for shenfen in shenfen_list:
        if shenfen in name:
            t = sx_df[sx_df['省'] == shenfen]
            for xianshi in t['NAME'].values:
                json_path = json_dir / f'{shenfen}-{xianshi}.json'
                if json_path.exists():
                    continue
                positions = find_value(df, xianshi)
                if len(positions) != 1:
                    positions = find_value(df, xianshi[:4])
                    if len(positions) != 1:
                        county_df.loc[(county_df['省'] == shenfen) & (
                            county_df['NAME'] == xianshi), f'drop-{year}'] = 1
                        continue
                county_df.loc[(county_df['省'] == shenfen) & (
                    county_df['NAME'] == xianshi), f'drop-{year}'] = ''
                py = 0
                for idx, p in enumerate(p5):
                    if p[0] - positions[0][0] > 0 and p[0] - positions[0][0] < 35:
                        py = idx
                        break
                try:
                    dqsczz = df.loc[p1[py][0], positions[0][1]]
                except:
                    dqsczz = ''
                try:
                    gddyyh = df.loc[p9[py][0], positions[0][1]]
                except:
                    gddyyh = ''
                data = {
                    '地区生产总值': dqsczz,
                    '火车站数量': '',
                    '高铁动车站数量': '',
                    '车站总数': '',
                    '地方一般公共预算收入': df.loc[p5[py][0], positions[0][1]],
                    '地方一般公共预算支出': df.loc[p6[py][0], positions[0][1]],
                    '住户存款余额': df.loc[p7[py][0], positions[0][1]],
                    '规模以上工业企业': df.loc[p8[py][0], positions[0][1]],
                    '固定电话用户': gddyyh,
                    '普通中学在村学生': df.loc[p10[py][0], positions[0][1]],
                    '小学在校学生': df.loc[p11[py][0], positions[0][1]],
                    '医疗卫生机构床位': df.loc[p12[py][0], positions[0][1]],
                    '提供住宿的民政服务机构': df.loc[p13[py][0], positions[0][1]],
                    '提供住宿的民政服务机构床位数': df.loc[p14[py][0], positions[0][1]],
                }
                with open(json_path, 'w', encoding='utf8') as f:
                    f.write(json.dumps(data))
                print(year, f'{xianshi}.json')
    county_df.to_excel(dir_path / 'county2.xlsx', index=None)


def open_excel(file_path, year):
    with pool_sema:
        with xw.App(visible=True) as app:
            book = app.books.open(file_path)
            for sheet in book.sheets:
                if sheet.name == 'CNKI':
                    continue
                row1 = sheet.range('A1').end('down').row
                row2 = sheet.range('A6').end('down').row
                max_row = max(row1, row2)
                while 1:
                    try:
                        last_row = sheet.range(
                            'A{}'.format(max_row+5)).end('down').row
                    except:
                        break
                    if last_row > 2000:
                        break
                    if last_row - max_row >= 50:
                        break
                    max_row = last_row
                max_column = sheet.range('A1').end('right').column
                while 1:
                    try:
                        last_col2 = sheet.range('{}1'.format(
                            col_number_to_name(max_column+4))).end('right').column
                        last_col1 = sheet.range('{}3'.format(
                            col_number_to_name(max_column+4))).end('right').column
                        last_col = max(last_col1, last_col2)
                    except:
                        break
                    if last_col > 250:
                        break
                    if last_col - max_column >= 15:
                        break
                    max_column = last_col
                if max_row < 2000:
                    max_row = max_row + 5
                if max_row > 10000:
                    max_row = 10000
                # print(max_row, max_column)
                # flag = False
                # if max_row >= 50:
                #     da_shu_biao(file_path.stem, year, sheet, max_row, max_column)
                # else:
                #     da_kuang_biao(file_path.stem, year, sheet, max_row, max_column)
                #     flag = True
                max_row = 3000
                max_column = 256
                if year >= 2013:  # and flag:
                    da_shu_biao(file_path.stem, year,
                                sheet, max_row, max_column)
                else:
                    da_kuang_biao(file_path.stem, year,
                                  sheet, max_row, max_column)


process_dir = dir_path / 'process_data'
for year in range(2008, 2021):
    county_df[f'drop-{year}'] = ''
    json_dir = process_dir / str(year)
    json_dir.mkdir(exist_ok=True, parents=True)
    year_dir = start_dir / str(year)
    pool_sema = threading.BoundedSemaphore(1)
    threads = []
    for fpath in year_dir.rglob('*'):
        for sub_str in shenfen_list:
            if str(fpath).find(sub_str) > -1:
                break
        else:
            continue
        print(fpath)
        t = threading.Thread(target=open_excel, args=(fpath, year))
        threads.append(t)
    for t in threads:
        t.start()
    for t in threads:
        t.join()

print('--- end ---')
print()
