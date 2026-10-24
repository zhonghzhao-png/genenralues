# ==================== 主程序.py - 所有核心函数 ====================
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference, LineChart
from openpyxl.chart.label import DataLabelList
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)




def _安全保存工作簿(wb, 路径='数据分析.xlsx'):
    """先保存到同目录临时 xlsx，再原子替换目标文件，降低中断造成文件损坏的风险。"""
    临时路径 = 路径 + '.tmp.xlsx'
    try:
        wb.save(临时路径)
        os.replace(临时路径, 路径)
    finally:
        if os.path.exists(临时路径):
            try:
                os.remove(临时路径)
            except OSError:
                pass


# ==================== 颜色定义 ====================
fille1 = PatternFill('solid', fgColor='54FF9F')
fille2 = PatternFill('solid', fgColor='00FF00')
fille3 = PatternFill('solid', fgColor='FFFF00')
fille4 = PatternFill('solid', fgColor='FFC125')
fille5 = PatternFill('solid', fgColor='FFC0CB')
fille6 = PatternFill('solid', fgColor='FFB6C1')
fille7 = PatternFill('solid', fgColor='00FFFF')
fille8 = PatternFill('solid', fgColor='F0E68C')


# ==================== 姓名修正函数 ====================

def 修正姓名(df, 参考名单路径=None, 打印详情=True):
    """用参考名单按学号修正 df 中的姓名。"""
    if df is None or df.empty:
        return df
    if '学号' not in df.columns:
        return df

    if 参考名单路径 is None:
        try:
            from 数据分析 import 控制页配置
            参考名单路径 = 控制页配置.get('名单路径', None)
        except:
            参考名单路径 = None

    if not 参考名单路径 or not os.path.exists(参考名单路径):
        return df

    try:
        df_名单 = pd.read_excel(参考名单路径)
    except Exception as e:
        print(f"    ⚠️ 读取参考名单失败: {e}")
        return df

    名单学号列 = None
    名单姓名列 = None
    for col in df_名单.columns:
        if 名单学号列 is None and ('学号' in str(col) or 'id' in str(col).lower() or '学籍号' in str(col)):
            名单学号列 = col
        elif 名单姓名列 is None and ('姓名' in str(col) or 'name' in str(col).lower()):
            名单姓名列 = col

    if 名单学号列 is None or 名单姓名列 is None:
        print(f"    ⚠️ 参考名单中未找到学号或姓名列，跳过修正")
        return df

    df_名单 = df_名单[[名单学号列, 名单姓名列]].copy()
    df_名单 = df_名单.rename(columns={名单学号列: '学号', 名单姓名列: '参考姓名'})
    df_名单['学号'] = pd.to_numeric(df_名单['学号'], errors='coerce')
    df_名单 = df_名单.dropna(subset=['学号'])
    df_名单['学号'] = df_名单['学号'].astype(int)
    df_名单 = df_名单.drop_duplicates(subset=['学号'], keep='first')

    df = df.copy()
    df['学号_原'] = df['学号']
    df['学号'] = pd.to_numeric(df['学号'], errors='coerce')
    df = df.dropna(subset=['学号'])
    df['学号'] = df['学号'].astype(int)

    df = pd.merge(df, df_名单, on='学号', how='left')

    if '参考姓名' in df.columns:
        if '姓名' in df.columns:
            旧姓名 = df['姓名'].copy()
            df['姓名'] = df['参考姓名'].fillna(df['姓名'])

            if 打印详情:
                修正掩码 = (旧姓名.fillna('') != df['姓名'].fillna('')) & (df['参考姓名'].notna())
                修正数 = int(修正掩码.sum())
                if 修正数 > 0:
                    print(f"    ℹ️ 修正了 {修正数} 个姓名（参考名单）:")
                    修正列表 = df[修正掩码].head(20)
                    for i, r in 修正列表.iterrows():
                        原名 = 旧姓名.loc[i]
                        新名 = r['姓名']
                        学号 = r['学号']
                        print(f"      ✏️ 学号 {学号}: '{原名}' → '{新名}'")
                    if 修正数 > 20:
                        print(f"      ... 还有 {修正数 - 20} 个")
        else:
            df['姓名'] = df['参考姓名']
        df = df.drop(columns=['参考姓名'])

    if '学号_原' in df.columns:
        df = df.drop(columns=['学号_原'])

    return df


# ==================== 分层核心函数 ====================

def apply_layer_colors(ws, start_idx, end_idx, n_columns, color):
    for idx in range(start_idx, end_idx):
        for j in range(1, n_columns + 1):
            ws.cell(row=idx + 2, column=j).fill = color


def apply_parallel_distribution(ws, df_students, start_idx, end_idx, colors, n_columns, parallel_groups=2):
    if start_idx >= end_idx:
        return
    if len(colors) == 0:
        return

    total = end_idx - start_idx
    num_groups = len(colors)

    if num_groups == 1:
        for idx in range(start_idx, end_idx):
            for j in range(1, n_columns + 1):
                ws.cell(row=idx + 2, column=j).fill = colors[0]
        return

    df_remaining = df_students.iloc[start_idx:end_idx].copy()
    df_remaining_sorted = df_remaining.sort_values(by=['加权平均'], ascending=True).reset_index(drop=True)

    mod_value = 2 * parallel_groups
    group_indices = [[] for _ in range(num_groups)]

    for i in range(total):
        mod_val = i % mod_value
        if parallel_groups == 2:
            if mod_val in [0, 1]:
                group_idx = 0
            else:
                group_idx = 1
        elif parallel_groups == 3:
            if mod_val in [0, 1]:
                group_idx = 0
            elif mod_val in [2, 5]:
                group_idx = 1
            else:
                group_idx = 2
        else:
            if mod_val < parallel_groups:
                if mod_val == 0 or mod_val == 1:
                    group_idx = 0
                else:
                    group_idx = mod_val - 1
            else:
                reverse_idx = mod_value - mod_val
                if reverse_idx <= 1:
                    group_idx = 0
                else:
                    group_idx = reverse_idx - 1
            if group_idx >= num_groups:
                group_idx = num_groups - 1
            if group_idx < 0:
                group_idx = 0
        group_indices[group_idx].append(i)

    all_groups = []
    for g_idx in range(num_groups):
        if group_indices[g_idx]:
            group_students = df_remaining_sorted.iloc[group_indices[g_idx]].copy()
            group_students = group_students.sort_values(by=['加权平均'], ascending=False).reset_index(drop=True)
            all_groups.append((group_students, g_idx))
        else:
            all_groups.append((pd.DataFrame(), g_idx))

    for idx in range(start_idx, end_idx):
        for j in range(1, n_columns + 1):
            ws.cell(row=idx + 2, column=j).value = None
            ws.cell(row=idx + 2, column=j).fill = PatternFill()

    current_pos_write = start_idx
    for group_students, g_idx in all_groups:
        if not group_students.empty:
            group_color = colors[g_idx % len(colors)]
            for idx, row in group_students.iterrows():
                actual_row = current_pos_write + idx
                if actual_row < end_idx:
                    for col_idx, col_name in enumerate(df_students.columns):
                        if col_idx < n_columns:
                            ws.cell(row=actual_row + 2, column=col_idx + 1, value=row[col_name])
                    for j in range(1, n_columns + 1):
                        ws.cell(row=actual_row + 2, column=j).fill = group_color
            current_pos_write += len(group_students)

    return group_indices


def parse_layer_config(layer_config, default_parallel_groups=2):
    parsed_config = []
    current_pos = 0

    for item in layer_config:
        if isinstance(item, (int, float)):
            parsed_config.append({
                'type': 'direct',
                'count': int(item),
                'start': current_pos,
                'end': current_pos + int(item)
            })
            current_pos += int(item)
        elif isinstance(item, (tuple, list)) and len(item) == 2:
            count, groups = item
            parsed_config.append({
                'type': 'parallel',
                'count': int(count),
                'groups': int(groups),
                'start': current_pos,
                'end': current_pos + int(count)
            })
            current_pos += int(count)
        elif isinstance(item, dict) and 'count' in item and 'groups' in item:
            count = item.get('count', 0)
            groups = item.get('groups', default_parallel_groups)
            parsed_config.append({
                'type': 'parallel',
                'count': int(count),
                'groups': int(groups),
                'start': current_pos,
                'end': current_pos + int(count)
            })
            current_pos += int(count)
        else:
            raise ValueError(f"无法识别的配置项: {item}")

    return parsed_config, current_pos


def setup_direct_layers(layer_counts, df_all=None, layer_colors=None):
    from 数据分析 import 显示选项

    wb = openpyxl.load_workbook('数据分析.xlsx')
    ws = wb['分层设置']
    n = ws.max_column

    if df_all is None:
        df_all = pd.read_excel('数据分析.xlsx', sheet_name='分层设置')

    total_students = df_all.shape[0]
    current_pos = 0

    if layer_colors is None:
        default_colors = [fille1, fille3, fille7, fille3, fille7]
        layer_colors = default_colors[:len(layer_counts) + 1]
    else:
        while len(layer_colors) < len(layer_counts) + 1:
            layer_colors.append(fille7)

    layer_names = ['A层', 'B层', 'C层', 'D层', 'E层', 'F层', 'G层', 'H层']

    if 显示选项.get('显示分层进度', True):
        print(f"直接分层模式:共{len(layer_counts) + 1}层")
        print("="*50)

    for i, count in enumerate(layer_counts):
        start_pos = current_pos
        end_pos = min(current_pos + count, total_students)
        if start_pos < end_pos:
            if 显示选项.get('显示分层进度', True):
                print(f"{layer_names[i]}: {end_pos - start_pos}人 (位置 {start_pos}-{end_pos-1})")
            apply_layer_colors(ws, start_pos, end_pos, n, layer_colors[i])
            current_pos = end_pos

    if current_pos < total_students:
        if 显示选项.get('显示分层进度', True):
            print(f"剩余层: {total_students - current_pos}人 (位置 {current_pos}-{total_students-1})")
        apply_layer_colors(ws, current_pos, total_students, n, layer_colors[len(layer_counts)])

    if 显示选项.get('显示分层进度', True):
        print("="*50)

    _安全保存工作簿(wb, '数据分析.xlsx')
    return total_students


def setup_parallel_layers(layer_config, df_all=None, parallel_groups=2, layer_colors=None):
    from 数据分析 import 显示选项

    wb = openpyxl.load_workbook('数据分析.xlsx')
    ws = wb['分层设置']
    n = ws.max_column

    if df_all is None:
        df_all = pd.read_excel('数据分析.xlsx', sheet_name='分层设置')

    total_students = df_all.shape[0]
    df_students = df_all.iloc[:total_students].copy()

    parsed_config, total_planned = parse_layer_config(layer_config, parallel_groups)

    total_layers = 0
    for config in parsed_config:
        if config['type'] == 'direct':
            total_layers += 1
        else:
            total_layers += config['groups']

    remaining_students = total_students - total_planned
    if remaining_students > 0:
        total_layers += parallel_groups
        if 显示选项.get('显示分层进度', True):
            print(f"剩余 {remaining_students} 人将平行分 {parallel_groups} 层")

    if layer_colors is None:
        default_colors = [fille1, fille3, fille7, fille5, fille2, fille4, fille6, fille8,
                         fille1, fille3, fille7, fille5, fille2, fille4, fille6, fille8]
        layer_colors = default_colors[:total_layers]
    else:
        while len(layer_colors) < total_layers:
            layer_colors.append(fille7)

    layer_names = ['A层', 'B层', 'C层', 'D层', 'E层', 'F层', 'G层', 'H层',
                   'I层', 'J层', 'K层', 'L层', 'M层', 'N层', 'O层', 'P层']

    current_pos = 0
    color_idx = 0

    if 显示选项.get('显示分层进度', True):
        print(f"平行分层模式: 总配置 {len(parsed_config)} 项")
        print(f"总学生数: {total_students}, 计划分配: {total_planned}, 剩余: {remaining_students}")
        print("="*60)

    for config in parsed_config:
        start_pos = config['start']
        end_pos = min(config['end'], total_students)
        if start_pos >= end_pos:
            continue

        if config['type'] == 'direct':
            layer_name = layer_names[color_idx] if color_idx < len(layer_names) else f"层{color_idx+1}"
            count = end_pos - start_pos
            if 显示选项.get('显示分层进度', True):
                print(f"{layer_name}: {count}人 (位置 {start_pos}-{end_pos-1}) [直接分层]")
            apply_layer_colors(ws, start_pos, end_pos, n, layer_colors[color_idx])
            current_pos = end_pos
            color_idx += 1
        else:
            groups = config['groups']
            count = end_pos - start_pos
            segment_colors = layer_colors[color_idx:color_idx + groups]
            while len(segment_colors) < groups:
                segment_colors.append(fille7)

            color_names = ', '.join([layer_names[color_idx+i] if color_idx+i < len(layer_names)
                                    else f'层{color_idx+i+1}' for i in range(groups)])
            if 显示选项.get('显示分层进度', True):
                print(f"平行分层: {count}人 (位置 {start_pos}-{end_pos-1})")
                print(f"  分{groups}层: {color_names}")

            apply_parallel_distribution(ws, df_students, start_pos, end_pos,
                                       segment_colors[:groups], n, groups)
            current_pos = end_pos
            color_idx += groups

    if current_pos < total_students:
        remaining = total_students - current_pos
        if 显示选项.get('显示分层进度', True):
            print(f"\n剩余 {remaining} 人平行分 {parallel_groups} 层 (位置 {current_pos}-{total_students-1})")

        remaining_colors = layer_colors[color_idx:color_idx + parallel_groups]
        while len(remaining_colors) < parallel_groups:
            remaining_colors.append(fille7)

        color_names = ', '.join([layer_names[color_idx+i] if color_idx+i < len(layer_names)
                                else f'层{color_idx+i+1}' for i in range(parallel_groups)])
        if 显示选项.get('显示分层进度', True):
            print(f"  分{parallel_groups}层: {color_names}")

        apply_parallel_distribution(ws, df_students, current_pos, total_students,
                                   remaining_colors, n, parallel_groups)
        color_idx += parallel_groups

    if 显示选项.get('显示分层进度', True):
        print("="*60)
        print(f"总学生数: {total_students}, 已处理: {total_students}")

    _安全保存工作簿(wb, '数据分析.xlsx')
    return total_students


# ==================== 辅助函数 ====================

def get_class_info(_分层数):
    class_list = []
    for a in range(1, 21):
        # 只判断文件是否存在，不把整份成绩表读入内存；真正处理时再读取。
        if os.path.isfile(_分层数(a)):
            class_list.append(a)
        else:
            break

    n_classes = len(class_list)

    if n_classes == 8:
        class_order = [1, 5, 2, 6, 3, 7, 4, 8]
    elif n_classes == 10:
        class_order = [1, 6, 2, 7, 3, 8, 4, 9, 5, 10]
    else:
        class_order = list(range(1, n_classes + 1))

    return n_classes, class_order


def get_qk_score(学号, 学期名称, 科目):
    """获取缺考赋分。优先级：学号列表赋分 > 学期赋分 > 默认赋分"""
    from 数据分析 import 分层排名配置

    缺考配置 = 分层排名配置.get('缺考赋分', {})

    学号列表配置 = 缺考配置.get('学号列表赋分', [])
    for 配置 in 学号列表配置:
        学号字符串 = 配置.get('学号', '')
        学号列表 = []
        for x in 学号字符串.replace('，', ',').split(','):
            x = x.strip()
            if x.isdigit():
                学号列表.append(int(x))
        if 学号 in 学号列表:
            学期赋分 = 配置.get('赋分', {})
            if 学期名称 in 学期赋分:
                科目赋分 = 学期赋分[学期名称]
                if 科目 in 科目赋分:
                    return 科目赋分[科目]

    学期赋分 = 缺考配置.get('学期赋分', {})
    if 学期名称 in 学期赋分:
        科目赋分 = 学期赋分[学期名称]
        if 科目 in 科目赋分:
            return 科目赋分[科目]

    默认赋分 = 缺考配置.get('默认赋分', {})
    if 科目 in 默认赋分:
        return 默认赋分[科目]

    return 0


# ==================== 知识点失分率函数 ====================

def 计算知识点失分率(学生姓名, 考试名称列表, 考试配置, 考试路径):
    知识点统计 = {}

    for 考试名称 in 考试名称列表:
        if 考试名称 not in 考试配置:
            continue

        config = 考试配置[考试名称]
        if '知识点' not in config:
            continue

        知识点列表 = config['知识点']
        分值列表 = config['各项分值']

        if len(知识点列表) != len(分值列表):
            continue

        try:
            df_exam = pd.read_excel(考试路径[考试名称])
            df_exam = df_exam.rename(columns={'学生校内学号': '学号'})
        except Exception as e:
            continue

        if '姓名' not in df_exam.columns:
            continue

        学生行 = df_exam[df_exam['姓名'] == 学生姓名]
        if 学生行.empty:
            continue

        学生行 = 学生行.iloc[0]

        题号列 = []
        for col in df_exam.columns:
            col_str = str(col).split('.')[0]
            if col_str.isdigit():
                题号列.append(col)

        try:
            题号列 = sorted(题号列, key=lambda x: int(str(x).split('.')[0]))
        except:
            pass

        if len(题号列) < len(分值列表):
            continue

        for i, 知识点 in enumerate(知识点列表):
            if i >= len(题号列):
                break
            分值 = 分值列表[i]
            列名 = 题号列[i]
            try:
                得分 = float(学生行[列名])
            except:
                得分 = 0

            if 知识点 not in 知识点统计:
                知识点统计[知识点] = {'总分值': 0, '总得分': 0}

            知识点统计[知识点]['总分值'] += 分值
            知识点统计[知识点]['总得分'] += 得分

    if not 知识点统计:
        return pd.DataFrame()

    结果列表 = []
    for 知识点, 数据 in 知识点统计.items():
        总分值 = 数据['总分值']
        总得分 = 数据['总得分']
        if 总分值 > 0:
            失分率 = round((总分值 - 总得分) / 总分值, 4)
        else:
            失分率 = 0
        结果列表.append({
            '知识点': 知识点,
            '总分值': 总分值,
            '总得分': round(总得分, 2),
            '失分率': 失分率
        })

    结果df = pd.DataFrame(结果列表)
    结果df = 结果df.sort_values(by='失分率', ascending=False).reset_index(drop=True)
    结果df = 结果df[结果df['失分率'] > 0]
    return 结果df


# ==================== 数据录入函数 ====================

def process_exam_data(考试名称, is_first=False):
    from 数据分析 import 考试配置, 题目难度分段配置

    config = 考试配置[考试名称]
    _总成绩 = config['总成绩']
    _分层数 = config['分层数']
    _各项分值 = config['各项分值']
    _表格 = config['表格']

    n_classes, class_order = get_class_info(_分层数)

    col0 = ['学号', '姓名', '排名', '总分']
    dfchushu = pd.read_excel(_总成绩)
    dfchushu = dfchushu.rename(columns={'学生校内学号': '学号'})
    dfchushu = dfchushu.rename(columns={'总排名': '排名'})

    dfchushu = 修正姓名(dfchushu, 参考名单路径=None, 打印详情=True)

    colchushu = [column for column in dfchushu.columns if (column.split(".")[0]).isdigit() or column.isdigit() or (column.split('-')[-1]).split(".")[-1].isdigit()]
    # 直接按题号列取数，避免对所有列逐列调用筛选函数。
    dfselchushu = dfchushu.loc[:, col0]
    dfchushu = dfchushu.loc[:, colchushu]

    if is_first:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', engine='openpyxl')
    else:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='overlay')

    with excel_writer as writer:
        dfchushu.to_excel(writer, sheet_name=_表格[0], index=False)

    wb = openpyxl.load_workbook('数据分析.xlsx')
    wbchushu = wb[_表格[0]]
    dchushu = max_rowbchushu = wbchushu.max_row
    max_columns = wbchushu.max_column

    valuechushu = _各项分值
    valchushu = sum(valuechushu)

    for (x, y) in enumerate(colchushu, start=1):
        # Excel 题目列名可能是字符串（如 '1'），不能用 dfchushu[x] 按整数标签取列；
        # 使用 iloc 按位置读取，Excel 列号及分值索引也统一使用位置 x。
        avg_x = round(pd.to_numeric(dfchushu.iloc[:, x - 1], errors='coerce').mean(), 5)
        wbchushu.cell(row=max_rowbchushu + 1, column=x, value=avg_x)
        wbchushu.cell(row=max_rowbchushu + 2, column=x, value=valuechushu[x - 1])
        wbchushu.cell(row=max_rowbchushu + 3, column=x, value=round(avg_x / valuechushu[x - 1], 5))
    _安全保存工作簿(wb, '数据分析.xlsx')

    dfchushu = pd.read_excel('数据分析.xlsx', sheet_name=_表格[0])
    dfchushu = dfchushu.sort_values(by=dchushu + 1, axis=1, ascending=False)

    if is_first:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', engine='openpyxl')
    else:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='overlay')

    with excel_writer as writer:
        dfchushu.to_excel(writer, sheet_name=_表格[0], startcol=3, index=False)

    wb = openpyxl.load_workbook('数据分析.xlsx')
    wbchushu = wb[_表格[0]]
    dchushu = max_rowbchushu = wbchushu.max_row
    max_columns = wbchushu.max_column

    # 配置比例可调整；归一化后形成累计分值阈值，默认仍为 60% / 90%。
    比例 = [
        max(0.0, float(题目难度分段配置.get('基础题比例', 0.6))),
        max(0.0, float(题目难度分段配置.get('中档题比例', 0.3))),
        max(0.0, float(题目难度分段配置.get('难题比例', 0.1))),
    ]
    比例总和 = sum(比例)
    if 比例总和 <= 0:
        比例 = [0.6, 0.3, 0.1]
        比例总和 = 1.0
    基础累计阈值 = 比例[0] / 比例总和
    中档累计阈值 = (比例[0] + 比例[1]) / 比例总和

    # 题目已按得分率从高到低排序；按累计分值比例确定切分点。
    题目数 = min(len(colchushu), len(_各项分值))
    if 题目数 == 0:
        raise ValueError(f'{考试名称}没有识别到有效题目列或各项分值为空')

    题目分值 = pd.to_numeric(dfchushu.iloc[dchushu - 3, :题目数], errors='coerce').fillna(0)
    总题目分值 = float(题目分值.sum())
    if 总题目分值 <= 0:
        # 数据异常时按题目数量切分，避免边界变量未定义。
        j_1 = max(1, min(题目数, round(题目数 * 基础累计阈值)))
        j_2 = max(j_1, min(题目数, round(题目数 * 中档累计阈值)))
    else:
        累计比例 = (题目分值.cumsum() / 总题目分值).to_numpy()
        j_1 = next((i + 1 for i, v in enumerate(累计比例) if v >= 基础累计阈值), 题目数)
        j_2 = next((i + 1 for i, v in enumerate(累计比例) if v >= 中档累计阈值), 题目数)
        j_2 = max(j_1, j_2)

    j_3 = 题目数
    wbchushu.cell(row=1, column=1, value='基础题')
    wbchushu.cell(row=1, column=2, value='中档题')
    wbchushu.cell(row=1, column=3, value='难题')

    list_a = list(dfchushu.iloc[0:0, 0:j_1])
    list_a = ', '.join(list_a)
    list_b = list(dfchushu.iloc[0:0, j_1:j_2])
    list_b = ', '.join(list_b)
    list_c = list(dfchushu.iloc[0:0, j_2:j_3])
    list_c = ', '.join(list_c)

    for i in range(0, max_rowbchushu - 2):
        list_0 = dfchushu.iloc[i, 0:j_1]
        list_1 = dfchushu.iloc[i, j_1:j_2]
        list_2 = dfchushu.iloc[i, j_2:j_3]
        sumh_0 = list_0.sum(skipna=False)
        sumh_1 = list_1.sum(skipna=False)
        sumh_2 = list_2.sum(skipna=False)
        wbchushu.cell(row=i + 2, column=1, value=sumh_0)
        wbchushu.cell(row=i + 2, column=2, value=sumh_1)
        wbchushu.cell(row=i + 2, column=3, value=sumh_2)

    v11 = wbchushu.cell(row=max_rowbchushu - 2, column=1).value
    v12 = wbchushu.cell(row=max_rowbchushu - 2, column=2).value
    v13 = wbchushu.cell(row=max_rowbchushu - 2, column=3).value
    v21 = wbchushu.cell(row=max_rowbchushu - 1, column=1).value
    v22 = wbchushu.cell(row=max_rowbchushu - 1, column=2).value
    v23 = wbchushu.cell(row=max_rowbchushu - 1, column=3).value
    wbchushu.cell(row=max_rowbchushu, column=1, value=round(v11 / v21, 5) if v21 not in (None, 0) else 0)
    wbchushu.cell(row=max_rowbchushu, column=2, value=round(v12 / v22, 5) if v22 not in (None, 0) else 0)
    wbchushu.cell(row=max_rowbchushu, column=3, value=round(v13 / v23, 5) if v23 not in (None, 0) else 0)
    _安全保存工作簿(wb, '数据分析.xlsx')

    dfchushu = pd.read_excel('数据分析.xlsx', sheet_name=_表格[0])
    dfchushu = pd.concat([dfselchushu, dfchushu], axis=1, join="outer")

    if is_first:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', engine='openpyxl')
    else:
        excel_writer = pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='replace')

    with excel_writer as writer:
        dfchushu.to_excel(writer, sheet_name=_表格[0], index=False)

    wb = openpyxl.load_workbook('数据分析.xlsx')
    wbchushu = wb[_表格[0]]
    dchushu = max_rowbchushu = wbchushu.max_row
    max_columns = wbchushu.max_column

    for i in range(1, max_rowbchushu + 1):
        for j in range(8, 8 + len(list_0)):
            wbchushu.cell(row=i, column=j).fill = fille1
        for j in range(8 + len(list_0), 8 + len(list_0 + list_1)):
            wbchushu.cell(row=i, column=j).fill = fille3
        for j in range(8 + len(list_0 + list_1), 8 + len(list_0 + list_1 + list_2)):
            wbchushu.cell(row=i, column=j).fill = fille5

    wbchushu.cell(row=max_rowbchushu - 2, column=1, value=_表格[2])
    wbchushu.cell(row=max_rowbchushu - 1, column=1, value=_表格[3])
    wbchushu.cell(row=max_rowbchushu, column=1, value=_表格[4])
    wbchushu.cell(row=dchushu + 2, column=1, value='基础题号')
    wbchushu.cell(row=dchushu + 2, column=7, value=list_a)
    wbchushu.cell(row=dchushu + 3, column=1, value='中档题号')
    wbchushu.cell(row=dchushu + 3, column=7, value=list_b)
    wbchushu.cell(row=dchushu + 4, column=1, value='难题题号')
    wbchushu.cell(row=dchushu + 4, column=7, value=list_c)

    list_y = dfchushu.columns.tolist()
    l_y = len(list_y)

    wbchushu.cell(row=dchushu + 7, column=1, value='分层班级')
    wbchushu.cell(row=dchushu + 7, column=4, value='均分')
    wbchushu.cell(row=dchushu + 7, column=5, value='基础均分')
    wbchushu.cell(row=dchushu + 7, column=6, value='中档均分')
    wbchushu.cell(row=dchushu + 7, column=7, value='难题均分')

    wbchushu.merge_cells(start_row=dchushu + 7, start_column=1, end_row=dchushu + 7, end_column=3)

    for idx, a in enumerate(class_order):
        current_row = dchushu + 8 + idx
        wbchushu.cell(row=current_row, column=1, value=f'M{a}')
        wbchushu.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)

    for idx, a in enumerate(class_order):
        current_row = dchushu + 8 + idx
        dfchushu_a = pd.read_excel(_分层数(a))
        dfchushu_a = dfchushu_a.rename(columns={'学生校内学号': '学号'})
        student_data = dfchushu.iloc[:dchushu - 3, :]
        dfchushu_a = student_data[student_data['学号'].isin(dfchushu_a['学号'])]

        for j in range(4, l_y + 1):
            avg_a_j = round(dfchushu_a[list_y[j-1]].mean(), 5)
            wbchushu.cell(row=current_row, column=j, value=avg_a_j)

    # 分层表统计信息
    layer_sheet_name = _表格[1]
    if layer_sheet_name not in wb.sheetnames:
        wb.create_sheet(layer_sheet_name)
    ws_layer = wb[layer_sheet_name]

    for row in ws_layer.iter_rows():
        for cell in row:
            cell.value = None

    cols_per_class = 4
    class_data_list = []
    max_students = 0

    for idx, a in enumerate(class_order):
        try:
            class_df = pd.read_excel(_分层数(a))
            class_df = class_df.rename(columns={'学生校内学号': '学号'})
            class_df = 修正姓名(class_df, 参考名单路径=None, 打印详情=False)
            merged_df = pd.merge(
                class_df[['学号', '姓名']],
                dfchushu[['学号', '总分', '排名']],
                on='学号',
                how='left'
            )
            merged_df = merged_df.sort_values(by=['总分'], ascending=False).reset_index(drop=True)
            class_data_list.append({
                'class_num': a,
                'data': merged_df,
                'count': len(merged_df),
                'avg': round(merged_df['总分'].mean(), 3) if len(merged_df) > 0 else 0
            })
            if len(merged_df) > max_students:
                max_students = len(merged_df)
        except Exception as e:
            print(f"处理分层表 M{a} 时出错: {e}")
            class_data_list.append({
                'class_num': a,
                'data': pd.DataFrame(),
                'count': 0,
                'avg': 0
            })

    avg0 = round(dfchushu['总分'].mean(), 3)

    header_row = 1
    for idx, class_info in enumerate(class_data_list):
        a = class_info['class_num']
        start_col = cols_per_class * idx + 1
        ws_layer.cell(row=header_row, column=start_col, value=f'数{a}')
        ws_layer.merge_cells(start_row=header_row, start_column=start_col,
                            end_row=header_row, end_column=start_col + cols_per_class - 1)
        ws_layer.cell(row=header_row, column=start_col).font = Font(name='黑体', size=14, bold=True)
        ws_layer.cell(row=header_row + 1, column=start_col, value='学号')
        ws_layer.cell(row=header_row + 1, column=start_col + 1, value='姓名')
        ws_layer.cell(row=header_row + 1, column=start_col + 2, value='总分')
        ws_layer.cell(row=header_row + 1, column=start_col + 3, value='排名')
        header_fill = fille3 if idx % 2 == 0 else fille2
        for j in range(start_col, start_col + cols_per_class):
            ws_layer.cell(row=header_row + 1, column=j).fill = header_fill

    data_start_row = header_row + 2
    for idx, class_info in enumerate(class_data_list):
        a = class_info['class_num']
        start_col = cols_per_class * idx + 1
        merged_df = class_info['data']
        color_fill = fille3 if idx % 2 == 0 else fille2
        for row_idx, row_data in merged_df.iterrows():
            actual_row = data_start_row + row_idx
            ws_layer.cell(row=actual_row, column=start_col, value=row_data['学号'])
            ws_layer.cell(row=actual_row, column=start_col + 1, value=row_data['姓名'])
            ws_layer.cell(row=actual_row, column=start_col + 2, value=row_data['总分'])
            ws_layer.cell(row=actual_row, column=start_col + 3, value=row_data['排名'])
            for j in range(start_col, start_col + cols_per_class):
                ws_layer.cell(row=actual_row, column=j).fill = color_fill

    stat_start_row = data_start_row + max_students + 2
    ws_layer.cell(row=stat_start_row, column=1, value='分层')
    for idx, class_info in enumerate(class_data_list):
        a = class_info['class_num']
        col = idx + 2
        ws_layer.cell(row=stat_start_row, column=col, value=f'数{a}')
        ws_layer.cell(row=stat_start_row, column=col).fill = fille3 if idx % 2 == 0 else fille2

    stat_row_avg = stat_start_row + 1
    ws_layer.cell(row=stat_row_avg, column=1, value='班均分')
    for idx, class_info in enumerate(class_data_list):
        col = idx + 2
        ws_layer.cell(row=stat_row_avg, column=col, value=class_info['avg'])
        ws_layer.cell(row=stat_row_avg, column=col).fill = fille3 if idx % 2 == 0 else fille2

    stat_row_diff = stat_start_row + 2
    ws_layer.cell(row=stat_row_diff, column=1, value='与均分差')
    for idx, class_info in enumerate(class_data_list):
        col = idx + 2
        diff = round(class_info['avg'] - avg0, 3)
        ws_layer.cell(row=stat_row_diff, column=col, value=diff)
        ws_layer.cell(row=stat_row_diff, column=col).fill = fille3 if idx % 2 == 0 else fille2

    stat_row_grade = stat_start_row + 3
    ws_layer.cell(row=stat_row_grade, column=1, value='年级均分')
    total_cols = n_classes + 1
    ws_layer.merge_cells(start_row=stat_row_grade, start_column=2,
                        end_row=stat_row_grade, end_column=total_cols)
    ws_layer.cell(row=stat_row_grade, column=2, value=avg0)
    ws_layer.cell(row=stat_row_grade, column=2).font = Font(bold=True)

    _安全保存工作簿(wb, '数据分析.xlsx')


def _create_multi_test_layer_ranking(分层排名配置, 显示选项):
    """按用户手动配置的多次测试权重生成分层总排名和分层设置。"""
    from 数据分析 import 考试配置

    tests = 分层排名配置.get('多次测试', [])
    if len(tests) < 2:
        raise ValueError('多次测试加权排名至少需要配置两次测试。')

    # 防止考试总表中选入了尚未配置分段工作表的考试，避免后续出现隐晦的 KeyError/worksheet 错误。
    missing_exam_config = [item.get('名称', '<未命名>') for item in tests if item.get('名称') not in 考试配置]
    if missing_exam_config:
        raise ValueError('以下参与分层的考试尚未在数据分析.py 的“考试配置”中定义，无法确定分段工作表：' + '、'.join(missing_exam_config))

    normalized_tests = []
    total_exam_weight = 0.0
    for item in tests:
        name = item['名称']
        weight = float(item['总权重'])
        difficulty = [float(x) for x in item['难度权重']]
        if weight <= 0 or len(difficulty) != 3 or any(x < 0 for x in difficulty) or sum(difficulty) <= 0:
            raise ValueError(f'考试“{name}”的权重配置无效。')
        difficulty_sum = sum(difficulty)
        difficulty = [x / difficulty_sum for x in difficulty]
        normalized_tests.append({
            'name': name,
            'weight': weight,
            'difficulty': difficulty,
            'sheet': 考试配置.get(name, {}).get('表格', [name + '分段'])[0],
        })
        total_exam_weight += weight
    if total_exam_weight <= 0:
        raise ValueError('多次测试总权重之和必须大于 0。')
    for item in normalized_tests:
        item['weight'] /= total_exam_weight

    print('多次测试加权排名：')
    for item in normalized_tests:
        print(f"  {item['name']}: 总权重={item['weight']:.4f}, 难度权重={item['difficulty']}")

    merged = None
    exam_means = {}
    exam_columns = {}
    for item in normalized_tests:
        name = item['name']
        raw = pd.read_excel('数据分析.xlsx', sheet_name=item['sheet'])
        marker = next((i for i, value in enumerate(raw.iloc[:, 0]) if str(value).strip() == '各项均分'), None)
        if marker is None:
            raise ValueError(f'工作表“{item["sheet"]}”中找不到“各项均分”行，请确认考试分段表已生成。')
        if raw.shape[1] < 7:
            raise ValueError(f'工作表“{item["sheet"]}”列数不足，无法读取基础/中档/难题数据。')

        # 与旧版 A/B 逻辑保持相同的列位置约定：第 1、2、5、6、7 列。
        selected = raw.iloc[:marker, [0, 1, 4, 5, 6]].copy()
        selected.columns = ['学号', '姓名', f'{name}基础', f'{name}中档', f'{name}难题']
        selected['学号'] = pd.to_numeric(selected['学号'], errors='coerce')
        selected = selected.dropna(subset=['学号'])
        selected['学号'] = selected['学号'].astype('int64')
        selected = selected.drop_duplicates(subset=['学号'], keep='first')
        exam_columns[name] = [f'{name}基础', f'{name}中档', f'{name}难题']

        try:
            # “各项均分”标记行之后一行是“各项分值”，即基础/中档/难题各自的满分。
            mean_row = raw.iloc[min(marker + 1, len(raw) - 1), [4, 5, 6]]
            exam_means[name] = [
                float(pd.to_numeric(v, errors='coerce')) if pd.notna(v) else 0.0
                for v in mean_row.tolist()
            ]
        except Exception:
            exam_means[name] = [0.0, 0.0, 0.0]

        if merged is None:
            merged = selected
        else:
            merged = pd.merge(merged, selected, on='学号', how='outer', suffixes=('', '_重复'))
            duplicate_name = '姓名_重复'
            if duplicate_name in merged.columns:
                if '姓名' in merged.columns:
                    merged['姓名'] = merged['姓名'].fillna(merged[duplicate_name])
                merged.drop(columns=[duplicate_name], inplace=True)

    if merged is None or merged.empty:
        raise ValueError('多次测试没有读取到有效学生数据。')

    # 使用配置名单补入缺考学生；姓名优先采用参考名单。
    缺考配置 = 分层排名配置.get('缺考赋分', {})
    if 缺考配置.get('启用', False):
        for setting in 缺考配置.get('学号列表赋分', []):
            for raw_id in str(setting.get('学号', '')).replace('，', ',').split(','):
                raw_id = raw_id.strip()
                if raw_id.isdigit() and int(raw_id) not in set(merged['学号'].tolist()):
                    merged = pd.concat([merged, pd.DataFrame([{'学号': int(raw_id), '姓名': f'学号{raw_id}'}])], ignore_index=True)

    名单路径 = 分层排名配置.get('名单路径') or 分层排名配置.get('参考名单')
    if not 名单路径:
        try:
            from 数据分析 import 控制页配置
            名单路径 = 控制页配置.get('名单路径')
        except Exception:
            名单路径 = None
    if 名单路径 and os.path.exists(名单路径):
        try:
            roster = pd.read_excel(名单路径)
            id_col = next((c for c in roster.columns if '学号' in str(c) or 'id' in str(c).lower() or '学籍号' in str(c)), None)
            name_col = next((c for c in roster.columns if '姓名' in str(c) or 'name' in str(c).lower()), None)
            if id_col and name_col:
                roster = roster[[id_col, name_col]].copy()
                roster.columns = ['学号', '参考姓名']
                roster['学号'] = pd.to_numeric(roster['学号'], errors='coerce')
                roster = roster.dropna(subset=['学号'])
                roster['学号'] = roster['学号'].astype('int64').drop_duplicates()
                roster = roster.drop_duplicates(subset=['学号'], keep='first')
                merged = pd.merge(merged, roster, on='学号', how='outer')
                if '姓名' not in merged.columns:
                    merged['姓名'] = merged['参考姓名']
                else:
                    merged['姓名'] = merged['参考姓名'].fillna(merged['姓名'])
                merged.drop(columns=['参考姓名'], inplace=True)
        except Exception as exc:
            print(f'  ⚠️ 参考名单读取失败：{exc}')

    merged['姓名'] = merged['姓名'].fillna('未知').astype(str).str.strip()
    merged['学号'] = pd.to_numeric(merged['学号'], errors='coerce')
    merged = merged.dropna(subset=['学号'])
    merged['学号'] = merged['学号'].astype('int64')

    # 缺考处理：优先使用本次测试名称 + 动态难度列名；旧配置别名仍兼容 A/B。
    for item in normalized_tests:
        name = item['name']
        for difficulty_name, col in zip(['基础', '中档', '难题'], exam_columns[name]):
            if col not in merged.columns:
                merged[col] = 0
            missing = merged[col].isna()
            if missing.any():
                alias = f'{name}{difficulty_name}'
                legacy_alias = None
                if normalized_tests.index(item) == 0:
                    legacy_alias = f'基准A{difficulty_name}'
                elif normalized_tests.index(item) == 1:
                    legacy_alias = f'基准B{difficulty_name}'
                for row_index in merged.index[missing]:
                    student_id = int(merged.at[row_index, '学号'])
                    score = get_qk_score(student_id, name, alias)
                    if score == 0 and legacy_alias:
                        score = get_qk_score(student_id, name, legacy_alias)
                    merged.at[row_index, col] = score

    # 方案A：每场考试先按基础/中档/难题各自满分计算得分率，再按本场难度权重合成为百分制。
    # 不同考试满分不同不会直接影响结果；考试间再按“总权重”加权平均。
    exam_percent_scores = {}
    for item in normalized_tests:
        name = item['name']
        difficulty = item['difficulty']
        full_marks = exam_means[name]
        active_categories = [i for i, mark in enumerate(full_marks) if mark > 0]
        if not active_categories:
            raise ValueError(f'考试“{name}”的基础/中档/难题满分均为 0，无法换算成百分制。')

        # 某个难度类别没有题目时，该类别不参与本场合成，并对剩余难度权重重新归一化。
        active_weight_sum = sum(difficulty[i] for i in active_categories)
        if active_weight_sum <= 0:
            raise ValueError(f'考试“{name}”在有题目的难度类别上没有有效权重，请调整难度权重。')

        exam_score = pd.Series(0.0, index=merged.index)
        for i in active_categories:
            col = exam_columns[name][i]
            student_points = pd.to_numeric(merged[col], errors='coerce').fillna(0.0)
            rate = (student_points / float(full_marks[i])).clip(lower=0.0, upper=1.0)
            exam_score += rate * (difficulty[i] / active_weight_sum)
        exam_percent_scores[name] = exam_score * 100.0

    # 所有单场成绩均已转换为百分制，因此直接按考试间权重合成。
    merged['加权平均'] = 0.0
    for item in normalized_tests:
        merged['加权平均'] += item['weight'] * exam_percent_scores[item['name']]
    merged['加权平均'] = merged['加权平均'].round(3)

    # 同分优先级可在配置中用考试名称列表指定；未列出的考试按多次测试配置顺序补齐。
    priority = 分层排名配置.get('多次测试同分优先级', [])
    ordered_names = []
    for name in priority:
        if name in exam_columns and name not in ordered_names:
            ordered_names.append(name)
    for item in normalized_tests:
        if item['name'] not in ordered_names:
            ordered_names.append(item['name'])
    sort_cols = ['加权平均']
    for difficulty_name in ['难题', '中档', '基础']:
        sort_cols.extend([f'{name}{difficulty_name}' for name in ordered_names])
    sort_cols.append('学号')
    ascending = [False] * (len(sort_cols) - 1) + [True]
    sort_df = merged.copy()
    sort_df['_加权平均排序键'] = pd.to_numeric(sort_df['加权平均'], errors='coerce').round(3)
    sort_cols[0] = '_加权平均排序键'
    sort_df = sort_df.sort_values(by=sort_cols, ascending=ascending, kind='mergesort', na_position='last')
    merged = sort_df.drop(columns=['_加权平均排序键']).reset_index(drop=True)

    # 先保留完整排名，再按分界学号拆成左右两组用于分层设置。
    with pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
        merged.to_excel(writer, sheet_name='分层总排名', index=False)

    exclude_text = 分层排名配置.get('排除学号', '')
    excluded = [int(x.strip()) for x in str(exclude_text).replace('，', ',').split(',') if x.strip().isdigit()]
    layer_df = merged[~merged['学号'].isin(excluded)].copy() if excluded else merged.copy()
    boundary = int(分层排名配置.get('分界学号', 0))
    group_a = layer_df.loc[layer_df['学号'] <= boundary].copy().reset_index(drop=True)
    group_b = layer_df.loc[layer_df['学号'] > boundary].copy().reset_index(drop=True)
    layer_setting = pd.concat([group_a, group_b], axis=1)
    layer_setting.insert(len(group_a.columns), '分隔', None)

    with pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
        layer_setting.to_excel(writer, sheet_name='分层设置', index=False)

    if 显示选项.get('显示数据统计', True):
        print(f'  参与排名人数：{len(merged)}')
        print(f'  排除后参与分层人数：{len(layer_df)}')
        print(f'  分界学号：{boundary}；A组：{len(group_a)}；B组：{len(group_b)}')
    print('多次测试加权排名和分层设置已生成。')
    return layer_setting


def create_layer_ranking():
    """创建分层总排名和分层设置工作表"""
    from 数据分析 import 分层排名配置, 显示选项

    # 只要存在新版“多次测试”键，就绝不回退到旧版 A/B 分层逻辑；即使列表为空也要明确报错。
    if '多次测试' in 分层排名配置:
        if not 分层排名配置.get('启用', True):
            try:
                from 数据分析 import 多次测试配置错误
                详情 = '；'.join(多次测试配置错误) if 多次测试配置错误 else '请检查考试总表中参与分层的考试是否均已录入。'
            except Exception:
                详情 = '请检查数据分析.py 中的多次测试配置诊断信息。'
            raise ValueError(f'多次测试分层配置未通过校验：{详情}')
        if not 分层排名配置.get('多次测试'):
            raise ValueError('没有可用于分层的考试。请至少设置两场考试“录入”: True 且“基准”非 None。')
        return _create_multi_test_layer_ranking(分层排名配置, 显示选项)

    if not 分层排名配置.get('启用', True):
        print("\n⚠️ 分层排名未启用（分层排名配置['启用'] = False），跳过创建")
        return None

    # 防止新旧版本的配置文件混用：新版已取消基准 A/B，不应进入旧版逻辑。
    if '基准A分段' not in 分层排名配置 or '基准B分段' not in 分层排名配置:
        print("\n⚠️ 分层排名配置与主程序版本不一致：未检测到有效的‘多次测试’配置，且旧版基准 A/B 配置已移除。")
        print("请将主程序.py 和 数据分析.py 一并更新为同一版本，并确认考试总表中至少有两场考试设置为‘录入’: True。")
        return None

    print("\n正在创建分层总排名和分层设置...")

    name_1 = 分层排名配置['基准A分段']
    name_2 = 分层排名配置['基准B分段']
    基准A名称 = 分层排名配置.get('基准A名称', '八下期中')
    基准B名称 = 分层排名配置.get('基准B名称', '八下期末')
    a0 = 分层排名配置['基准A权重']
    b0 = 分层排名配置['基准B权重']
    分界 = 分层排名配置['分界学号']
    排序字段 = 分层排名配置['排序字段']
    排序顺序 = 分层排名配置['排序顺序']

    缺考配置 = 分层排名配置.get('缺考赋分', {})
    启用缺考赋分 = 缺考配置.get('启用', False)

    def extract_student_data(sheet_name, exam_prefix):
        df = pd.read_excel('数据分析.xlsx', sheet_name=sheet_name)
        for i in range(len(df)):
            if df.iloc[i, 0] == '各项均分':
                student_data = df.iloc[:i+2, [0, 1, 4, 5, 6]]
                student_data = student_data.dropna(subset=['学号'])
                break
        student_data = student_data.rename(columns={
            '基础题': f'{exam_prefix}基础',
            '中档题': f'{exam_prefix}中档',
            '难题': f'{exam_prefix}难题'
        })
        return student_data

    df1 = extract_student_data(name_1, '基准A')
    df2 = extract_student_data(name_2, '基准B')
    df1_1 = df1.iloc[-1, [2, 3, 4]]
    df2_1 = df2.iloc[-1, [2, 3, 4]]

    if 显示选项.get('显示数据统计', True):
        print(f"基准A: {基准A名称}（{name_1}）")
        print(f"基准B: {基准B名称}（{name_2}）")
        print(f"基准A权重: 基础={a0[0]}, 中档={a0[1]}, 难题={a0[2]}")
        print(f"基准B权重: 基础={b0[0]}, 中档={b0[1]}, 难题={b0[2]}")
        if 启用缺考赋分:
            学号列表配置 = 缺考配置.get('学号列表赋分', [])
            if 学号列表配置:
                print(f"\n学号列表赋分配置:")
                for 配置 in 学号列表配置:
                    学号列表 = 配置.get('学号', '')
                    print(f"  学号 [{学号列表}]")

    df1['学号'] = pd.to_numeric(df1['学号'], errors='coerce')
    df1 = df1.dropna(subset=['学号'])
    df1['学号'] = df1['学号'].astype(int)
    df1 = df1.sort_values(by=['学号'], ascending=True)

    df2['学号'] = pd.to_numeric(df2['学号'], errors='coerce')
    df2 = df2.dropna(subset=['学号'])
    df2['学号'] = df2['学号'].astype(int)
    df2 = df2.sort_values(by=['学号'], ascending=True)

    df = pd.merge(df1, df2, on=['学号'], how='outer', suffixes=('_A', '_B'))

    if '姓名_A' in df.columns and '姓名_B' in df.columns:
        df['姓名'] = df['姓名_A'].fillna(df['姓名_B'])
        df = df.drop(columns=['姓名_A', '姓名_B'])
    elif '姓名_A' in df.columns:
        df = df.rename(columns={'姓名_A': '姓名'})
    elif '姓名_B' in df.columns:
        df = df.rename(columns={'姓名_B': '姓名'})

    名单路径 = 分层排名配置.get('名单路径', None) or 分层排名配置.get('参考名单', None)
    if not 名单路径:
        try:
            from 数据分析 import 控制页配置
            名单路径 = 控制页配置.get('名单路径', None)
        except:
            pass

    if 名单路径 and os.path.exists(名单路径):
        try:
            df_名单 = pd.read_excel(名单路径)
            名单学号列 = None
            名单姓名列 = None
            for col in df_名单.columns:
                if 名单学号列 is None and ('学号' in str(col) or 'id' in str(col).lower() or '学籍号' in str(col)):
                    名单学号列 = col
                elif 名单姓名列 is None and ('姓名' in str(col) or 'name' in str(col).lower()):
                    名单姓名列 = col
            if 名单学号列 and 名单姓名列:
                df_名单 = df_名单[[名单学号列, 名单姓名列]].copy()
                df_名单 = df_名单.rename(columns={名单学号列: '学号', 名单姓名列: '参考姓名'})
                df_名单['学号'] = pd.to_numeric(df_名单['学号'], errors='coerce')
                df_名单 = df_名单.dropna(subset=['学号'])
                df_名单['学号'] = df_名单['学号'].astype(int)
                df_名单 = df_名单.drop_duplicates(subset=['学号'], keep='first')
                df = pd.merge(df, df_名单, on='学号', how='left')
                df['姓名'] = df['参考姓名'].fillna(df['姓名'])
                df = df.drop(columns=['参考姓名'])
                if 显示选项.get('显示数据统计', True):
                    print(f"  姓名已参考名单文件: {名单路径}")
        except Exception as e:
            print(f"  ⚠️ 读取名单文件失败: {e}")

    df['姓名'] = df['姓名'].fillna('未知')
    df['姓名'] = df['姓名'].astype(str).str.strip()

    其他列 = [c for c in df.columns if c not in ['学号', '姓名']]
    df = df[['学号', '姓名'] + 其他列]

    if 启用缺考赋分:
        学号列表配置 = 缺考配置.get('学号列表赋分', [])
        for 配置 in 学号列表配置:
            学号字符串 = 配置.get('学号', '')
            学号列表 = []
            for x in 学号字符串.replace('，', ',').split(','):
                x = x.strip()
                if x.isdigit():
                    学号列表.append(int(x))
            for 学号 in 学号列表:
                if 学号 not in df['学号'].values:
                    新行 = pd.DataFrame([{'学号': 学号, '姓名': f'学号{学号}'}])
                    df = pd.concat([df, 新行], ignore_index=True)

    缺考_期中 = set()
    缺考_期末 = set()

    for idx in df.index:
        学号 = df.loc[idx, '学号']
        是否缺考_期中 = False
        是否缺考_期末 = False

        for 科目 in ['基准A基础', '基准A中档', '基准A难题']:
            if pd.isna(df.loc[idx, 科目]):
                if 启用缺考赋分:
                    赋分 = get_qk_score(学号, 基准A名称, 科目)
                    df.loc[idx, 科目] = 赋分
                else:
                    df.loc[idx, 科目] = 0
                是否缺考_期中 = True

        for 科目 in ['基准B基础', '基准B中档', '基准B难题']:
            if pd.isna(df.loc[idx, 科目]):
                if 启用缺考赋分:
                    赋分 = get_qk_score(学号, 基准B名称, 科目)
                    df.loc[idx, 科目] = 赋分
                else:
                    df.loc[idx, 科目] = 0
                是否缺考_期末 = True

        if 是否缺考_期中:
            缺考_期中.add(学号)
        if 是否缺考_期末:
            缺考_期末.add(学号)

    if 显示选项.get('显示数据统计', True):
        print(f"\n缺考统计:")
        print(f"  基准A({基准A名称})缺考: {len(缺考_期中)}人")
        print(f"  基准B({基准B名称})缺考: {len(缺考_期末)}人")

    df = df.fillna(0)

    k0 = round(100/(df1_1[0]*a0[0] + df1_1[1]*a0[1] + df1_1[2]*a0[2] +
                     df2_1[0]*b0[0] + df2_1[1]*b0[1] + df2_1[2]*b0[2]), 3)
    x0 = [k0, k0]

    df['加权平均'] = x0[0]*(df['基准A基础']*a0[0] + df['基准A中档']*a0[1] + df['基准A难题']*a0[2]) + \
                     x0[1]*(df['基准B基础']*b0[0] + df['基准B中档']*b0[1] + df['基准B难题']*b0[2])

    # 同分判定按控制表显示的 3 位小数进行，避免 71.742 与
    # 71.74199999999999 这类浮点误差阻止后续字段决胜。
    # 实际写入的“加权平均”原值不修改；仅使用临时排序键比较。
    def 按分层规则排序(数据):
        排序数据 = 数据.copy()
        当前排序字段 = list(排序字段)
        if '加权平均' in 当前排序字段:
            临时排序列 = '_加权平均排序键'
            排序数据[临时排序列] = pd.to_numeric(
                排序数据['加权平均'], errors='coerce'
            ).round(3)
            当前排序字段[当前排序字段.index('加权平均')] = 临时排序列
        排序数据 = 排序数据.sort_values(
            by=当前排序字段,
            ascending=排序顺序,
            kind='mergesort',
            na_position='last'
        )
        return 排序数据.drop(columns=['_加权平均排序键'], errors='ignore')

    df = 按分层规则排序(df)

    with pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='分层总排名', index=False)

    df = pd.read_excel('数据分析.xlsx', sheet_name='分层总排名')

    排除学号字符串 = 分层排名配置.get('排除学号', '')
    if 排除学号字符串:
        排除学号列表 = []
        for x in 排除学号字符串.replace('，', ',').split(','):
            x = x.strip()
            if x.isdigit():
                排除学号列表.append(int(x))
        if 排除学号列表:
            df['学号'] = pd.to_numeric(df['学号'], errors='coerce')
            原始人数 = len(df)
            df = df[~df['学号'].isin(排除学号列表)].copy()
            df = df.dropna(subset=['学号'])
            df['学号'] = df['学号'].astype(int)
            排除人数 = 原始人数 - len(df)
            if 显示选项.get('显示数据统计', True):
                print(f"\n已排除 {排除人数} 人（不参与分层设置）")

    df_a = 按分层规则排序(df.loc[df['学号'] <= 分界].copy())
    df_a.reset_index(drop=True, inplace=True)

    df_b = 按分层规则排序(df.loc[df['学号'] > 分界].copy())
    df_b.reset_index(drop=True, inplace=True)

    df = pd.concat([df_a, df_b], axis=1)
    df.insert(9, '分隔', None)

    with pd.ExcelWriter('数据分析.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='分层设置', index=False)

    if 显示选项.get('显示数据统计', True):
        print(f"\n分界学号: {分界}")
        print(f"A组人数: {len(df_a)}")
        print(f"B组人数: {len(df_b)}")
        print(f"总参与排名人数: {len(df_a) + len(df_b)}")

    print("分层总排名和分层设置创建完成！")
    return df


def 录入数据(考试名称, is_first=False, apply_layers=False):
    from 数据分析 import 分层颜色配置, 分层排名配置

    process_exam_data(考试名称, is_first)

    if apply_layers:
        if not 分层排名配置.get('启用', True):
            print("\n⚠️ 分层排名未启用，跳过分层设置")
            return

        df_layer = create_layer_ranking()

        layer_mode = 分层颜色配置.get('mode', 'parallel')
        layer_config = 分层颜色配置.get('config', [42, 40, 32])
        parallel_groups = 分层颜色配置.get('parallel_groups', 2)

        print(f"\n正在应用分层设置 (模式: {layer_mode})...")

        df_all = pd.read_excel('数据分析.xlsx', sheet_name='分层设置')

        if layer_mode == 'direct':
            setup_direct_layers(layer_config, df_all)
        else:
            setup_parallel_layers(layer_config, df_all, parallel_groups)

        print("分层设置已应用完成！\n")


# ==================== 排名变化函数 ====================

def 生成排名变化(考试字典=None, 考试名称=None, 名单路径=None,
                  列名映射=None, 名单列名=None, 输出sheet="排名变化"):
    print("\n正在生成排名变化...")

    if 考试字典 is not None:
        考试列表 = list(考试字典.keys())
        考试路径 = 考试字典
    else:
        考试列表 = []
        考试路径 = {}

    if 考试名称 is not None and len(考试名称) > 0:
        考试索引 = [考试列表.index(名) for 名 in 考试名称 if 名 in 考试列表]
    else:
        考试索引 = list(range(len(考试列表)))

    if 名单路径 is None or 列名映射 is None or 名单列名 is None:
        try:
            from 数据分析 import 排名变化配置
            if 名单路径 is None:
                名单路径 = 排名变化配置.get('名单路径', None)
            if 列名映射 is None:
                列名映射 = 排名变化配置.get('列名映射', {})
            if 名单列名 is None:
                名单列名 = 排名变化配置.get('名单列名', {})
        except ImportError:
            print("⚠️ 无法导入排名变化配置")
            return None

    if 列名映射 is None:
        列名映射 = {}
    if 名单列名 is None:
        名单列名 = {}

    学号列名 = 列名映射.get('学号列', '学号')
    姓名列名 = 列名映射.get('姓名列', '姓名')
    总分列映射 = 列名映射.get('总分列', {})
    排名列映射 = 列名映射.get('排名列', {})

    # 至少需要 2 个考试才能对比
    if 考试索引 is None or len(考试索引) < 2:
        print(f"⚠️ 排名变化至少需要 2 个考试（当前 {0 if 考试索引 is None else len(考试索引)} 个），跳过排名变化")
        return None

    for idx in 考试索引:
        if idx < 0 or idx >= len(考试列表):
            print(f"⚠️ 考试索引 {idx} 超出范围，跳过排名变化")
            return None

    考试名称 = [考试列表[i] for i in 考试索引]
    print(f"选择的考试: {考试名称}")

    # 读取名单
    if 名单路径 and os.path.exists(名单路径):
        df_名单 = pd.read_excel(名单路径)
        名单学号列 = 名单列名.get('学号列', 学号列名) if 名单列名 else 学号列名
        名单姓名列 = 名单列名.get('姓名列', 姓名列名) if 名单列名 else 姓名列名

        if 名单学号列 not in df_名单.columns:
            for col in df_名单.columns:
                if '学号' in col or 'id' in col.lower() or '学籍号' in col:
                    名单学号列 = col
                    break
        if 名单姓名列 not in df_名单.columns:
            for col in df_名单.columns:
                if '姓名' in col or 'name' in col.lower():
                    名单姓名列 = col
                    break

        df_名单 = df_名单[[名单学号列, 名单姓名列]].copy()
        df_名单 = df_名单.rename(columns={名单学号列: '学号', 名单姓名列: '姓名'})
        print(f"已加载名单: {名单路径}, 共 {len(df_名单)} 人")
    else:
        try:
            df_temp = pd.read_excel(考试路径[考试名称[0]])
            学号列 = 学号列名 if 学号列名 in df_temp.columns else None
            姓名列 = 姓名列名 if 姓名列名 in df_temp.columns else None
            if 学号列 is None:
                for col in df_temp.columns:
                    if '学号' in col or 'id' in col.lower():
                        学号列 = col
                        break
            if 姓名列 is None:
                for col in df_temp.columns:
                    if '姓名' in col or 'name' in col.lower():
                        姓名列 = col
                        break

            if 学号列 is not None and 姓名列 is not None:
                df_名单 = df_temp[[学号列, 姓名列]].copy()
                df_名单 = df_名单.rename(columns={学号列: '学号', 姓名列: '姓名'})
                print(f"从 {考试名称[0]} 中获取名单,共 {len(df_名单)} 人")
            else:
                print(f"❌ 无法识别学号和姓名列")
                return None
        except Exception as e:
            print(f"❌ 读取考试文件失败: {e}")
            return None

    # 读取每次考试的数据
    df_list = []
    for i, 名称 in enumerate(考试名称):
        try:
            df_temp = pd.read_excel(考试路径[名称])

            学号列 = 学号列名 if 学号列名 in df_temp.columns else None
            姓名列 = 姓名列名 if 姓名列名 in df_temp.columns else None
            if 学号列 is None:
                for col in df_temp.columns:
                    if '学号' in col or 'id' in col.lower():
                        学号列 = col
                        break
            if 姓名列 is None:
                for col in df_temp.columns:
                    if '姓名' in col or 'name' in col.lower():
                        姓名列 = col
                        break

            总分列 = 总分列映射.get(名称) if 总分列映射 else None
            if 总分列 is None or 总分列 not in df_temp.columns:
                总分列 = None
                for col in df_temp.columns:
                    if '总分' in col or '总成绩' in col or '合计' in col:
                        总分列 = col
                        break

            排名列 = 排名列映射.get(名称) if 排名列映射 else None
            if 排名列 is None or 排名列 not in df_temp.columns:
                排名列 = None
                for col in df_temp.columns:
                    if '排名' in col or '名次' in col:
                        排名列 = col
                        break

            if 学号列 is None or 姓名列 is None:
                print(f"  ⚠️ {名称} 中未找到学号或姓名列")
                continue
            if 总分列 is None:
                print(f"  ⚠️ {名称} 中未找到总分列")
                continue
            if 排名列 is None:
                print(f"  ⚠️ {名称} 中未找到排名列")
                continue

            df_temp = df_temp[[学号列, 姓名列, 总分列, 排名列]].copy()
            df_temp = df_temp.rename(columns={
                学号列: '学号',
                姓名列: '姓名',
                总分列: 名称,
                排名列: 名称 + "排"
            })
            df_list.append(df_temp)
            print(f"  ✅ 已读取: {名称}")
        except Exception as e:
            print(f"  ❌ 读取 {名称} 时出错: {e}")
            continue

    if len(df_list) < 2:
        print(f"❌ 成功读取的考试数据不足2个")
        return None

    # 合并数据
    df_result = df_名单[['学号', '姓名']].copy()
    for i, df_temp in enumerate(df_list):
        df_result = pd.merge(
            df_result,
            df_temp[['姓名', 考试名称[i], 考试名称[i] + "排"]],
            on='姓名',
            how='left'
        )

    # 👇 核心修改：最近一次考试 vs 前面 N-1 次考试
    最近一次名称 = 考试名称[-1]
    
    # 生成动态列
    对比列名列表 = []
    for i in range(len(考试名称) - 1):
        历史名称 = 考试名称[i]
        列名 = f"较{历史名称}"
        # 计算进步/退步名次（正数=进步，负数=退步）
        df_result[列名] = df_result[历史名称 + "排"] - df_result[最近一次名称 + "排"]
        对比列名列表.append(列名)

    # 整理输出列的顺序
    cols = ['姓名', '学号']
    for 名称 in 考试名称:
        cols.extend([名称, 名称 + "排"])
    cols.extend(对比列名列表)  # 加入所有对比列
    
    df_result = df_result[cols]
    
    # 默认按"较倒数第二次考试"的进步情况降序排
    if len(对比列名列表) > 0:
        排序基准 = 对比列名列表[-1]
        df_result = df_result.sort_values(by=[排序基准], ascending=False)
    
    df_result = df_result.reset_index(drop=True)

    print(f"\n排名变化统计:")
    print(f"  总人数: {len(df_result)}")
    for 列名 in 对比列名列表:
        print(f"  {列名}: 最多进步 {df_result[列名].max()} 名，最多退步 {df_result[列名].min()} 名")

    try:
        if os.path.exists('数据分析.xlsx'):
            with pd.ExcelWriter('数据分析.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_result.to_excel(writer, sheet_name=输出sheet, index=False)
        else:
            with pd.ExcelWriter('数据分析.xlsx', engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name=输出sheet, index=False)
        print(f"\n✅ 排名变化已写入 '数据分析.xlsx' 的 '{输出sheet}' Sheet")
        
        # ==================== 👇 新增：调整格式 ====================
        import openpyxl
        from openpyxl.styles import Alignment
        
        wb_fmt = openpyxl.load_workbook('数据分析.xlsx')
        if 输出sheet in wb_fmt.sheetnames:
            ws_fmt = wb_fmt[输出sheet]
            
            # 1. 遍历第一行（表头），给“较”开头的列设置合适的宽度和自动换行
            for col_idx, cell in enumerate(ws_fmt[1], start=1):
                # 如果要调整所有表头，或者只针对对比列
                if cell.value and str(cell.value).startswith('较'):
                    # 获取列字母并设置宽度（比如宽度设为 12，配合换行刚好显示两行）
                    from openpyxl.utils import get_column_letter
                    col_letter = get_column_letter(col_idx)
                    ws_fmt.column_dimensions[col_letter].width = 12
                    
                    # 设置该列所有单元格自动换行、居中
                    for row_idx in range(1, ws_fmt.max_row + 1):
                        target_cell = ws_fmt.cell(row=row_idx, column=col_idx)
                        target_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            _安全保存工作簿(wb_fmt, '数据分析.xlsx')
            print("  ✅ 已调整排名变化页对比列的宽度与自动换行")
        # ==========================================================

    except Exception as e:
        print(f"❌ 写入Excel时出错: {e}")
        return None

    return df_result


# ==================== 控制页函数 ====================

def 生成控制页(名单路径=None, 考试字典=None, 考试名称=None,
               查询范围="B2:G333", 图表位置=None, 隐藏前列=True):
    print("\n正在生成控制页...")

    if 考试字典 is not None:
        考试列表 = list(考试字典.keys())
        考试路径 = 考试字典
    else:
        考试列表 = []
        考试路径 = {}

    if 考试名称 is not None and len(考试名称) > 0:
        考试索引 = [考试列表.index(名) for 名 in 考试名称 if 名 in 考试列表]
    else:
        考试索引 = list(range(min(3, len(考试列表))))

    # 简化兜底：只取"名单路径、查询范围、图表位置、隐藏前列"（考试列表/考试路径已通过参数传入）
    if 名单路径 is None or 查询范围 is None or 图表位置 is None or 隐藏前列 is None:
        try:
            from 数据分析 import 控制页配置
            if 名单路径 is None:
                名单路径 = 控制页配置.get('名单路径', r'./统考数学成绩/0-2024名单.xlsx')
            if 查询范围 is None:
                查询范围 = 控制页配置.get('查询范围', 'B2:G333')
            if 图表位置 is None:
                图表位置 = 控制页配置.get('图表位置', {'柱状图': 'E3', '折线图': 'P3'})
            if 隐藏前列 is None:
                隐藏前列 = 控制页配置.get('隐藏前列', True)
        except ImportError:
            print("⚠️ 无法导入控制页配置")
            return None

    if 考试路径 is None:
        考试路径 = {}
    if 图表位置 is None:
        图表位置 = {'柱状图': 'E3', '折线图': 'P3'}

    if 考试索引 is None or len(考试索引) == 0:
        raise ValueError("考试索引不能为空")

    for idx in 考试索引:
        if idx < 0 or idx >= len(考试列表):
            raise ValueError(f"索引 {idx} 超出范围")

    考试名称 = [考试列表[i] for i in 考试索引]
    n_考试 = len(考试名称)
    print(f"控制页使用的考试: {考试名称}（共 {n_考试} 个）")

    if 名单路径 and os.path.exists(名单路径):
        df0 = pd.read_excel(名单路径)
        姓名列 = None
        for col in df0.columns:
            if '姓名' in col or 'name' in col.lower():
                姓名列 = col
                break
        if 姓名列 is None:
            print(f"⚠️ 名单中未找到姓名列")
            return None
        mingdan = df0[姓名列].to_frame()
        print(f"已加载名单: {名单路径}, 共 {len(mingdan)} 人")
    else:
        print(f"⚠️ 名单文件不存在: {名单路径}")
        return None

    from 数据分析 import 考试配置 as _考试配置

    def 取分段名(名称):
        if 名称 in _考试配置 and '表格' in _考试配置[名称]:
            return _考试配置[名称]['表格'][0]
        return 名称 + "分段"

    分段列表 = [取分段名(名) for 名 in 考试名称]

    val列表 = []
    for 分段名 in 分段列表:
        try:
            df_tmp = pd.read_excel('数据分析.xlsx', sheet_name=分段名)
            val列表.append(df_tmp.iloc[-7].values)
        except:
            val列表.append([0, 0, 0, 0, 0, 0, 0])

    wb = openpyxl.load_workbook('数据分析.xlsx')
    if '控制页' in wb.sheetnames:
        del wb['控制页']

    sheet = wb.create_sheet('控制页')
    wk = wb['控制页']

    for row in dataframe_to_rows(mingdan, index=False, header=True):
        wk.append(row)

    dv = DataValidation(
        type="list",
        formula1=f'=OFFSET($A$1,MATCH("*"&C5&"*",A:A,0)-1,0,COUNTIFS(A:A,"*"&C5&"*"),1)',
        allow_blank=True
    )
    wk.add_data_validation(dv)
    dv.add(wk["C5"])
    wk.merge_cells('C5:D6')

    fontText1 = Font(name='宋体', size=18, color='C00000', bold=True)
    wk.cell(5, 3).font = fontText1

    wk.cell(row=5, column=5, value='rk')
    wk.cell(row=5, column=6, value='学期')
    wk.cell(row=5, column=7, value='基础')
    wk.cell(row=5, column=8, value='基参')
    wk.cell(row=5, column=9, value='中档')
    wk.cell(row=5, column=10, value='中参')
    wk.cell(row=5, column=11, value='难题')
    wk.cell(row=5, column=12, value='难参')

    for j, 名 in enumerate(考试名称):
        wk.cell(row=6 + j, column=6, value=名)

    for j, 分段名 in enumerate(分段列表):
        行号 = 6 + j
        sheet[f"E{行号}"] = f'=VLOOKUP(C{5},{分段名}!{查询范围},2,0)'
        sheet[f"G{行号}"] = f'=VLOOKUP(C{5},{分段名}!{查询范围},4,0)'
        sheet[f"I{行号}"] = f'=VLOOKUP(C{5},{分段名}!{查询范围},5,0)'
        sheet[f"K{行号}"] = f'=VLOOKUP(C{5},{分段名}!{查询范围},6,0)'

    for j, val in enumerate(val列表):
        行号 = 6 + j
        try:
            wk.cell(row=行号, column=8, value=val[4])
            wk.cell(row=行号, column=10, value=val[5])
            wk.cell(row=行号, column=12, value=val[6])
        except:
            pass

    align = Alignment(horizontal='center', vertical='center')
    for i in range(1, 200):
        for j in range(1, 60):
            wk.cell(i, j).alignment = align

    wb.active = len(wb.sheetnames) - 1

    # 创建柱状图
    try:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.x_axis.title = sheet.cell(column=6, row=6).value
        chart.y_axis.delete = False
        chart.height = 9
        chart.width = 18

        data = Reference(sheet, min_col=6, max_col=12, min_row=5, max_row=5 + n_考试)
        label = Reference(sheet, min_col=6, min_row=6, max_row=5 + n_考试)
        chart.set_categories(label)
        chart.add_data(data, from_rows=False, titles_from_data=True)
        sheet.add_chart(chart, 图表位置.get('柱状图', 'E3'))
    except Exception as e:
        print(f"  ⚠️ 创建柱状图时出错: {e}")

    # 创建折线图
    try:
        linechart = LineChart()
        linechart.style = 10
        linechart.title = "rk"
        linechart.x_axis.delete = False
        linechart.y_axis.delete = False
        linechart.height = 9
        linechart.width = 18

        data = Reference(sheet, min_col=5, max_col=5, min_row=5, max_row=5 + n_考试)
        linechart.add_data(data, titles_from_data=True)
        label = Reference(sheet, min_col=6, max_col=6, min_row=6, max_row=5 + n_考试)
        linechart.set_categories(label)

        linechart.y_axis.scaling.orientation = "maxMin"

        linechart.dataLabels = DataLabelList()
        linechart.dataLabels.showVal = True
        linechart.dataLabels.showSerName = False
        linechart.dataLabels.showCatName = False
        linechart.dataLabels.showLegendKey = False

        sheet.add_chart(linechart, 图表位置.get('折线图', 'N3'))
    except Exception as e:
        print(f"  ⚠️ 创建折线图时出错: {e}")

    if 隐藏前列:
        try:
            sheet.column_dimensions.group('A', 'B', hidden=True)
        except:
            pass

     # ==================== 各题得分率（每次考试占 3 行：题号 / 知识点 / 得分率） ====================
    try:
        from 数据分析 import 考试配置, 控制页配置

        知识点配置 = 控制页配置.get('各题得分率', {})
        if 知识点配置.get('启用', True):
            起始行 = 知识点配置.get('起始行', 23)
            起始列 = 知识点配置.get('起始列', 3)

            print(f"\n  正在录入各题得分率(每次考试占3行:题号/知识点/得分率)...")

            # 最大题数
            最大题数 = 0
            for 考试名称_i in 考试名称:
                if 考试名称_i in 考试配置 and '知识点' in 考试配置[考试名称_i]:
                    最大题数 = max(最大题数, len(考试配置[考试名称_i]['知识点']))
            if 最大题数 == 0:
                最大题数 = 28
            print(f"  最大题数: {最大题数}")

            # ===== 第一趟：收集所有考试的信息 =====
            考试信息列表 = []
            for i, 考试名称_i in enumerate(考试名称):
                if 考试名称_i not in 考试配置:
                    print(f"  ⚠️ {考试名称_i} 不在考试配置中")
                    continue
                config = 考试配置[考试名称_i]
                if '知识点' not in config or '各项分值' not in config:
                    print(f"  ⚠️ {考试名称_i} 缺少知识点或分值配置")
                    continue
                知识点列表 = config['知识点']
                分值列表 = config['各项分值']
                if len(知识点列表) != len(分值列表):
                    print(f"  ⚠️ {考试名称_i} 知识点与分值数量不一致")
                    continue

                sheet_name = 考试配置[考试名称_i]['表格'][0] if '表格' in config else 考试名称_i + "分段"
                if sheet_name not in wb.sheetnames:
                    print(f"  ⚠️ {sheet_name} 不存在,跳过")
                    continue

                try:
                    df_exam = pd.read_excel('数据分析.xlsx', sheet_name=sheet_name)
                except Exception as e:
                    print(f"  ⚠️ 读取 {sheet_name} 失败: {e}")
                    continue

                题号列名列表 = []
                for col in df_exam.columns:
                    col_str = str(col).split('.')[0]
                    if col_str.isdigit():
                        题号列名列表.append(col)

                if len(题号列名列表) < len(分值列表):
                    print(f"  ⚠️ {sheet_name} 题号列不足({len(题号列名列表)} < {len(分值列表)})")
                    continue

                try:
                    题号列名列表 = sorted(题号列名列表, key=lambda x: int(str(x).split('.')[0]))
                except:
                    pass

                姓名列名 = None
                for col in df_exam.columns:
                    if '姓名' in str(col):
                        姓名列名 = col
                        break
                if 姓名列名 is None:
                    print(f"  ⚠️ {sheet_name} 没有姓名列")
                    continue

                考试信息列表.append({
                    '名称': 考试名称_i,
                    '题号列名': 题号列名列表,
                    '知识点列表': 知识点列表,
                    '分值列表': 分值列表,
                    '数据': df_exam,
                    '姓名列': 姓名列名,
                })

            # ===== 第二趟：写隐藏区 =====
            隐藏起始行 = 200
            隐藏起始列 = 27   # AA 列

            wk.cell(row=隐藏起始行 - 1, column=隐藏起始列, value='拼接键')
            wk.cell(row=隐藏起始行 - 1, column=隐藏起始列 + 1, value='学生')
            wk.cell(row=隐藏起始行 - 1, column=隐藏起始列 + 2, value='考试')
            for k in range(最大题数):
                wk.cell(row=隐藏起始行 - 1, column=隐藏起始列 + 3 + k * 3, value=f'题号{k+1}')
                wk.cell(row=隐藏起始行 - 1, column=隐藏起始列 + 4 + k * 3, value=f'知识点{k+1}')
                wk.cell(row=隐藏起始行 - 1, column=隐藏起始列 + 5 + k * 3, value=f'得分率{k+1}')

            当前隐藏行 = 隐藏起始行
            学生列表 = mingdan.iloc[:, 0].tolist()

            for 学生 in 学生列表:
                for 考试信息 in 考试信息列表:
                    df_exam = 考试信息['数据']
                    姓名列名 = 考试信息['姓名列']
                    题号列名列表 = 考试信息['题号列名']
                    知识点列表 = 考试信息['知识点列表']
                    分值列表 = 考试信息['分值列表']
                    考试名称_i = 考试信息['名称']

                    row_in_exam = df_exam[df_exam[姓名列名] == 学生]
                    if row_in_exam.empty:
                        continue
                    row_in_exam = row_in_exam.iloc[0]

                    wk.cell(row=当前隐藏行, column=隐藏起始列, value=f'{学生}|{考试名称_i}')
                    wk.cell(row=当前隐藏行, column=隐藏起始列 + 1, value=学生)
                    wk.cell(row=当前隐藏行, column=隐藏起始列 + 2, value=考试名称_i)

                    本次题数 = len(题号列名列表)

                    错题列表 = []
                    for k, 知识点 in enumerate(知识点列表):
                        if k >= 本次题数:
                            break
                        分值 = 分值列表[k]
                        列名 = 题号列名列表[k]
                        try:
                            得分 = float(row_in_exam[列名])
                        except:
                            得分 = 0
                        if 分值 > 0:
                            得分率 = round(得分 / 分值, 4)
                        else:
                            得分率 = 1
                        if 得分率 < 1:
                            题号 = str(题号列名列表[k]).split('.')[0]
                            错题列表.append((题号, 知识点, 得分率))

                    错题列表.sort(key=lambda x: x[2])

                    for k, (题号, 知识点, 得分率) in enumerate(错题列表):
                        if k >= 最大题数:
                            break
                        基列 = 隐藏起始列 + 3 + k * 3
                        wk.cell(row=当前隐藏行, column=基列, value=题号)
                        wk.cell(row=当前隐藏行, column=基列 + 1, value=知识点)
                        wk.cell(row=当前隐藏行, column=基列 + 2, value=得分率)

                    当前隐藏行 += 1

            隐藏结束行 = 当前隐藏行

            # ===== 第三趟：写可见区 =====
            当前行 = 起始行
            数据起始列 = 起始列 + 1      # D 列

            for i, 考试名称_i in enumerate(考试名称):
                题号行 = 当前行
                知识点行 = 当前行 + 1
                得分率行 = 当前行 + 2

                # C 列写标签
                wk.cell(row=题号行, column=起始列, value=考试名称_i)
                wk.cell(row=题号行, column=起始列).font = Font(bold=True)
                wk.cell(row=知识点行, column=起始列, value='知识点')
                wk.cell(row=得分率行, column=起始列, value='得分率')

                for k in range(最大题数):
                    列 = 数据起始列 + k
                    # 隐藏区列号：拼接键=1, 学生=2, 考试=3, 题号1=4, 知识点1=5, 得分率1=6, ...
                    题号列号 = 4 + k * 3
                    知识点列号 = 5 + k * 3
                    得分率列号 = 6 + k * 3
                    查找值 = f'$C$5&"|"&"{考试名称_i}"'

                    # 第 1 行：第几题
                    wk.cell(
                        row=题号行, column=列,
                        value=(
                            f'=IFERROR('
                            f'IF(VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{题号列号},0)="",'
                            f'"",'
                            f'"第"&VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{题号列号},0)&"题"'
                            f'),"")'
                        )
                    )
                    # 第 2 行：知识点
                    wk.cell(
                        row=知识点行, column=列,
                        value=(
                            f'=IFERROR('
                            f'IF(VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{知识点列号},0)="",'
                            f'"",'
                            f'VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{知识点列号},0)'
                            f'),"")'
                        )
                    )
                    # 第 3 行：得分率
                    wk.cell(
                        row=得分率行, column=列,
                        value=(
                            f'=IFERROR('
                            f'IF(VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{得分率列号},0)="",'
                            f'"",'
                            f'VLOOKUP({查找值},$AA${隐藏起始行}:$ZZ${隐藏结束行},{得分率列号},0)'
                            f'),"")'
                        )
                    )

                当前行 += 3

            # ===== 条件格式（只作用在有值的单元格） =====
            from openpyxl.formatting.rule import FormulaRule
            from openpyxl.utils import get_column_letter

            当前行 = 起始行
            for i, 考试名称_i in enumerate(考试名称):
                得分率行 = 当前行 + 2
                起始列字母 = get_column_letter(数据起始列)
                结束列字母 = get_column_letter(数据起始列 + 最大题数 - 1)
                区域 = f'{起始列字母}{得分率行}:{结束列字母}{得分率行}'

                # 红色 <0.5（且不为空）
                sheet.conditional_formatting.add(
                    区域,
                    FormulaRule(
                        formula=[f'AND({起始列字母}{得分率行}<>"", {起始列字母}{得分率行}<0.5)'],
                        fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
                        font=Font(color='9C0006')
                    )
                )
                # 棕色 0.5~0.8（且不为空）
                sheet.conditional_formatting.add(
                    区域,
                    FormulaRule(
                        formula=[f'AND({起始列字母}{得分率行}<>"", {起始列字母}{得分率行}>=0.5, {起始列字母}{得分率行}<0.8)'],
                        fill=PatternFill(start_color='FFD9A0', end_color='FFD9A0', fill_type='solid'),
                        font=Font(color='974706')
                    )
                )
                # 绿色 >=0.8（且不为空）
                sheet.conditional_formatting.add(
                    区域,
                    FormulaRule(
                        formula=[f'AND({起始列字母}{得分率行}<>"", {起始列字母}{得分率行}>=0.8)'],
                        fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'),
                        font=Font(color='006100')
                    )
                )

                当前行 += 3

            try:
                sheet.column_dimensions.group('AA', 'BZ', hidden=True)
            except:
                pass

            print(f"  ✅ 各题得分率已写入（每次考试占3行，共{len(考试信息列表)}次考试）")

    except Exception as e:
        print(f"  ⚠️ 添加各题得分率时出错: {e}")
        import traceback
        traceback.print_exc()

    # ==================== 设置C至N列宽度 ====================
    try:
        from openpyxl.utils import get_column_letter
        for col_idx in range(3, 15):  # C=3, N=14
            col_letter = get_column_letter(col_idx)
            sheet.column_dimensions[col_letter].width = 10
        print("  ✅ 已设置控制页 C至N 列宽度为 10")
    except Exception as e:
        print(f"  ⚠️ 设置列宽时出错: {e}")

    # ==================== 设置23至38行行高为20 ====================
    try:
        for row_idx in range(23, 38):  # 23 到 32 行（含）
            sheet.row_dimensions[row_idx].height = 25
        print("  ✅ 已设置控制页 23至32 行行高为 25")
    except Exception as e:
        print(f"  ⚠️ 设置行高时出错: {e}")

    # ==================== 设置D23到W列自动换行 ====================
    # ==================== 新增：生成历次综合学习建议（真实得分率动态版） ====================
    from 数据分析 import 控制页配置 as _控制页配置
    if _控制页配置.get('综合建议启用', True):
        try:
            from 数据分析 import 考试配置 as _考试配置
            from openpyxl.styles import Border, Side
            
            # 1. 动态定位建议区行号 (3次=32, 4次=35, 5次=38...)
            建议标题行 = 23 + n_考试 * 3
            建议表头行 = 建议标题行 + 1
            建议数据行 = 建议表头行 + 1

            print(f"\n  正在后台计算所有学生的综合学习建议 (固定在行 {建议标题行} 起)...")
            
            # 2. 收集所有学生的汇总数据
            学生汇总 = []

            # 性能优化：考试工作表只读取一次，并预先建立姓名索引。
            # 原实现把 read_excel 放在“学生 × 考试”的双重循环内，名单较大时会反复
            # 解压同一 xlsx 数百/数千次；这里缓存数据，不改变统计口径。
            _建议考试缓存 = {}
            for _考试名 in 考试名称:
                if _考试名 not in _考试配置:
                    continue
                _sheet_name = _考试配置[_考试名]['表格'][0]
                try:
                    _df_exam = pd.read_excel('数据分析.xlsx', sheet_name=_sheet_name)
                    if '姓名' in _df_exam.columns:
                        _姓名索引 = _df_exam.drop_duplicates(subset=['姓名'], keep='first').set_index('姓名', drop=False)
                    else:
                        _姓名索引 = None
                    _题号列 = [c for c in _df_exam.columns if str(c).split('.')[0].isdigit()]
                    _题号列 = sorted(_题号列, key=lambda x: int(str(x).split('.')[0]))
                    _基满 = _中满 = _难满 = 0.0
                    for _r_idx in range(len(_df_exam)-1, -1, -1):
                        if str(_df_exam.iloc[_r_idx, 0]).strip() == '各项分值':
                            try:
                                _基满, _中满, _难满 = (float(_df_exam.iloc[_r_idx, _c]) for _c in (4, 5, 6))
                            except (ValueError, TypeError, IndexError):
                                pass
                            break
                    if _基满 == 0:
                        _单次总分 = sum(_考试配置[_考试名].get('各项分值', [100]))
                        _基满, _中满, _难满 = _单次总分 * 0.6, _单次总分 * 0.3, _单次总分 * 0.1
                    _建议考试缓存[_考试名] = (_df_exam, _姓名索引, _题号列, (_基满, _中满, _难满))
                except Exception:
                    continue

            for idx, name_cell in enumerate(mingdan['姓名']):
                if pd.isna(name_cell):
                    continue
                学生 = str(name_cell).strip()
                
                累计得分 = {'基础': 0, '中档': 0, '难题': 0}
                累计满分 = {'基础': 0, '中档': 0, '难题': 0}
                知识点失分统计 = {}
                
                for i, 考试名 in enumerate(考试名称):
                    if 考试名 not in _考试配置:
                        continue
                    
                    _缓存项 = _建议考试缓存.get(考试名)
                    if _缓存项 is None:
                        continue
                    df_exam, _姓名索引, 题号列, (_基满缓存, _中满缓存, _难满缓存) = _缓存项
                    if _姓名索引 is None or 学生 not in _姓名索引.index:
                        continue
                    行 = _姓名索引.loc[学生]
                    if isinstance(行, pd.DataFrame):
                        行 = 行.iloc[0]
                    
                    基得 = float(行.get('基础题', 0)) if pd.notna(行.get('基础题')) else 0
                    中得 = float(行.get('中档题', 0)) if pd.notna(行.get('中档题')) else 0
                    难得 = float(行.get('难题', 0)) if pd.notna(行.get('难题')) else 0
                    
                    # 满分已在读取工作表时预先提取，避免每个学生都反向扫描整张表。
                    基满, 中满, 难满 = _基满缓存, _中满缓存, _难满缓存

                    累计得分['基础'] += 基得
                    累计得分['中档'] += 中得
                    累计得分['难题'] += 难得
                    累计满分['基础'] += 基满
                    累计满分['中档'] += 中满
                    累计满分['难题'] += 难满
                    
                    # 统计知识点
                    if '知识点' in _考试配置[考试名] and '各项分值' in _考试配置[考试名]:
                        知识点列表 = _考试配置[考试名]['知识点']
                        分值列表 = _考试配置[考试名]['各项分值']
                        # 题号列已在缓存阶段排序。
                        for k, 知识点 in enumerate(知识点列表):
                            if k >= len(题号列): break
                            得分 = float(行.get(题号列[k], 0)) if pd.notna(行.get(题号列[k])) else 0
                            满分 = 分值列表[k]
                            if 满分 > 0:
                                失分率 = (满分 - 得分) / 满分
                                if 失分率 > 0.3:
                                    if 知识点 not in 知识点失分统计:
                                        知识点失分统计[知识点] = 0
                                    知识点失分统计[知识点] += 失分率
                
                # 计算真实的得分率
                综合率 = {}
                for 类别 in ['基础', '中档', '难题']:
                    if 累计满分[类别] > 0:
                        综合率[类别] = 累计得分[类别] / 累计满分[类别]
                    else:
                        综合率[类别] = 0
                
                def 评级(率, 优, 良, 合):
                    if 率 >= 优: return "优秀"
                    elif 率 >= 良: return "良好"
                    elif 率 >= 合: return "合格"
                    else: return "薄弱"
                    
                基评 = 评级(综合率['基础'], 0.90, 0.75, 0.60)
                中评 = 评级(综合率['中档'], 0.80, 0.65, 0.50)
                难评 = 评级(综合率['难题'], 0.60, 0.40, 0.25)
                
                建议列表 = []
                if 基评 == "薄弱":
                    建议列表.append(f"【急需补基础】基础综合率仅 {综合率['基础']*100:.1f}%，建议暂停刷难题，回归课本死磕概念和计算。")
                elif 中评 == "薄弱" and 基评 != "薄弱":
                    建议列表.append(f"【中档有断层】基础尚可（{综合率['基础']*100:.1f}%），但中档率 {综合率['中档']*100:.1f}%，需多总结题型模型。")
                elif 基评 in ["优秀","良好"] and 中评 in ["优秀","良好"]:
                    if 难评 == "薄弱":
                        建议列表.append(f"【需突破压轴】基础({综合率['基础']*100:.1f}%)与中档({综合率['中档']*100:.1f}%)扎实，但难题仅 {综合率['难题']*100:.1f}%。")
                    else:
                        建议列表.append(f"【全面发展】基础、中档、难题表现优异，继续保持！")
                else:
                    建议列表.append("【整体良好】成绩较为均衡，建议针对自身薄弱环节进行专题突破。")
                    
                if 知识点失分统计:
                    薄弱点 = sorted(知识点失分统计.items(), key=lambda x: x[1], reverse=True)[:3]
                    点名 = "、".join([k for k, v in 薄弱点])
                    建议列表.append(f"【重点攻克】多次在「{点名}」上丢分，请找对应专项卷进行限时集训。")
                
                建议文本 = " ".join(建议列表)
                
                学生汇总.append({
                    '姓名': 学生,
                    '基础率': 综合率['基础'],
                    '中档率': 综合率['中档'],
                    '难题率': 综合率['难题'],
                    '评级': f"基:{基评}/中:{中评}/难:{难评}",
                    '建议': 建议文本
                })

            # 3. 创建独立隐藏 Sheet 存数据
            if '评价数据' in wb.sheetnames:
                del wb['评价数据']
            ws_data = wb.create_sheet('评价数据')
            ws_data.append(['姓名', '基础率', '中档率', '难题率', '评级', '建议'])
            for data in 学生汇总:
                ws_data.append([data['姓名'], data['基础率'], data['中档率'], data['难题率'], data['评级'], data['建议']])
            ws_data.sheet_state = 'hidden'

            # 4. 可见区写表头及公式
            # 标题以“=”开头，openpyxl 可能将其识别为公式，导致 Excel 打开时删除公式。
            # 显式将单元格数据类型设为字符串，保留原显示内容和合并格式。
            标题单元格 = wk.cell(row=建议标题行, column=3)
            标题单元格.value = "=== 历次考试综合能力评价与学习建议 ==="
            标题单元格.data_type = 's'
            wk.merge_cells(start_row=建议标题行, start_column=3, end_row=建议标题行, end_column=12) 
            wk.cell(row=建议标题行, column=3).font = Font(bold=True, size=12, color='C00000')
            wk.cell(row=建议标题行, column=3).alignment = Alignment(horizontal='left', vertical='center')
            
            表头 = ['姓名', '基础得分率', '中档得分率', '难题得分率', '综合评价', '学习建议']
            表头列映射 = [3, 4, 5, 6, 7, 8]
            for c_idx, h_text in zip(表头列映射, 表头):
                c = wk.cell(row=建议表头行, column=c_idx, value=h_text)
                c.fill = PatternFill('solid', fgColor='87CEFA')
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal='center', vertical='center')

            wk.merge_cells(start_row=建议表头行, start_column=8, end_row=建议表头行, end_column=13)
            wk.merge_cells(start_row=建议数据行, start_column=8, end_row=建议数据行, end_column=13)

            查找区域 = f"'评价数据'!$A$1:$F${1+len(学生汇总)}"
            
            wk.cell(row=建议数据行, column=3, value='=C5')
            wk.cell(row=建议数据行, column=4, value=f'=IFERROR(VLOOKUP($C$5,{查找区域},2,FALSE),"")')
            wk.cell(row=建议数据行, column=5, value=f'=IFERROR(VLOOKUP($C$5,{查找区域},3,FALSE),"")')
            wk.cell(row=建议数据行, column=6, value=f'=IFERROR(VLOOKUP($C$5,{查找区域},4,FALSE),"")')
            wk.cell(row=建议数据行, column=7, value=f'=IFERROR(VLOOKUP($C$5,{查找区域},5,FALSE),"")')
            wk.cell(row=建议数据行, column=8, value=f'=IFERROR(VLOOKUP($C$5,{查找区域},6,FALSE),"")')

            for col_idx in [4, 5, 6]:
                wk.cell(row=建议数据行, column=col_idx).number_format = '0.0%'

            for c in range(3, 8):
                wk.cell(row=建议数据行, column=c).alignment = Alignment(horizontal='center', vertical='center')
            wk.cell(row=建议数据行, column=8).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            wk.row_dimensions[建议数据行].height = 60 
            
            细边框 = Border(
                left=Side(style='thin', color='E0E0E0'),
                right=Side(style='thin', color='E0E0E0'),
                top=Side(style='thin', color='E0E0E0'),
                bottom=Side(style='thin', color='E0E0E0')
            )
            for r in [建议表头行, 建议数据行]:
                for c in range(3, 14):
                    wk.cell(row=r, column=c).border = 细边框
            
            try:
                sheet.ignored_errors = openpyxl.worksheet.errors.IgnoredErrors()
                sheet.ignored_errors.numberStoredAsText = True
                sheet.ignored_errors.formula = True
                sheet.ignored_errors.formulaRange = True
            except:
                pass

            print(f"  ✅ 已追加综合建议动态展示 (可见区在第 {建议标题行} 行)")

        except Exception as e:
            print(f"  ⚠️ 生成综合建议时出错: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("\n  ⏭️ 综合学习建议已关闭，跳过生成。")
    # ==========================================================

    _安全保存工作簿(wb, '数据分析.xlsx')
    print("✅ 控制页已生成完成！")
    return True


def 应用格式(边框颜色="FF000000", 边框样式="thin", 控制页边框颜色="E0E0E0"):
    print("\n正在应用格式...")

    try:
        wb = openpyxl.load_workbook('数据分析.xlsx')

        border_style = Side(border_style=边框样式, color=边框颜色)
        border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        控制页边框样式 = Side(border_style=边框样式, color=控制页边框颜色)
        控制页边框 = Border(left=控制页边框样式, right=控制页边框样式,
                           top=控制页边框样式, bottom=控制页边框样式)

        align = Alignment(horizontal='center', vertical='center')
        换行对齐 = Alignment(horizontal='center', vertical='center', wrap_text=True)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.freeze_panes = 'A2'

            is_控制页 = (sheet_name == '控制页')

            if is_控制页:
                # 控制页只格式化可见区，避免遍历到隐藏的 AA~BZ 超大区域
                max_rows = min(ws.max_row, 200)
                max_columns = min(ws.max_column, 26)  # A~Z
            else:
                max_rows = ws.max_row
                max_columns = ws.max_column

            for i in range(1, max_rows + 1):
                for j in range(1, max_columns + 1):
                    cell = ws.cell(i, j)
                    if is_控制页 and i >= 23 and 3 <= j <= 23:
                        cell.alignment = 换行对齐   # D23:W 自动换行
                    else:
                        cell.alignment = align
                    if is_控制页:
                        cell.border = 控制页边框
                    else:
                        cell.border = border

        _安全保存工作簿(wb, '数据分析.xlsx')
        print("✅ 格式应用完成！")
        print(f"  普通页面边框: {边框颜色}")
        print(f"  控制页边框: {控制页边框颜色}")

    except Exception as e:
        print(f"⚠️ 应用格式时出错: {e}")
        raise