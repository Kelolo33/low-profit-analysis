import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

def process_subscription_file(subscription_file):
    # 读取海运订阅文件，移除 encoding 参数
    df = pd.read_excel(subscription_file)
    
    print("海运订阅文件的列名:", df.columns.tolist(), flush=True)
    print(f"原始数据行数: {len(df)}", flush=True)
    
    # 检查必需的列
    required_columns = ['二级部门', '委托客户', '客户约价', '是否低负', '未税人民币总毛利', '未税人民币总收入', '业务大类名称']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"海运订阅文件缺少以下列: {', '.join(missing_columns)}")
    
    # 筛选业务大类为海运的数据
    df = df[df['业务大类名称'] == '海运']
    print(f"筛选海运业务后的数据行数: {len(df)}", flush=True)
    
    # 选择所需列
    subscription_data = df[required_columns].copy()
    print(f"选择所需列后的数据行数: {len(subscription_data)}", flush=True)
    
    # 计算约价的负毛利票数和非约价的低负票数
    subscription_data['约价负毛利'] = ((subscription_data['客户约价'].notna() & (subscription_data['客户约价'] != 'N')) & 
                                  (subscription_data['是否低负'] == '负毛利'))
    subscription_data['非约价低负'] = ((subscription_data['客户约价'].isna() | (subscription_data['客户约价'] == 'N')) & 
                                  ((subscription_data['是否低负'] == '低毛利') | (subscription_data['是否低负'] == '负毛利')))
    
    print(f"约价负毛利的记录数: {subscription_data['约价负毛利'].sum()}", flush=True)
    print(f"非约价低负的记录数: {subscription_data['非约价低负'].sum()}", flush=True)
    
    # 分别计算约价和非约价的数据，只考虑 '是否低' 不为 'N' 的数据
    yue_data = subscription_data[(subscription_data['约价负毛利']) & (subscription_data['是否低负'] != 'N')].groupby(['二级部门', '委托客户']).agg({
        '未税人民币总毛利': 'sum',
        '未税人民币总收入': 'sum',
        '约价负毛利': 'count'
    }).reset_index()
    
    non_yue_data = subscription_data[(subscription_data['非约价低负']) & (subscription_data['是否低负'] != 'N')].groupby(['二级部门', '委托客户']).agg({
        '未税人民币总毛利': 'sum',
        '未税人民币总收入': 'sum',
        '非约价低负': 'count'
    }).reset_index()
    
    print(f"约价数据行数: {len(yue_data)}")
    print(f"非约价数据行数: {len(non_yue_data)}")
    
    # 所有委托客户
    all_customers = df[['二级部门', '委托客户']].drop_duplicates()
    
    # 合并约价和非约价数据，并确保包含所有委托客户
    grouped_data = pd.merge(all_customers, yue_data, on=['二级部门', '委托客户'], how='left')
    grouped_data = pd.merge(grouped_data, non_yue_data, on=['二级部门', '委托客户'], how='left')
    
    # 重命名列
    grouped_data = grouped_data.rename(columns={
        '约价负毛利': '约价负毛利票数',
        '非约价低负': '非约价低负票数',
        '未税人民币总毛利_x': '约价未税人民币总毛利',
        '未税人民币总收入_x': '约价未税人民币总收入',
        '未税人民币总毛利_y': '非约价未税人民币总毛利',
        '未税人民币总收入_y': '非约价未税人民币总收入',
    })
    
    # 填充NaN值为0
    fill_columns = ['约价负毛利票数', '非约价低负票数', '约价未税人民币总毛利', '约价未税人民币总收入', '非约价未税人民币总毛利', '非约价未税人民币总收入']
    grouped_data[fill_columns] = grouped_data[fill_columns].fillna(0)
    
    # 过滤掉约价负毛利票数和非约价低负票数都为0的记录
    grouped_data = grouped_data[
        (grouped_data['约价负毛利票数'] > 0) | 
        (grouped_data['非约价低负票数'] > 0)
    ]
    
    # 重新计算毛利率
    grouped_data['约价毛利率'] = np.where(
        grouped_data['约价未税人民币总收入'] != 0,
        grouped_data['约价未税人民币总毛利'] / grouped_data['约价未税人民币总收入'],
        -1
    )
    grouped_data['非约价毛利率'] = np.where(
        grouped_data['非约价未税人民币总收入'] != 0,
        grouped_data['非约价未税人民币总毛利'] / grouped_data['非约价未税人民币总收入'],
        -1
    )

    # 将毛利率转换为百分比格式的数值，票数为0时显示为空
    def format_rate(rate, count):
        if pd.isnull(rate) or pd.isnull(count) or count == 0:
            return None
        elif rate == -1:
            return -1  # 返回 -1，表示 -100%
        else:
            return rate  # 保持原始的小数形式
    
    grouped_data['约价毛利率'] = grouped_data.apply(lambda row: format_rate(row['约价毛利率'], row['约价负毛利票数']), axis=1)
    grouped_data['非约价毛利率'] = grouped_data.apply(lambda row: format_rate(row['非约价毛利率'], row['非约价低负票数']), axis=1)
    
    grouped_data['约价负毛利票数'] = grouped_data['约价负毛利票数'].apply(lambda x: '' if x == 0 else int(x))
    grouped_data['非约价低负票数'] = grouped_data['非约价低负票数'].apply(lambda x: '' if x == 0 else int(x))
    
    print("grouped_data 的前几行:")
    print(grouped_data.head().to_string())
    
    # 打印一些统信息
    print(f"\n约价负毛利票数不为空的记录数: {grouped_data['约价负毛利票数'].astype(bool).sum()}")
    print(f"非约价低负票数不为空的记录数: {grouped_data['非约价低负票数'].astype(bool).sum()}")
    print(f"约价毛利率不为空的记录数: {grouped_data['约价毛利率'].astype(bool).sum()}")
    print(f"非约价毛利率不为空的记录数: {grouped_data['非约价毛利率'].astype(bool).sum()}")
    
    # 计算每个委托客户的总毛利率
    total_stats = df.groupby(['二级部门', '委托客户']).agg({
        '未税人民币总毛利': 'sum',
        '未税人民币总收入': 'sum'
    }).reset_index()
    
    # 安全地计算总利润率
    def calculate_profit_rate(row):
        if pd.isna(row['未税人民币总收入']) or row['未税人民币总收入'] == 0:
            return 0  # 或者返回其他默认值
        return row['未税人民币总毛利'] / row['未税人民币总收入']
    
    total_stats['总利润率'] = total_stats.apply(calculate_profit_rate, axis=1)
    
    # 过滤掉异常值（如果需要的话）
    total_stats['总利润率'] = total_stats['总利润率'].apply(
        lambda x: x if abs(x) <= 1 else (1 if x > 1 else -1)
    )
    
    # 合并总利润率到grouped_data
    grouped_data = pd.merge(grouped_data, total_stats[['二级部门', '委托客户', '总利润率']], 
                          on=['二级部门', '委托客户'], how='left')
    
    return grouped_data

def analyze_excel_data(input_file, output_file, subscription_file, status_callback=None):
    if status_callback:
        status_callback("开始读取海运订阅文件...")
    
    # 读取海运订阅文件
    subscription_df = pd.read_excel(subscription_file)
    
    if status_callback:
        status_callback("处理海运订阅数据...")
    subscription_data = process_subscription_file(subscription_file)

    if input_file:
        if status_callback:
            status_callback("读取预对账文件...")
        df = pd.read_excel(input_file)
        
        if status_callback:
            status_callback("分析数据中...")
        # 确保必要的列存在
        required_columns = ['法人部门', '委托客户', '别名', '应收应付', '本位币金额', '费率单号', '币种']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"缺少必要的列: {col}")

        # 按法人部门、委托客户、费率单号和别名进行汇总
        grouped = df.groupby(['法人部门', '委托客户', '费率单号', '别名', '应收应付', '币种'])['本位币金额'].sum().unstack(level='应收应付').fillna(0)
        grouped = grouped.rename(columns={'应收': '应收金额', '应付': '应付金额'})
        grouped['费目利润'] = grouped['应收金额'] - grouped['应付金额']

        # 重置索引，使得所有列都变成普通列
        grouped = grouped.reset_index()

        # 先计算每个费率单号的总毛利和毛利率
        rate_totals = grouped.groupby('费率单号').agg({
            '应收金额': 'sum',
            '应付金额': 'sum'
        }).reset_index()
        
        rate_totals['单票毛利'] = rate_totals['应收金额'] - rate_totals['应付金额']
        rate_totals['单票毛利率'] = rate_totals.apply(
            lambda x: x['单票毛利'] / x['应收金额'] if x['应收金额'] != 0 else -1, 
            axis=1
        )
        
        # 只保留需要的列
        rate_totals = rate_totals[['费率单号', '单票毛利', '单票毛利率']]

        # 创建结果 DataFrame
        results = []
        seen_rate_nos = {}  # 用于跟踪已经见过的费率单号

        for _, row in grouped.iterrows():
            # 检查是否已经见过这个费率单号
            is_first_occurrence = row['费率单号'] not in seen_rate_nos
            if is_first_occurrence:
                seen_rate_nos[row['费率单号']] = True

            # 获取该费率单号的总计数据
            rate_total = rate_totals[rate_totals['费率单号'] == row['费率单号']].iloc[0]

            result = {
                '法人部门': row['法人部门'],
                '委托客户': row['委托客户'] if is_first_occurrence else '',
                '费率单号': row['费率单号'] if is_first_occurrence else '',
                '别名': row['别名'],
                '币种': row['币种'],
                '应收金额': row['应收金额'],
                '应付金额': row['应付金额'],
                '费目利润': row['费目利润'],
                '类型': '',
                '单票毛利': rate_total['单票毛利'] if is_first_occurrence else '',  # 使用费率单总毛利
                '单票毛利率': rate_total['单票毛利率'] if is_first_occurrence else ''  # 使用费率单总毛利率
            }

            if row['应付金额'] > 0 and row['应收金额'] == 0:
                result['类型'] = '无应收'
            elif row['应收金额'] < row['应付金额']:
                result['类型'] = '倒挂'

            results.append(result)

        # 创建结果 DataFrame
        result_df = pd.DataFrame(results)
        
        # 创建客户公司分析数据
        customer_analysis = result_df.groupby('委托客户').agg({
            '费率单号': lambda x: len(x.unique()),  # 统计唯一费率单号的数量
            '费目利润': 'sum',  # 总金额
            '法人部门': 'first',  # 获取法人部门
        }).reset_index()

        customer_analysis['初步分析'] = result_df.groupby('委托客户').apply(format_analysis).reset_index(drop=True)
        customer_analysis = customer_analysis.rename(columns={
            '费率单号': '总票数',
            '费目利润': '总金额',
        })
    else:
        # 如果没有预对账文件，创建一个空的customer_analysis DataFrame
        customer_analysis = pd.DataFrame(columns=['委托客户', '总票数', '总金额', '初步分析'])

    # 将海运订阅文件中的所有二级部门和委托客户信息合并到客户分析结果中
    full_analysis = pd.merge(subscription_data, customer_analysis, on='委托客户', how='left')
    
    # 填充NaN值
    full_analysis = full_analysis.fillna({'总票数': 0, '总金额': 0, '总利润率': 0, '初步分析': ''})

    # 对full_analysis进行排序
    full_analysis = full_analysis.sort_values(by=['二级部门', '委托客户'])

    # 保存结果到 Excel 文件
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 首先保存海运订阅文件的原始数据
        subscription_df.to_excel(writer, index=False, sheet_name='海运订阅原始数据')
        
        if input_file:
            # 保存预对账原始数据和分析结果
            df.to_excel(writer, index=False, sheet_name='预对账原始数据')
            result_df.to_excel(writer, index=False, sheet_name='分析结果')
        
        # 创建客户公司分析sheet
        workbook = writer.book
        analysis_sheet = workbook.create_sheet(title='客户公司分析')
        
        # 设置列宽
        column_widths = {
            'A': 15,  # 二级部门
            'B': 30,  # 客户公司
            'C': 12,  # 约价负毛利票数
            'D': 10,  # 约价毛利率
            'E': 12,  # 非约价低负票数
            'F': 10,  # 非约价毛利率
            'G': 10,  # 总票数
            'H': 10,  # 总利润率
            'I': 40,  # 初步分析
            'J': 40,  # 业务部门反馈具体原因
            'K': 15,  # 原因类别
            'L': 12,  # 损调利润
            'M': 40,  # 计划采取的措施
            'N': 15,  # 是否完成价格备案表
            'O': 15,  # 是否联合磋商
            'P': 15,  # 督办任务
            'Q': 10,  # 责任人
            'R': 15,  # 督办时间点
        }

        # 设置列宽
        for col, width in column_widths.items():
            analysis_sheet.column_dimensions[col].width = width

        # 添加表头
        headers = [
            ['二级部门', '委托公司', '约价', '约价', '非约价', '非约价', '总票数', '总利润率', '初步分析', '业务部门反馈具体原因', 
             '原因类别', '损调利润', '计划采取的措施', '是否完成价格备案表', '是否联合磋商', '督办任务', 
             '责任人', '督办时间点'],
            ['', '', '负毛利票数', '毛利率', '低负票数', '毛利率', '', '', '', '', '', '', '', '', '', '', '', '']
        ]
        
        # 设置表头样式
        for row_index, row in enumerate(headers, start=1):
            for col_index, value in enumerate(row, start=1):
                cell = analysis_sheet.cell(row=row_index, column=col_index, value=value)
                cell.font = Font(bold=True, size=9)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

        # 合并单元格
        merge_ranges = [
            'A1:A2',  # 二级部门
            'B1:B2',  # 委托公司
            'C1:D1',  # 约价
            'E1:F1',  # 非约价
            'G1:G2',  # 总票数
            'H1:H2',  # 总利润率
            'I1:I2',  # 初步分析
            'J1:J2',  # 业务部门反馈具体原因
            'K1:K2',  # 原因类别
            'L1:L2',  # 损调利润
            'M1:M2',  # 计划采取的措施
            'N1:N2',  # 是否完成价格备案表
            'O1:O2',  # 是否联合磋商
            'P1:P2',  # 督办任务
            'Q1:Q2',  # 责任人
            'R1:R2',  # 督办时间点
        ]

        # 执行单元格合并
        for cell_range in merge_ranges:
            analysis_sheet.merge_cells(cell_range)

        # 设置表头行高
        header_height = min(30, 50)  # 表头行高上限为50
        analysis_sheet.row_dimensions[1].height = header_height
        analysis_sheet.row_dimensions[2].height = header_height

        # 添加客户公司分析数据
        for index, row in full_analysis.iterrows():
            row_num = analysis_sheet.max_row + 1
            
            # 设置单元格值和样式
            cells = [
                (1, row['二级部门'], 'left'),
                (2, row['委托客户'], 'left'),
                (3, row['约价负毛利票数'], 'center'),
                (4, row['约价毛利率'], 'center'),
                (5, row['非约价低负票数'], 'center'),
                (6, row['非约价毛利率'], 'center'),
                (7, row['总票数'], 'center'),
                (8, row['总利润率'], 'center'),
                (9, row['初步分析'], 'left'),
            ]
            
            for col, value, align in cells:
                cell = analysis_sheet.cell(row=row_num, column=col, value=value)
                cell.font = Font(size=9)  # 设置字体大小为9
                cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # 设置百分比格式
                if col in [4, 6, 8]:  # 毛利率列
                    cell.number_format = '0.00%'

            # 添加空白列
            for col in range(10, 19):
                cell = analysis_sheet.cell(row=row_num, column=col, value='')
                cell.font = Font(size=9)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

        # 设置原始数据sheet的列宽
        if input_file:
            for sheet in ['预对账原始数据', '分析结果']:
                worksheet = writer.sheets[sheet]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 40)  # 限制最大宽度为40
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # 设置海运订阅原始数据sheet的列宽
        worksheet = writer.sheets['海运订阅原始数据']
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 40)  # 限制最大宽度为40
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"分析完成，结果已保存到 {output_file}")

def format_analysis(group):
    """
    格式化分析结果，将无应收和倒挂的情况整理成文本描述
    """
    no_receivable = group[group['类型'] == '无应收']['别名'].unique()
    reversed = group[group['类型'] == '倒挂']['别名'].unique()
    
    result = []
    if len(no_receivable) > 0:
        result.append(f"无应收：{', '.join(no_receivable)}")
    if len(reversed) > 0:
        result.append(f"倒挂：{', '.join(reversed)}")
    
    return '\n'.join(result)

if __name__ == "__main__":
    from gui import run_gui
    input_file, output_file, subscription_file = run_gui()
    if subscription_file:
        analyze_excel_data(input_file, output_file, subscription_file)
    else:
        print("未选择海运订阅文件，程序退出")
