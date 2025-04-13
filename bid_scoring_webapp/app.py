import os
from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import numpy as np
import re
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制

def custom_secure_filename(filename):
    # 保留中文、字母、数字、下划线、点和短横线
    filename = re.sub(r'[^\w\u4e00-\u9fff\-\.]', '', filename)
    filename = filename.strip().replace(' ', '_')
    return filename

def calculate_bid_scores(input_file, output_excel, W1, W2, n1, n2, C):
    """
    从Excel读取多个sheet的投标数据，处理非数字价格，按分包名称分组计算得分
    并将结果保存到Excel的不同sheet中
    
    参数:
        input_file: 输入Excel文件路径
        output_excel: 输出Excel文件路径
        W1, W2: 区间参数
        n1, n2: 得分计算参数
        C: 基准价浮动系数
    """
    # 1. 读取Excel文件中的所有sheet
    try:
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        print(f"Excel文件中包含的sheet: {sheet_names}")
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return False
    
    # 确保至少有一个sheet
    if not sheet_names:
        print("Excel文件中没有包含任何sheet")
        return False
    
    # 创建一个Excel writer对象，用于写入多个sheet
    writer = pd.ExcelWriter(output_excel, engine='openpyxl')
    
    processed_sheets = 0
    
    for sheet_name in sheet_names:
        try:
            print(f"\n正在处理sheet: {sheet_name}")
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            print(f"成功读取sheet '{sheet_name}'，原始数据行数:", len(df))
        except Exception as e:
            print(f"读取sheet '{sheet_name}'失败: {e}")
            continue
        
        # 检查必要列是否存在
        required_columns = ['分包名称', '投标人名称', '投标价格']
        if not all(col in df.columns for col in required_columns):
            print(f"警告: sheet '{sheet_name}'中必须包含{required_columns}列，跳过此sheet")
            continue
        
        # 处理非数字投标价格
        df['投标价格'] = pd.to_numeric(df['投标价格'], errors='coerce')
        df = df[df['投标价格'].notna()]
        print(f"有效数字数据行数:", len(df))
        
        if len(df) == 0:
            print(f"警告: sheet '{sheet_name}'中没有有效数据，跳过此sheet")
            continue
        
        # 定义计算函数
        def process_group(group):
            if len(group) == 0:
                return pd.DataFrame()
                
            # 按投标价格排序
            sorted_group = group.sort_values('投标价格')
            bids = sorted_group['投标价格'].values
            
            # 确定要去掉的最低和最高数量
            remove_low = 3  # 去掉3个最低
            remove_high = 4  # 去掉4个最高
            
            # 标记被剔除的投标
            sorted_group['是否被剔除'] = False
            if len(sorted_group) > remove_low:
                sorted_group.iloc[:remove_low, sorted_group.columns.get_loc('是否被剔除')] = True
            if len(sorted_group) > remove_high:
                sorted_group.iloc[-remove_high:, sorted_group.columns.get_loc('是否被剔除')] = True
            
            # 获取有效投标(未被剔除的)
            valid_bids = sorted_group[~sorted_group['是否被剔除']]
            
            # 计算A1(有效投标的平均值)
            A1 = valid_bids['投标价格'].mean() if not valid_bids.empty else 0
            
            # 计算有效投标区间
            lower_bound = A1 * (1 + W1)
            upper_bound = A1 * (1 + W2)
            
            # 标记有效投标人(在区间内的)
            is_valid = (valid_bids['投标价格'] > lower_bound) & (valid_bids['投标价格'] < upper_bound)
            valid_bids['有效投标人标记'] = np.where(is_valid, "有效", "无效")
            
            # 计算A2(有效投标区间内的平均值)
            valid_in_range = valid_bids[is_valid]
            A2 = valid_in_range['投标价格'].mean() if not valid_in_range.empty else A1
            
            # 计算基准价
            benchmark = A2 * (1 - C)
            
            # 计算偏离度
            valid_bids['偏离度'] = (valid_bids['投标价格'] - benchmark) / benchmark
            
            # 计算得分(只对有效投标计算)
            def calculate_score(row):
                bid = row['投标价格']
                if bid >= benchmark:
                    return 100 - 100 * n1 * abs(bid - benchmark) / benchmark
                else:
                    return 100 - 100 * n2 * abs(bid - benchmark) / benchmark
            
            valid_bids['价格得分'] = valid_bids.apply(calculate_score, axis=1)
            
            # 合并结果(有效投标在前，被剔除的在后)
            valid_bids['A1'] = A1
            valid_bids['A2'] = A2
            valid_bids['基准价'] = benchmark
            
            # 被剔除的投标不计算得分
            eliminated = sorted_group[sorted_group['是否被剔除']].copy()
            eliminated['有效投标人标记'] = "被剔除"
            eliminated['A1'] = A1
            eliminated['A2'] = A2
            eliminated['基准价'] = benchmark
            eliminated['价格得分'] = np.nan
            
            combined = pd.concat([valid_bids, eliminated])
            
            # 计算排名(只对有效投标排名)
            combined['排名'] = np.nan
            if not valid_bids.empty:
                valid_bids_sorted = valid_bids.sort_values('价格得分', ascending=False)
                valid_bids_sorted['原始排名'] = valid_bids_sorted['价格得分'].rank(method='first', ascending=False)
                valid_bids_sorted['排名组'] = valid_bids_sorted['价格得分'].rank(method='dense', ascending=False)
                rank_adjustment = valid_bids_sorted.groupby('排名组')['原始排名'].min()
                valid_bids_sorted['排名'] = valid_bids_sorted['排名组'].map(rank_adjustment)
                combined.update(valid_bids_sorted[['排名']])
            
            return combined
        
        # 按分包名称分组处理
        result_dfs = []
        for bid_no, group in df.groupby('分包名称'):
            print(f"处理分包名称: {bid_no} (数据量: {len(group)})")
            result_group = process_group(group)
            if not result_group.empty:
                result_dfs.append(result_group)
        
        if not result_dfs:
            print("\n警告: 没有有效数据可供计算，跳过此sheet")
            continue
        
        # 合并所有分组结果
        result_df = pd.concat(result_dfs)
        
        # 选择需要的列并格式化
        final_columns = ['分包名称', '投标人名称', '投标价格', '偏离度','价格得分', '排名']
        result_df = result_df[final_columns]
        
        # 格式化数值列
        float_cols = ['投标价格', '价格得分']
        result_df[float_cols] = result_df[float_cols].round(3)
        result_df['偏离度'] = result_df['偏离度'].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "")
        result_df['排名'] = result_df['排名'].astype('Int64')
        
        # 创建自然排序键
        def make_sort_key(col):
            if col.name == '分包名称':
                return col.map(lambda x: tuple(
                    int(text) if text.isdigit() else text.lower()
                    for text in re.split('([0-9]+)', str(x))
                    if text
                ))
            return col

        # 按分包名称和排名排序
        result_df = result_df.sort_values(
            by=['分包名称', '投标价格'],
            key=lambda x: make_sort_key(x),
            ascending=[True, False])  
        
        # 在不同标包之间插入空行
        def insert_empty_rows(df, group_column):
            grouped = df.groupby(group_column, sort=False).groups
            empty_row = pd.Series({col: None for col in df.columns}, name=None)
            result = pd.DataFrame(columns=df.columns)
            for group_name, indices in grouped.items():
                group_df = df.loc[indices]
                result = pd.concat([result, group_df, empty_row.to_frame().T], ignore_index=True)
            return result.iloc[:-1]
        
        result_df = insert_empty_rows(result_df, '分包名称')
        
        # 将结果保存到当前sheet对应的输出sheet中
        output_sheet_name = f"{sheet_name}_得分结果"
        result_df.to_excel(writer, sheet_name=output_sheet_name, index=False)
        print(f"结果已保存到sheet: {output_sheet_name}")
        processed_sheets += 1
    
    # 如果没有处理任何sheet，返回False
    if processed_sheets == 0:
        print("没有处理任何sheet，可能所有sheet都不符合要求")
        return False
    
    # 保存Excel文件
    try:
        writer.close()
        print(f"\n所有结果已保存到 {output_excel}")
        return True
    except Exception as e:
        print(f"保存Excel文件失败: {e}")
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        
        if file and file.filename.lower().endswith(('.xlsx', '.xls')):
            # 使用自定义的 secure_filename 保留中文
            filename = custom_secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            try:
                W1 = float(request.form.get('W1', -0.2))
                W2 = float(request.form.get('W2', 0.05))
                n1 = float(request.form.get('n1', 1.8))
                n2 = float(request.form.get('n2', 0.0))
                C = float(request.form.get('C', 0.0025))
            except ValueError:
                return "参数必须是数字", 400
            
            # 分离文件名和扩展名，确保输出文件是 .xlsx
            name, ext = os.path.splitext(filename)
            output_filename = f"{name}_得分结果.xlsx"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            if calculate_bid_scores(input_path, output_path, W1, W2, n1, n2, C):
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                return "处理文件时出错，请检查文件格式和内容是否符合要求", 400
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
    # app.run(host='0.0.0.0', port=5000, debug=True)