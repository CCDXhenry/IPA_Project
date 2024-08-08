import pandas as pd

# 定义一个函数来检查每个号码的记录是否存在异常维度1
def check_abnormality(group):
    group = group.sort_values(by='业务受理时间')
    prev_institution = None
    prev_time = None

    for i, row in group.iterrows():
        if prev_institution is not None and prev_institution != row['组织机构名称']:
            if isinstance(row['业务受理时间'], pd.Timestamp) and isinstance(prev_time, pd.Timestamp):
                time_diff = (row['业务受理时间'] - prev_time).total_seconds() / 60
                if time_diff <= 60:  # 检查是否在60分钟内
                    if i < len(group):  # 确保i不会超出group的长度
                        # 标记当前行
                        combined_df.at[group.index[i], '是否异常'] = '是'
                        combined_df.at[group.index[i], '异常维度'] = '异常维度1'
                        if i - 1 >= 0:  # 确保不会越界
                            # 标记前一行
                            combined_df.at[group.index[i - 1], '是否异常'] = '是'
                            combined_df.at[group.index[i - 1], '异常维度'] = '异常维度1'
        prev_institution = row['组织机构名称']
        prev_time = row['业务受理时间']

source_path = f'combined_file.xlsx'
combined_df = pd.read_excel(source_path)

# 添加新列初始化为False
combined_df['是否异常'] = '否'
combined_df['异常维度'] = None
# 将业务受理时间列转换为datetime类型
combined_df['业务受理时间'] = pd.to_datetime(combined_df['业务受理时间'], errors='coerce')
# 应用函数检查每个电话号码的记录
combined_df.groupby('客户手机号码').apply(check_abnormality)

# 保存处理后的数据
combined_df.to_excel('processed_data.xlsx', index=False)