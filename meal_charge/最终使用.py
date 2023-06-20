'''
此程序相对于‘餐费计算.py’在实现‘everyone_record’函数时使用了不同的方法
'''
# 导入pandas模块
import pandas as pd

excel_name = 'file.xlsx'


# 此代码段对所有符合条件（吃饭时间内的刷脸记录）的人员进行筛选，生成一个表格。
def all_record():
    df = pd.read_excel(excel_name, header=None)
    df.drop(0, inplace=True)
    df.rename(columns=df.iloc[0], inplace=True)
    df.drop(1, inplace=True)
    df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]], axis=1)
    df.rename(columns={
        '姓名\nName': '姓名',
        '部门\nDepartment': '部门',
        '人员编号\nPerson ID': '工号',
        '识别时间\nRecognition time': '识别时间'
    },
              inplace=True)

    df['部门'] = df['部门'].str.replace('-', '')

    df['识别时间'] = pd.to_datetime(df['识别时间'])

    # 早上数据
    mask = (df['识别时间'].dt.time >= pd.to_datetime('08:00').time()) & (
        df['识别时间'].dt.time <= pd.to_datetime('13:00').time())
    df1 = df.loc[mask]
    # 按照日期（去掉时间）排序，每人每时间段只取一条
    df1 = df1.groupby(['姓名', df1['识别时间'].dt.date]).head(1)

    # 晚上数据
    mask = (df['识别时间'].dt.time >= pd.to_datetime('13:01').time()) & (
        df['识别时间'].dt.time <= pd.to_datetime('19:00').time())
    df2 = df.loc[mask]
    df2 = df2.groupby(['姓名', df2['识别时间'].dt.date]).head(1)

    # 合并两个时间段（早晚餐）的数据
    df = pd.concat([df1, df2])
    
    sorted_df = df.sort_values(by=['姓名', '识别时间'])

    with pd.ExcelWriter('本月餐费记录.xlsx') as writer:
        sorted_df.to_excel(writer, sheet_name='全部记录', index=False)


all_record()


# 此函数会分别计算每个人的早、晚餐费用
def meal_charges(start_time, end_time, xls_name, price):
    # 读取Excel文件,删除第一行，将第二行作为列名
    df = pd.read_excel(excel_name, header=None)
    df.drop(0, inplace=True)
    df.rename(columns=df.iloc[0], inplace=True)
    df.drop(1, inplace=True)

    # 删除不需要的列
    df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]], axis=1)

    # 重命名列名
    df.rename(columns={
        '姓名\nName': '姓名',
        '部门\nDepartment': '部门',
        '人员编号\nPerson ID': '工号',
        '识别时间\nRecognition time': '识别时间'
    }, inplace=True)

    df['部门'] = df['部门'].str.replace('-', '')
    # 转换时间格式
    # df = df.drop(df[df['姓名'] == '人员未注册'].index, inplace=True)
    df['识别时间'] = pd.to_datetime(df['识别时间'])

    # 采集起始、结束时间段
    start_time = start_time
    end_time = end_time

    # 获取start_time到end_time的行,此处可以筛选到分钟数，
    mask = (df['识别时间'].dt.time >= pd.to_datetime(start_time).time()) & (
        df['识别时间'].dt.time <= pd.to_datetime(end_time).time())
    df = df.loc[mask]

    # 按照姓名和时间进行分组，并删除重复行
    df = df.groupby(['姓名', '识别时间']).apply(lambda x: x.drop_duplicates())

    # 将‘时间’转换回字符串格式
    df['识别时间'] = df['识别时间'].dt.strftime('%Y-%m-%d')

    # 删除姓名和时间相同的行
    df.drop_duplicates(subset=['姓名', '识别时间'], keep='first', inplace=True)

    # 写入新的excel表格中，不含索引
    df.to_excel(xls_name + '.xlsx', index=False)

    # 写入新的excel表格中，含索引
    df.to_excel(xls_name + '1.xlsx', index=True)

    # 以下代码块实现对上面生成的excel表格，按照姓名去重，并计算餐费
    df = pd.read_excel(xls_name + '.xlsx')

    # 计算同一姓名的行数和，即为吃了几顿早餐或晚餐
    counts = df['姓名'].value_counts()

    # 将计算的结果添加到新的列中
    df[xls_name] = counts[df['姓名']].values

    # 列数*每餐费用=本月总费用
    df[xls_name] *= price

    # 每个人只保留第一行
    df.drop_duplicates(subset='姓名', keep='first', inplace=True)

    # 保存到新的表格中，此费用用于计算总费用
    df.to_excel(xls_name + '2.xlsx', index=False)


# 4个参数分别是：开始时间-结束时间-表格名称-每餐费用
meal_charges('08:00', '13:00', '早餐费', 4)
meal_charges('13:01', '18:30', '晚餐费', 12)


# 此函数计算‘总费用=早餐费+晚餐费’
def count_charges():
    df1 = pd.read_excel('早餐费2.xlsx')
    df2 = pd.read_excel('晚餐费2.xlsx')

    df1 = df1.drop(df1.columns[[3]], axis=1)
    df2 = df2.drop(df2.columns[[3]], axis=1)

    # 合并两个表格中，姓名、部门、工号相同的行，此处的how='outer'类似于数据库的外连接，确保df1和df2的并集
    df = pd.merge(df1, df2, on=['姓名', '部门', '工号'], how='outer')
    df.fillna(0, inplace=True)

    # 计算总费用，并保存到新的列中
    df['总费用'] = df['早餐费'] + df['晚餐费']
    # df['总费用'] = df.groupby(['姓名'])['费用'].transform('sum')

    # 删除姓名重复的行，只保留第一行
    df.drop_duplicates(subset='姓名', keep='first', inplace=True)

    # 按部门进行分表（部门列中如果出现异常，此处会失败）
    grouped = df.groupby('部门')
    
    with pd.ExcelWriter('本月餐费记录.xlsx', mode='a') as writer:
        for name, group in grouped:
            group.to_excel(writer, sheet_name=name, index=False)


count_charges()




# 此函数实现给每人生成一个只含费用的excel表格
# def everyone_charge():
#     df = pd.read_excel('总费用.xlsx')

#     # 将每个人的行保存到新的excel文件中
#     def save_to_excel(group):
#         group.to_excel(f"{group['姓名'].iloc[0]}.xlsx", index=False)

#     df.groupby('姓名', group_keys=False).apply(save_to_excel)

# everyone_charge()