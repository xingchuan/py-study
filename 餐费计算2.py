'''
此程序相对于‘餐费计算.py’在实现‘everyone_record’函数时使用了不同的方法
'''
# 导入pandas模块
import pandas as pd

excel_name = '11.xlsx'


def meal_charges(start_time, end_time, xls_name, price):
    # 读取Excel文件
    df = pd.read_excel(excel_name)

    # 删除不需要的列
    df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]], axis=1)

    # 重命名列名
    df.rename(columns={
        '姓名\nName': '姓名',
        '部门\nDepartment': '部门',
        '人员编号\nPerson ID': '工号',
        '识别时间\nRecognition time': '识别时间'
    },
              inplace=True)

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
meal_charges('08:00', '09:10', '早餐费', 4)
meal_charges('16:50', '18:30', '晚餐费', 12)


# 此函数计算‘早餐费+晚餐费=总费用’
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
    df.to_excel('总费用.xlsx', index=False)


count_charges()


# 此代码段给每个人生成一个excel打卡记录表格,分成两个时间段进行筛选，最后在合并。
def everyone_record():
    df = pd.read_excel(excel_name)
    df = df.drop(df.columns[[3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14]], axis=1)
    df.rename(columns={
        '姓名\nName': '姓名',
        '部门\nDepartment': '部门',
        '人员编号\nPerson ID': '工号',
        '识别时间\nRecognition time': '识别时间'
    },
              inplace=True)

    df['识别时间'] = pd.to_datetime(df['识别时间'])

    # 早上数据
    mask = (df['识别时间'].dt.time >= pd.to_datetime('08:00').time()) & (
        df['识别时间'].dt.time <= pd.to_datetime('10:00').time())
    df1 = df.loc[mask]
    # 按照日期（去掉时间）排序，每人每时间段只取一条
    df1 = df1.groupby(['姓名', df1['识别时间'].dt.date]).head(1)

    # 晚上数据
    mask = (df['识别时间'].dt.time >= pd.to_datetime('16:00').time()) & (
        df['识别时间'].dt.time <= pd.to_datetime('19:00').time())
    df2 = df.loc[mask]
    df2 = df2.groupby(['姓名', df2['识别时间'].dt.date]).head(1)

    # 合并两个时间段（早晚餐）的数据
    df = pd.concat([df1, df2])
    grouped = df.groupby('姓名')
    for name, group in grouped:
        # if len(group) == 2: //加上此代码，输出的表格中排序是所有日期的早上打卡，再是晚上打卡
        group.sort_values(by=['识别时间'], inplace=True)
        # group.drop_duplicates(subset=['时间'], keep='first', inplace=True)
        group.to_excel(f'{name}食堂刷脸记录.xlsx', index=False)


everyone_record()
