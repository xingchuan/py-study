# 导入pandas模块
import pandas as pd


def meal_charges(start_time, end_time, xls_name, price):
    # 读取Excel文件
    df = pd.read_excel('1.xls')

    # 删除不需要的列，此处删除第2列
    df = df.drop(df.columns[[1]], axis=1)

    # 转换时间格式
    df['时间'] = pd.to_datetime(df['时间'])

    start_time = start_time
    end_time = end_time

    # 获取start_time到end_time的行
    mask = (df['时间'].dt.time >= pd.to_datetime(start_time).time()) & (
        df['时间'].dt.time <= pd.to_datetime(end_time).time())
    df = df.loc[mask]

    # 按照姓名和时间进行分组，并删除重复行
    df = df.groupby(['姓名', '时间']).apply(lambda x: x.drop_duplicates())

    # 将‘时间’转换回字符串格式
    df['时间'] = df['时间'].dt.strftime('%Y-%m-%d')

    # 删除姓名和时间相同的行
    df.drop_duplicates(subset=['姓名', '时间'], keep='first', inplace=True)

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

    # 删除第2列
    df = df.drop(df.columns[[2]], axis=1)

    # 保存到新的表格中
    df.to_excel(xls_name + '2.xlsx', index=False)


# 4个参数分别是：开始时间-结束时间-表格名称-每餐费用
meal_charges('06:00', '10:00', '早餐费', 4)
meal_charges('16:30', '19:00', '晚餐费', 12)


# 此函数计算‘早餐费+晚餐费=总费用’
def count_charges():
    df1 = pd.read_excel('早餐费2.xlsx')
    df2 = pd.read_excel('晚餐费2.xlsx')

    # 合并两个表格中，姓名和工号相同的行，此处请注意：公司人脸系统工号并非完全正确
    df = pd.merge(df1, df2, on=['姓名', '工号'])

    # 计算总费用，并保存到新的列中
    df['总费用'] = df['早餐费'] + df['晚餐费']
    # df['总费用'] = df.groupby(['姓名'])['费用'].transform('sum')

    # 删除姓名重复的行，只保留第一行
    df.drop_duplicates(subset='姓名', keep='first', inplace=True)
    df.to_excel('总费用.xlsx', index=False)


count_charges()


# 此代码段给每个人生成一个excel打卡记录表格
def everyone_record():
    df = pd.read_excel('1.xls')
    df['时间'] = pd.to_datetime(df['时间'])

    # 此处无法对分钟和秒数进行筛选,如果加上"& (df['时间'].dt.minute <= 10)",会出错
    df1 = df.loc[(df['时间'].dt.hour >= 5) & (df['时间'].dt.hour <= 10)]
    df1 = df1.groupby(['姓名', df1['时间'].dt.date]).head(1)
    df2 = df.loc[(df['时间'].dt.hour >= 16) & (df['时间'].dt.hour <= 19)]
    df2 = df2.groupby(['姓名', df2['时间'].dt.date]).head(1)
    df = pd.concat([df1, df2])
    grouped = df.groupby('姓名')
    for name, group in grouped:
        # if len(group) == 2: //加上此代码，排序是所有日期的早上打卡，再是晚上打卡
        group.sort_values(by=['时间'], inplace=True)
        # group.drop_duplicates(subset=['时间'], keep='first', inplace=True)
        group.to_excel(f'{name}.xlsx', index=False)


everyone_record()


# 此函数实现给每人生成一个只含费用的excel表格
# def everyone_charge():
#     df = pd.read_excel('总费用.xlsx')

#     # 将每个人的行保存到新的excel文件中
#     def save_to_excel(group):
#         group.to_excel(f"{group['姓名'].iloc[0]}.xlsx", index=False)

#     df.groupby('姓名', group_keys=False).apply(save_to_excel)

# everyone_charge()
