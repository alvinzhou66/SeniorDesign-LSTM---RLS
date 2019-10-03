import numpy as np
import csv
import xlrd
import json
import pandas as pd
'''
问题1 correlation的计算程序实现，确定输入与输出（完成）
问题2 热点图的实现(完成)
问题3 混淆矩阵（完成）
'''


# 读取相应的CSV文件，将数据存储到字典中并返回
def read_csv(url):
    dic = {}
    with open(url, 'r', encoding='utf-8-sig') as csvFile:
        reader = csv.reader(csvFile)
        # cvs文件第一个单元格格式会有前缀,需要去除BOM
        for row in reader:
            dic[row[0]] = row[1:]
    return dic


# 对列表进行平均长度分割
def divide_chunks(l, n):
    for i in range(0, len(l), n):
        yield l[i:i + n]


def cal_correlation(arg1, arg2):
    ans = np.corrcoef(np.array(arg1), np.array(arg2))
    return ans[0][1]


def correlation_matrix(url):
    dic = read_csv(url)
    # 获得字典后使用循环遍历输出算出的correlation
    temp = []
    for key in dic:
        for sec_key in dic:
            arg_1 = map(float, dic[key])
            arg_2 = map(float, dic[sec_key])
            temp.append(cal_correlation(list(arg_1), list(arg_2)))
    corr_matrix = list(divide_chunks(temp, len(dic)))
    return corr_matrix


def data_write_csv(file_name, data_list):
    # file_name为写入CSV文件的路径，data_list为要写入的数据列表
    file_csv = open(file_name, 'w')

    for data in data_list:
        for x in data:
            file_csv.write(str(x))
            file_csv.write(',')
        file_csv.write('\n')

    file_csv.close()
    print('handle success')


def main():
    url = '/Users/wzc/Desktop/air_1.csv'
    data = correlation_matrix(url)
    filename = '/Users/wzc/Desktop/correlation.csv'
    data_write_csv(filename, data)


if __name__ == '__main__':
    main()
