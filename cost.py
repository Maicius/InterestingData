#coding=utf-8

import xlrd

def getData():
    data = xlrd.open_workbook('2017年部属高校国家级大学生创新创业训练计划项目名单.xlsx')
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    project_num_dic = {}
    people_num_dic = {}
    avg_num_dic = {}
    for i in range(3, nrows):
        row = table.row_values(i)
        if row[1] in project_num_dic:
            project_num_dic[row[1]] += 1
            people_num_dic[row[1]] += int(row[7])
        else:
            # print(row[7])
            project_num_dic[row[1]] = 1
            people_num_dic[row[1]] = int(row[7])
        # print(table.row_values(i))
        # print(table.col_values(i))
    for key in project_num_dic.keys():
        avg_num_dic[key] = people_num_dic[key] / project_num_dic[key]

    print("项目数量：" + str(sorted(project_num_dic.items(), key= lambda item: item[1], reverse=True)))
    print("参与人数" + str(sorted(people_num_dic.items(), key= lambda item: item[1], reverse=True)))
    print("平均人数" + str(sorted(avg_num_dic.items(), key=lambda item: item[1], reverse=False)))

if __name__ == '__main__':
    getData()