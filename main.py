#coding=utf-8

import xlrd
import jieba
from wordcloud import WordCloud, ImageColorGenerator, STOPWORDS
from scipy.misc import imread
import matplotlib.pyplot as plt

def getData():
    # university ="清华大学	北京大学	厦门大学 南京大学	复旦大学	天津大学 浙江大学	南开大学	西安交通大学 东南大学	武汉大学	上海交通大学 山东大学	湖南大学	中国人民大学 吉林大学	重庆大学	电子科技大学 四川大学	中山大学	华南理工大学 兰州大学	东北大学	西北工业大学 哈尔滨工业大学	华中科技大学	中国海洋大学 北京理工大学	大连理工大学	北京航空航天大学 北京师范大学	同济大学	中南大学 中国科学技术大学 中国农业大学	国防科学技术大学	中央民族大学 华东师范大学	西北农林科技大学 北京邮电大学 上海财经大学 西南财经大学 中国政法大学 中央财经大学"
    university = ""
    data = xlrd.open_workbook('2015年部属高校国家级大学生创新创业训练计划项目名单.xls')
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    project_num_dic = {}
    people_num_dic = {}
    avg_num_list = []
    project_num_list = []
    people_num_list = []
    university_list = []
    project_names_str = ""
    cost_total = 0
    ministry_cost = 0
    university_cost = 0
    university_cost_list = []
    university_cost_dic = {}

    for i in range(3, nrows):
        row = table.row_values(i)
        # project_names_str += row[4]
        # print(row[12])
        cost_total += int(row[11])
        ministry_cost += int(row[12])
        university_cost += int(row[13])
        if university.find(row[1]) == -1:
            if row[1] in university_cost_dic:
                university_cost_dic[row[1]] += int(row[11])
            else:
                university_cost_dic[row[1]] = int(row[11])
        # if university.find(row[1]) != -1:
        #     if row[1] in project_num_dic:
        #         project_num_dic[row[1]] += 1
        #         people_num_dic[row[1]] += int(row[7])
        #     else:
        #         print(row[7])
        #         project_num_dic[row[1]] = 1
        #         people_num_dic[row[1]] = int(row[7])
        # print(table.row_values(i))
        # print(table.col_values(i))
    data2 = xlrd.open_workbook('2016年部属高校国家级大学生创新创业训练计划项目名单.xlsx')
    table = data2.sheets()[1]
    nrows = table.nrows
    for i in range(2, nrows):
        row = table.row_values(i)
        # project_names_str += row[4]
        print(row[13])
        cost_total += int(row[15]) if row[15] != "" else 0
        ministry_cost += int(row[14]) if row[14] != "" else 0
        university_cost += int(row[13]) if row[13] != "" else 0
        if university.find(row[3]) == -1:
            if row[3] in university_cost_dic:
                university_cost_dic[row[3]] += int(row[15])
            else:
                university_cost_dic[row[3]] = int(row[15])

    for item in university_cost_dic.items():
        university_cost_list.append(list(item))
    university_cost_list.sort(key=lambda x:x[1], reverse=True)
    university_cost_list = university_cost_list[0:40]
    university_cost_list.sort(key=lambda x: x[1], reverse=False)
    print(university_cost_list)
    # for key in project_num_dic.keys():
    #     temp = round(people_num_dic[key] / project_num_dic[key], 2)
    #     avg_num_list.append([key, temp])
    #
    # for item in project_num_dic.items():
    #     project_num_list.append(list(item))
    # for item in people_num_dic.items():
    #     people_num_list.append(list(item))
    #
    # for i in range (len(project_num_list)):
    #     university_list.append([project_num_list[i][0], project_num_list[i][1], people_num_list[i][1], avg_num_list[i][1]])
    #
    print([cost_total, ministry_cost, university_cost])
    # word_text = splitWords(project_names_str)
    # drawWordCloud(word_text, 'project_name.png')
    # university_list.sort(key=lambda x:x[3], reverse=True)
    # print(university_list)
    # # print(project_num_list)
    # print(project_num_list)
    # print(people_num_list)
    # print(avg_num_list)
    # print(sorted(project_num_dic.items(), key= lambda item: item[1], reverse=True))
    # print(sorted(people_num_dic.items(), key= lambda item: item[1], reverse=True))
    # print(sorted(avg_num_dic.items(), key=lambda item: item[1], reverse=False))

def getData2016():
    university = "清华大学	北京大学	厦门大学 南京大学	复旦大学	天津大学 浙江大学	南开大学	西安交通大学 东南大学	武汉大学	上海交通大学 山东大学	湖南大学	中国人民大学 吉林大学	重庆大学	电子科技大学 四川大学	中山大学	华南理工大学 兰州大学	东北大学	西北工业大学 哈尔滨工业大学	华中科技大学	中国海洋大学 北京理工大学	大连理工大学	北京航空航天大学 北京师范大学	同济大学	中南大学 中国科学技术大学 中国农业大学	国防科学技术大学	中央民族大学 华东师范大学	西北农林科技大学 北京邮电大学 上海财经大学 西南财经大学 中国政法大学 中央财经大学"
    data = xlrd.open_workbook('2016年部属高校国家级大学生创新创业训练计划项目名单.xlsx')
    table = data.sheets()[1]
    nrows = table.nrows
    ncols = table.ncols
    project_num_dic = {}
    people_num_dic = {}
    avg_num_list = []
    project_num_list = []
    people_num_list = []
    university_list = []
    project_names_str = ""
    for i in range(3, nrows):
        row = table.row_values(i)
        print(row)
        print("row5" + str(row[5]))
        project_names_str += row[5]
        if university.find(row[3]) != -1:
            if row[1] in project_num_dic:
                project_num_dic[row[3]] += 1
                people_num_dic[row[3]] += int(row[9])
            else:
                print(row[9])
                project_num_dic[row[3]] = 1
                people_num_dic[row[3]] = int(row[9])
                # print(table.row_values(i))
                # print(table.col_values(i))
    for key in project_num_dic.keys():
        temp = round(people_num_dic[key] / project_num_dic[key], 2)
        avg_num_list.append([key, temp])

    for item in project_num_dic.items():
        project_num_list.append(list(item))
    for item in people_num_dic.items():
        people_num_list.append(list(item))

    for i in range(len(project_num_list)):
        university_list.append(
            [project_num_list[i][0], project_num_list[i][1], people_num_list[i][1], avg_num_list[i][1]])

    word_text = splitWords(project_names_str)
    drawWordCloud(word_text, 'project_name.png')

def splitWords(word):
    print('begin')
    word_list = jieba.cut(word, cut_all=False)
    word_list2 = []
    waste_word_str = "基于 研究 技术 方法 理论 为例 实验 影响 模拟 作用 应用 工程 探究 探索 浅析 机制 坚定 分析 调研 构建 特征 设计 方案 新型 及其 系统 公司 组织 平台 用于 对策 不同 使用 调查 合作 辅助 我国 地区"
    for word in word_list:
        # print(word)
        if len(word) >= 2 and waste_word_str.find(word) == -1:
            print(word)
            word_list2.append(word)
    # print(word_list2)
    word_text = " ".join(word_list2)
    # print(word_text)
    return word_text

def drawWordCloud(word_text, filename):
    mask = imread('pic.png')
    my_wordcloud = WordCloud(
        background_color='white',  # 设置背景颜色
        mask=mask,  # 设置背景图片
        max_words=2000,  # 设置最大现实的字数
        stopwords=STOPWORDS,  # 设置停用词
        font_path='/System/Library/Fonts/Hiragino Sans GB W6.ttc',  # 设置字体格式，如不设置显示不了中文
        max_font_size=50,  # 设置字体最大值
        random_state=30,  # 设置有多少种随机生成状态，即有多少种配色方案
        scale=1.5
    ).generate(word_text)
    image_colors = ImageColorGenerator(mask)
    my_wordcloud.recolor(color_func=image_colors)
    # 以下代码显示图片
    plt.imshow(my_wordcloud)
    plt.axis("off")
    plt.show()
    # 保存图片
    my_wordcloud.to_file(filename=filename)
    print("finish")


if __name__ == '__main__':
    getData()

