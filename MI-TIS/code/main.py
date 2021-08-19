
import function1 as ft
import read
import pandas as pd
import os
import time

folders = []                                                                    # 文件夹
data1 = []                                                                      # MI数据
data2 = []                                                                      # TIS数据
data3 = []                                                                      # powerlever数据
count = 0                                                                       # 记录次数
read_path_0 = 'C:\\Users\\admin\\PycharmProjects\\pythonProject\\处理完的数据\\'  # 读取已处理好的数据文件路径
save_path_0 = 'C:\\Users\\admin\\PycharmProjects\\pythonProject\\拟合数据\\'     # 保存拟合数据文件路径

# 对文件进行排序
def sort_file(file):
    files = []
    lists = []
    for item in file:
        item1 = (item.split('_')[-1])
        item2 = item1.replace('.xlsx', '')
        item3 = item2.replace('dp','')
        lists.append(int(item3))
    lists.sort()
    file_name = ''
    for a in lists:
        b = file[0].split('_')
        b.pop()
        b.append('dp'+str(a) + '.xlsx')
        for c in b:
            file_name = file_name + c + '_'
        list_new_file = list(file_name)
        list_new_file.pop(-1)
        string1 = ''.join(list_new_file)
        file_name = ''
        files.append(string1)
    lists = []
    return files

# 保存文件夹名称
def save_folders(a):
    b = []
    for root, dirs, file in os.walk(a, topdown=False):
        b = dirs
    return b


# 创建文件夹
def creat_folder(path,folder):
    folder_path = path + folder
    a = os.path.exists(folder_path)
    if a:
        a = 1
    else:
        os.mkdir(folder_path)

read.creat_excel_allfile_TIS_MI()                                  # 对需要处理的数据进行提取并保存在处理完的数据的文件夹中

folders1 = save_folders(read_path_0)                               # 需要处理的文件夹

for folder_1 in folders1:        # 循环文件
    read_path_1 = read_path_0 + folder_1 + '\\'                    # 探头文件路径
    save_path_1 = save_path_0 + folder_1 + '\\'                    # 保存探头文件路径
    creat_folder(save_path_0, folder_1)                            # 创建探头文件夹
    folders2 = save_folders(read_path_1)                           # 模式文件路径

    for folder_2 in folders2:

        new_path = read_path_1 + folder_2                          # 文件存在的文件夹路径
        creat_folder(save_path_1, folder_2)                        # 创建模式文件夹
        file = os.listdir(new_path)                                # 提取文件夹中所需的文件
        file = sort_file(file)                                     # 对文件进行排序
        print(file)

        for item in file:  # 循环文件
            file_name = new_path + '\\' + item                     # 文件的路径及名称
            save_file_path = save_path_1 + folder_2 + '\\' + item  # 保存文件的路径
            item1 = item.replace('.xlsx', '')                      # 修改文件后缀
            df = pd.read_excel(file_name, sheet_name=None)         # 读取相关文档

            for sheet, number in df.items():
                data1 = number.loc[:, 'MI'].values                 # 提取文档中MI的数据
                data2 = number.loc[:, 'TIS'].values                # 提取文档中TIS的数据
                data3 = number.loc[:, 'powerlevel'].values         # 提取文档中powerlevel的数据

                data_MI = ft.creat_image(data3, data1, 'MI')       # 保存数据
                data_TIS = ft.creat_image(data3, data2, 'TIS')     # 保存数据

                ft.creat_excel(data_MI, data_TIS, save_file_path , sheet, count)           # 创建excel表
                count += 1                                                                 # 下一个sheet表
            count = 0                                                                      # sheet表数清零