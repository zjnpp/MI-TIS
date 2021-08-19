import read_function as re                                                            # 读取需要的函数库

def creat_excel_allfile_TIS_MI():
	count = 0                                                                         # 记录创建ROI或RES的sheet表次数
	read_path_0 = 'C:\\Users\\admin\\PycharmProjects\\pythonProject\\需要处理的数据\\'   # 提取需要处理的TISMI数据文件路径
	save_path_0 = 'C:\\Users\\admin\\PycharmProjects\\pythonProject\\处理完的数据\\'    # 保存处理好的TISMI数据文件路径

	folders1 = re.save_folders(read_path_0)                                           # 提取探头文件夹
	for folders_1 in folders1:

		read_path_1 = read_path_0 + folders_1 + '\\'                                  # 增加探头文件夹路径
		folders2 = re.save_folders(read_path_1)                                       # 提取探头文件夹中的模式文件夹
		re.creat_folder(save_path_0, folders_1)                                       # 创建保存相应探头的文件夹
		save_path_1 = save_path_0 + folders_1 + '\\'                                  # 增加相应探头文件夹路径

		for folders_2 in folders2:

			read_path_2 = read_path_1 + folders_2 + '\\'                              # 增加模式文件夹路径
			folders3 = re.save_folders(read_path_2)                                   # 提取模式文件夹中的深度文件夹
			re.creat_folder(save_path_1, folders_2)                                   # 创建保存相应模式的文件夹
			save_path_2 = save_path_1 + folders_2 + '\\'                              # 增加相应模式文件夹路径

			for folders_3 in folders3:

				read_path_3 = read_path_2 + folders_3 + '\\'                          # 增加深度文件夹路径
				folders4 = re.save_folders(read_path_3)                               # 提取深度文件夹中的ROI或者RES文件夹
				folders4 = re.sort(folders4,folders_2)                                # 对提取的ROI或者RES文件夹进行排序

				for folders_4 in folders4:
					data_MI = []                                                      # 保存MI数据
					data_TIS = []                                                     # 保存TIS数据
					data_powerlevel = []                                              # 保存powerlevel
					new_read_path = read_path_3 + folders_4 + '\\'                    # 增加ROI或者RES文件夹路径
					file = re.save_file(new_read_path)                                # 提取相应ROI或者RES文件夹下的全部文件
					file = re.sort_file(file)                                         # 对文件进行排序

					for item in file:
						file_path = new_read_path + 'Z' + '\\' + item                 # 文件路径
						item2 = item.replace('%.xls', '')                             # 删除文件后缀%.xls
						number = int(item2.split('_')[-1])                            # 提取powerlevel
						a = re.extract_MI_TIS(file_path)                              # 提取文件中MI和TIS数据
						data_MI.append(a[0])                                          # 增加MI
						data_TIS.append(a[1])                                         # 增加TIS
						data_powerlevel.append(number)                                # 增加powerlevel
					string1 = file[0][::-1]                                           # 对文件名进行颠倒
					string2 = string1.split('_', 3)[3]                                # 提取相应的字符串
					string3 = string2[::-1]                                           # 对相应字符串进行颠倒操作
					new_save_path = save_path_2 + string3 + '.xlsx'                   # 文件保存路径
					re.creat_excel(data_MI, data_TIS, data_powerlevel, count, folders_4, new_save_path)  # 创建excel表
					count += 1                                                        # 创建ROI或RESsheet表的次数
				count = 0                                                             # 次数清零
	return 1
