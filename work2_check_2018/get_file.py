import os
file_num = 0
dir_1 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/昆安（31座）/昆安每座桥记录表汇总'
dir_2 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/楚广（40座）'
dir_3 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/武昆（176座）'
dir_4 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/永武（433座）'

dir_5 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/安楚（306座）'
dir_6 = '/home/jusaka/work3/2018年桥梁记录表（昆西）/楚大（132座）'

file_ass = []
dir_file = dir_1    #切换
dir_file_more = [dir_1, dir_2, dir_3, dir_4]

# for root, dirs, files in os.walk(dir_file):
#     for file in files:
#         (filename, extension) = os.path.splitext(file)          #将文件名拆分为文件名与后缀
#         if (extension == '.xlsx'):                             #判断该后缀是否为.c文件
#             file_ass.append(file)
              
def choose_content(i):
    global file_ass
    global dir_file
    file_ass.clear()
    dir_file = dir_file_more[i]
    for root, dirs, files in os.walk(dir_file):
        for file in files:
            (filename, extension) = os.path.splitext(file)          #将文件名拆分为文件名与后缀
            if (extension == '.xlsx'):                             #判断该后缀是否为.c文件
                file_ass.append(file)
    return file_ass




# print(file_ass)
# for i in range(0, len(file_ass)):
#     if dir_file == dir_1:
#         open_contentfile = '2018年桥梁记录表（昆西）/昆安（31座）/昆安每座桥记录表汇总/'+file_ass[i]
#     elif dir_file == dir_2:
#         open_contentfile = '2018年桥梁记录表（昆西）/楚广（40座）/'+file_ass[i]
#     elif dir_file == dir_3:
#         open_contentfile = '2018年桥梁记录表（昆西）/武昆（176座）/'+file_ass[i]
#     elif dir_file == dir_4:
#         open_contentfile = '2018年桥梁记录表（昆西）/永武（433座）/'+file_ass[i]
   
#     print(i, open_contentfile)

# def test():
#     print("test")















# dir_list = [dir_1, dir_2, dir_3, dir_4]
# item = 0
# for item in range(0, len(dir_list)):
#     for root, dirs, files in os.walk(dir_list[item]):
#         for file in files:
#             (filename, extension) = os.path.splitext(file)          #将文件名拆分为文件名与后缀
#             if (extension == '.xlsx'):                             #判断该后缀是否为.c文件
#                 file_num= file_num+1                      #记录.c文件的个数为对应文件号
#                 # print(file_num, os.path.join(root,filename)) #输出文件号以及对应的路径加文件名
#                 file_ass.append(file)
                # print(file)
                # print("PLACE_RAM(" + filename + ')')
                