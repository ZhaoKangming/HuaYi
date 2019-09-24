'''
Function：帮家宜将易铭天的打款凭证PDF，每页上下等分为两张图片并以医生姓名命名图片
Author: ZhaoKangming
E-mail: zhaokm0@gmail.com
Version: V0.1
'''

from PIL import Image
import sys
import os


def get_name(txt_path: str) ->list:
    namelist = []
    with open(txt_path,'r') as txt:
        for ln in txt:
            ln_part = ln.replace("\n", "").replace(" ","")
            if ln_part[0:3] == '付款人':
                name_start_pst: int = ln_part.find('收款人户名')
                namelist.append(ln_part[name_start_pst + 5 :])
    return namelist


def cut_image(image_path: str):
    image = Image.open(image_path)
    width, height = image.size
    item_height = int(height / 2)
    box_list = []
    # (left, upper, right, lower)
    for i in range(0,2): 
        box = (0, i*item_height, width, (i+1)*item_height)
        box_list.append(box)
    
    image_list = [image.crop(box) for box in box_list]
    return image_list



txt_path: str = r'D:\tools\PJY_Cut_Pdf2Pic\source.txt'   # pdf 导出的 txt 文件路径
src_folder: str = r'D:\tools\PJY_Cut_Pdf2Pic\img'        # pdf 导出的 jpg 文件路径
dst_folder_path: str = r'D:\tools\PJY_Cut_Pdf2Pic\医生打款凭证'       # 输出
os.makedirs(dst_folder_path)

# 获取并打印名字列表
namelist: list = get_name(txt_path)
print('-'*20 + '【付款名单：' + str(len(namelist)) + '人】' + '-'*20)
print(namelist)

# 遍历图片文件夹
for file in os.listdir(src_folder):
    src_pic: str = os.path.join(src_folder,file)
    if os.path.isfile(src_pic) == True:
        # 获取原始图片的序号
        src_pic_seq: int = int(file.split('.jpg')[0].split('页面')[1].replace('_0', '').replace('_', ''))
        # 图片上下等分
        image_list = cut_image(src_pic)
        # 保存图片并重命名
        for i in range(0,len(image_list)):
            img_name: str = namelist[(src_pic_seq - 1) * 2 + i] + '.jpg'
            image_list[i].save(os.path.join(dst_folder_path, img_name))

print("Finished！")
