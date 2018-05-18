import os
import xlrd
import math
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39, code128, code93


## 尺寸mm
ID_SIZE = (297,210)
## 水印文本
WATERMARK_TXT = "泗县教体局"
## mm
POSITIONS = [
    [36,137],[53,137],# y+1
    [31,131],# y+1
    [38,123],[57,123],[63.5,123],# y+1
    [25,116],[43,116],# y+1
    [22,109],[38,109],# y+1
    [35,102],[54,102],# y+1
    [22,94],[37,94],
    [42,50],[55,50],

    [110,72],[116,50],[130,50],# y+1

    [205,138],[238,138],# y+1
    [173,129.5],[195,129.5],[220,129.5],# y+1
    [190,121.5],[220,121.5],# y+1
    [180,113],[220,113],# y+1
    [205,104],[230,104],# y+1
    [185,96],[205,96],
    [192,48],[215,48],
    ]
## 以上为以下七项的输出位置

IMG_PATH = ".\\gpdf"


# 照片打印位置
IMG_X = 108 
IMG_Y = 76

PAGE_SIZE = 50

def get_img_height(file,wid):
    im = Image.open(file)
    size = im.size
    im.close()
    height = (size[1]/size[0]) * wid
    return height

def confirm_path(path):
    if not os.path.exists(path):
        os.makedirs(path)

def set_font(canv,size,font_name='msyh',font_file='msyh.ttf'):
    pdfmetrics.registerFont(TTFont(font_name,font_file))
    canv.setFont(font_name,size)

# def draw_barcode(canv,idcode,codetype='code128'):
#     ## 绘制条形码
#     ## codetype have:code39,code93,code128
#     barcd = BAR_METHODS[codetype](idcode,barWidth=1,humanReadable=True)
#     barcd.drawOn(canv,BAR_X*mm,BAR_Y*mm)
#     set_font(canv,16,font_name='simsun',font_file='simsun.ttc')
#     canv.drawString(BAR_X*mm+10,BAR_Y*mm,"‡")
#     canv.drawString(BAR_X*mm+42*mm,BAR_Y*mm,"‡")
#     # barcode39 = code39.Extended39('34322545666',barHeight=1*cm,barWidth=0.8)
#     # barcode39.drawOn(c,20,20)
#     # barcode93 = code93.Standard93('34322545666')
#     # barcode93.drawOn(c,20,60)
#     # barcode128 = code128.Code128('34322545666')
#     # barcode128.drawOn(c,20,100)

def draw_page(canv,stud):
    ## 背景图
    # canv.drawImage('bg.jpg',POSITIONS[-1][0]*mm,POSITIONS[-1][1]*mm)
    set_font(canv,8)
    for data,pos in zip(stud[0],POSITIONS[:16]):
        x,y = pos
        canv.drawString(x*mm,y*mm,data)

    set_font(canv,11)
    for data,pos in zip(stud[1],POSITIONS[16:19]):
        x,y = pos
        canv.drawString(x*mm,y*mm,data)

    for data,pos in zip(stud[2],POSITIONS[19:]):
        x,y = pos
        canv.drawString(x*mm,y*mm,data)

    img_file = os.path.join('pho','.'.join((stud[-1],'jpg')))

    # 绘制照片
    width = 36 #+6
    height = get_img_height(img_file,width)
    canv.drawImage(img_file,IMG_X*mm,IMG_Y*mm,width=width*mm,height=height*mm)

    width = 17
    height = get_img_height(img_file,width)
    canv.drawImage(img_file,18.5*mm,46*mm,width=width*mm,height=height*mm)
    # # 绘制条形码
    # draw_barcode(canv,stud[0])
    # 绘制水印
    set_font(canv,10)
    canv.setFillColorRGB(180,180,180,alpha=0.3)
    canv.drawString(IMG_X*mm,IMG_Y*mm+9*mm,WATERMARK_TXT)
    canv.showPage()

def gen_pdf(dir_name,sch_name,studs,page):
    confirm_path(dir_name)
    path = os.path.join(dir_name,sch_name + str(page) + '.pdf')
    canv = canvas.Canvas(path,pagesize=(ID_SIZE[0]*mm,ID_SIZE[1]*mm))
    for stud in studs:
        draw_page(canv,stud)
    canv.save()

def get_space(data):
    length = len(data)
    space = '  '
    i = 4 - length
    return space * (i+1)

# 学号 姓名 身份证号
def gen(file='aa.xls'):
    studs = []
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(1,nrows):
        datas = ws.row_values(i)
        birth_year = datas[17][6:10]
        birth_month = datas[17][10:12]
        strs_left = [datas[1],datas[12],
            datas[3],
            datas[5],'\\'*(1 if int(datas[6])==2 else 0),'\\'*(0 if int(datas[6])==2 else 1),
            birth_year,birth_month,
            datas[8],datas[9],
            '2015','09',
            '2018','06',
            '2018','06']

        strs_mid = [datas[3],'18',datas[12]]

        strs_right = [
            datas[5],'\\' if int(datas[6])==2 else '',
            '' if int(datas[6])==2 else '\\',birth_year,birth_month,
            datas[8],datas[9],
            datas[10],datas[11],
            '2015','09',
            '2018','06',
            '2018','06']
        studs.append((strs_left,strs_mid,strs_right,datas[17]))
    pages = math.ceil(len(studs)/PAGE_SIZE)
    for i in range(pages):
        gen_pdf('.\\idsd','sz',studs[i*PAGE_SIZE:(i+1)*PAGE_SIZE],i)

def check_pho(file='aa.xls'):
    studs = []
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(1,nrows):
        datas = ws.row_values(i)
        file = os.path.join('pho','.'.join((datas[17],'jpg')))
        if not os.path.exists(file):
            studs.append((datas[17],datas[5]))
    print(studs)

if __name__ == '__main__':
    gen()
    # check_pho()

# def gen_all_pdfs(dir_name):
#     schs = getdata.get_schs()
#     for sch in schs:
#         studs = getdata.get_studs(sch)
#         gen_pdf(dir_name,sch,studs)


#以下各函数中照片文件未生成

# # 按学校生成准考证
# def gen_examid_sch(dir_name):
#     schs = select(s.sch for s in StudPh)
#     for sch in schs:
#         datas = select((s.phid,s.name,s.sex,s.exam_addr,s.sch,''.join(("Z",s.signid,'.jpg')))
#          for s in StudPh if s.sch==sch).order_by(StudPh.classcode,StudPh.phid)
#         gen_pdf(dir_name,sch,datas)

# # # 按时间段（半日）和考点生成准考证
# def gen_bak_examid(dir_name):
#     exam_addrs = select(s.exam_addr for s in StudPh)
#     exam_dates = select(s.exam_date for s in StudPh)
#     for exam_date in exam_dates:
#         for exam_addr in exam_addrs:
#             studs = select((s.phid,s.name,s.sex,s.exam_addr,s.sch,''.join(("Z",s.signid,'.jpg')))
#                 for s in StudPh if s.exam_addr==exam_addr and s.exam_date==exam_date).order_by(StudPh.phid)
#             gen_pdf(dir_name,exam_addr+exam_date,studs)
