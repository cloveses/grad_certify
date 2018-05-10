import os
import xlrd
import math
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39, code128, code93


## 尺寸mm
ID_SIZE = (184,139)
## 水印文本
WATERMARK_TXT = "泗县教体局"
## mm
POSITIONS = ((110,12),(94,59),(94,66),(94,76),(94,84),(98,93),(123,100))
## 以上为以下七项的输出位置
ROWS = ['',] * 5
ROWS[0] = ROWS[1] = '2018&nbsp;&nbsp;&nbsp;&nbsp;6'
ROWS[2] = '2015&nbsp;&nbsp;&nbsp;&nbsp;9'
ROWS[3] = '泗'
ROWS[4] = '安徽&nbsp;&nbsp;&nbsp;&nbsp;宿州'
ROWS5 = '{}&nbsp;&nbsp;&nbsp;&nbsp;{}&nbsp;&nbsp;&nbsp;&nbsp;{}'
ROWS6 = '{}&nbsp;&nbsp;&nbsp;&nbsp;{}'

IMG_PATH = ".\\gpdf"

# BAR_METHODS = {'code39':code39.Extended39, 
#             'code128':code128.Code128,
#             'code93':code93.Standard93}

# # 条形码打印位置
# BAR_X = 120
# BAR_Y = 14

# 照片打印位置
IMG_X = 120
IMG_Y = 24

STUD_NO_X = STUD_NO_Y = 35

PAGE_SIZE = 50

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
    # 绘制标题
    set_font(canv,11)
    canv.drawString(*TITLE)
    set_font(canv,18)
    canv.drawString(*ID_NAME)
    set_font(canv,10)
    ## 背景图
    # canv.drawImage('bg.jpg',POSITIONS[-1][0]*mm,POSITIONS[-1][1]*mm)
    for index,info in enumerate(stud[:5]):
        canv.drawString(POS_X*mm,POS_Y[index]*mm,''.join((ITEM_NAMES[index],info)))
    # 绘制照片
    canv.drawImage(stud[-1],IMG_X*mm,IMG_Y*mm)
    # # 绘制条形码
    # draw_barcode(canv,stud[0])
    # 绘制水印
    set_font(canv,10)
    canv.setFillColorRGB(180,180,180,alpha=0.3)
    canv.drawString(IMG_X*mm+mm,IMG_Y*mm+0.5*mm,WATERMARK_TXT)
    canv.showPage()

def gen_pdf(dir_name,sch_name,studs):
    confirm_path(dir_name)
    path = os.path.join(dir_name,sch_name + '.pdf')
    canv = canvas.Canvas(path,pagesize=(ID_SIZE[0]*mm,ID_SIZE[1]*mm))
    for stud in studs:
        draw_page(canv,stud)
    canv.save()


# 学号 姓名 身份证号
def gen(file='aa.xls'):
    studs = []
    wb = xlrd.open_workbook(file)
    ws = wb.sheets()[0]
    nrows = ws.nrows
    for i in range(1,nrows):
        datas = ws.row_values(i)
        birth_year = datas[2][6:10]
        birth_month = datas[2][10:12]
        sex = int(datas[2][-2]) % 2 == 1
        row_6 = ROWS6.format(datas[1],'\\' if sex else '')
        row_5 = ROWS5.format('' if sex else '\\',birth_year,birth_month)
        data = []
        data.extend(ROWS)
        data.append(row_5)
        data.append(row_6)
        data.append(datas[0])
        data.append('.'.join(datas[2],'png'))
        studs.append(data)
    pages = math.ceil(len(studs)/PAGE_SIZE)
    for i in range(pages):
        gen_pdf('.\\idsd','sz',studs[i*PAGE_SIZE:(i+1)*PAGE_SIZE])

if __name__ == '__main__':
    gen()

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
