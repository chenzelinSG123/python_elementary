from PIL import ImageGrab, Image
from win32com.client import Dispatch, DispatchEx
import pythoncom, os, itertools, time
import openpyxl as vb
import datetime
path3 = r"F:/excel/尺码表最终.xlsx"
now_date = datetime.datetime.now().strftime('%Y-%m-%d')
path4 = r"F:/excel/"+now_date+"尺码图"
os.mkdir(path4)
def get_alldata(row):
    path1 = r"F:/excel/尺码表.xlsx"
    wb1 = vb.load_workbook(path1)
    sheet1 = wb1['总表']
    a = sheet1['A'+str(row):'K'+str(row)]
    # print(a)
    list1=[]
    for date in a:
        for box in date:
            if box.value == None:
                continue
            else:
                list1.append(box.value)
    return list1

def get_list (row):
    alllist=[]
    for r in range(2,row+1):
        alllist.append(get_alldata(r))
    return alllist

def get_dic(row):
    dic ={}
    alllist = get_list(row)
    for onelist in alllist:
        dic[onelist[0]] = onelist[1:4]
    return dic

def st_list(row):
    dic = get_dic(row)
    picture_list=[]
    for k,v in dic.items():
        picture_list.append(k)
    return picture_list

def size_list(row):
    dic = get_dic(row)
    size=[]
    for k,v in dic.items():
        size.append(v)
    return size

def dowimg(row):
    sheet_list = st_list(row)
    sizelist= size_list(row)
    file_name = os.path.abspath(path3)  # 把相对路径转成绝对路径
    pythoncom.CoInitialize()  # 开启多线程
    # 创建Excel对象
    excel = DispatchEx('excel.application')
    excel.visible = False  # 不显示Excel
    excel.DisplayAlerts = 0  # 关闭系统警告(保存时不会弹出窗口)
    excel.ScreenUpdating = 1    # 关闭屏幕刷新
    collist = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J','K']
    rowlist = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]
    for i in range(0,len(sheet_list)):
        workbook = excel.workbooks.Open(file_name)       # 打开Excel文件
        size = sizelist[i]
        sheetname = sheet_list[i]
        row1= rowlist[int(size[0])+1]
        col1= collist[int(size[2])]
        workbook.Sheets(sheetname).select
        sheet = workbook.worksheets[sheetname]
        img_name = sheetname
        sheet.Range('A1:'+col1+str(row1)).CopyPicture() # 复制图片区域
        sheet.Paste()  # 粘贴
        excel.Selection.ShapeRange.Name = img_name  # 将刚刚选择的Shape重命名，避免与已有图片混淆
        sheet.Shapes(img_name).Copy()  # 选择图片
        img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
        img = img.convert('RGB')
        w = img.size[0] - 2
        h = img.size[1] - 2
        cropped = img.crop((2, 2, w, h))  # (left, upper, right, lower)
        cropped.save(path4+'/'+img_name + ".JPG")
        workbook.Close()  # 关闭Excel文件,不保存
        excel.Quit() # 退出Excel
    pythoncom.CoUninitialize() # 关闭多线程

def gen(row):
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' : 截图开始')

    try:
        flag = 'Y'
        while flag == 'Y':  # 循环调用截图函数
            try:
                dowimg(row)
                flag = 'N'
            except Exception as e:
                flag == 'Y'
    except Exception as e:
        print('main error is:', e)

    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' : 截图结束')


gen(11)