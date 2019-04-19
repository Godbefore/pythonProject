import re
import xlrd
import xlwt
# xlrd和xlwt处理的是xls格式文件,单个sheet最大行数65535,如果超过65535就会报ValueError
# 如果需要读取更大数据量,可以使用openpyxl,最大行数1048576(1024*1024)


class Operate_excel:
    #初始化文件地址和sheet页
    #一个实例就是操作一个sheet对象
    def __init__(self, file, sheet):
        self.readbook = xlrd.open_workbook(file)
        # 以索引的方式，从0开始
        if isinstance(sheet,int):
            self.sheet = self.readbook.sheet_by_index(sheet)  
        # 以名字的方式
        else:
            self.sheet = self.readbook.sheet_by_name(sheet)

    def cell(self, x, y=None):
        # 用EXCEL自带的 大写字母+数字 的方式定位
        if y==None:
            return self._get_cell_by_name(x)
            
        # 用(x,y)坐标定位,从(0,0)开始
        else:
            return self.sheet.cell(x, y)
    
    def _get_cell_by_name(self, value):
        if re.match(r"[A-Z]+\d+$", value):
            column = int(re.findall(r"\d+", value)[0])
            line = re.findall(r"[A-Z]+", value)[0]
            # 把连续的大写字母转换为数字
            # 把26个大写字母组合看成26进制数,则num=n0*26**0+n1*26**1+n2*26**2+nn*26**n...
            # ord() 返回ASCII码对应的十进制整数
            row = 0
            for i, j in enumerate(line[::-1]):
                row += (ord(j) - 64) * 26 ** i
            return self.sheet.cell(row - 1, column - 1)
        else:
            raise ValueError("cell argument format error")

    # 最大列数
    @property
    def cols(self):
        return self.sheet.ncols
        
    # 最大行数
    @property
    def rows(self):
        return self.sheet.nrows



if __name__=="__main__":
    sh=Operate_excel("test.xls",0)
    print(sh.cell("B2"))
    print(type(sh.cell(1, 0)))
    print(sh.rows)
    print(sh.cols)
    # 遍历EXCEL表
    for x in range(sh.rows):
        for y in range(sh.cols):
            print(sh.cell(x, y).value,end="\t")
        print()
