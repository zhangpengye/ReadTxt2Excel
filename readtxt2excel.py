# coding = utf-8
import os
import xlwt
def readTxt2Excel():
    PathList = []
    valueList = []
    TimeList = []
    WaterList = []
    path = "E:\\gongzuo\\468zdtest"
    num = 1
    #存储年月日
    Water={}
    #默认传的文件夹里面都是txt，需要被读取的
    for filename in os.listdir(path):
        #print (filename)
        name = filename.split('.')[0]
        #此代码存放名称，用于excel名称
        #PathList.append(name)
        filepath = os.path.join(path, filename)
        shuju = []
        #以下代码用于处理每个txt，分别为对水量进行一个排除，分为舍掉换行符以及对五位数的截取后两位，先对水量，后对日期
        with open(filepath,'r',encoding="utf-8") as f:
            line = f.readline()
            while line:
                shuju = line.split(' ')
                shuju = [i for i in shuju if(len(str(i))!=0)]
                #通过strip来去掉换行符，这个试用于不知道是否有换行符的情况
                shuju[3] = shuju[3].strip('\n')
                if len(shuju[3])==5:
                    shuju[3]=shuju[3][-2:]
                    shuju[3]= int(shuju[3])
                    #print(shuju[3])
                shuju[3] = int(shuju[3])
                #print(shuju[3])
                valueList.append(shuju)
                #print(valueList)
                line = f.readline()
    
    #对日期进行筛选，water保存字典，key是年月日，value是另外
        for h in valueList:
        #print(h[2])
            temp = h[2][-4:]
        
        #print(temp)
            tempValue = str(h[0])+'_' + str(h[1])+'_' + str(h[3])
            aa = {h[2]:tempValue}
            Water.update(aa)
        
    #需要将字典转化为list，才能del字典里面的key
        for c in list(Water.keys()):
            LowerValue = 601
            HighValue = 831
            d = c[-4:]
            d = int(d)
        #print(d)
        #print(c)
            if d< LowerValue:
            #Water.pop(c)
                del Water[c]
            if d> HighValue:
            #Water.pop(c)
                del Water[c]
    #重新从字典中取出要素，组装成列表，value_存储目前往excel表里书写的字符，value_是二维数组
        value_ = []
        sub = []
        for key,value in Water.items():
            sub = value.split('_')
            sub.insert(0,key)
            value_.append(sub)
        #print(sub)
        #print('开始')
        #print(value_)

        
        #print(PathList)
        
        xls = os.path.join(path, name + '.xls')
        ToExcel(xls,value_)
            
        
        Water.clear()
        valueList = []
        value_ = []
        shuju = []
        f.close()
    
def ToExcel(path,value):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1',cell_overwrite_ok = True)
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = True
    style.font = font
    index = len(value)
    #可以按照顺序来
    
    for x in range(0,index):
        for y in range(0,len(value[x])):
            sheet.write(x,y,value[x][y])
    workbook.save(path)
    print('写入成功')
def main():
    readTxt2Excel()
if __name__ == '__main__':
    main()
