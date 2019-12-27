import pandas as pd
import numpy as np
import os

def examin(column, column_name, length, lower_equal, flag, exam_column, istype):
#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)
    i = 0
    count = 0
    while(i < df.shape[0]-1):
        item = df.iloc[i][column]
        if(i >= 3):
            if (pd.isnull(item) == False):
                if (istype == 1):
                    if type(item) is not str:
                        print("%s型態有錯!請改為文字,位置:%s%d" %(str(item), column_name, i+1))
                        count += 1
                        
                if (lower_equal == 'equal'):        
                    if(len(str(item)) < length):
                        print("%s字數%d,少於%d!位置:%s%d" %(str(item), len(str(item)), length, column_name, i+1))
                        count += 1
                    elif(len(str(item)) > length):
                        print("%s字數%d,多於%d!位置:%s%d" %(str(item), len(str(item)), length, column_name, i+1))
                        count += 1
                else:
                    if(len(str(item)) > length):
                        print("%s字數%d,多於%d!位置:%s%d" %(str(item), len(str(item)), length, column_name, i+1))
                        count += 1    
            if(flag == 1 and pd.isnull(item)):
                #print(exam_column)
                #os.system("pause")
                print("%s%d不可空白!" %(column_name, i+1))
                count += 1
        i += 1
    if(count > 0):
        print("%s錯誤,共%d筆!" %(exam_column, count))
        print("\n")


def duplicate(column, **else_column):
    count_duplicate = 0
    i = 3
    while(i < df.shape[0]-2):#最後一個不用再跟別人比
        j = i+1
        if (else_column
        item = df.iloc[i][column]
        while(j <= df.shape[0]-1):
            if (item == df.iloc[j][column]):
                print("A%d與A%d重複" %(i+1,j+1))
                count_duplicate += 1
            j += 1
        i += 1
    print("共%d筆重複" %count_duplicate)    
    print("\n")        

def continuous(column):
    count_uncontinue = 0


            
#index_col 索引欄
#print(df.head(10))
#print(df.tail())
###JRCF020坐落地址檔
df = pd.read_excel('D:\sample.xlsx', sheet_name = 'JRCF020坐落地址檔', index_col = None, header = None, skiprow = 1)

#檢查稅號是否重複
duplicate(0)

#檢查結尾*字號
if(df.iloc[df.shape[0]-1][0] != '*'):
    print("結尾沒有*")
    print("\n")

#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)    
examin(0, 'A', 11, 'equal', 1, '房屋稅籍編號', 1)        
examin(1, 'B', 3, 'equal', 0, '郵遞區號', 0)
examin(2, 'C', 5, 'equal', 1, '縣市鄉鎮村里', 0)
examin(3, 'D', 3, 'equal', 0, '鄰', 0) 
examin(4, 'E', 4, 'equal', 1, '街路', 0)
examin(5, 'F', 2, 'equal', 0, '段', 0)
examin(6, 'G', 4, 'equal', 0, '巷', 0)
examin(7, 'H', 3, 'equal', 0, '中文巷', 0)
examin(8, 'I', 3, 'equal', 0, '弄', 0)
examin(9, 'J', 3, 'equal', 0, '衖', 0)
examin(10, 'K', 4, 'equal', 1, '號', 0)
examin(11, 'L', 4, 'equal', 0, '號之', 0)
examin(12, 'M', 4, 'equal', 0, '之', 0)
examin(13, 'N', 4, 'equal', 0, '之號', 0)
examin(14, 'O', 4, 'equal', 0, '樓', 0)
examin(15, 'P', 4, 'equal', 0, '樓之', 0)
examin(16, 'Q', 2, 'equal', 0, '室', 0)

###HOUF100
df = pd.read_excel('D:\sample.xlsx', sheet_name = 'HOUF100中文主檔', index_col = None, header = None, skiprow = 1)

#檢查稅號是否重複
duplicate(0)

#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)examin(0, 'A', 11, 1, '稅籍編號', 0)
examin(0, 'A', 11, 'equal', 1, '稅籍編號', 1)
examin(1, 'B', 1, 'equal', 1, '公私有別', 1)
examin(2, 'C', 1, 'equal', 0, '歸屬別', 0)
examin(3, 'D', 24, 'equal', 0, '設籍文號', 1)
examin(4, 'E', 7, 'equal', 1, '設籍日期', 0)
examin(5, 'F', 14, 'equal', 0, '使用執照編號', 1)

###HOUF110課稅主檔,111加減項檔
df = pd.read_excel('D:\sample.xlsx', sheet_name = 'HOUF110課稅主檔,111加減項檔', index_col = None, header = None, skiprow = 1)

#檢查稅號+卡序是否重複
duplicate(0)

#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)examin(0, 'A', 11, 1, '稅籍編號', 1)
examin(0, 'A', 11, 'equal', 1, '稅籍編號', 1)
examin(1, 'B', 14, 'equal', 1, '使照編號', 1)
examin(2, 'C', 3, 1, 'lower', '層次', 1)
examin(3, 'D', 1, 'equal', 1, '卡序', 1)
examin(4, 'E', 1 'equal',, 1, '建物類別', 0)
examin(5, 'F', 1, 'equal', 0, '公設名稱', 0)
examin(6, 'G', 1, 'equal', 1, '構造別', 1)
examin(7, 'H', 1, 'equal', 1, '用途類別', 0)
examin(8, 'I', 2, 'equal', 1, '用途細類', 0)
examin(9, 'J', 3, 'lower', 1, '總層數', 0)
examin(10, 'K', 6, 'lower', 0, '標準單價', 0)
examin(11, 'L', 3, 'equal', 1, '樓層高度', 1)
examin(12, 'M', 8, 'lower', 0, '核定單價', 0)
examin(13, 'N', 3, 'equal', 1, '折舊率', 0)
examin(14, 'O', 1, 'equal', 1, '折舊率序號', 0)
examin(15, 'P', 2, 'lower', 1, '經歷年數', 0)
examin(16, 'Q', 3, 'lower', 1, '地段調整率', 0)
examin(17, 'R', 8, 'lower', 0, '評定總值', 0)
examin(18, 'S', 8, 'lower', 1, '營業', 0)
examin(19, 'T', 8, 'lower', 1, '營業減半', 0)
examin(20, 'U', 8, 'lower', 1, '住家', 0)
examin(21, 'V', 8, 'lower', 1, '住家用減半', 0)
examin(22, 'W', 8, 'lower', 1, '非住營', 0)
examin(23, 'X', 8, 'lower', 1, '非住營減半', 0)
examin(24, 'Y', 2, 'lower', 1, '課稅月數', 0)
examin(25, 'Z', 1, 'equal', 1, '特殊課稅', 1)
examin(26, 'AA', 1, 'equal', 0, '免稅代號', 0)
examin(27, 'AB', 13, 'equal', 0, '減免法令條款', 1)
examin(28, 'AC', 7, 'equal', 0, '減免稅起日', 1)
examin(29, 'AD', 7, 'equal', 0, '減免稅迄日', 1)
examin(30, 'AE', 5, 'equal', 1, '起課年月', 1)
examin(31, 'AF', 1, 'equal', 0, '折算註記', 0)
examin(32, 'AG', 3, 'lower', 0, '折算率', 0)
examin(33, 'AH', 1, 'equal', 0, '加減項1', 1)
examin(34, 'AI', 2, 'equal', 0, '加減項1', 1)
examin(35, 'AJ', 1, 'equal', 0, '加減項2', 1)
examin(36, 'AK', 2, 'equal', 0, '加減項2', 1)
examin(37, 'AL', 1, 'equal', 0, '加減項3', 1)
examin(38, 'AM', 2, 'equal', 0, '加減項3', 1)
examin(39, 'AN', 1, 'equal', 0, '加減項4', 1)
examin(40, 'AO', 2, 'equal', 0, '加減項4', 1)
examin(41, 'AP', 1, 'equal', 0, '加減項5', 1)
examin(42, 'AQ', 2, 'equal', 0, '加減項5', 1)
examin(43, 'AR', 1, 'equal', 0, '加減項6', 1)
examin(44, 'AS', 2, 'equal', 0, '加減項6', 1)
examin(45, 'AT', 1, 'equal', 0, '加減項7', 1)
examin(46, 'AU', 2, 'equal', 0, '加減項7', 1)
examin(47, 'AV', 1, 'equal', 0, '加減項8', 1)
examin(48, 'AW', 2, 'equal', 0, '加減項8', 1)
examin(49, 'AX', 1, 'equal', 0, '折算序號', 0)
examin(50, 'AY', 2, 'lower', 0, '分層分攤率', 0)
examin(51, 'AZ', 1, 'equal', 1, '稅率代號', 0)
examin(52, 'BA', 1, 'equal', 0, '使用別', 0)


###HOUF120持分人檔
df = pd.read_excel('D:\sample.xlsx', sheet_name = 'HOUF120持分人檔', index_col = None, header = None, skiprow = 1)

#檢查稅號+卡別是否重複
duplicate(0)

#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)
examin(0, 'A', 11, 'equal', 1, '稅籍編號', 1)
examin(1, 'B', 7, 'equal', 0, '移轉日期', 1)
examin(2, 'C', 1, 'equal', 1, '卡別', 0)
examin(3, 'D', 1, 'equal', 0, '群組', 1)
examin(4, 'E', 7, 'lower', 1, '持分分子', 0)
examin(5, 'F', 7, 'lower', 1, '持分分母', 0)
examin(6, 'G', 2, 'equal', 1, '課稅月數', 1)
examin(7, 'H', 1, 'equal', 1, '是否分單', 1)
examin(8, 'I', 1, 'equal', 0, '是否分管', 1)
examin(9, 'J', 10, 'lower', 1, 'IDN_BAN', 1)
examin(10, 'K', 1, 'equal', 0, '身分代號', 1)
examin(11, 'L', 26, 'lower', 1, '姓名', 1)
examin(12, 'M', 26, 'lower', 0, '備註', 1)
examin(13, 'N', 54, 'lower', 1, '通訊地址', 1)

###HOUF151土地標示檔
df = pd.read_excel('D:\sample.xlsx', sheet_name = 'HOUF151土地標示檔', index_col = None, header = None, skiprow = 1)

#檢查稅號+土標是否重複
duplicate(0)

#參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 型態是否限定文字(是：1, 否：0)
examin(0, 'A', 11, 'equal', 0, '稅籍編號', 1)
examin(1, 'B', 14, 'equal', 0, '土地標示', 1)

'''
i = 0
while i < (df.shape[0]):    
    j = 0
    while j < (df.shape[1]):
        if(df.iloc[i][j] != "nan"):
            print(df.iloc[i][j], end = '  ')
            j += 1
    print("\n")        
    i += 1    
    os.system("pause")
'''
#print(df.info())
#print(df['稅籍編號'].head())
#print(df.稅籍編號)
#print(df.shape[1])
#print(df[0])
