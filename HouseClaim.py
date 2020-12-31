import pandas as pd
import numpy as np
import os
import traceback
import re

# iloc[列, 行]

def writeCsv(file_name, column_name, row_num, seq, wrong_type=None):
	# 存檔路徑
	file_name = str(filename).split(".")[0]
	sFilepath = "./%s.csv" %file_name
	
	# 開啟要寫入的檔案
	f = open(sFilepath, 'a+', encoding = 'UTF-8', newline ='')
	f_writer = csv.writer(f, dialect = 'excel')
	
	wrong_dict = {}
	fields = ['序號', '欄位名稱', '位置', '錯誤型態', '正確型態']

	wrong_dict = {'序號':seq, '欄位名稱':column_name, '位置':row_num, '錯誤型態':wrong_type, '正確型態':None}
	if seq == 1:
		f_writer.writerow(wrong_dict.keys())
	f_writer.writerow(wrong_dict.values())
	
	# 關閉檔案
	f.close()

def writeText(file_name, row_num, hou_losn, card, df_area, data_area, error_seq):
	# 存檔路徑
	file_name = str(file_name).split(".")[0]
	sFilepath = "./%s.txt" %file_name
	
	# 如果檔案已存在就刪除
	if os.path.isfile(sFilepath) and error_seq == 1: 
		os.remove(sFilepath)
	
	# 開啟要寫入的檔案
	f = open(sFilepath, 'a+', encoding = 'ANSI', newline ='')
	f.write("%d %s%d面積不同! 稅號%s, 卡序%s, 課稅主檔面積:%s, 公共設施檔面積:%s\n" %(error_seq, 'U', row_num, hou_losn, card, df_area, data_area))
	
	# 關閉檔案
	f.close()

def examin(dataFrame, column, column_name, length, lower_equal, necessary, exam_column, type_, length2 = None):
# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(整數：0, 文字：1, 小數：2), 第2種長度(針對IDN_BAN有2種長度)
	df = dataFrame

	lengthCount = 0
	emptyCount = 0
	count  = 0

	i = 0
	while i < df.shape[0]-1:
		item = df.iloc[i, column]
		
		if not pd.isnull(item):
			# type_ = 1 文字型態
			if type_ == 1:
				item = str(item).replace(".0", "")
				if length2 is None:
					if lower_equal == 'equal':
						if len(item) < length:
							print("位置:%s%d, %s 字數%d,少於%d!" %(column_name, i+4, item, len(item), length))
							lengthCount += 1
							
						elif len(item) > length:
							print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, item, len(item), length))
							lengthCount += 1        
					else:
						if len(item) > length:
							print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, item, len(item), length))
							lengthCount += 1
				
				else:	# 針對IDN有2種長度
					if len(item) != length and len(item) != length2:
						print("位置:%s%d, %s 字數%d, 不符合長度%d或%d" %(column_name, i+4, str(item), len(item), length, length2))
			
			# type_ = 0 整數型態
			elif type_ == 0:
				if int(item) == float(item):
					if lower_equal == 'equal':
						if len(str(int(item))) < length:
							print("位置:%s%d, %s 字數%d,少於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
							lengthCount += 1
						elif len(str(int(item))) > length:
							print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
							lengthCount += 1
					elif lower_equal == 'lower':
						if len(str(int(item))) > length:
							print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
							lengthCount += 1
				else:
					print("位置:%s%d, %s可能有小數!" %(column_name, i+4, str(item)))
					lengthCount += 1
			
			# type_ = 2 小數型態
			elif type_ == 2:
				if lower_equal == 'equal':
					if len(str(item)) < length:
						print("位置:%s%d, %s 字數%d,少於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
						lengthCount += 1
					elif len(str(item)) > length:
						print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
						lengthCount += 1
				
				elif lower_equal == 'lower':
					if len(str(item)) > length:
						print("位置:%s%d, %s 字數%d,多於%d!" %(column_name, i+4, str(item), len(str(int(item))), length))
						lengthCount += 1		
						
		if pd.isnull(item) and necessary == 1:
			print("%s%d不可空白!" %(column_name, i+4))
			emptyCount += 1
		i += 1
	count = lengthCount+emptyCount
	
	if count > 0:
		print("%s錯誤,共%d筆!\n" %(exam_column, count))
		os.system("pause")
	return count

def examArea(file_name):
	try:
		# step1:存取公共設施sheet資料
		df1 = pd.read_excel(files[readFile-1], sheet_name = '公共設施', index_col = None, header = 6, skiprows=0)
		
		# data 的欄位名稱
		data = []
		data.append([])
		data[0] = ['稅號', '地址']		
		columnNum = 0
		while pd.notnull(df1.iloc[1, 5+columnNum]):
			# 取得所有卡序名稱
			card = df1.iloc[1, 5+columnNum]		# df1.iloc 的位置從 header 開始往下數
			result = re.search("[A-Z]", card).span() 	# 正則表達式 找出卡序
			#print(card[result[0]:result[1]])
			if (int(card[0:result[0]]) < 990 and card[result[0]:result[1]+1] not in ('P', 'Y', 'Z')):	# 如果樓層<990 且 卡序不是 PYZ		
				card = card[result[0]:result[1]+1]	# 只抓卡不抓樓層
			data[0].append(card)
			#print(data[0])	
			columnNum += 1
			
		#os.system("pause")
		
		# areaData 放地址和面積
		areaData = []	
		i = 0
		while (i < df1.shape[0]-4):	# shape 大小要把 header 之前的都算進去
			areaData.append([])
			
			# append 地址
			for dash in ('-', '—'):	
				if str(df1.iloc[i+4, 1]).find(dash) > -1:
					item1 = str(df1.iloc[i+4, 1]).replace(dash, '之')
					break
				else:
					item1 = str(df1.iloc[i+4, 1])
			areaData[i].append(item1)	
			
			# append 各卡序面積
			for j in range(0, columnNum):
				areaData[i].append(df1.iloc[i+4, 5+j])	
	
			#print("areaData:", areaData[i])
			i += 1
		
		# step2:讀取 JRCF020坐落地址檔
		df2 = pd.read_excel(files[readFile-1], sheet_name = 'JRCF020坐落地址檔', index_col = None, header = 2, skiprows=0,
							dtype = {'稅籍編號':str, '號':str, '號之':str, '之':str, '之號':str, '樓':str, '樓之':str})	
		# data 記錄稅號與地址
		k = 0	# k 從 header 之後開始算
		while (k < df2.shape[0]-1):
			data.append([])
			data[k+1].append(df2.iloc[k, 0])	# 第0列放欄位名稱
			#print(data)
			#os.system("pause")
			if pd.notnull(df2.iloc[k, 10]):
				addr = str(df2.iloc[k, 10]) + "號"
				if pd.notnull(df2.iloc[k, 11]): 
					addr = addr + "之" + str(df2.iloc[k, 11])	# M號之N

			elif pd.notnull(df2.iloc[k, 12]):
				addr = str(df2.iloc[k, 12]) + "之"
				if pd.notnull(df2.iloc[k, 13]):
					addr = addr + str(df2.iloc[k, 13]) + "號"	# M之N號

			if pd.notnull(df2.iloc[k, 14]):
				addr = addr + str(df2.iloc[k, 14]) + "樓"

			data[k+1].append(addr)
			#print(data[k])	
			k += 1
		
		#print(data)
		#os.system("pause")
		#print(areaData, "\n")
		#os.system("pause")
		
		for sub_data in data:
			#print("sub_data", sub_data)
			for sub_areaData in areaData:
				#print("sub_areaData", sub_areaData)
				if sub_data[1] == sub_areaData[0]:
					secondsub = 1
					while secondsub < len(sub_areaData):
						sub_data.append(sub_areaData[secondsub])
						secondsub += 1
			#print(sub_data)
		# step3:讀取 HOUF110課稅主檔,111加減項檔
		df3 = pd.read_excel(files[readFile-1], sheet_name = 'HOUF110課稅主檔,111加減項檔', index_col = None, header = 2, skiprows=0,
					dtype = {'稅籍編號':str, '層次':str, '卡序':str})
		
		p = 0	# p 從 header 之後開始算
		error_seq = 0
		while (p < df3.shape[0]-1):
			#print(df3.iloc[p][0])
			#n = input("")
			#如果樓層<990 且 卡序不是 PYZ
			card_seq = str(df3.iloc[p, 3]) if (int(df3.iloc[p, 2]) < 990 and df3.iloc[p, 3] not in ('P', 'Y', 'Z')) else str(df3.iloc[p, 2]) + str(df3.iloc[p, 3])
			
			for q in range(1, len(data)):	# data 裡面放稅號和地址
				if df3.iloc[p, 0] == data[q][0]:	# HOUF110 的稅號比對 data 裡的稅號 
					for r in range(2, len(data[0])):	# 找卡序
						if card_seq == data[0][r]:	# 樓層+卡序相同
							#print("HOUF110卡序", card_seq, ",公共設施卡序:", data[0][r], ",課稅主檔面積:", df3.iloc[p, 20], ",公共設施:", data[q][r])
							if df3.iloc[p, 20] != data[q][r]:	# 比對面積 不同就印出欄位
								error_seq += 1
								print("%s%d面積不同! 稅號%s, 卡序%s, 課稅主檔面積:%s, 公共設施檔面積:%s" %('U', p+4, df3.iloc[p, 0], card_seq, df3.iloc[p, 20], data[q][r]))
								writeText(file_name, p+4, df3.iloc[p, 0], card_seq, df3.iloc[p, 20], data[q][r], error_seq)
							else:	# 面積相同 break 換下一筆
								break
			p += 1	
		del data
		if error_seq == 0:
			print("面積正確!")
		else:
			print("\n比對結束!詳細錯誤欄位請查看%s.txt\n" %file_name.split('.')[0])
	except:
		traceback.print_exc()		

def duplicate(column):
	count = 0
	i = 0
	while(i < df.shape[0]-2):	# 最後一個不用再跟別人比
		j = i+1
		item = df.iloc[i, column]
		while(j <= df.shape[0]-1):
			if (item == df.iloc[j, column]):
				print("資料重複!A%d與A%d重複" %(i+4,j+4))
				count += 1
			j += 1
		i += 1
	return count 

def duplicate2(column, column2):
	count = 0
	i = 0
	while(i < df.shape[0]-2):	# 最後一個不用再跟別人比
		j = i+1
		item = df.iloc[i, column]
		item2 = df.iloc[i, column2]
		while(j <= df.shape[0]-1):
			if (item == df.iloc[j, column] and item2 == df.iloc[j, column2]):
				print("資料重複!A%d與A%d重複" %(i+4,j+4))
				count += 1
			j += 1
		i += 1
	return count

def duplicate3(column, column2, column3):
	count = 0
	i = 0
	while i < df.shape[0]-2:	# 最後一個不用再跟別人比
		j = i+1
		item = df.iloc[i, column]
		item2 = df.iloc[i, column2]
		item3 = df.iloc[i, column3]
		while(j <= df.shape[0]-1):
			if (item == df.iloc[j, column]):
				if (item2 == df.iloc[j, column2]):
					if (item3 == df.iloc[j, column3]):
						print("資料重複!A%d與A%d重複" %(i+4,j+4))
						count += 1
			j += 1
		i += 1
	return count

def specialrate():
	i = 0
	while i < df.shape[0]-1:
		if(df.iloc[i, 25] == 'Y' and df.iloc[i, 51] != '2'):
			print("Z%d與AZ%d加重課稅註記與稅率不符" %(i+3, i+3))
		if(df.iloc[i, 25] == 'N' and df.iloc[i, 51] != '1'): 
			print("Z%d與AZ%d加重課稅註記與稅率不符" %(i+3, i+3))
		i += 1

def existHou(df, hou_list):
	i = 0
	count = 0
	
	while i < (df.shape[0]-1):
		item = df.iloc[i, 0]
		#print(item)
		if item not in hou_list:
			print("位置:A%d, 稅號%s不存在HOUF100" %(i, item))
			count += 1
		i += 1
	
	if count > 0:
		print("稅號不存在HOUF100共%d筆!\n" %count)
	
	return count

def listFiles(filetype_ = "xls", path = os.getcwd(), massage = None):
	files = os.listdir(path)
	qualifyFiles = []

	if (massage != None):
		print(massage)
	
	for f in files:
		# 要給完整路徑才能正確讀取
		allPath = os.path.join(path,f)
		if (os.path.isfile(allPath)):
			if (f[f.find(".")+1:] == filetype_):
				qualifyFiles.append(f)       
	for compareF, i in zip(qualifyFiles, range(1,len(qualifyFiles)+1)):
		print("(%d)%s" %(i, compareF)) 
	return qualifyFiles


errorhou_count = 0

try:
	while 1:
		###主程式            
		#index_col 索引欄
		#print(df.head(10))
		#print(df.tail())

		##列出檔案供選擇
		files = listFiles(filetype_ = "xls")
		readFile = int(input("請選擇檔案："))
		print("\n")
		
		##HOUF100
		df = pd.read_excel(files[readFile-1], sheet_name = 'HOUF100中文主檔', index_col = None, header = 2, skiprows = 0,
							dtype = {'稅籍編號':str, '公私有別':str, '歸屬別':str, '設籍文號':str, '設籍日期':str, '使用執照編號':str})

		print("\n開始檢查HOUF100...")
		columns = df.columns
		
		# 抓所有稅號，用來檢查其他表的稅號是否存在HOUF100
		hou = []
		i = 0
		while i < (df.shape[0]-1):
			hou.append(df.iloc[i, 0])
			i += 1

		# 檢查稅號是否重複
		duplicate(0)

		# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(數字：0, 文字：1, 小數：2)
		examin(df, 0, 'A', 11, 'equal', 1, '稅籍編號', 1)
		examin(df, 1, 'B', 1, 'equal', 1, '公私有別', 1)
		examin(df, 2, 'C', 1, 'equal', 0, '歸屬別', 1)
		examin(df, 3, 'D', 24, 'lower', 0, '設籍文號', 1)
		examin(df, 4, 'E', 7, 'equal', 0, '設籍日期', 1)
		examin(df, 5, 'F', 14, 'equal', 0, '使用執照編號', 1)


		##JRCF020坐落地址檔
		sheet = 'JRCF020坐落地址檔'
		df = pd.read_excel(files[readFile-1], sheet_name = sheet, index_col = None, header = 2, skiprows = 0,
							dtype = {'稅籍編號':str, '郵遞區號':str, '縣市鄉鎮村里':str, '鄰':str, '路街':str, '段':str, '巷':str,
									 '中文巷':str, '弄':str, '衖':str, '號':str, '號之':str, '之':str, '之號':str, '樓':str, '樓之':str,
									 '室':str, '房屋座落':str, '特殊地址註記':str})
		print("\n開始檢查JRCF020")

		# 檢查稅號是否存在HOUF100
		errorhou_count += existHou(df, hou)

		# 檢查稅號是否重複
		duplicate(0)

		# 檢查結尾*字號
		if df.iloc[df.shape[0]-1, 0] != '*':
			print("結尾沒有*")
			print("\n")

		# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(整數：0, 文字：1, 小數：2)  
		examin(df, 0, 'A', 11, 'equal', 1, '房屋稅籍編號', 1)        
		examin(df, 1, 'B', 3, 'equal', 0, '郵遞區號', 0)
		examin(df, 2, 'C', 5, 'equal', 1, '縣市鄉鎮村里', 1)
		examin(df, 3, 'D', 3, 'equal', 0, '鄰', 1) 
		examin(df, 4, 'E', 4, 'equal', 1, '街路', 1)
		examin(df, 5, 'F', 2, 'equal', 0, '段', 1)
		examin(df, 6, 'G', 4, 'equal', 0, '巷', 1)
		examin(df, 7, 'H', 0, 'equal', 0, '中文巷', 1)
		examin(df, 8, 'I', 3, 'equal', 0, '弄', 1)
		examin(df, 9, 'J', 3, 'equal', 0, '衖', 1)
		examin(df, 10, 'K', 4, 'lower', 1, '號', 0)
		examin(df, 11, 'L', 4, 'equal', 0, '號之', 0)
		examin(df, 12, 'M', 4, 'equal', 0, '之', 0)
		examin(df, 13, 'N', 4, 'equal', 0, '之號', 0)
		examin(df, 14, 'O', 3, 'lower', 0, '樓', 0)
		examin(df, 15, 'P', 4, 'equal', 0, '樓之', 0)
		examin(df, 16, 'Q', 2, 'equal', 0, '室', 1)


		##HOUF110課稅主檔,111加減項檔
		df = pd.read_excel(files[readFile-1], sheet_name = 'HOUF110課稅主檔,111加減項檔', index_col = None, header = 2, skiprows = 0,
							dtype = {'稅籍編號':str, '使照編號':str, '層次':str, '卡序':str, '建物類別':str,
									 '公設名稱':str, '構造別':str, '用途類別':str, '用途細類':str, '總層數':str,
									 '標準單價':str, '樓層高度':str, '核定單價':str, '折舊率':str, '折舊率序號':str,
									 '經歷年數':str, '地段調整率':str, '評定總值':str, '營業':str, '營業減半':str,
									 '住家':str, '住家用減半':str, '非住營':str, '非住營減半':str, '課稅月數':str,
									 '加重課稅':str, '免稅代號':str, '減免法令條款':str, '減免稅起日':str, '減免稅迄日':str,
									 '起課年月':str, '折算註記':str, '折算率':str, '加減項1':str, '加減項1':str, '加減項2':str, '加減項3':str,
									 '加減項4':str, '加減項5':str, '加減項6':str, '加減項7':str, '加減項8':str, '折算序號':str,
									 '分層分攤率':str, '稅率代號':str, '使用別':str})
		print("\n開始檢查HOUF110課稅主檔,111加減項檔")

		# 檢查稅號是否存在HOUF100
		errorhou_count += existHou(df, hou)

		# 檢查稅號+層次+卡序是否重複
		duplicate3(0, 2, 3)

		# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(整數：0, 文字：1, 小數：2), 第2種長度(針對IDN_BAN有2種長度)
		examin(df, 0, 'A', 11, 'equal', 1, '稅籍編號', 1)
		examin(df, 1, 'B', 14, 'equal', 0, '使照編號', 1)
		examin(df, 2, 'C', 3, 'lower', 1, '層次', 0)
		examin(df, 3, 'D', 1, 'equal', 1, '卡序', 1)
		examin(df, 4, 'E', 1, 'equal', 1, '建物類別', 0)
		examin(df, 5, 'F', 1, 'equal', 0, '公設名稱', 0)
		examin(df, 6, 'G', 1, 'equal', 1, '構造別', 1)
		examin(df, 7, 'H', 1, 'equal', 1, '用途類別', 0)
		examin(df, 8, 'I', 2, 'equal', 1, '用途細類', 0)
		examin(df, 9, 'J', 3, 'lower', 1, '總層數', 0)
		examin(df, 10, 'K', 6, 'lower', 0, '標準單價', 0)
		examin(df, 11, 'L', 3, 'lower', 1, '樓層高度', 0)
		examin(df, 12, 'M', 8, 'lower', 0, '核定單價', 0)
		examin(df, 13, 'N', 4, 'lower', 1, '折舊率',2)
		examin(df, 14, 'O', 1, 'equal', 1, '折舊率序號',0)
		examin(df, 15, 'P', 2, 'lower', 1, '經歷年數', 0)
		examin(df, 16, 'Q', 3, 'lower', 1, '地段調整率', 0)
		examin(df, 17, 'R', 8, 'lower', 0, '評定總值', 0)
		examin(df, 18, 'S', 8, 'lower', 1, '營業', 2)
		examin(df, 19, 'T', 8, 'lower', 1, '營業減半', 2)
		examin(df, 20, 'U', 8, 'lower', 1, '住家', 2)
		examin(df, 21, 'V', 8, 'lower', 1, '住家用減半', 2)
		examin(df, 22, 'W', 8, 'lower', 1, '非住營', 2)
		examin(df, 23, 'X', 8, 'lower', 1, '非住營減半', 2)
		examin(df, 24, 'Y', 2, 'lower', 1, '課稅月數', 0)
		examin(df, 25, 'Z', 1, 'equal', 1, '特殊課稅', 1)
		examin(df, 26, 'AA', 1, 'equal', 0, '免稅代號', 0)
		examin(df, 27, 'AB', 13, 'equal', 0, '減免法令條款', 1)
		examin(df, 28, 'AC', 7, 'equal', 0, '減免稅起日', 1)
		examin(df, 29, 'AD', 7, 'equal', 0, '減免稅迄日', 1)
		examin(df, 30, 'AE', 5, 'equal', 1, '起課年月', 1)
		examin(df, 31, 'AF', 1, 'equal', 0, '折算註記', 0)
		examin(df, 32, 'AG', 3, 'lower', 0, '折算率', 2)
		examin(df, 33, 'AH', 1, 'equal', 0, '加減項1', 1)
		examin(df, 34, 'AI', 2, 'equal', 0, '加減項1', 0)
		examin(df, 35, 'AJ', 1, 'equal', 0, '加減項2', 1)
		examin(df, 36, 'AK', 2, 'equal', 0, '加減項2', 0)
		examin(df, 37, 'AL', 1, 'equal', 0, '加減項3', 1)
		examin(df, 38, 'AM', 2, 'equal', 0, '加減項3', 0)
		examin(df, 39, 'AN', 1, 'equal', 0, '加減項4', 1)
		examin(df, 40, 'AO', 2, 'equal', 0, '加減項4', 0)
		examin(df, 41, 'AP', 1, 'equal', 0, '加減項5', 1)
		examin(df, 42, 'AQ', 2, 'equal', 0, '加減項5', 0)
		examin(df, 43, 'AR', 1, 'equal', 0, '加減項6', 1)
		examin(df, 44, 'AS', 2, 'equal', 0, '加減項6', 0)
		examin(df, 45, 'AT', 1, 'equal', 0, '加減項7', 1)
		examin(df, 46, 'AU', 2, 'equal', 0, '加減項7', 0)
		examin(df, 47, 'AV', 1, 'equal', 0, '加減項8', 1)
		examin(df, 48, 'AW', 2, 'equal', 0, '加減項8', 0)
		examin(df, 49, 'AX', 1, 'equal', 0, '折算序號', 0)
		examin(df, 50, 'AY', 2, 'lower', 0, '分層分攤率', 0)
		examin(df, 51, 'AZ', 1, 'equal', 1, '稅率代號', 0)
		examin(df, 52, 'BA', 1, 'equal', 0, '使用別', 0)

		# 檢查加重課稅註記與稅率是否相符
		specialrate()
		

		##HOUF120持分人檔
		df = pd.read_excel(files[readFile-1], sheet_name = 'HOUF120持分人檔', index_col = None, header = 2, skiprows = 0,
							dtype = {'稅籍編號':str, '移轉日期':str, '卡別':str, '群組':str, '持分分子':str, 
									 '持分分母':str, '課稅月數':str, '是否分單':str, '是否分管':str, 'IDN_BAN':str, 
									 '身分代號':str, '姓   名':str, '備註':str, '通 訊  地  址':str, '郵遞區號':str, 
									 '縣市鄉鎮村里':str, '鄰':str, '路街':str, '段':str, '巷':str, '弄':str, 
									 '衖':str, '號':str, '號之':str, '之':str, '之號':str, '樓':str, '樓之':str, 
									 '室':str, '特殊地址註記':str})
		print("\n開始檢查HOUF120持分人檔")

		# 檢查稅號是否存在HOUF100
		errorhou_count += existHou(df, hou)

		# 檢查稅號+卡別是否重複
		duplicate2(0, 2)

		# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(整數：0, 文字：1, 小數：2), 第2種長度(針對IDN_BAN有2種長度)
		examin(df, 0, 'A', 11, 'equal', 1, '稅籍編號', 1)
		examin(df, 1, 'B', 7, 'equal', 0, '移轉日期', 1)
		examin(df, 2, 'C', 1, 'equal', 1, '卡別', 1)
		examin(df, 3, 'D', 1, 'equal', 0, '群組', 1)
		examin(df, 4, 'E', 7, 'lower', 1, '持分分子', 0)
		examin(df, 5, 'F', 7, 'lower', 1, '持分分母', 0)
		examin(df, 6, 'G', 2, 'lower', 1, '課稅月數', 0)
		examin(df, 7, 'H', 1, 'equal', 1, '是否分單', 1)
		examin(df, 8, 'I', 1, 'equal', 0, '是否分管', 1)
		examin(df, 9, 'J', 10, 'lower', 1, 'IDN_BAN', 1, length2 = 8)
		examin(df, 10, 'K', 1, 'equal', 0, '身分代號', 1)
		examin(df, 11, 'L', 26, 'lower', 1, '姓名', 1)
		examin(df, 12, 'M', 26, 'lower', 0, '備註', 1)
		examin(df, 13, 'N', 54, 'lower', 1, '通訊地址', 1)

		##HOUF151土地標示檔

		df = pd.read_excel(files[readFile-1], sheet_name = 'HOUF151土地標示檔', index_col = None, header = 2, skiprows = 0,
							dtype = {'稅籍編號':str, '土地標示':str})
		print("\n開始檢查HOUF151土地標示檔")

		# 檢查稅號是否存在HOUF100
		errorhou_count += existHou(df, hou)

		# 檢查稅號+土地標示是否重複
		duplicate2(0, 1)

		# 參數：欄位, 欄位名稱, 字串長度, 是否可小於字串長度(equal,lower), 是否必填(是：1, 否：0), 欄位中文, 資料型態(整數：0, 文字：1, 小數：2), 第2種長度(針對IDN_BAN有2種長度)
		examin(df, 0, 'A', 11, 'equal', 0, '稅籍編號', 1)
		examin(df, 1, 'B', 14, 'equal', 0, '土地標示', 1)
		
		
		# 檢查面積
		examArea(files[readFile-1])

except:
	traceback.print_exc()

#print(df.info())
#print(df['稅籍編號'].head())
#print(df.稅籍編號)
#print(df.shape[1])
#print(df[0])
