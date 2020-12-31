from pandas import read_excel, notnull, isnull
from os import path, remove, listdir, getcwd, _exit, system
import traceback
import re


def writeText(file_name, row_num, hou_losn, card, df_area, data_area, error_seq):
	# 存檔路徑
	file_name = str(file_name).split(".")[0]
	sFilepath = "./%s.txt" %file_name
	
	# 如果檔案已存在就刪除
	if path.isfile(sFilepath) and error_seq == 1: 
		remove(sFilepath)
	
	# 開啟要寫入的檔案
	f = open(sFilepath, 'a+', encoding = 'ANSI', newline ='')
	f.write("%d %s%d面積不同! 稅號%s, 卡序%s, 課稅主檔面積:%s, 公共設施檔面積:%s\n" %(error_seq, 'U', row_num, hou_losn, card, df_area, data_area))
	
	# 關閉檔案
	f.close()

def listFiles(filetype_ = "xls", nowPath = getcwd(), massage = None):
	files = listdir(nowPath)
	qualifyFiles = []

	if (massage != None):
		print(massage)
	try:
		for f in files:
			# 要給完整路徑才能正確讀取
			allPath = path.join(nowPath,f)
			if (f[f.find(".")+1:] == filetype_):
				qualifyFiles.append(f)
		if len(qualifyFiles) > 0:
			for compareF, i in zip(qualifyFiles, range(1,len(qualifyFiles)+1)):
				print("(%d)%s" %(i, compareF))
		else:
			print("沒有檔案!請將檔案與執行檔放在同一個資料夾")
			s = input("press enter to exit...")
			_exit(0)
	except:
		pass
		
	return qualifyFiles

def examArea(file_name):
	try:
		# step1:存取公共設施sheet資料
		df1 = read_excel(files[readFile-1], sheet_name = '公共設施', index_col = None, header = 6, skiprows=0)
		# data 的欄位名稱
		data = []
		data.append([])
		data[0] = ['稅號', '地址']		
		columnNum = 0
		while notnull(df1.iloc[1, 5+columnNum]):
			# 取得所有卡序名稱
			card = df1.iloc[1, 5+columnNum]
			result = re.search("[A-Z]", card).span() 	# 正則表達式 找出卡序
			# 如果樓層 <990 且 卡序不是 PYZ 就只抓卡不抓樓層
			card = card[result[0]:result[1]+1] if (int(card[0:result[0]]) < 990 and card[result[0]:result[1]+1] not in ('P', 'Y', 'Z')) else df1.iloc[1, 5+columnNum]
			data[0].append(card)
			columnNum += 1

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
		df2 = read_excel(files[readFile-1], sheet_name = 'JRCF020坐落地址檔', index_col = None, header = 2, skiprows=0,
							dtype = {'稅籍編號':str, '號':str, '號之':str, '之':str, '之號':str, '樓':str, '樓之':str})	
		# data 記錄稅號與地址
		k = 0	# k 從 header 之後開始算
		while (k < df2.shape[0]-1):
			data.append([])
			data[k+1].append(df2.iloc[k, 0])	# 第0列放欄位名稱
			if notnull(df2.iloc[k, 10]):
				addr = str(df2.iloc[k, 10]) + "號"
				if notnull(df2.iloc[k, 11]): 
					addr = addr + "之" + str(df2.iloc[k, 11])	# M號之N

			elif notnull(df2.iloc[k, 12]):
				addr = str(df2.iloc[k, 12]) + "之"
				if notnull(df2.iloc[k, 13]):
					addr = addr + str(df2.iloc[k, 13]) + "號"	# M之N號

			addr = addr + str(df2.iloc[k, 14]) + "樓" if notnull(df2.iloc[k, 14]) else addr
				
			data[k+1].append(addr)
			#print(data[k])	
			k += 1
		
		# data 存取面積 
		for sub_data in data:
			for sub_areaData in areaData:
				#print("sub_areaData", sub_areaData)
				if sub_data[1] == sub_areaData[0]:
					secondsub = 1
					while secondsub < len(sub_areaData):
						sub_data.append(sub_areaData[secondsub])
						secondsub += 1
			#print(sub_data)
		
		# step3:讀取 HOUF110課稅主檔,111加減項檔
		df3 = read_excel(files[readFile-1], sheet_name = 'HOUF110課稅主檔,111加減項檔', index_col = None, header = 2, skiprows=0,
					dtype = {'稅籍編號':str, '層次':str, '卡序':str})

		p = 0	# p 從 header 之後開始算
		error_seq = 0
		while (p < df3.shape[0]-1):
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
	except FileNotFoundError:
		system("cls")
		print("檔案已被移除!請重新選擇\n")
		return 

###主程式
attention = '''注意事項：
(A) 執行時，本執行檔與網路申報檔案(限XLS檔)放於同一資料夾
(B) 比對完面積如有錯誤資料會產生txt檔
(C) 公共設施頁籤須和網路申報檔案放於同一XLS檔案內，且頁籤名稱須命名為「公共設施」
'''
try:
	while 1:
		files = listFiles(filetype_ = "xls", massage = attention)
		readFile = int(input("請選擇檔案："))
		print("\n")
		
		# 檢查面積
		examArea(files[readFile-1])
except:
	traceback.print_exc()		