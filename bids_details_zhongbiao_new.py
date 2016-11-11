# -*- coding:utf-8 -*-

import re
from bs4 import BeautifulSoup
import xlrd
import xlwt
import string

def u(s,encoding):
	if isinstance(s,unicode):
		return s
	else:
		return unicode(s,encoding)

def mysort(a):
	for k in range(len(a)):
		(a[k][0],a[k][1]) = (a[k][1],a[k][0])
	a.sort()
	for k in range(len(a)):
		(a[k][0],a[k][1]) = (a[k][1],a[k][0])

def myreverse(a):
	b = ''
	blist = []
	if len(a) > 0:
		for i in range(len(a)-1,-1,-1):
			blist.append(a[i])
		for i in range(0,len(blist)):
			b += blist[i]
	return b

dict_book = xlrd.open_workbook('./dict_total/dict_total_zhongbiao.xlsx')
dict_sheet = dict_book.sheet_by_name('dict_total')
dict_buyer_sheet = dict_book.sheet_by_name('buyer_dict')
dict_agency_sheet = dict_book.sheet_by_name('agency_dict')
nrows = dict_sheet.nrows
ncols = dict_sheet.ncols
nrows_buyer = dict_buyer_sheet.nrows
ncols_buyer = dict_buyer_sheet.ncols
nrows_agency = dict_agency_sheet.nrows
ncols_agency = dict_agency_sheet.ncols

#sheet:dict_total
buyername_dict = []#采购人
projname_dict = []#项目名称
projnum_dict = []#项目编号
projmatter_dict = []#项目内容
bidsmedia_dict = []#发布日期及媒体
situationill_dict = []#情况说明
bidsopeningdata_dict = []#开标时间
bidsopeningloc_dict = []#开标地点
bidsopeningmainer_dict = []#评审委员会负责人
bidsopeningmember_dict = []#评审委员会成员
reviewcomment_dict = []#评审意见
bidsresult_dict = []#中标信息
bidsman_dict = []#中标人
bidsprice_dict = []#中标金额
eqptdetails_dict = []#设备明细
bidsmanloc_dict = []#中标人地址
bidsdetails_dict = []#主要中标信息
agencyname_dict = []#代理机构名称
agencyloc_dict = []#代理机构地址
procmethod_dict = []#采购方式
ctrlprice_dict = []#招标控制价
buyermancontact_dict = []#采购人联系人
agencycontact_dict = []#代理机构联系人
supervisioncontact_dict = []#监督机构联系人
stopword_dict = []

myorder_dict = [buyername_dict,projname_dict,projnum_dict,projmatter_dict,bidsmedia_dict,\
situationill_dict,\
bidsopeningdata_dict,bidsopeningloc_dict,bidsopeningmainer_dict,bidsopeningmember_dict,\
reviewcomment_dict,\
bidsresult_dict,bidsman_dict,bidsprice_dict,eqptdetails_dict,bidsmanloc_dict,bidsdetails_dict,\
agencyname_dict,agencyloc_dict,\
procmethod_dict,ctrlprice_dict,buyermancontact_dict,agencycontact_dict,\
supervisioncontact_dict,stopword_dict]

for j in range(0,ncols):
	for i in range(1,nrows):
		if len(dict_sheet.cell_value(i,j)) != 0:
			myorder_dict[j].append(dict_sheet.cell_value(i,j))
		else:
			break

#sheet:buyer
buyerunitname_dict = []
buyermanname_dict = []
buyertel_dict = []
buyerloc_dict = []

myorder_buyer_dict = [buyerunitname_dict,buyermanname_dict,\
buyertel_dict,buyerloc_dict]

for j in range(0,ncols_buyer):
	for i in range(1,nrows_buyer):
		if len(dict_buyer_sheet.cell_value(i,j)) != 0:
			myorder_buyer_dict[j].append(dict_buyer_sheet.cell_value(i,j))
		else:
			break

#sheet:agency
agencyunitname_dict = []
agencymanname_dict = []
agencyunittel_dict = []
agencyunitloc_dict = []

myorder_agency_dict = [agencyunitname_dict,agencymanname_dict,\
agencyunittel_dict,agencyunitloc_dict]

for j in range(0,ncols_agency):
	for i in range(0,nrows_agency):
		if len(dict_agency_sheet.cell_value(i,j)) != 0:
			myorder_agency_dict[j].append(dict_agency_sheet.cell_value(i,j))
		else:
			break

class DEF_CONTENTS:
	def __init__(self,path,data):
		self.path = path
		self.data = data

	def readfile(self,path):
		myfile = open(path,'r')
		mycontents = []
		for line_temp in myfile.readlines():
			line = line_temp.strip('\n')
			if len(line) > 1:
				mycontents.append(line)

		return mycontents

	def prepro(self,mycontents):
		mycontents_new = []
		for i in range(0,len(mycontents)):
			flag = 0
			for k in range(1,9):
				if i-k >= 0:
					if mycontents[i] in mycontents[i-k]:
						flag += 1
			if flag == 0:
				mycontents_new.append(mycontents[i])
		return mycontents_new

	def getxlsarticle(self,contents,data):
		mycontents = []
		mycontents.append('article')
		mycontents.append(data)

		mydict_num = ['一、','二、','三、','四、','五、','六、','七、','八、','九、','有异议']
		flag = 0
		for index in range(0,len(contents)):
			if flag == 0:
				if mydict_num[0] in contents[index]:
					start = index
					flag = 1
			if flag > 0:
				if mydict_num[-1] in contents[index]:
					end = index
					break
		contents_main = contents[start:end]

		mystring = ''
		for item in contents_main:
			temp = u(item,'utf-8')
			mystring += temp
		#print mystring

		buyername_flag = 0
		for buyername_dict_temp in buyername_dict:
			if buyername_dict_temp in mystring:
				k = len(buyername_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == buyername_dict_temp:
						buyer_name = buyername_dict_temp
						buyername_index = i + k - 1
						break
				buyername_flag += 1
				break
			if buyername_flag == 0:
				buyer_name = 'none'
				buyername_index = len(mystring)
		projname_flag = 0
		for projname_dict_temp in projname_dict:
			if projname_dict_temp in mystring:
				k = len(projname_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == projname_dict_temp:
						proj_name = projname_dict_temp
						projname_index = i + k - 1
						break
				projname_flag += 1
				break
			if projname_flag == 0:
				proj_name = 'none'
				projname_index = len(mystring)
		projnum_flag = 0
		for projnum_dict_temp in projnum_dict:
			if projnum_dict_temp in mystring:
				k = len(projnum_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == projnum_dict_temp:
						proj_num = projnum_dict_temp
						projnum_index = i + k - 1
						break
				projnum_flag += 1
				break
			if projnum_flag == 0:
				proj_num = 'none'
				projnum_index = len(mystring)
		projmatter_flag = 0
		for projmatter_dict_temp in projmatter_dict:
			if projmatter_dict_temp in mystring:
				k = len(projmatter_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == projmatter_dict_temp:
						proj_matter = projmatter_dict_temp
						projmatter_index = i + k - 1
						break
				projmatter_flag += 1
				break
			if projmatter_flag == 0:
				proj_matter = 'none'
				projmatter_index = len(mystring)
		bidsmedia_flag = 0
		for bidsmedia_dict_temp in bidsmedia_dict:
			if bidsmedia_dict_temp in mystring:
				k = len(bidsmedia_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsmedia_dict_temp:
						bids_media = bidsmedia_dict_temp
						bidsmedia_index = i + k - 1
						break
				bidsmedia_flag += 1
				break
			if bidsmedia_flag == 0:
				bids_media = 'none'
				bidsmedia_index = len(mystring)
		bidsopeningdata_flag = 0
		for bidsopeningdata_dict_temp in bidsopeningdata_dict:
			if bidsopeningdata_dict_temp in mystring:
				k = len(bidsopeningdata_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsopeningdata_dict_temp:
						bidsopening_data = bidsopeningdata_dict_temp
						bidsopeningdata_index = i + k - 1
						break
				bidsopeningdata_flag += 1
				break
			if bidsopeningdata_flag == 0:
				bidsopening_data = 'none'
				bidsopeningdata_index = len(mystring)
		bidsopeningloc_flag = 0
		for bidsopeningloc_dict_temp in bidsopeningloc_dict:
			if bidsopeningloc_dict_temp in mystring:
				k = len(bidsopeningloc_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsopeningloc_dict_temp:
						bidsopening_loc = bidsopeningloc_dict_temp
						bidsopeningloc_index = i + k - 1
						break
				bidsopeningloc_flag += 1
				break
			if bidsopeningloc_flag == 0:
				bidsopening_loc = 'none'
				bidsopeningloc_index = len(mystring)
		bidsopeningmainer_flag = 0
		for bidsopeningmainer_dict_temp in bidsopeningmainer_dict:
			if bidsopeningmainer_dict_temp in mystring:
				k = len(bidsopeningmainer_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsopeningmainer_dict_temp:
						bidsopening_mainer = bidsopeningmainer_dict_temp
						bidsopeningmainer_index = i + k - 1
						break
				bidsopeningmainer_flag += 1
				break
			if bidsopeningmainer_flag == 0:
				bidsopening_mainer = 'none'
				bidsopeningmainer_index = len(mystring)
		bidsopeningmember_flag = 0
		for bidsopeningmember_dict_temp in bidsopeningmember_dict:
			if bidsopeningmember_dict_temp in mystring:
				k = len(bidsopeningmember_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsopeningmember_dict_temp:
						bidsopening_member = bidsopeningmember_dict_temp
						bidsopeningmember_index = i + k - 1
						break
				bidsopeningmember_flag += 1
				break
			if bidsopeningmember_flag == 0:
				bidsopening_member = 'none'
				bidsopeningmember_index = len(mystring)
		reviewcomment_flag = 0
		for reviewcomment_dict_temp in reviewcomment_dict:
			if reviewcomment_dict_temp in mystring:
				k = len(reviewcomment_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == reviewcomment_dict_temp:
						review_comment = reviewcomment_dict_temp
						reviewcomment_index = i + k - 1
						break
				reviewcomment_flag += 1
				break
			if reviewcomment_flag == 0:
				review_comment = 'none'
				reviewcomment_index = len(mystring)
		bidsresult_flag = 0
		for bidsresult_dict_temp in bidsresult_dict:
			if bidsresult_dict_temp in mystring:
				k = len(bidsresult_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsresult_dict_temp:
						bids_result = bidsresult_dict_temp
						bidsresult_index = i + k - 1
						break
				bidsresult_flag += 1
				break
			if bidsresult_flag == 0:
				bids_result = 'none'
				bidsresult_index = len(mystring)
		bidsman_flag = 0
		for bidsman_dict_temp in bidsman_dict:
			if bidsman_dict_temp in mystring:
				k = len(bidsman_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsman_dict_temp:
						bids_man = bidsman_dict_temp
						bidsman_index = i + k - 1
						break
				bidsman_flag += 1
				break
			if bidsman_flag == 0:
				bids_man = 'none'
				bidsman_index = len(mystring)
		bidsprice_flag = 0
		for bidsprice_dict_temp in bidsprice_dict:
			if bidsprice_dict_temp in mystring:
				k = len(bidsprice_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsprice_dict_temp:
						bids_price = bidsprice_dict_temp
						bidsprice_index = i + k - 1
						break
				bidsprice_flag += 1
				break
			if bidsprice_flag == 0:
				bids_price = 'none'
				bidsprice_index = len(mystring)
		bidsmanloc_flag = 0
		for bidsmanloc_dict_temp in bidsmanloc_dict:
			if bidsmanloc_dict_temp in mystring:
				k = len(bidsmanloc_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsmanloc_dict_temp:
						bidsman_loc = bidsmanloc_dict_temp
						bidsmanloc_index = i + k - 1
						break
				bidsmanloc_flag += 1
				break
			if bidsmanloc_flag == 0:
				bidsman_loc = 'none'
				bidsmanloc_index = len(mystring)
		bidsdetails_flag = 0
		for bidsdetails_dict_temp in bidsdetails_dict:
			if bidsdetails_dict_temp in mystring:
				k = len(bidsdetails_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == bidsdetails_dict_temp:
						bids_details = bidsdetails_dict_temp
						bidsdetails_index = i + k - 1
						break
				bidsdetails_flag += 1
				break
			if bidsdetails_flag == 0:
				bids_details = 'none'
				bidsdetails_index = len(mystring)	
		agencyname_flag = 0
		for agencyname_dict_temp in agencyname_dict:
			if agencyname_dict_temp in mystring:
				k = len(agencyname_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == agencyname_dict_temp:
						agency_name = agencyname_dict_temp
						agencyname_index = i + k - 1
						break
				agencyname_flag += 1
				break
			if agencyname_flag == 0:
				agency_name = 'none'
				agencyname_index = len(mystring)
		agencyloc_flag = 0
		for agencyloc_dict_temp in agencyloc_dict:
			if agencyloc_dict_temp in mystring:
				k = len(agencyloc_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == agencyloc_dict_temp:
						agency_loc = agencyloc_dict_temp
						agencyloc_index = i + k - 1
						break
				agencyloc_flag += 1
				break
			if agencyloc_flag == 0:
				agency_loc = 'none'
				agencyloc_index = len(mystring)
		procmethod_flag = 0
		for procmethod_dict_temp in procmethod_dict:
			if procmethod_dict_temp in mystring:
				k = len(procmethod_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == procmethod_dict_temp:
						proc_method = procmethod_dict_temp
						procmethod_index = i + k - 1
						break
				procmethod_flag += 1
				break
			if procmethod_flag == 0:
				proc_method = 'none'
				procmethod_index = len(mystring)
		ctrlprice_flag = 0
		for ctrlprice_dict_temp in ctrlprice_dict:
			if ctrlprice_dict_temp in mystring:
				k = len(ctrlprice_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == ctrlprice_dict_temp:
						ctrl_price = ctrlprice_dict_temp
						ctrlprice_index = i + k - 1
						break
				ctrlprice_flag += 1
				break
			if ctrlprice_flag == 0:
				ctrl_price = 'none'
				ctrlprice_index = len(mystring)
		buyermancontact_flag = 0
		for buyermancontact_dict_temp in buyermancontact_dict:
			if buyermancontact_dict_temp in mystring:
				k = len(buyermancontact_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == buyermancontact_dict_temp:
						buyerman_contact = buyermancontact_dict_temp
						buyermancontact_index = i + k - 1
						break
				buyermancontact_flag += 1
				break
			if buyermancontact_flag == 0:
				buyerman_contact = 'none'
				buyermancontact_index = len(mystring)
		agencycontact_flag = 0
		for agencycontact_dict_temp in agencycontact_dict:
			if agencycontact_dict_temp in mystring:
				k = len(agencycontact_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == agencycontact_dict_temp:
						agency_contact = agencycontact_dict_temp
						agencycontact_index = i + k - 1
						break
				agencycontact_flag += 1
				break
			if agencycontact_flag == 0:
				agency_contact = 'none'
				agencycontact_index = len(mystring)
		supervisioncontact_flag = 0
		for supervisioncontact_dict_temp in supervisioncontact_dict:
			if supervisioncontact_dict_temp in mystring:
				k = len(supervisioncontact_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == supervisioncontact_dict_temp:
						supervision_contact = supervisioncontact_dict_temp
						supervisioncontact_index = i + k - 1
						break
				supervisioncontact_flag += 1
				break
			if supervisioncontact_flag == 0:
				supervision_contact = 'none'
				supervisioncontact_index = len(mystring)
		eqptdetails_flag = 0
		for eqptdetails_dict_temp in eqptdetails_dict:
			if eqptdetails_dict_temp in mystring:
				k = len(eqptdetails_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == eqptdetails_dict_temp:
						eqpt_details = eqptdetails_dict_temp
						eqptdetails_index = i + k - 1
						break
				eqptdetails_flag += 1
				break
			if eqptdetails_flag == 0:
				eqpt_details = 'none'
				eqptdetails_index = len(mystring)
		situationill_flag = 0
		for situationill_dict_temp in situationill_dict:
			if situationill_dict_temp in mystring:
				k = len(situationill_dict_temp)
				for i in range(0,len(mystring)):
					if mystring[i:i+k] == situationill_dict_temp:
						situation_ill = situationill_dict_temp
						situationill_index = i + k - 1
						break
				situationill_flag += 1
				break
			if situationill_flag == 0:
				situation_ill = 'none'
				situationill_index = len(mystring)
		
		index_total = [[buyer_name,buyername_index],[proj_name,projname_index],\
		[proj_num,projnum_index],[proj_matter,projmatter_index],\
		[bids_media,bidsmedia_index],[bidsopening_data,bidsopeningdata_index],\
		[bidsopening_loc,bidsopeningloc_index],[bidsopening_mainer,bidsopeningmainer_index],\
		[bidsopening_member,bidsopeningmember_index],[bids_result,bidsresult_index],\
		[bids_man,bidsman_index],[review_comment,reviewcomment_index],\
		[bids_price,bidsprice_index],[bidsman_loc,bidsmanloc_index],\
		[bids_details,bidsdetails_index],[buyerman_contact,buyermancontact_index],\
		[eqpt_details,eqptdetails_index],[situation_ill,situationill_index],\
		[proc_method,procmethod_index],[ctrl_price,ctrlprice_index],\
		[agency_contact,agencycontact_index],[supervision_contact,supervisioncontact_index]]
		mysort(index_total)

		mycontents_new = []
		for i in range(0,len(index_total)):
			string_temp = ''
			if i < len(index_total)-1:
				if index_total[i+1][1] < len(mystring)-1:
					string_temp = mystring[index_total[i][1]+1:index_total[i+1][1]-len(index_total[i+1][0])+1]
				elif index_total[i+1][1] >= len(mystring)-1:
					string_temp = mystring[index_total[i][1]+1:index_total[i+1][1]+1]
				string_my = index_total[i][0] + string_temp
				mycontents_new.append(string_my)
			else:
				string_temp = mystring[index_total[i][1]+1:len(mystring)]
				string_my = index_total[i][0] + string_temp
				mycontents_new.append(string_my)

		wordfig_dict = [u'一',u'二',u'三',u'四',u'五',u'六',u'七',u'八',u'九',u'十']
		arabicfig_dict = [u'1',u'2',u'3',u'4',u'5',u'6',u'7',u'8',u'9',u'0']
		stopword_dict = [u'中标信息第一包',u'评审信息',u'中标信息',u'专家信息',u'中标结果',\
		u'联系事项',u'成交信息',u'中标（成交）信息',u'成交结果',u'评标信息',u'开标信息',\
		u'谈判信息',u'项目内容及招标控制价',u'评审结果']
		mycontents_extra2 = []
		for item in mycontents_new:
			if len(item) > 0:
				if item[-1] == u'、' or item[-1] == u'.':
					item = item[:-1]
					mycontents_extra2.append(item)
				else:
					mycontents_extra2.append(item)

		mycontents_end1 = []
		for item in mycontents_extra2:
			main_flag = 0
			item_flag = 0
			for arabicfig_dict_temp in arabicfig_dict:
				if item[-2] == arabicfig_dict_temp:
					item_flag += 1
					break
			if item_flag == 0:
				for arabicfig_dict_temp in arabicfig_dict:
					if item[-1] == arabicfig_dict_temp:
						item = item[:-1]
						mycontents_end1.append(item)
						main_flag += 1
						break
				if main_flag == 0:
					mycontents_end1.append(item)
			else:
				mycontents_end1.append(item)

		mycontents_end1 = map(string.strip,mycontents_end1)
		mycontents_end2 = []
		for item in mycontents_end1:
			item_temp = item.rstrip()
			mycontents_end2.append(item_temp)

		mycontents_extra1 = []
		for item in mycontents_end2:
			if len(item) > 0:
				if item[-1] == u'：' or item[-1] == u':':
					item = item[:-1]
					mycontents_extra1.append(item)
				else:
					mycontents_extra1.append(item)

		mycontents_end3 = []
		for item in mycontents_extra1:
			if len(item) > 0:
				item_flag = 0
				for stopword_dict_temp in stopword_dict:
					k = len(stopword_dict_temp)
					if item[-k:] == stopword_dict_temp:
						item = item[:-k]
						mycontents_end3.append(item)
						item_flag += 1
						break
				if item_flag == 0:
					mycontents_end3.append(item)

		mycontents_end4 = []
		for item in mycontents_end3:
			if len(item) > 0:
				if item[-1] == u'、':
					item = item[:-1]
					mycontents_end4.append(item)
				else:
					mycontents_end4.append(item)

		mycontents_end5 = []
		for item in mycontents_end4:
			if len(item) > 0:
				item_flag = 0
				for wordfig_dict_temp in wordfig_dict:
					if item[-1] == wordfig_dict_temp:
						item = item[:-1]
						mycontents_end5.append(item)
						item_flag += 1
						break
				if item_flag == 0:
					mycontents_end5.append(item)

		for item in mycontents_end5:
			mycontents.append(item)

		return mycontents

	def indexmaker(self,mycontents):
		
		myorder = [buyername_dict,projname_dict,projnum_dict,projmatter_dict,bidsmedia_dict,\
		situationill_dict,\
		bidsopeningdata_dict,bidsopeningloc_dict,bidsopeningmainer_dict,bidsopeningmember_dict,\
		reviewcomment_dict,\
		bidsresult_dict,bidsman_dict,bidsprice_dict,eqptdetails_dict,bidsmanloc_dict,bidsdetails_dict,\
		procmethod_dict,ctrlprice_dict,buyermancontact_dict,agencycontact_dict,\
		supervisioncontact_dict]
		
		mycontents_end = []
		mycontents_end.append(mycontents[0])
		mycontents_end.append(mycontents[1])
		for order_big_temp in myorder:
			big_flag = 0
			for order_temp in order_big_temp:
				k = len(order_temp)
				item_flag = 0
				for item in mycontents[2:]:
					if item[:k] == order_temp:
						mycontents_end.append(item)
						item_flag += 1
						big_flag += 1
						break
				if item_flag > 0:
					break
			if big_flag == 0:
				mycontents_end.append(u'')

		return mycontents_end

	def bidsopeningmainer_maker(self,mycontents):
		mycontents_end = []
		for i in range(0,10):
			mycontents_end.append(mycontents[i])

		item = mycontents[10]
		end_temp = []
		if u'评审委员会' in item:
			split_temp = item.split(u'评审委员会',2)
			if len(split_temp) == 3:
				end_temp.append(split_temp[1])
				end_temp.append(split_temp[2])
			elif len(split_temp) == 2:
				end_temp.append(split_temp[1])
			else:
				end_temp.append(split_temp[0])
		elif u'评标委员会' in item:
			split_temp = item.split(u'评标委员会',2)
			if len(split_temp) == 3:
				end_temp.append(split_temp[1])
				end_temp.append(split_temp[2])
			elif len(split_temp) == 2:
				end_temp.append(split_temp[1])
			else:
				end_temp.append(split_temp[0])
		else:
			end_temp.append(item)
		for item in end_temp:
			if len(item) > 0:
				if item[-1] == u'.':
					item = item[:-1]
			if len(item) > 0:
				if item[-1] == u'4':
					item = item[:-1]
			mycontents_end.append(item)

		if mycontents[11] != u'评审委员会' and mycontents[11] != u'评标委员会':
			mycontents_end.append(mycontents[11])
		for i in range(12,len(mycontents)):
			mycontents_end.append(mycontents[i])

		return mycontents_end

	def bidsopeningmember_maker(self,mycontents):
		mycontents_end = []
		for i in range(0,11):
			mycontents_end.append(mycontents[i])

		if u'：' in mycontents[11]:
			split_temp = mycontents[11].split(u'：',1)
			split_my = split_temp[1]
		elif u':' in mycontents[11]:
			split_temp = mycontents[11].split(u':',1)
			split_my = split_temp[1]
		else:
			flag_temp = 0
			for dict_temp in bidsopeningmember_dict:
				if dict_temp in mycontents[11]:
					split_temp = mycontents[11].split(dict_temp,1)
					split_my = split_temp[1]
					flag_temp += 1
					break
			if flag_temp == 0:
				split_my = mycontents[11]

		temp_flag = 0
		string_temp = u''
		for dict_temp in bidsresult_dict:
			if dict_temp in split_my:
				split_temp = split_my.split(dict_temp)
				string_temp = dict_temp + split_temp[1]
				end_temp = split_temp[0]
				temp_flag += 1
				break
		if temp_flag == 0:
			end_temp = split_my
		if u'、' in end_temp:
			end_temp = end_temp.replace(u' ',u'')
		newstring = end_temp.replace(u'、',u' ')

		mycontents_end.append(newstring)

		mycontents_end.append(mycontents[12])

		if u'公示如下' in mycontents[13]:
			if len(string_temp) > 0:
				mycontents_end.append(string_temp)
			else:
				mycontents_end.append(mycontents[13])
		else:
			mycontents_end.append(mycontents[13])

		for i in range(14,len(mycontents)):
			mycontents_end.append(mycontents[i])

		return mycontents_end

	def buyermancontact_maker(self,mycontents):
		mycontents_end = []
		for i in range(0,21):
			mycontents_end.append(mycontents[i])

		if len(mycontents[21]) > 0:
			buyer = mycontents[21]

			buyerunitname_flag = 0
			for buyerunitname_dict_temp in buyerunitname_dict:
				if buyerunitname_dict_temp in buyer:
					k = len(buyerunitname_dict_temp)
					for i in range(0,len(buyer)):
						if buyer[i:i+k] == buyerunitname_dict_temp:
							buyerunit_name = buyerunitname_dict_temp
							buyerunitname_index = i + k - 1
							break
					buyerunitname_flag += 1
					break
				if buyerunitname_flag == 0:
					buyerunit_name = 'enon'
					buyerunitname_index = len(buyer)
			buyermanname_flag = 0
			for buyermanname_dict_temp in buyermanname_dict:
				if buyermanname_dict_temp in buyer:
					k = len(buyermanname_dict_temp)
					for i in range(0,len(buyer)):
						if buyer[i:i+k] == buyermanname_dict_temp:
							buyerman_name = buyermanname_dict_temp
							buyermanname_index = i + k - 1
							break
					buyermanname_flag += 1
					break
				if buyermanname_flag == 0:
					buyerman_name = 'enon'
					buyermanname_index = len(buyer)
			buyertel_flag = 0
			for buyertel_dict_temp in buyertel_dict:
				if buyertel_dict_temp in buyer:
					k = len(buyertel_dict_temp)
					for i in range(0,len(buyer)):
						if buyer[i:i+k] == buyertel_dict_temp:
							buyer_tel = buyertel_dict_temp
							buyertel_index = i + k - 1
							break
					buyertel_flag += 1
					break
				if buyertel_flag == 0:
					buyer_tel = 'enon'
					buyertel_index = len(buyer)
			buyerloc_flag = 0
			for buyerloc_dict_temp in buyerloc_dict:
				if buyerloc_dict_temp in buyer:
					k = len(buyerloc_dict_temp)
					for i in range(0,len(buyer)):
						if buyer[i:i+k] == buyerloc_dict_temp:
							buyer_loc = buyerloc_dict_temp
							buyerloc_index = i + k - 1
							break
					buyerloc_flag += 1
					break
				if buyerloc_flag == 0:
					buyer_loc = 'enon'
					buyerloc_index = len(buyer)

			index_total = [[buyerunit_name,buyerunitname_index],[buyerman_name,buyermanname_index],\
			[buyer_tel,buyertel_index],[buyer_loc,buyerloc_index]]
			mysort(index_total)
			
			mycontents_new = []
			for i in range(0,len(index_total)):
				string_temp = ''
				if i < len(index_total)-1:
					if index_total[i+1][1] < len(buyer)-1:
						string_temp = buyer[index_total[i][1]+1:index_total[i+1][1]-len(index_total[i+1][0])+1]
					elif index_total[i+1][1] >= len(buyer)-1:
						string_temp = buyer[index_total[i][1]+1:index_total[i+1][1]+1]
					string_my = index_total[i][0] + string_temp
					mycontents_new.append(string_my)
				else:
					string_temp = buyer[index_total[i][1]+1:len(buyer)]
					string_my = index_total[i][0] + string_temp
					mycontents_new.append(string_my)

			myorder = [buyerunitname_dict,buyermanname_dict,buyertel_dict,buyerloc_dict]
			mycontents_extra1 = []
			for order_big_temp in myorder:
				big_flag = 0
				for order_temp in order_big_temp:
					k = len(order_temp)
					item_flag = 0
					for item in mycontents_new:
						if item[:k] == order_temp:
							mycontents_extra1.append(item)
							item_flag += 1
							big_flag += 1
							break
					if item_flag > 0:
						break
				if big_flag == 0:
					mycontents_extra1.append(u'')

			split_temp = mycontents_extra1[2].split(u'联系电话')
			if len(split_temp) == 3:
				mycontents_extra1[2] = u'联系电话' + split_temp[2]
				extra_tel = u'联系电话' + split_temp[1]
			else:
				extra_tel = u''

			if u'1.2' in mycontents_extra1[1]:
				mycontents_extra1[1] = mycontents_extra1[1].replace(u'1.2',u'')
			if u'1.3' in mycontents_extra1[2]:
				mycontents_extra1[2] = mycontents_extra1[2].replace(u'1.3',u'')

			for dict_temp in buyerunitname_dict:
				if mycontents_extra1[0] == dict_temp:
					mycontents_extra1[0] = u''
					break
		else:
			mycontents_extra1 = [u'',u'',u'',u'']
			extra_tel = u''

		for item in mycontents_extra1:
			mycontents_end.append(item)

		mycontents_end.append(mycontents[22])
		mycontents_end.append(mycontents[23])
		return mycontents_end,extra_tel

	def agencycontact_maker(self,mycontents,extra_tel):
		mycontents_end = []
		for i in range(0,25):
			mycontents_end.append(mycontents[i])

		
		if len(mycontents[25]) > 0:
			agency = mycontents[25]
			
			agencyunitname_flag = 0
			for agencyunitname_dict_temp in agencyunitname_dict:
				if agencyunitname_dict_temp in agency:
					k = len(agencyunitname_dict_temp)
					for i in range(0,len(agency)):
						if agency[i:i+k] == agencyunitname_dict_temp:
							agencyunit_name = agencyunitname_dict_temp
							agencyunitname_index = i + k - 1
							break
					agencyunitname_flag += 1
					break
				if agencyunitname_flag == 0:
					agencyunit_name = 'none'
					agencyunitname_index = len(agency)
			agencymanname_flag = 0
			for agencymanname_dict_temp in agencymanname_dict:
				if agencymanname_dict_temp in agency:
					k = len(agencymanname_dict_temp)
					for i in range(0,len(agency)):
						if agency[i:i+k] == agencymanname_dict_temp:
							agencyman_name = agencymanname_dict_temp
							agencymanname_index = i + k - 1
							break
					agencymanname_flag += 1
					break
				if agencymanname_flag == 0:
					agencyman_name = 'none'
					agencymanname_index = len(agency)
			agencyunittel_flag = 0
			for agencyunittel_dict_temp in agencyunittel_dict:
				if agencyunittel_dict_temp in agency:
					k = len(agencyunittel_dict_temp)
					for i in range(0,len(agency)):
						if agency[i:i+k] == agencyunittel_dict_temp:
							agencyunit_tel = agencyunittel_dict_temp
							agencyunittel_index = i + k - 1
							break
					agencyunittel_flag += 0
					break
				if agencyunittel_flag == 0:
					agencyunit_tel = 'none'
					agencyunittel_index = len(agency)
			agencyunitloc_flag = 0
			for agencyunitloc_dict_temp in agencyunitloc_dict:
				if agencyunitloc_dict_temp in agency:
					k = len(agencyunitloc_dict_temp)
					for i in range(0,len(agency)):
						if agency[i:i+k] == agencyunitloc_dict_temp:
							agencyunit_loc = agencyunitloc_dict_temp
							agencyunitloc_index = i + k - 1
							break
					agencyunitloc_flag += 1
					break
				if agencyunitloc_flag == 0:
					agencyunit_loc = 'none'
					agencyunitloc_index = len(agency)

			index_total = [[agencyunit_name,agencyunitname_index],[agencyman_name,agencymanname_index],\
			[agencyunit_tel,agencyunittel_index],[agencyunit_loc,agencyunitloc_index]]
			mysort(index_total)

			mycontents_new = []
			for i in range(0,len(index_total)):
				string_temp = ''
				if i < len(index_total)-1:
					if index_total[i+1][1] < len(agency)-1:
						string_temp = agency[index_total[i][1]+1:index_total[i+1][1]-len(index_total[i+1][0])+1]
					elif index_total[i+1][1] >= len(agency)-1:
						string_temp = agency[index_total[i][1]+1:index_total[i+1][1]+1]
					string_my = index_total[i][0] + string_temp
					mycontents_new.append(string_my)
				else:
					string_temp = agency[index_total[i][1]+1:len(agency)]
					string_my = index_total[i][0] + string_temp
					mycontents_new.append(string_my)

			myorder = [agencyunitname_dict,agencymanname_dict,agencyunittel_dict,agencyunitloc_dict]
			mycontents_extra1 = []
			for order_big_temp in myorder:
				big_flag = 0
				for order_temp in order_big_temp:
					k = len(order_temp)
					item_flag = 0
					for item in mycontents_new:
						if item[:k] == order_temp:
							mycontents_extra1.append(item)
							item_flag += 1
							big_flag += 1
							break
					if item_flag > 0:
						break
				if big_flag == 0:
					mycontents_extra1.append(u'')

			if len(mycontents_extra1[2]) == 0 and len(extra_tel) > 0:
				mycontents_extra1[2] = extra_tel

			if u'2.2' in mycontents_extra1[1]:
				mycontents_extra1[1] = mycontents_extra1[1].replace(u'2.2',u'')
			if u'2.3' in mycontents_extra1[2]:
				mycontents_extra1[2] = mycontents_extra1[2].replace(u'2.3',u'')

			for dict_temp in agencyunitname_dict:
				if mycontents_extra1[0] == dict_temp:
					mycontents_extra1[0] = u''
					break
		else:
			mycontents_extra1 = [u'',u'',u'',u'']

		for item in mycontents_extra1:
			mycontents_end.append(item)

		mycontents_end.append(mycontents[26])
		return mycontents_end

	def getxlstable(self,contents,data):
		mycontents = []
		mycontents.append('table')
		mycontents.append(data)
		mydict = ['采购单位：','项目名称：','项目编号：','发布日期：','竞价开始时间：','采购预算：','竞价规则：','报名条件：','交货方式：','补充说明：','联系方式：','成交结果：','采购商品列表','有异议']
		for i in range(0,len(contents)):
			for j in range(0,len(mydict)-1):
				if mydict[j] in contents[i]:
					string_temp = ''
					split_temp = contents[i].split(mydict[j],1)
					mysplit = split_temp[1].strip()
					string_temp += mysplit
					for k in range(1,len(contents)-1-i):
						if mydict[j+1] not in contents[i+k]:
							string_temp += contents[i+k]
						else:
							break
					string_temp = string_temp.split()
					mycontents.append(string_temp)
		mycontents_end = []
		for item in mycontents:
			flag = 0
			if len(item) > 0:
				if item[0] == '：':
					item = item[1:]
					mycontents_end.append(item)
					flag += 1
			if flag == 0:
				mycontents_end.append(item)

		return mycontents_end

	def start(self):
		path = self.path
		data = self.data
		#【抓取信息读取、预处理】
		mycontents = self.readfile(path)
		mycontents_new = self.prepro(mycontents)

		#【结构化】
		flag = 0
		for item in mycontents_new:
			if '有异议' in item:
				xlscontents_temp1 = self.getxlsarticle(mycontents_new,data)
				xlscontents_temp2 = self.indexmaker(xlscontents_temp1)
				xlscontents_temp3 = self.bidsopeningmainer_maker(xlscontents_temp2)
				#xlscontents = self.bidsopeningmember_maker(xlscontents_temp3)
				xlscontents_temp4 = self.bidsopeningmember_maker(xlscontents_temp3)
				xlscontents_temp5,extra_tel = self.buyermancontact_maker(xlscontents_temp4)
				xlscontents = self.agencycontact_maker(xlscontents_temp5,extra_tel)
				flag += 1
				break
		if flag == 0:
			xlscontents = self.getxlstable(mycontents_new,data)
			flag += 1

		return xlscontents

menu_book = xlrd.open_workbook('.\menu_links\menu_links_zhongbiaogonggao_qj.xls')
menu_sheet = menu_book.sheet_by_name('menu')
nrows = menu_sheet.nrows
anno_datas = []
index_total = []
class_total = []
for i in range(0,nrows):
	anno_datas.append(menu_sheet.cell_value(i,1))
	index_total.append(menu_sheet.cell_value(i,3))
	和.append(menu_sheet.cell_value(i,4))

newworkbook = xlwt.Workbook(encoding = 'utf-8')
newworksheet_article = newworkbook.add_sheet('article')
newworksheet_table = newworkbook.add_sheet('table')

length = nrows + 1
#length = 2
#i start from 1
for i in range(1,length):
	#if class_total[i-1] == 'new' or class_total[i-1] == 'old':
	if class_total[i-1] == 'new':
		mypath = './txt/zhongbiaogonggao_qj/' + anno_datas[i-1] + '_' + str(int(index_total[i-1])) + '.txt'
		data = anno_datas[i-1]
		bid_contents = DEF_CONTENTS(mypath,data)
		mycontents = bid_contents.start()
		mycontents.insert(0,str(int(index_total[i-1])))

		col_article = 0
		col_table = 0
		if mycontents[1] == u'article':
			flag_my = 1
		elif mycontents[1] == u'table':
			flag_my = 0
		for item in mycontents:
			if flag_my == 1:
				newworksheet_article.write(i-1,col_article,label = item)
				col_article += 1
			elif flag_my == 0:
				newworksheet_table.write(i-1,col_table,label = item)
				col_table += 1

		print u'完成【' + str(int(index_total[i-1])) + u'】'

newworkbook.save('bids_test.xls')