#!/usr/bin/python

import os

import sys
reload(sys)
sys.setdefaultencoding('utf8')

import sqlite3
import re
import types
import xlrd
import datetime, time
from Utils import *
from pyExcelerator import *
import pyExcelerator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from UserInfo import *
import pyExcelerator
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from time import strftime, localtime
from operator import itemgetter

prj_team_sum_begin = [[0 for col in range(50)]for row in range(5)]
prj_team_sum_end = [[0 for col in range(50)]for row in range(5)]

class dataCollection(UserInfo):
	def __init__(self, name):
		self.toList = []
		self.ccList = []
		self.pieChart = {}
		self.lineChart = {}
		self.titleList = ['DATE']
		self.addHtmlTopAnnotate = False
		self.projectName = ''
		self.platformName = ''

	def addTo(self, mailAddr):
		self.toList.append(mailAddr)

	def addCc(self, mailAddr):
		self.ccList.append(mailAddr)

	def getDaulftBorder(self):
		borders = pyExcelerator.Borders()
		borders.left = 1
		borders.right = 1
		borders.top = 1
		borders.bottom = 1
		return borders

	def getTitleStyle(self):
		fnt = pyExcelerator.Font()
		fnt.colour_index = 0
		fnt.bold = True
		al = pyExcelerator.Alignment()
		al.horz = pyExcelerator.Alignment.HORZ_CENTER
		al.vert = pyExcelerator.Alignment.VERT_CENTER
		pattern = pyExcelerator.Pattern()
		pattern.pattern=2
		pattern.pattern_fore_colour = 30
		pattern.pattern_back_colour = 95
		style = pyExcelerator.XFStyle()
		style.font = fnt
		style.borders = self.getDaulftBorder()
		style.alignment = al
		style.pattern = pattern
		return style

	def getCellStyle(self):
		fnt = pyExcelerator.Font()
		fnt.colour_index = 0
		al = pyExcelerator.Alignment()
		al.horz = pyExcelerator.Alignment.HORZ_CENTER
		al.vert = pyExcelerator.Alignment.VERT_CENTER
		pattern = pyExcelerator.Pattern()
		style = pyExcelerator.XFStyle()
		style.font = fnt
		style.borders = self.getDaulftBorder()
		style.alignment = al
		style.pattern = pattern
		return style

	def exTime(self, t):
		Y,m,d = time.strptime(t, "%Y-%m-%d")[0:3]
		tp = datetime.datetime(Y,m,d)
		return tp

	def isopenstate(self, item):
		if item == 'ASSIGNED' or item == 'OPENED' or item == 'NEW' or item == 'INVESTIGATED' or item == 'RESOLVED':
			return True
		return False

	def isclosestate(self, item):
		if item == 'VERIFIED_SW' or item == 'DELIVERED' or item == 'VERIFIED' or item == 'CLOSED':
			return True
		return False

	def isblocked(self, item):
		if 'TOP' in item.upper() or 'BLOCK' in item.upper() in item.upper():
			return True
		return False

	def readSheet(self, sh, sh_days, name):
		ogCount = 0
		ogHomo = 0
		toVsw = 0
		fixed = 0
		regTotal = 0
		engBugInfor = []
		app_list = ["Alarm","ADService","AlcatelHelp","Calculator","Calendar","Camera","Camcoder","Call","CloudBackup","Cloud Backup","Compass","Contacts","Contact","CustomerVoice","Dialer","Elabel","Email","Exchange","FileManager","File Manager","FM","FMRadio","FM Radio","Fota","Gallery","JrdAssetsFileOperation","Launcher","Mms","SMS","Music","NfcSwitch","NFC Switch","Note","PGService","SalesTracker","SetupWizard","SmartCare","Smart care","SoundRecorder","Tethering","Timelapse","TimeTool","Torch","UserCenterService","Weather","WifiTransfer","Wifi Transfer","Wi-fi transfer"]

		#get the record from the pr list
		for rownum in range(sh.nrows):
			if sh.row_values(rownum)[4] == name:
				if self.isopenstate(sh.row_values(rownum)[5].upper()) == True:
					ogCount = ogCount + 1

				if self.isopenstate(sh.row_values(rownum)[5].upper()) == True and self.isblocked(sh.row_values(rownum)[9]) == True:
					ogHomo = ogHomo + 1
			
				if sh.row_values(rownum)[5] == 'RESOLVED':
					toVsw = toVsw + 1

		for rownum in range(sh_days.nrows):
			if sh_days.row_values(rownum)[4] == name:
				isGenericApp = False
				for appName in app_list:
					if appName.upper() in sh_days.row_values(rownum)[6].upper():
						isGenericApp = True
						break

				if self.isclosestate(sh_days.row_values(rownum)[5].upper()) == True:
					fixed = fixed + 1
			
				if sh_days.row_values(rownum)[8] == 'YES':
					if isGenericApp != True:
						regTotal = regTotal + 1

		engBugInfor = [ogCount,ogHomo,toVsw,fixed,regTotal]
		return engBugInfor

	def writeTitle(self, worksheet, index, begin):
		worksheet.write(index, begin, 'OG_ALL', self.getTitleStyle())
		worksheet.write(index, begin + 1, 'OG_Block', self.getTitleStyle())
		worksheet.write(index, begin + 2, 'To V_SW', self.getTitleStyle())
		worksheet.write(index, begin + 3, 'Fixed_All', self.getTitleStyle())
		worksheet.write(index, begin + 4, 'REG Total', self.getTitleStyle())

	def writeSum(self, worksheet, beginIndex, i):
		list1 = ['0','0','0','0','0']
		list2 = ['D','E','F','G','H']
		j = 0
		for item in list1:
			if beginIndex > i:
				worksheet.write(i, j+3, Formula(item), self.getTitleStyle())
			else:
				worksheet.write(i, j+3, Formula('SUM(' + list2[j] + str(beginIndex) + ':' + list2[j] + str(i) + ')'), self.getTitleStyle())
			j = j + 1

	def writeSumAll(self, worksheet, beginIndex, i):
		list1 = ['0','0','0','0','0']
		list2 = ['C','D','E','F','G']
		j = 0
		for item in list1:
			if beginIndex > i:
				worksheet.write(i, j+2, Formula(item), self.getTitleStyle())
			else:
				worksheet.write(i, j+2, Formula('SUM(' + list2[j] + str(beginIndex) + ':' + list2[j] + str(i) + ')'), self.getTitleStyle())
			j = j + 1
         
	def writeTeamSum(self, worksheet, i, lineNum):
		list2 = ['D','E','F','G','H']
		j = 0
		for item in list2:
			worksheet.write(i, j+3, Formula(item + str(lineNum)), self.getCellStyle())
			j = j + 1

	def writeSumAllD(self, worksheet, beginIndex, i):
		list1 = ['0','0','0','0','0']
		list2 = ['D','E','F','G','H']
		j = 0
		for item in list1:
			if beginIndex > i:
				worksheet.write(i, j+3, Formula(item), self.getTitleStyle())
			else:
				worksheet.write(i, j+3, Formula('SUM(' + list2[j] + str(beginIndex) + ':' + list2[j] + str(i) + ')'), self.getTitleStyle())
			j = j + 1

	def createDB(self, sqliteConn):
		cur = sqliteConn.cursor()
		cur.execute('select name from sqlite_master where type="table"')
		tableList = cur.fetchall()
		found = False
		for table in tableList:
			if table[0] == 'engreport':
				found = True
				cur.execute('delete from engreport')
		if not found:
			cur.execute('create table engreport(name VARCHAR(30),email VARCHAR(30),team VARCHAR(30),platform VARCHAR(30),ogAll INTEGER,ogHomo INTEGER,ToSW INTEGER, fixed INTEGER,reg INTEGER)')
			sqliteConn.commit()
		cur.close()

	def hasData(self, fetchList):
		if(type(fetchList[0]) == types.NoneType or type(fetchList[1]) == types.NoneType or type(fetchList[2]) == types.NoneType or type(fetchList[3]) == types.NoneType or type(fetchList[4]) == types.NoneType):
			return False

		if(int(fetchList[0]) == 0 and int(fetchList[1]) == 0 and int(fetchList[2]) == 0 and int(fetchList[3]) == 0 and int(fetchList[4]) == 0 and int(fetchList[5]) == 0 and int(fetchList[6]) == 0 and int(fetchList[7]) == 0 and int(fetchList[8]) == 0 and int(fetchList[9]) == 0 and int(fetchList[10]) == 0):
			return False
		return True

	def insert(self, sqliteConn, dbName, dbItem):
		cur = sqliteConn.cursor()
		keystr = "'"
		for tmp in dbItem:
			if type(tmp) == int or type(tmp) == float:
				keystr += str(tmp) + "','"
			else:
				keystr += tmp + "','"
		cur.execute('insert into ' + dbName + ' values (' + keystr[:-2] + ')')
		sqliteConn.commit()
		cur.close()

	def getTitleHtml(self, title_name):
		return '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" height="15" valign="middle"><b>'+ title_name + '</b></td>'

	def getHtml(self, type_name, project_name, worksheet, teams, teamNameList, sheetCount):
		html = '<font color="#000000" face="Arial" size="3"><strong>Team ' + type_name + ' status on:</strong>'
		html += '<table border="1" cellspacing="0" cols="12" rules="none">'
		html += '<tr>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" width = "100" height="15" valign="middle" colspan="2"><b>'+ project_name.upper() +'</b></td>'
		html += self.getTitleHtml('OG_All')
		html += self.getTitleHtml('OG_Block')
		html += self.getTitleHtml('To V_SW')
		html += self.getTitleHtml('Fixed_All')
		html += self.getTitleHtml('REG Total')
		index = 0
		appSum = [0 for col in range(5)]
		for teamName in teamNameList:
			htmlRow = ''
			isZero = True
			if teamName not in teams:
				index = index + 1
				continue
			else:
				if '_' in teamName and 'NJ' not in teamName:
					teamName = teamName[3:]
				htmlRow += '<tr>'
				htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" width = "100" height="15" valign="middle" colspan="2"><b>'+ teamName +'</b></td>'

			sumCount = 0
			for j in range(prj_team_sum_begin[sheetCount][index] - 1, prj_team_sum_end[sheetCount][index]):
				sumCount = sumCount + int(worksheet.row_values(j)[3])
			htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#D2E9FF" height="15" valign="middle"><b>'+ str(sumCount) +'</b></td>'
			appSum[0] = appSum[0] + sumCount
			if(0 != sumCount):
				isZero = False

			sumCount = 0
			for j in range(prj_team_sum_begin[sheetCount][index] - 1, prj_team_sum_end[sheetCount][index]):
				sumCount = sumCount + int(worksheet.row_values(j)[4])
			htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ str(sumCount) +'</b></td>'
			appSum[1] = appSum[1] + sumCount
			if(0 != sumCount):
				isZero = False

			sumCount = 0
			for j in range(prj_team_sum_begin[sheetCount][index] - 1, prj_team_sum_end[sheetCount][index]):
				sumCount = sumCount + int(worksheet.row_values(j)[5])
			htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor= "#D2E9FF" height="15" valign="middle"><b>'+ str(sumCount) +'</b></td>'
			appSum[2] = appSum[2] + sumCount
			if(0 != sumCount):
				isZero = False

			sumCount = 0
			for j in range(prj_team_sum_begin[sheetCount][index] - 1, prj_team_sum_end[sheetCount][index]):
				sumCount = sumCount + int(worksheet.row_values(j)[6])
			htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ str(sumCount) +'</b></td>'
			appSum[3] = appSum[3] + sumCount
			if(0 != sumCount):
				isZero = False

			sumCount = 0
			for j in range(prj_team_sum_begin[sheetCount][index] - 1, prj_team_sum_end[sheetCount][index]):
				sumCount = sumCount + int(worksheet.row_values(j)[7])
			htmlRow += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#D2E9FF" height="15" valign="middle"><b>'+ str(sumCount) +'</b></td>'
			appSum[4] = appSum[4] + sumCount
			if(0 != sumCount):
				isZero = False

			index = index + 1
			htmlRow += '</tr>'

			if isZero == False:
				html += htmlRow

		html += '<tr>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" width = "200" height="15" valign="middle" colspan="2"><b>'+"APP_SUM"+'</b></td>'
		for num in appSum:
			html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" width = "80" height="15" valign="middle"><b>'+ str(num) +'</b></td>'
		html += '</tr>'
		html += '</table>'
		html += '<br><br>'
		return html

	def getTopOwnerHtml(self):
		html = '<font color="#000000" face="Arial" size="3"><strong>Top Bug Owner List (total >= 5):</strong>'
		html += '<table border="1" cellspacing="0" cols="12" rules="none">'
		html += '<tr>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" height="15" width = "200" valign="middle"><b>Assigneer</b></td>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" height="15" width = "80" valign="middle"><b>Total</b></td>'

		sqliteConn = sqlite3.connect('/local/reportpr/pr_report.db')
		cursor = sqliteConn.cursor()
		cursor2 = sqliteConn.cursor()
		cursor.execute("SELECT email FROM engreport")
		emails = cursor.fetchall()
		email_total_list = []
		for email in emails:
			cursor2.execute("SELECT ogAll FROM engreport WHERE email = '" + email[0] + "'")
			ogNums = cursor2.fetchall()
			for ogNum in ogNums:
				if(ogNum[0] >= 10):
					email_total_list.append((email[0],ogNum[0]))
		cursor2.close()
		cursor.close()

		email_totals = sorted(email_total_list, key=itemgetter(1), reverse = True)
		for email_total in email_totals:
			html += '<tr>'
			html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ email_total[0] +'</b></td>'
			html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ str(email_total[1]) +'</b></td>'
			html += '</tr>'

		html += '</tr>'
		html += '</table>'
		html += '<br><br>'
		return html

	def getTopFunctionHtml(self):
		html = '<font color="#000000" face="Arial" size="3"><strong>Top Fuction List (total >= 5):</strong>'
		html += '<table border="1" cellspacing="0" cols="12" rules="none">'
		html += '<tr>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" height="15" width = "200" valign="middle"><b>Function</b></td>'
		html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" bgcolor="#97CBFF" height="15" width = "80" valign="middle"><b>Total</b></td>'

		sqliteConn = sqlite3.connect('/local/reportpr/pr_report.db')
		cursor = sqliteConn.cursor()
		cursor2 = sqliteConn.cursor()
		cursor.execute("SELECT id, function FROM functionreport")
		infos = cursor.fetchall()
		functions = []
		func_total_list = []
		for info in infos:
			if info[1] not in functions:
				functions.append(info[1])
				cursor2.execute("SELECT id FROM functionreport WHERE function = '" + info[1] + "'")
				ids = cursor2.fetchall()
				if(len(ids) >= 10):
					func_total_list.append((info[1],len(ids)))
		cursor2.close()
		cursor.close()

		func_totals = sorted(func_total_list, key=itemgetter(1), reverse = True)
		for func_total in func_totals:
			html += '<tr>'
			html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ func_total[0] +'</b></td>'
			html += '<td style="font-size:13px; border: 1px solid rgb(0, 0, 0);" align="center" height="15" valign="middle"><b>'+ str(func_total[1]) +'</b></td>'
			html += '</tr>'

		html += '</tr>'
		html += '</table>'
		html += '<br><br>'
		return html

	def addNATeamMember(self, targetFile):
		wb = xlrd.open_workbook('/local/reportpr/pr.xls')
		sheetNames = wb.sheet_names()
		NAMembers = []
		for sheetName in sheetNames:
			if sheetName != 'PR Status' and '(' not in sheetName:
				sh = wb.sheet_by_name(sheetName)
				for rownum in range(sh.nrows):
					if sh.row_values(rownum)[10] == '#NA' and sh.row_values(rownum)[4] not in NAMembers:
						NAMembers.append(sh.row_values(rownum)[4])

		NAFile = open(targetFile, 'w')
		for member in NAMembers:
			NAFile.write(member[:member.find('@')])
			NAFile.write(":")
			NAFile.write(member)
			NAFile.write("\n")

		NAFile.close()

	def deleteNATeamMember(self, targetFile):
		if os.path.isfile(targetFile): 
			os.remove(targetFile)

	def createFunctionDB(self, sqliteConn):
		cur = sqliteConn.cursor()
		cur.execute('select name from sqlite_master where type="table"')
		tableList = cur.fetchall()
		found = False
		for table in tableList:
			if table[0] == 'functionreport':
				found = True
				cur.execute('delete from functionreport')
		if not found:
			cur.execute('create table functionreport(id VARCHAR(30), function VARCHAR(30))')
			sqliteConn.commit()
		cur.close()

		wb = xlrd.open_workbook('/local/reportpr/pr.xls')
		sheetNames = wb.sheet_names()
		for sheetName in sheetNames:
			if sheetName != 'PR Status' and -1 == sheetName.find('('):
				sh = wb.sheet_by_name(sheetName)
				for rownum in range(sh.nrows):
					if self.isopenstate(sh.row_values(rownum)[5].upper()) == True:
						dbItem = []
						dbItem.append(sh.row_values(rownum)[0])
						dbItem.append(sh.row_values(rownum)[11])
						self.insert(sqliteConn, 'functionreport', dbItem)

	def start(self, project_name, days_value):
		sqliteConn = sqlite3.connect('/local/reportpr/pr_report.db')
		self.createDB(sqliteConn)
		self.createFunctionDB(sqliteConn)

		wb = xlrd.open_workbook('/local/reportpr/pr.xls')
		sheetNames = wb.sheet_names()

		now = datetime.datetime.today().date()
		report_days = now.strftime('%Y-%m-%d')

		if days_value != 0:
			date_days = datetime.timedelta(days=days_value)
			last = now - date_days
			report_days = '(' + last.strftime('%Y-%m-%d') + '~' + now.strftime('%Y-%m-%d') + ')'

		sheetCount = 0
		workbook = pyExcelerator.Workbook()
		for sheetName in sheetNames:
			if sheetName != 'PR Status' and -1 == sheetName.find('('):
				sh = wb.sheet_by_name(sheetName)
				worksheet = workbook.add_sheet(sheetName + '_SUM')
				worksheet.frmla_opts = RecalcAlways
				worksheet.row(0).height = 500

				#line number
				i = 3
				sumDict = {}
				teamNameList = []
				#get all Team member
				files = []
				files = os.listdir('%s/alm_report_teammember/conf/TeamMember' % getToolPath())
				files.sort()
				count = 0
				for fileName in files:
					#write the title
					worksheet.write_merge(i, i, 1, 2, fileName.decode('utf8'),self.getTitleStyle())
					self.writeTitle(worksheet, i , 3)
					i = i + 1
					beginIndex = i + 1
					for line in file('%s/alm_report_teammember/conf/TeamMember/%s' % (getToolPath(),fileName)):
						if line.strip() == '':
							continue
						dbItem = []
						line = line.split(':')

						countInfor = []
						if days_value != 0:
							sh_days = wb.sheet_by_name(sheetName + '_' + report_days)
							countInfor = self.readSheet(sh, sh_days, (line[1].strip()).lower())
						else:
							countInfor = self.readSheet(sh, sh, (line[1].strip()).lower())
						
						if(countInfor[0] == 0 and countInfor[1] == 0 and countInfor[2] == 0 and countInfor[3] == 0 and countInfor[4] == 0):
							continue

						dbItem.append(line[0].decode('utf8'))
						dbItem.append((line[1].strip()).lower())
						dbItem.append(fileName.decode('utf8'))
						dbItem.append(sheetName)
						worksheet.write(i, 1, line[0].decode('utf8'),self.getCellStyle())
						worksheet.write(i, 2, (line[1].strip()).lower(),self.getCellStyle())
						j = 3

						for tmpItem in countInfor:
							worksheet.write(i, j, tmpItem,self.getCellStyle())
							dbItem.append(tmpItem)
							j = j + 1

						self.insert(sqliteConn, 'engreport', dbItem)
                    
						i = i + 1

					#write the team sum
					teamNameList.append(fileName)
					sumDict[fileName] = i + 1
					worksheet.write_merge(i, i, 1, 2, fileName.decode('utf8') + '_SUM',self.getTitleStyle())
					self.writeSum(worksheet, beginIndex, i)
					prj_team_sum_begin[sheetCount][count] = beginIndex
					prj_team_sum_end[sheetCount][count] = i
					count = count + 1
					i = i + 2

				sheetCount = sheetCount + 1

				#write the project sum
				worksheet.write(i, 2, sheetName,self.getTitleStyle())
				self.writeTitle(worksheet, i, 3)

				i = i + 1
				beginIndex = i + 1
				for teamName in teamNameList:
					lineNum = sumDict[teamName]
					worksheet.write(i, 2, teamName.decode('utf8'),self.getCellStyle())
					self.writeTeamSum(worksheet, i, lineNum)
					i = i + 1
				worksheet.write(i, 2, 'APP_SUM',self.getTitleStyle())
				self.writeSum(worksheet, beginIndex, i)

		#write the Team&Project
		worksheet = workbook.add_sheet('TEAM&PROJECT')
		worksheet.frmla_opts = RecalcAlways

		i = 3
		teamNameList.sort()
		for teamName in teamNameList:
			worksheet.write(i, 1, teamName.decode('utf8'), self.getTitleStyle())
			self.writeTitle(worksheet, i, 2)
			i = i + 1
			beginIndex = i + 1
			for sheetName in sheetNames:
				if sheetName != 'PR Status':
					#get the data from DB
					cur = sqliteConn.cursor()
					haveData = True
					cur.execute("SELECT SUM(ogAll),SUM(ogHomo),SUM(ToSW),SUM(fixed),SUM(reg) FROM engreport WHERE platform = '" + sheetName + "'  AND  team = '" + teamName + "'")
					while (True):
						fetchList = cur.fetchone()
						if type(fetchList) == types.NoneType:
							break
						haveData = self.hasData(fetchList)
						if(haveData == False):
							break
						worksheet.write(i, 1, sheetName, self.getCellStyle())
						worksheet.write(i, 2, fetchList[0], self.getCellStyle())
						worksheet.write(i, 3, fetchList[1], self.getCellStyle())
						worksheet.write(i, 4, fetchList[2], self.getCellStyle())
						worksheet.write(i, 5, fetchList[3], self.getCellStyle())
						worksheet.write(i, 6, fetchList[4], self.getCellStyle())
					cur.close()
					if(haveData == True):
						i = i + 1
			worksheet.write(i, 1, teamName.decode('utf8') + '_SUM',self.getTitleStyle())
			self.writeSumAll(worksheet, beginIndex, i)

			i = i + 2

		#write the TEAM&Member
		worksheet = workbook.add_sheet('TEAM&Member')
		worksheet.frmla_opts = RecalcAlways

		i = 3
		files = []
		files = os.listdir('%s/alm_report_teammember/conf/TeamMember' % getToolPath())
		files.sort()
		for fileName in files:
			worksheet.write_merge(i, i, 1, 2, fileName.decode('utf8'), self.getTitleStyle())
			self.writeTitle(worksheet, i, 3)
			i = i + 1
			beginIndex = i + 1

			for line in file('%s/alm_report_teammember/conf/TeamMember/%s' % (getToolPath(),fileName)):
				if line.strip() == '':
					continue
				dbItem = []
				line = line.split(':')
				haveData = True
				#get the data from DB
				cur = sqliteConn.cursor()

				cur.execute("SELECT SUM(ogAll),SUM(ogHomo),SUM(ToSW),SUM(fixed),SUM(reg) FROM engreport WHERE email = '" + (line[1].strip()).lower() + "'")
				while (True):
					fetchList = cur.fetchone()
					if type(fetchList) == types.NoneType:
						break
					haveData = self.hasData(fetchList)
					if(haveData == False):
						break
					worksheet.write_merge(i, 1, 1, 2, line[0].decode('utf8'), self.getCellStyle())
					worksheet.write(i, 3, fetchList[0], self.getCellStyle())
					worksheet.write(i, 4, fetchList[1], self.getCellStyle())
					worksheet.write(i, 5, fetchList[2], self.getCellStyle())
					worksheet.write(i, 6, fetchList[3], self.getCellStyle())
					worksheet.write(i, 7, fetchList[4], self.getCellStyle())
				cur.close()
				if(haveData == True):
					i = i + 1

			worksheet.write_merge(i, i, 1, 2, fileName.decode('utf8') + '_SUM', self.getTitleStyle())
			self.writeSumAllD(worksheet, beginIndex, i)
			i = i + 2
		#write the all data

		worksheet.write(i, 2, 'ALL', self.getTitleStyle())
		self.writeTitle(worksheet, i, 3)
		i = i + 1
		beginIndex = i + 1
		for teamName in teamNameList:
			worksheet.write(i, 2, teamName.decode('utf8'), self.getCellStyle())
			#get the data from DB
			cur = sqliteConn.cursor()

			cur.execute("SELECT SUM(ogAll),SUM(ogHomo),SUM(ToSW),SUM(fixed),SUM(reg) FROM engreport WHERE team = '" + teamName + "'")
			while (True):
				fetchList = cur.fetchone()
				if type(fetchList) == types.NoneType:
					break
				haveData = self.hasData(fetchList)
				if(haveData == False):
					break
				worksheet.write(i, 3, fetchList[0], self.getCellStyle())
				worksheet.write(i, 4, fetchList[1], self.getCellStyle())
				worksheet.write(i, 5, fetchList[2], self.getCellStyle())
				worksheet.write(i, 6, fetchList[3], self.getCellStyle())
				worksheet.write(i, 7, fetchList[4], self.getCellStyle())
			cur.close()
			if(haveData == True):
				i = i + 1

		worksheet.write(i, 2, 'APP_SUM', self.getTitleStyle())
		self.writeSumAllD(worksheet, beginIndex, i)

		i = i + 2

		workbook = workbook.save('/local/reportpr/bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls')
		print 'Create the REPORT IN : /local/reportpr/bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls'

		html = '<html xmlns="http://www.w3.org/1999/xhtml">'
		html += '<head>'
		html += '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />'
		html += '</head>'
		html += '<body>'
		html += '<font color="#000000" face="Arial" size="3">Dears,<br><br>Please review the team Task/Defect report as follows: <br><br>'

		SHTeams = []
		NBTeams = []
		OtherTeams = []
		for teamName in teamNameList:
			if('SH_' in teamName or 'NJ_' in teamName):
				SHTeams.append(teamName)
			elif ('NB_' in teamName):
				NBTeams.append(teamName)
			else:
				OtherTeams.append(teamName)
		html += '<table width="100%" border="0" cellspacing="0" cellpadding="0">'
		html += '<tr><td >'

		wb = xlrd.open_workbook('/local/reportpr/bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls')
		worksheet = wb.sheet_by_index(0)

		html += self.getHtml('SH', project_name, worksheet, SHTeams, teamNameList, 0)
		html += self.getHtml('NB', project_name, worksheet, NBTeams, teamNameList, 0)
		html += self.getHtml('Others', project_name, worksheet, OtherTeams, teamNameList, 0)
		html += '</td><td valign="top">'
		html += self.getTopOwnerHtml()
		html += self.getTopFunctionHtml()
		html += '</td></tr>'
		html += '</table>'

		html += '<font color="#000000" face="Arial" size="3"><br><strong>Note:</strong>'
		html += '<font color="#000000" face="Arial" size="3">&nbsp;&nbsp;"To V_SW" column means those Task/Defect in resolved state. Resolved Task/Defect can not be taken as it has been fixed by SW, owner should verify it and change its state to VEVRIFIED_SW.<br><br>'
		html += '<font color="#000000" face="Arial" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>"Fixed_All"/"REG Total"</strong> column that caculates data for the days <strong>' + report_days + '</strong>.<br><br>'
		html += '<font color="#000000" face="Arial" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;From attached file pr.xls, you can see all of OG & Fixed Task/Defect,  you can also filter team info by "Team" column.<br><br>'
		html += '<font color="#000000" face="Arial" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;From attached file '
		html += '<font color="#000000" face="Arial" size="3">bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls,'+ '&nbsp;you can see data statistics of one team or one person.<br><br>'
		html += '<font color="#000000" face="Arial" size="3"><strong>Any mistake please let me know.&nbsp;Thanks.</strong><br><br>'
		html += '<font color="#000000" face="Arial" size="3">BR,<br>'
		html += '<font color="#000000" face="Arial" size="3">' + self.getFullName() + '<br><br>'
		html += '</html>'

		msg = MIMEMultipart('mixed')
		msg['Date'] = strftime("%a, %d %b %Y %T", localtime()) + ' +0800'
		msg['Subject'] = project_name.upper() + " Task/Defect Report by Team " + now.strftime('%Y-%m-%d')

		if days_value == 1:
			msg['Subject'] = project_name.upper() + " Task/Defect Daily Report by Team (" + now.strftime('%Y-%m-%d') + ")"
		elif days_value == 7:
			msg['Subject'] = project_name.upper() + " Task/Defect Weekly Report by Team " + report_days
		else:
			msg['Subject'] = project_name.upper() + " Task/Defect Report by Team " + now.strftime('%Y-%m-%d')

		msg['From'] = '"\'%s\'" <%s>' % (self.getFullName(), self.getMail())
		for to in self.toList:
			msg['To'] = to
		for cc in self.ccList:
			msg['Cc'] = cc

		contMsg = MIMEMultipart('related')

		htmlPart = MIMEBase('text', 'html', charset="utf-8")
		htmlPart.set_payload(html)
		encoders.encode_base64(htmlPart)
		contMsg.attach(htmlPart)

		msg.attach(contMsg)

		attach = MIMEBase('application', 'octet-stream')
		fp = open('/local/reportpr/pr.xls', 'rb')
		attach.set_payload(fp.read())
		fp.close()
		encoders.encode_base64(attach)
		attach["Content-Disposition"] = 'attachment; filename="''pr'+'.xls"'
		msg.attach(attach)
	
		attach = MIMEBase('application', 'octet-stream')
		fp = open('/local/reportpr/bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls', 'rb')
		attach.set_payload(fp.read())
		fp.close()
		encoders.encode_base64(attach)
		attach["Content-Disposition"] = 'attachment; filename="''bugs_monitor_' + now.strftime('%Y-%m-%d') + '.xls"'
		msg.attach(attach)

		s = smtplib.SMTP('172.24.61.92')
		s.set_debuglevel(0)
		s.sendmail(self.getMail(), self.toList + self.ccList, msg.as_string())
		s.quit()
