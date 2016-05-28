#!/usr/bin/python

# coding=gbk
import codecs

import re
import sqlite3
import types
from Utils import *
import MySQLdb
import datetime, time
import email
import pyExcelerator

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from UserInfo import *
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from time import strftime, localtime
from Integrity import IntegrityClient

class PrStatic(UserInfo):
    def __init__(self, name):
		self.now = datetime.datetime.today().date()
		self.name = name
		self.toList = []
		self.ogList = []

    def addTo(self, mailAddr):
        self.toList.append(mailAddr)

    def initUserInfo(self):
		for line in file('%s/alm_report_teammember/conf/platform.conf' % getToolPath()):
			line = line.split('=')
			self.ogList.append({line[0].strip() : line[1][1:].strip()})


    def createPRSheet(self, worksheetOg, items):
		worksheetOg.write(0, 0, 'ID')
		worksheetOg.write(0, 1, 'Type')
		worksheetOg.write(0, 2, 'Branch')
		worksheetOg.write(0, 3, 'Priority')
		worksheetOg.write(0, 4, 'Assigner')
		worksheetOg.write(0, 5, 'State')
		worksheetOg.write(0, 6, 'Summary')
		worksheetOg.write(0, 7, 'deadline')
		worksheetOg.write(0, 8, 'regression')
		worksheetOg.write(0, 9, 'comment_from_cea')
		worksheetOg.write(0, 10, 'Team')
		worksheetOg.write(0, 11, 'Function')
		lineNum = 1

		for item in items:
			try:
				branch = item.branch
				if self.projectName == branch:
					worksheetOg.write(lineNum, 2, item.branch)
				else:
					continue
			except AttributeError:
				continue
			except AssertionError:
				branches = item.branch
				for branch in branches:
					if self.projectName == branch:
						worksheetOg.write(lineNum, 2, item.branch)

			worksheetOg.write(lineNum, 0, item.id)
			worksheetOg.write(lineNum, 1, item.type)
			worksheetOg.write(lineNum, 3, item.priority)
			worksheetOg.write(lineNum, 4, item.assigned_user.email)
			worksheetOg.write(lineNum, 5, item.state)
			worksheetOg.write(lineNum, 6, item.summary)

			if 'Deadline' in item.keys():
				worksheetOg.write(lineNum, 7, item['Deadline'][1].value)

			if 'Regression' in item.keys():
				worksheetOg.write(lineNum, 8, item['Regression'][1].value)

			if 'All CEA comments' in item.keys():
				worksheetOg.write(lineNum, 9, item['All CEA comments'][1].value)

#			worksheetOg.write(lineNum, 9, item.comment_from_cea)
			worksheetOg.write(lineNum, 11, item.function)

			files = []
			files = os.listdir('%s/alm_report_teammember/conf/TeamMember' % getToolPath())
			haveTeam = False
			for fileName in files:
				for line in file('%s/alm_report_teammember/conf/TeamMember/%s' % (getToolPath(), fileName)):
					if line.strip() == '':
						break
					line = line.split(':')
					if cmp((line[1].strip()).lower(), (item.assigned_user.email).lower()) == 0:
						worksheetOg.write(lineNum, 10, (str(fileName)).upper().decode("utf-8"))
						haveTeam = True
						break
			if haveTeam == False:
				worksheetOg.write(lineNum, 10, '#NA')

			lineNum += 1

    def sendPRStatic(self, full_name, project, days_value):
		wsdl = "http://alm.tclcom.com:7001/webservices/10/2/Integrity/?wsdl"
		intclient = IntegrityClient(wsdl=wsdl,
                                    credential_username="scm_tools",
                                    credential_password="SCM_TOOLS123!")
		if self.initUserInfo:
			self.initUserInfo()

		workbook = pyExcelerator.Workbook()

		tmpPlatformName = 'null'
		tmpDict = {}
		continueFlag = True;
		for tmpDict in self.ogList:
			for tmpA, tmpB in tmpDict.items():
				if type(tmpA) == types.NoneType:
					break
				elif cmp(tmpA[:tmpA.index('_')].lower(), full_name.lower()) == 0 :
					self.platformName = tmpA
					self.projectName = tmpB
					continueFlag = False;

			if continueFlag == True:
				continue;

			if self.platformName != tmpPlatformName:
				tmpPlatformName = self.platformName
				worksheetOg = workbook.add_sheet(self.platformName)
				fields=["ID","Type","Priority","Assigned User","State","Summary","Deadline","Regression","All CEA comments","Function"]

				query = ('((field[Type]="Defect","Task") and ' \
                 		'(field[Project]="%s") and ' \
                        '(field[Assigned User]="ying.tan","tyin","wenhua.tu","sichao.hu","yuanxing.tan"))') %(project)
				items = intclient.getItemsByCustomQuery(fields=fields, query=query)
				self.createPRSheet(worksheetOg, items)

				if days_value != 0:
					now = datetime.datetime.today().date()
					report_days = now.strftime('%Y-%m-%d')
					date_days = datetime.timedelta(days=days_value)
					last = now - date_days
					report_days = '(' + last.strftime('%Y-%m-%d') + '~' + now.strftime('%Y-%m-%d') + ')'
					worksheetOg_2 = workbook.add_sheet(self.platformName + '_' + report_days)

					datetime_format = '%b %d, %Y %I:%M:%S %p'
					date_days = datetime.timedelta(days=days_value)
					date = datetime.datetime.today().date() - date_days
					start_str = date.strftime(datetime_format)
					end_str = self.now.strftime(datetime_format)
					print "start time: ",start_str
					print "end time: ",end_str
					query_2 = ('((field[Type]="Defect","Task") and ' \
                 		'(field[Project]="%s") and ' \
                 		'(histdate[State] was changed between time %s and %s))')%(project, start_str, end_str)

					items_2 = intclient.getItemsByCustomQuery(fields=fields, query=query_2)
					self.createPRSheet(worksheetOg_2, items_2)

		workbook = workbook.save('/local/reportpr/pr.xls')

